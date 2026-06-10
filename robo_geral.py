"""
robo_geral.py — Certidoes Federais, FGTS, Trabalhista e Estadual MT

Uso:
  python robo_geral.py federal
  python robo_geral.py federal,fgts,trabalhista,estadual_mt
  python robo_geral.py todas
"""

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeout
from datetime import datetime
import openpyxl
import unicodedata
import subprocess
import tempfile
import base64
import shutil
import random
import time
import sys
import re
import os

try:
    import pypdf
    _PYPDF_OK = True
except ImportError:
    _PYPDF_OK = False

try:
    from playwright_stealth import stealth_sync as _stealth_sync
    _HAS_STEALTH = True
except ImportError:
    _HAS_STEALTH = False

# ── Configuracao ──────────────────────────────────────────────────────────────

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ARQUIVO_EXCEL  = os.path.join(BASE_DIR, "cnpjs.xlsx.xlsx")
COLUNA_CNPJ    = "CNPJ (MF) N.º"
COLUNA_NOME    = "Razao Social"
COLUNA_CIDADE  = "Cidade / UF"
ABA_PRODUTORES = "produtores ativos"

WEB_MODE        = os.environ.get("ROBO_WEB") == "1"
_JOB_ID         = os.environ.get("ROBO_JOB_ID", "")
_BASE_CERT      = os.path.join(os.path.expanduser("~"), "Documents", "certidões")
PASTA_BASE      = os.path.join(_BASE_CERT, _JOB_ID) if _JOB_ID else _BASE_CERT
PASTA_NEGATIVAS = os.path.join(PASTA_BASE, "negativas")
PASTA_POSITIVAS = os.path.join(PASTA_BASE, "positivas")
os.makedirs(PASTA_NEGATIVAS, exist_ok=True)
os.makedirs(PASTA_POSITIVAS, exist_ok=True)

HEADLESS = os.environ.get("ROBO_HEADLESS") == "1" or WEB_MODE

TIPOS_VALIDOS = ["federal", "fgts", "trabalhista", "estadual_mt"]

URL_FEDERAL     = "https://servicos.receitafederal.gov.br/servico/certidoes/#/home"
URL_FGTS        = "https://consulta-crf.caixa.gov.br/consultacrf/pages/consultaEmpregador.jsf"
URL_TRABALHISTA = "https://cndt-certidao.tst.jus.br/inicio.faces"
URL_SEFAZ_MT    = "https://www.sefaz.mt.gov.br/cnd/certidao/servlet/ServletRotdAberto?origem=60"

ANTI_CAPTCHA_KEY = os.environ.get("ANTI_CAPTCHA_KEY", "")
RF_TS_COOKIE     = os.environ.get("RF_TS_COOKIE", "")  # ex: TS3750c27d027=abc123

NOMES_TIPO = {
    "federal":     "Federal (CND)",
    
    "fgts":        "FGTS (CRF)",
    "trabalhista": "Trabalhista (CNDT)",
    "estadual_mt": "Estadual MT (SEFAZ)",
}

# ── Utilitarios ───────────────────────────────────────────────────────────────

def _norm(s):
    s = unicodedata.normalize("NFD", str(s))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()

def formatar_doc(doc):
    d = re.sub(r"\D", "", str(doc))
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    return d

def detectar_tipo_certidao(caminho):
    if not _PYPDF_OK:
        return "negativa"
    try:
        reader = pypdf.PdfReader(caminho)
        texto = " ".join(p.extract_text() or "" for p in reader.pages).upper()
        if "CERTID" in texto and "NEGATIVA" in texto:
            return "negativa"
        if "CERTID" in texto and "POSITIVA" in texto:
            return "positiva"
    except Exception:
        pass
    return "negativa"

# ── Leitura do Excel ──────────────────────────────────────────────────────────

def _ler_sheet(sheet):
    """Converte uma aba do openpyxl em lista de dicts (primeira linha = cabecalho)."""
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return [], []
    headers = [str(c).strip() if c is not None else "" for c in rows[0]]
    data = []
    for row in rows[1:]:
        d = {headers[i]: (str(row[i]).strip() if row[i] is not None else "") for i in range(len(headers))}
        data.append(d)
    return headers, data

def _ler_empresas():
    """Le CNPJs da aba principal."""
    try:
        wb = openpyxl.load_workbook(ARQUIVO_EXCEL, read_only=True, data_only=True)
        sheet = wb.active
        headers, data = _ler_sheet(sheet)
        wb.close()

        col_cnpj = None
        col_nome = None
        for h in headers:
            if _norm(h) == _norm(COLUNA_CNPJ):
                col_cnpj = h
            if _norm(h) == _norm(COLUNA_NOME):
                col_nome = h

        if col_cnpj is None:
            print(f"  AVISO: coluna '{COLUNA_CNPJ}' nao encontrada na aba principal.")
            return []

        out = []
        for row in data:
            cnpj = re.sub(r"\D", "", row.get(col_cnpj, "")).strip()
            if len(cnpj) == 14 and cnpj.isdigit():
                out.append({
                    "doc":  cnpj,
                    "nome": row.get(col_nome, "").strip() if col_nome else "",
                })
        return out
    except Exception as e:
        print(f"  ERRO ao ler empresas: {e}")
        return []

def _ler_produtores():
    """Le CPFs da aba 'produtores ativos'."""
    try:
        wb = openpyxl.load_workbook(ARQUIVO_EXCEL, read_only=True, data_only=True)
        if ABA_PRODUTORES not in wb.sheetnames:
            wb.close()
            return []
        sheet = wb[ABA_PRODUTORES]
        headers, data = _ler_sheet(sheet)
        wb.close()

        col_doc = None
        for h in headers:
            n = _norm(h)
            if "cpf" in n or _norm(COLUNA_CNPJ) == n:
                col_doc = h
                break
        if col_doc is None:
            for h in headers:
                if any(len(re.sub(r"\D", "", row.get(h, ""))) == 11 for row in data[:5]):
                    col_doc = h
                    break
        if col_doc is None:
            print(f"  AVISO: coluna de CPF nao encontrada na aba '{ABA_PRODUTORES}'.")
            return []

        col_nome = None
        for h in headers:
            n = _norm(h)
            if "nome" in n or "razao" in n or _norm(COLUNA_NOME) == n:
                col_nome = h
                break

        out = []
        for row in data:
            cpf = re.sub(r"\D", "", row.get(col_doc, "")).strip()
            if len(cpf) == 11 and cpf.isdigit():
                out.append({
                    "doc":  cpf,
                    "nome": row.get(col_nome, "").strip() if col_nome else "",
                })
        return out
    except Exception as e:
        print(f"  AVISO: aba '{ABA_PRODUTORES}' nao disponivel ou com erro: {e}")
        return []

# ── Helpers de download ───────────────────────────────────────────────────────

def _capturar_popup_como_pdf(page, context, caminho):
    """Captura o primeiro popup aberto e salva como PDF. Retorna True se ok."""
    popups = [p for p in context.pages if p != page]
    if not popups:
        return False
    pop = popups[-1]
    try:
        pop.wait_for_load_state("networkidle", timeout=20000)
        pop.emulate_media(media="print")
        pop.pdf(path=caminho, format="A4", print_background=True)
        pop.close()
        return os.path.exists(caminho) and os.path.getsize(caminho) > 3000
    except Exception:
        try:
            pop.close()
        except Exception:
            pass
        return False



# ── Fluxo: CND Federal ────────────────────────────────────────────────────────

def _baixar_federal_http(doc_limpo, caminho, px_token=""):
    """Tenta baixar via requests usando o token PerimeterX capturado do browser."""
    import requests
    eh_cpf = len(doc_limpo) == 11
    tipo = "CPF" if eh_cpf else "PJ"
    tipo_enum = "CPF" if eh_cpf else "CNPJ"
    base = "https://servicos.receitafederal.gov.br/servico/certidoes"

    hdrs = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "pt-BR,pt;q=0.9",
        "Origin": "https://servicos.receitafederal.gov.br",
        "Referer": base + "/",
        "Content-Type": "application/json",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Dest": "empty",
    }
    if px_token:
        hdrs["x-captcha-token"] = px_token

    sess = requests.Session()
    sess.headers.update(hdrs)

    try:
        sess.get(base + "/", timeout=15)
    except Exception:
        pass

    try:
        r = sess.post(f"{base}/api/consulta/validar-contribuinte",
                      json={"ni": doc_limpo, "tipoContribuinte": tipo}, timeout=30)
        print(f"  [HTTP] validar: {r.status_code} {r.text[:300]}")
        if r.status_code == 200:
            dv = r.json()
            if dv.get("statusValidacao") in ("Invalido", "invalido"):
                return "com_debitos"
        elif r.status_code != 400:
            return None
    except Exception as e:
        print(f"  [HTTP] validar erro: {e}")
        return None

    try:
        r2 = sess.post(f"{base}/api/Emissao/verificar",
                       json={"ni": doc_limpo, "tipoContribuinte": tipo,
                             "tipoContribuinteEnum": tipo_enum}, timeout=30)
        print(f"  [HTTP] verificar: {r2.status_code} {r2.text[:400]}")
        if r2.status_code != 200:
            return None
        dados = r2.json()
    except Exception as e:
        print(f"  [HTTP] verificar erro: {e}")
        return None

    # Tenta download pelo ID retornado
    for chave in ("id", "idCertidao", "numeroCertidao", "hash"):
        cert_id = dados.get(chave)
        if cert_id:
            for ep in [f"{base}/api/Emissao/download/{cert_id}",
                       f"{base}/api/certidao/{cert_id}/download"]:
                try:
                    rd = sess.get(ep, timeout=30)
                    if rd.status_code == 200 and len(rd.content) > 3000:
                        with open(caminho, "wb") as f:
                            f.write(rd.content)
                        return "ok"
                except Exception:
                    pass

    # Tenta emitir nova certidao
    try:
        r3 = sess.post(f"{base}/api/Emissao/emitir",
                       json={"ni": doc_limpo, "tipoContribuinte": tipo,
                             "tipoContribuinteEnum": tipo_enum}, timeout=60)
        print(f"  [HTTP] emitir: {r3.status_code} len={len(r3.content)} {r3.text[:300]}")
        if r3.status_code == 200 and len(r3.content) > 3000:
            with open(caminho, "wb") as f:
                f.write(r3.content)
            return "ok"
    except Exception as e:
        print(f"  [HTTP] emitir erro: {e}")

    return None

def _set_angular_input(page, selector, value):
    """Define valor em input Angular disparando todos os eventos necessarios."""
    page.evaluate("""([sel, val]) => {
        const el = document.querySelector(sel);
        if (!el) return;
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(el, val);
        el.dispatchEvent(new Event('input',  { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.dispatchEvent(new Event('blur',   { bubbles: true }));
    }""", [selector, value])

def baixar_federal(page, context, doc, caminho):
    """CND Federal — Receita Federal (servicos.receitafederal.gov.br)."""
    doc_limpo = re.sub(r"\D", "", doc)

    # HTTP sem token sempre falha (022) E envenena o IP no PerimeterX.
    # Fazemos HTTP apenas depois de capturar o token do browser.

    eh_cpf = len(doc_limpo) == 11

    url_hash = "/home/cpf" if eh_cpf else "/home/cnpj"
    url = f"https://servicos.receitafederal.gov.br/servico/certidoes/#{url_hash}"

    # Se já estamos no domínio RF, navega só o hash — PerimeterX não re-inicializa
    ja_no_rf = "receitafederal.gov.br/servico/certidoes" in page.url
    if ja_no_rf:
        page.evaluate(f"location.hash = '{url_hash}'")
        page.wait_for_timeout(800)
    else:
        page.goto(url, wait_until="domcontentloaded", timeout=30000)
        # Primeira carga: aguarda PerimeterX inicializar
        page.wait_for_timeout(random.randint(4000, 5000))

    # Limpa estado Angular residual
    page.evaluate("() => { try { sessionStorage.clear(); } catch(e) {} }")
    page.wait_for_timeout(300)

    # Campo CPF ou CNPJ
    label = "CPF" if eh_cpf else "CNPJ"
    campo = page.get_by_label(label, exact=False)
    try:
        campo.wait_for(state="visible", timeout=15000)
    except PlaywrightTimeout:
        return "nao_encontrado"

    # Move mouse antes de clicar (comportamento humano — apenas na primeira vez)
    if not ja_no_rf:
        try:
            box_c = campo.bounding_box()
            if box_c:
                page.mouse.move(random.randint(100, 400), random.randint(100, 300))
                page.wait_for_timeout(random.randint(200, 400))
                page.mouse.move(
                    box_c["x"] + box_c["width"] * random.uniform(0.3, 0.7),
                    box_c["y"] + box_c["height"] * random.uniform(0.3, 0.7),
                    steps=random.randint(5, 8),
                )
                page.wait_for_timeout(random.randint(100, 250))
        except Exception:
            pass

    # Digita o documento com delay reduzido (ainda parece humano)
    campo.click()
    campo.fill("")  # Limpa campo antes de digitar
    for ch in doc_limpo:
        campo.press(ch)
        page.wait_for_timeout(random.randint(40, 90))
    page.wait_for_timeout(random.randint(200, 400))

    # Garante que o Angular registrou o valor formatado
    sel = "input[id*='cnpj' i], input[name*='cnpj' i], input[formcontrolname*='cnpj' i]"
    if eh_cpf:
        sel = "input[id*='cpf' i], input[name*='cpf' i], input[formcontrolname*='cpf' i]"
    doc_fmt = formatar_doc(doc)
    _set_angular_input(page, sel, doc_fmt)
    page.wait_for_timeout(random.randint(300, 500))

    # Captura o token PerimeterX gerado pelo browser
    _px_token = [""]
    _api_resps = []
    def _on_req_px(req):
        if "certidoes/api" in req.url:
            tok = req.headers.get("x-captcha-token", "")
            if tok and not _px_token[0]:
                _px_token[0] = tok
    def _on_resp_api(resp):
        if "certidoes/api" in resp.url:
            try:
                body_resp = resp.text()[:400]
            except Exception:
                body_resp = ""
            _api_resps.append(f"{resp.status} {resp.url.split('api/')[-1]} | {body_resp}")
    page.on("request", _on_req_px)
    page.on("response", _on_resp_api)

    # Aguarda botao ficar habilitado
    btn_emitir = page.get_by_role("button", name="Emitir Certidão")
    try:
        btn_emitir.wait_for(state="visible", timeout=5000)
    except PlaywrightTimeout:
        return "nao_encontrado"

    # Move mouse brevemente antes de clicar
    try:
        box = btn_emitir.bounding_box()
        if box:
            page.mouse.move(
                box["x"] + box["width"] * random.uniform(0.3, 0.7),
                box["y"] + box["height"] * random.uniform(0.3, 0.7),
                steps=random.randint(4, 7),
            )
            page.wait_for_timeout(random.randint(80, 200))
    except Exception:
        pass

    # Captura downloads via event listener
    _downloads = []
    _on_dl = lambda d: _downloads.append(d)
    context.on("download", _on_dl)

    _RE_ERRO_TX = re.compile(
        r"nao foi poss[íi]vel|insuficientes para emitir|dados insuf|nao.*localizado",
        re.I,
    )

    def _tem_erro_visivel():
        try:
            return page.get_by_text(_RE_ERRO_TX, exact=False).first.is_visible()
        except Exception:
            return False

    def _tem_erro_api():
        for r in _api_resps[-3:]:
            if r.startswith("400") or " 023" in r or " 022" in r:
                return True
        return False

    def _tem_popup():
        return any(pg != page for pg in context.pages)

    def _finalizar(status):
        try:
            context.remove_listener("download", _on_dl)
        except Exception:
            pass
        for r in _api_resps:
            print(f"  [API] {r}")
        return status

    def _salvar_dl():
        try:
            _downloads[0].save_as(caminho)
            return os.path.exists(caminho) and os.path.getsize(caminho) > 3000
        except Exception:
            return False

    def _aguardar_resultado(max_ticks, label=""):
        """Aguarda download/popup/erro por até max_ticks × 500ms. Retorna 'ok','erro','timeout'."""
        for tick in range(max_ticks):
            page.wait_for_timeout(500)
            if _downloads:
                print(f"  Download detectado (tick {tick+1})")
                return "ok" if _salvar_dl() else "timeout"
            if _tem_popup():
                print(f"  Popup/nova-aba detectado (tick {tick+1})")
                return "ok" if _capturar_popup_como_pdf(page, context, caminho) else "timeout"
            if _tem_erro_api() or _tem_erro_visivel():
                print(f"  Erro detectado{' ' + label if label else ''} (tick {tick+1})")
                return "erro"
            # Pagina de resultado sem download = informacoes insuficientes
            if "/resultado" in page.url and not _downloads:
                print(f"  Pagina /resultado sem download (tick {tick+1}) — insuficiente")
                return "erro"
        return "timeout"

    # ── Passo 1: clica "Emitir Certidão" ─────────────────────────────────────
    print("  Clicando Emitir Certidao...")
    btn_emitir.click()

    # Aguarda até 6s: download direto, erro ou modal
    _modal_nome = None
    for _tick in range(12):  # 12 × 500ms = 6s
        page.wait_for_timeout(500)
        if _downloads:
            print(f"  Download direto (tick {_tick+1})")
            return _finalizar("ok") if _salvar_dl() else _finalizar("nao_encontrado")
        if _tem_popup():
            print(f"  Popup detectado apos Emitir (tick {_tick+1})")
            r = _capturar_popup_como_pdf(page, context, caminho)
            return _finalizar("ok") if r else _finalizar("nao_encontrado")
        if _tem_erro_api() or _tem_erro_visivel():
            print(f"  Erro apos Emitir (tick {_tick+1})")
            return _finalizar("nao_encontrado")
        # Modal de certidão válida?
        for _nome in ["Emitir Nova Certidão", "Consultar Certidão"]:
            try:
                if page.get_by_role("button", name=_nome).is_visible():
                    _modal_nome = _nome
                    break
            except Exception:
                pass
        if _modal_nome:
            print(f"  Modal detectado (tick {_tick+1}): '{_modal_nome}'")
            break

    # ── Passo 2: clica botão do modal se apareceu ─────────────────────────────
    if _modal_nome:
        # Aguarda o backdrop/loading sumir antes de tentar clicar
        try:
            page.locator("br-loading").wait_for(state="hidden", timeout=10000)
            print("  Loading desapareceu, clicando...")
        except Exception:
            pass  # Se nao achou br-loading, tenta assim mesmo

        # Tenta o botao azul primeiro; cai no que foi detectado como fallback
        _clicado = False
        for _nome_click in ["Emitir Nova Certidão", _modal_nome]:
            try:
                btn = page.get_by_role("button", name=_nome_click)
                btn.wait_for(state="visible", timeout=3000)
                btn.click()
                print(f"  Clicou '{_nome_click}', aguardando resultado...")
                _clicado = True
                break
            except Exception:
                continue

        if not _clicado:
            print("  Nao conseguiu clicar botao do modal")
            return _finalizar("nao_encontrado")

        res = _aguardar_resultado(max_ticks=40, label="apos modal")
        return _finalizar("ok") if res == "ok" else _finalizar("nao_encontrado")

    # Sem modal e sem resultado em 6s
    if _tem_popup():
        r = _capturar_popup_como_pdf(page, context, caminho)
        return _finalizar("ok") if r else _finalizar("nao_encontrado")

    print("  Sem modal, sem download e sem erro em 6s — pulando")
    return _finalizar("nao_encontrado")

# ── Fluxo: FGTS ───────────────────────────────────────────────────────────────

def baixar_fgts(page, context, doc, caminho):
    """CRF (FGTS) — portal Caixa (consulta-crf.caixa.gov.br)."""
    cnpj_limpo = re.sub(r"\D", "", doc)

    page.goto(URL_FGTS, wait_until="load", timeout=30000)
    page.wait_for_timeout(2000)

    # CPF (produtor rural) exige selecionar tipo "3" antes de preencher
    if len(cnpj_limpo) == 11:
        page.locator('[id="mainForm:tipoEstabelecimento"]').select_option("3")
        page.wait_for_timeout(500)

    campo = page.locator('[id="mainForm:txtInscricao1"]')
    try:
        campo.wait_for(state="visible", timeout=15000)
    except PlaywrightTimeout:
        return "nao_encontrado"
    campo.click()
    campo.fill(cnpj_limpo)

    page.get_by_role("button", name="Consultar").click()

    # Aguarda o link de resultado aparecer
    try:
        link = page.get_by_role("link", name="Certificado de Regularidade")
        link.wait_for(state="visible", timeout=15000)
    except PlaywrightTimeout:
        try:
            if page.get_by_text(re.compile(
                r"irregular|pendenc|devedor|n.o.*regular", re.I
            ), exact=False).first.is_visible(timeout=2000):
                return "com_debitos"
        except Exception:
            pass
        return "nao_encontrado"

    # Captura popup que pode abrir ao clicar "Certificado de Regularidade"
    popups_cap = []
    downloads_cap = []
    def _on_page(p): popups_cap.append(p)
    def _on_dl(d): downloads_cap.append(d)
    context.on("page", _on_page)
    context.on("download", _on_dl)

    link.click()
    page.wait_for_timeout(3000)

    pagina_cert = popups_cap[-1] if popups_cap else page

    try:
        btn_vis = pagina_cert.get_by_role("button", name="Visualizar")
        btn_vis.wait_for(state="visible", timeout=8000)
    except PlaywrightTimeout:
        context.remove_listener("page", _on_page)
        context.remove_listener("download", _on_dl)
        return "nao_encontrado"

    btn_vis.click()
    page.wait_for_timeout(4000)

    context.remove_listener("page", _on_page)
    context.remove_listener("download", _on_dl)

    pagina_impr = popups_cap[-1] if popups_cap else pagina_cert

    # Impede dialogo nativo e clica Imprimir para preparar conteudo JSF
    try:
        pagina_impr.evaluate("window.print = function() {}")
        pagina_impr.get_by_text("Imprimir").click()
        page.wait_for_timeout(1500)
    except Exception:
        pass

    # Gera PDF via CDP com CSS de impressão
    try:
        pagina_impr.emulate_media(media="print")
        cdp = context.new_cdp_session(pagina_impr)
        result = cdp.send("Page.printToPDF", {
            "printBackground": True,
            "landscape": False,
            "paperWidth": 8.27,
            "paperHeight": 11.69,
            "marginTop": 0.4,
            "marginBottom": 0.4,
            "marginLeft": 0.4,
            "marginRight": 0.4,
        })
        pdf_data = base64.b64decode(result["data"])
        with open(caminho, "wb") as f:
            f.write(pdf_data)
        if len(pdf_data) > 3000:
            if pagina_impr != page:
                pagina_impr.close()
            return "ok"
    except Exception:
        pass

    return "nao_encontrado"

# ── Fluxo: Trabalhista ────────────────────────────────────────────────────────

def _capturar_img_captcha_tst(page):
    """Extrai imagem do captcha TST direto do src base64. Retorna bytes ou None."""
    try:
        img = page.locator('[id="idImgBase64"]')
        img.wait_for(state="visible", timeout=10000)
        # Aguarda o src ser preenchido com a imagem base64 real
        page.wait_for_function(
            "() => { const el = document.getElementById('idImgBase64'); "
            "return el && el.src && el.src.length > 200; }",
            timeout=10000,
        )
        src = img.get_attribute("src", timeout=5000)
        if not src or "," not in src:
            return None
        return base64.b64decode(src.split(",", 1)[1].strip())
    except Exception as e:
        print(f"  AVISO: imagem captcha nao encontrada ({e})")
        return None


def _enviar_captcha_anticaptcha(img_bytes):
    """Envia bytes da imagem para Anti-Captcha e retorna o texto. Roda em thread."""
    if not ANTI_CAPTCHA_KEY:
        print("  AVISO: ANTI_CAPTCHA_KEY nao configurada — defina a variavel de ambiente.")
        return ""
    try:
        from anticaptchaofficial.imagecaptcha import imagecaptcha
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        tmp.write(img_bytes)
        tmp.close()
        solver = imagecaptcha()
        solver.set_verbose(0)
        solver.set_key(ANTI_CAPTCHA_KEY)
        resultado = solver.solve_and_return_solution(tmp.name)
        os.unlink(tmp.name)
        if resultado and resultado != 0:
            return str(resultado)
    except Exception as e:
        print(f"  AVISO: captcha nao resolvido ({e})")
    return ""


def baixar_trabalhista(page, context, doc, caminho):
    """CNDT — portal TST (cndt-certidao.tst.jus.br)."""
    doc_limpo = re.sub(r"\D", "", doc)

    for _tentativa in range(3):
        page.goto(URL_TRABALHISTA, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(1000)

        page.get_by_role("button", name="Emitir Certidão").click()

        campo = page.get_by_role("textbox", name="Registro no Cadastro Nacional")
        try:
            campo.wait_for(state="visible", timeout=10000)
        except PlaywrightTimeout:
            return "nao_encontrado"
        campo.click()

        # Captura screenshot do captcha (thread principal) e envia ao Anti-Captcha
        # em background — enquanto isso digita o CNPJ (ambos em paralelo)
        img_bytes = _capturar_img_captcha_tst(page)
        if not img_bytes:
            return "nao_encontrado"

        with ThreadPoolExecutor(max_workers=1) as ex:
            future = ex.submit(_enviar_captcha_anticaptcha, img_bytes)
            campo.press_sequentially(doc_limpo, delay=285)
            try:
                captcha_text = future.result(timeout=90)
            except FutureTimeout:
                captcha_text = ""

        if not captcha_text:
            continue

        page.get_by_role("textbox", name="* Digite os caracteres").fill(captcha_text)

        try:
            with page.expect_download(timeout=30000) as dl:
                page.get_by_role("button", name="Emitir Certidão").click()
            dl.value.save_as(caminho)
            if os.path.exists(caminho) and os.path.getsize(caminho) > 3000:
                return "ok"
        except Exception:
            pass

        # Verifica debitos na pagina resultante
        try:
            if page.get_by_text(re.compile(
                r"devedor|pendenc|inscrit|irregular", re.I
            ), exact=False).first.is_visible(timeout=3000):
                return "com_debitos"
        except Exception:
            pass

        page.wait_for_timeout(1000)

    return "nao_encontrado"

# ── Clique real de OS (isTrusted=true, cursor restaurado) ─────────────────────

def _clicar_os(page, seletor):
    """Clica no elemento via CDP do Playwright (sem mover cursor do OS).
    Fallback para SendInput caso a página rejeite o clique CDP."""
    btn = page.locator(seletor).first
    box = btn.bounding_box()
    if not box:
        return False
    cx = box["x"] + box["width"]  / 2
    cy = box["y"] + box["height"] / 2

    # SendInput: único método que produz isTrusted=true no SEFAZ-MT
    if os.name != "nt":
        return False
    try:
        pos = page.evaluate("""() => {
            const bh = window.outerHeight - window.innerHeight;
            return { sx: window.screenX, sy: window.screenY + bh };
        }""")
        sx = int(pos["sx"] + cx)
        sy = int(pos["sy"] + cy)

        import ctypes
        PUL = ctypes.POINTER(ctypes.c_ulong)
        class MouseInput(ctypes.Structure):
            _fields_ = [("dx", ctypes.c_long), ("dy", ctypes.c_long),
                         ("mouseData", ctypes.c_ulong), ("dwFlags", ctypes.c_ulong),
                         ("time", ctypes.c_ulong), ("dwExtraInfo", PUL)]
        class Input(ctypes.Structure):
            class _I(ctypes.Union):
                _fields_ = [("mi", MouseInput)]
            _anonymous_ = ("_i",)
            _fields_  = [("type", ctypes.c_ulong), ("_i", _I)]

        sw = ctypes.windll.user32.GetSystemMetrics(0)
        sh = ctypes.windll.user32.GetSystemMetrics(1)

        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
        orig = POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(orig))

        def _send(flags, dx, dy):
            i = Input()
            i.type = 0
            i.mi.dx, i.mi.dy = int(dx * 65535 / sw), int(dy * 65535 / sh)
            i.mi.dwFlags = flags
            ctypes.windll.user32.SendInput(1, ctypes.byref(i), ctypes.sizeof(i))

        _send(0x0001 | 0x8000, sx, sy)
        time.sleep(0.015)
        _send(0x0002 | 0x8000, sx, sy)
        time.sleep(0.008)
        _send(0x0004 | 0x8000, sx, sy)
        _send(0x0001 | 0x8000, orig.x, orig.y)   # restaura imediatamente após MOUSEUP
        print(f"  SendInput click ({sx},{sy}) → cursor restaurado")
        return True
    except Exception as e:
        print(f"  SendInput falhou: {e}")
        return False

# ── Fluxo: SEFAZ-MT ───────────────────────────────────────────────────────────

def baixar_sefaz_mt(page, context, doc, caminho):
    """CND Estadual MT — SEFAZ-MT (sefaz.mt.gov.br)."""
    # Retry para ERR_CONNECTION_RESET (servidor bloqueia navegações muito rápidas)
    for tentativa in range(3):
        try:
            page.goto(URL_SEFAZ_MT, wait_until="domcontentloaded", timeout=30000)
            break
        except Exception as e:
            if tentativa < 2:
                espera = (tentativa + 1) * 12
                print(f"  SEFAZ-MT: conexão recusada, aguardando {espera}s...")
                page.wait_for_timeout(espera * 1000)
            else:
                print("  SEFAZ-MT: servidor indisponível após 3 tentativas")
                return "nao_encontrado"
    # Espera Cloudflare Turnstile resolver e formulário aparecer (até 40s)
    _sel_radio = 'input[type="radio"]'
    _sel_campo = 'input[type="text"]:visible'
    _formulario_ok = False
    for _cf in range(40):
        try:
            if page.locator(_sel_radio).first.is_visible(timeout=600):
                _formulario_ok = True
                break
        except Exception:
            pass
        # Se o checkbox do Cloudflare aparecer, clica com OS click
        try:
            cf_frame = page.frame_locator('iframe[src*="challenges.cloudflare.com"]')
            cf_cb = cf_frame.locator('input[type="checkbox"]')
            if cf_cb.is_visible(timeout=300):
                cf_cb.scroll_into_view_if_needed()
                _clicar_os(page, 'iframe[src*="challenges.cloudflare.com"]')
        except Exception:
            pass
        page.wait_for_timeout(1000)

    if not _formulario_ok:
        print("  SEFAZ-MT: formulário não apareceu após Cloudflare")
        return "nao_encontrado"

    # Aguarda TSPD/ThreatMetrix carregar tokens antes de submeter
    page.wait_for_timeout(3500)

    eh_cpf = len(re.sub(r"\D", "", doc)) == 11

    # Seleciona radio CPF ou CNPJ pelo value/label
    try:
        radio_val = "CPF" if eh_cpf else "CNPJ"
        page.locator(f'input[type="radio"][value="{radio_val}" i]').first.check()
        page.wait_for_timeout(400)
    except Exception:
        try:
            lbl = re.compile(r"cpf|f[íi]sica", re.I) if eh_cpf else re.compile(r"cnpj|j[uú]ridica", re.I)
            page.get_by_role("radio", name=lbl).first.check()
            page.wait_for_timeout(400)
        except Exception:
            pass

    # Campo de documento — exclui radio/checkbox explicitamente
    campo = page.locator(
        'input[type="text"][name*="cpfcnpj" i], '
        'input[type="text"][name*="documento" i], '
        'input[type="text"][name*="numero" i], '
        'input[type="text"][placeholder*="CPF" i], '
        'input[type="text"][placeholder*="CNPJ" i], '
        'input[type="text"]:visible'
    ).first
    try:
        campo.wait_for(state="visible", timeout=8000)
    except PlaywrightTimeout:
        return "nao_encontrado"
    campo.click()
    campo.fill(formatar_doc(doc))
    page.wait_for_timeout(500)

    # Listeners de download antes do clique
    _dl_sefaz = []
    _on_dl = lambda d: _dl_sefaz.append(d)
    context.on("download", _on_dl)

    sel_btn = 'input[type="submit"], button[type="submit"], input[type="button"][value*="ok" i]'

    # Tenta click CDP (sem mover cursor); se não navegar em 7s usa SendInput como fallback
    _nav_ok = False
    try:
        with page.expect_navigation(wait_until="domcontentloaded", timeout=7000):
            page.locator(sel_btn).first.click(timeout=3000)
        _nav_ok = True
        print("  SEFAZ-MT: submit via CDP (sem cursor)")
    except Exception:
        pass

    if not _nav_ok:
        try:
            with page.expect_navigation(wait_until="domcontentloaded", timeout=20000):
                _clicar_os(page, sel_btn)
        except Exception as e:
            print(f"  SEFAZ-MT: navegação não detectada ({e})")
            context.remove_listener("download", _on_dl)
            return "nao_encontrado"

    # Aguarda rede estabilizar
    try:
        page.wait_for_load_state("networkidle", timeout=8000)
    except Exception:
        pass
    page.wait_for_timeout(800)

    context.remove_listener("download", _on_dl)

    # Se a página resultante for erro de conexão, não prosseguir
    try:
        if page.locator("text=ERR_CONNECTION_RESET").is_visible(timeout=1000) or \
           page.locator("text=A ligação foi reposta").is_visible(timeout=500):
            print("  SEFAZ-MT: servidor resetou conexão no submit — nao_encontrado")
            return "nao_encontrado"
    except Exception:
        pass

    # Download direto (raro, mas possível)
    if _dl_sefaz:
        try:
            _dl_sefaz[0].save_as(caminho)
            return "ok" if os.path.exists(caminho) and os.path.getsize(caminho) > 3000 else "nao_encontrado"
        except Exception:
            pass

    # Popup já aberto antes de qualquer clique
    outras = [pg for pg in context.pages if pg != page]
    if outras:
        return "ok" if _capturar_popup_como_pdf(page, context, caminho) else "nao_encontrado"

    # Página intermediária: prefere "Emitir nova Certidão", aceita "Reimprimir"
    _link_cert = None
    for _sel_link in ['a:has-text("Emitir")', 'a:has-text("Reimprimir")']:
        try:
            if page.locator(_sel_link).first.is_visible(timeout=2000):
                _link_cert = _sel_link
                break
        except Exception:
            pass

    if _link_cert:
        # Log atributos do link para diagnóstico
        try:
            _dbg_href   = page.locator(_link_cert).first.get_attribute("href")    or "N/A"
            _dbg_target = page.locator(_link_cert).first.get_attribute("target")  or "N/A"
            _dbg_onclick= page.locator(_link_cert).first.get_attribute("onclick") or "N/A"
            print(f"  SEFAZ-MT: Emitir href={_dbg_href!r} target={_dbg_target!r} onclick={_dbg_onclick!r}")
        except Exception:
            pass

        # Coleta respostas application/pdf via listener de contexto (não bloqueia).
        # Polling a cada 500ms: verifica PDF real e débitos na página simultaneamente.
        _DEBITO_RE = re.compile(r"pendenc|devedor|inscrit|irregular|positiv", re.I)

        _pdf_queue = []
        def _on_pdf_resp(r):
            try:
                if r.status == 200 and "pdf" in r.headers.get("content-type", "").lower():
                    _pdf_queue.append(r)
            except Exception:
                pass

        _dl_cert = []
        def _on_dl_cert(d):
            _dl_cert.append(d)
            print(f"  SEFAZ-MT: download: {d.suggested_filename!r}")

        context.on("response", _on_pdf_resp)
        context.on("download", _on_dl_cert)

        try:
            page.evaluate("clickEmitirNova()")
            print("  SEFAZ-MT: JS clickEmitirNova() chamado")
        except Exception as _je:
            print(f"  SEFAZ-MT: JS retornou ({_je})")

        _real_pdf = None
        # Polling: 180 ticks × 500ms = 90s máximo
        for _tick in range(180):
            page.wait_for_timeout(500)

            # Processa respostas PDF acumuladas
            while _pdf_queue and not _real_pdf:
                _resp = _pdf_queue.pop(0)
                try:
                    _body = _resp.body()
                    print(f"  SEFAZ-MT: resp pdf: {len(_body)} bytes inicio={_body[:8]!r}")
                    if _body and len(_body) > 5000 and _body[:4] == b'%PDF':
                        _real_pdf = _body
                    else:
                        _txt = _body.decode("utf-8", errors="ignore")
                        if _DEBITO_RE.search(_txt):
                            context.remove_listener("response", _on_pdf_resp)
                            context.remove_listener("download", _on_dl_cert)
                            print("  SEFAZ-MT: debitos na resposta — com_debitos")
                            return "com_debitos"
                except Exception as _re:
                    print(f"  SEFAZ-MT: resp.body() erro: {_re}")

            if _real_pdf:
                break

            # Verifica débitos na página a cada ~2s (ticks pares)
            if _tick % 4 == 3:
                try:
                    if page.get_by_text(_DEBITO_RE, exact=False).first.is_visible(timeout=200):
                        context.remove_listener("response", _on_pdf_resp)
                        context.remove_listener("download", _on_dl_cert)
                        print("  SEFAZ-MT: debitos na pagina — com_debitos")
                        return "com_debitos"
                except Exception:
                    pass

        context.remove_listener("response", _on_pdf_resp)
        context.remove_listener("download", _on_dl_cert)

        if _dl_cert and not _real_pdf:
            try:
                _dl_cert[0].save_as(caminho)
                if os.path.exists(caminho) and os.path.getsize(caminho) > 5000:
                    return "ok"
            except Exception as _de:
                print(f"  SEFAZ-MT: erro ao salvar download: {_de}")

        if _real_pdf:
            with open(caminho, "wb") as _f:
                _f.write(_real_pdf)
            if os.path.exists(caminho) and os.path.getsize(caminho) > 5000:
                print(f"  SEFAZ-MT: certidao salva ({os.path.getsize(caminho)} bytes)")
                return "ok"
            return "nao_encontrado"

        # Sem PDF real — ultima verificacao de debitos na pagina atual
        try:
            if page.get_by_text(_DEBITO_RE, exact=False).first.is_visible(timeout=2000):
                print("  SEFAZ-MT: debitos detectados — com_debitos")
                return "com_debitos"
        except Exception:
            pass
        return "nao_encontrado"

    # Verifica débitos na página resultante
    try:
        if page.get_by_text(re.compile(
            r"devedor|pendenc|inscrit|irregular", re.I
        ), exact=False).first.is_visible(timeout=2000):
            return "com_debitos"
    except Exception:
        pass

    # Renderiza a página de resultado como PDF
    try:
        page.emulate_media(media="print")
        page.pdf(path=caminho, format="A4", print_background=True)
        return "ok" if os.path.exists(caminho) and os.path.getsize(caminho) > 3000 else "nao_encontrado"
    except Exception as e:
        print(f"  SEFAZ-MT: erro ao gerar PDF: {e}")
        return "nao_encontrado"

# ── Argumentos ────────────────────────────────────────────────────────────────

if len(sys.argv) < 2:
    print("Uso: python robo_geral.py federal,fgts,trabalhista,estadual_mt")
    print("     python robo_geral.py todas")
    sys.exit(1)

arg_tipos = sys.argv[1].lower().strip()
if arg_tipos == "todas":
    tipos = TIPOS_VALIDOS[:]
else:
    tipos = [t.strip() for t in arg_tipos.split(",") if t.strip() in TIPOS_VALIDOS]

if not tipos:
    print(f"Nenhum tipo valido em '{arg_tipos}'. Opcoes: {', '.join(TIPOS_VALIDOS)}")
    sys.exit(1)

# ── Leitura das fontes de dados ───────────────────────────────────────────────

empresas   = _ler_empresas()
produtores = _ler_produtores()

# Monta lista de tarefas: (tipo_certidao x docs)
todos = []
for tipo in tipos:
    for reg in empresas:
        todos.append({**reg, "certidao": tipo})
    for reg in produtores:
        todos.append({**reg, "certidao": tipo})

# Remove duplicatas (mesmo doc + mesmo tipo)
vistos = set()
todos_dedup = []
for item in todos:
    key = (item["doc"], item["certidao"])
    if key not in vistos:
        vistos.add(key)
        todos_dedup.append(item)
todos = todos_dedup

total = len(todos)
print(f"Modo: {arg_tipos.upper()} | {total} CNPJs selecionados.\n")

if total == 0:
    print("Nenhum documento encontrado. Verifique a planilha.")
    sys.exit(0)

# Cria subpastas por tipo (compativel com o scanner do server.py)
for tipo in tipos:
    os.makedirs(os.path.join(PASTA_NEGATIVAS, tipo), exist_ok=True)
    os.makedirs(os.path.join(PASTA_POSITIVAS, tipo), exist_ok=True)

BAIXAR_FN = {
    "federal":     baixar_federal,
    "fgts":        baixar_fgts,
    "trabalhista": baixar_trabalhista,
    "estadual_mt": baixar_sefaz_mt,
}

erros            = []
nao_encontrados  = []
com_debitos_list = []
ja_existentes    = []

# No Railway (WEB_MODE), usa display virtual para rodar headed e evitar captcha Cloudflare
_vdisplay = None
if WEB_MODE and HEADLESS:
    try:
        from pyvirtualdisplay import Display
        _vdisplay = Display(visible=False, size=(1920, 1080), color_depth=24)
        _vdisplay.start()
        HEADLESS = False
        print("  Display virtual (Xvfb) iniciado — modo headed ativo")
    except Exception as e:
        print(f"  AVISO: Xvfb indisponivel ({e}) — mantendo headless")

def _find_chrome():
    import glob as _glob
    caminhos = [
        # Windows
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.join(os.path.expanduser("~"), r"AppData\Local\Google\Chrome\Application\chrome.exe"),
        # Linux
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/snap/bin/chromium",
        "/snap/bin/google-chrome",
        # macOS
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    ]
    for c in caminhos:
        if os.path.exists(c):
            return c
    # Fallback: Chromium embutido do Playwright (Linux / Windows)
    for pattern in [
        os.path.expanduser("~/.cache/ms-playwright/chromium-*/chrome-linux/chrome"),
        "/root/.cache/ms-playwright/chromium-*/chrome-linux/chrome",
        "/ms-playwright/chromium-*/chrome-linux/chrome",
        os.path.join(os.path.expanduser("~"),
                     r"AppData\Local\ms-playwright\chromium-*\chrome-win\chrome.exe"),
    ]:
        matches = _glob.glob(pattern)
        if matches:
            return sorted(matches)[-1]
    return None

_chrome_exe = _find_chrome()
if _chrome_exe:
    print(f"  Usando Chrome: {_chrome_exe}")
else:
    print("  Chrome nao encontrado, usando Chromium do Playwright")

_CHROME_ARGS = [
    "--disable-blink-features=AutomationControlled",
    "--no-first-run",
    "--no-default-browser-check",
    "--disable-infobars",
    "--disable-extensions",
    "--disable-sync",
    "--hide-crash-restore-bubble",
    "--no-restore-state",
]
_CONTEXT_KWARGS = dict(
    accept_downloads=True,
    locale="pt-BR",
    timezone_id="America/Sao_Paulo",
    viewport={"width": 1366, "height": 768},
)

def _criar_perfil_chrome_minimo():
    """Copia apenas cookies + Local State do Chrome real para um dir temporario.
    Isso permite usar cookies autenticos sem session restore (sem hang)."""
    if os.name != "nt" or not _chrome_exe:
        return None
    src_ud = os.path.join(os.path.expanduser("~"),
                          r"AppData\Local\Google\Chrome\User Data")
    if not os.path.isdir(src_ud):
        return None
    try:
        tmp = tempfile.mkdtemp(prefix="robo_chrome_")
        # Local State contem a chave AES para decriptar cookies
        ls = os.path.join(src_ud, "Local State")
        if os.path.exists(ls):
            shutil.copy2(ls, os.path.join(tmp, "Local State"))
        # Copia o banco de cookies
        ck_src = os.path.join(src_ud, "Default", "Network", "Cookies")
        if os.path.exists(ck_src):
            ck_dst = os.path.join(tmp, "Default", "Network")
            os.makedirs(ck_dst, exist_ok=True)
            shutil.copy2(ck_src, os.path.join(ck_dst, "Cookies"))
        print(f"  Perfil Chrome minimo criado em: {tmp}")
        return tmp
    except Exception as e:
        print(f"  AVISO: nao foi possivel copiar perfil Chrome ({e})")
        return None

def _launch_chrome_cdp(porta=9222, exe=None):
    """Lança Chrome via subprocess com remote-debugging (sem flags de automação Playwright).
    Usa --user-data-dir isolado para garantir nova instância mesmo que Chrome já esteja aberto.
    Copia cookies do Chrome real para o perfil CDP."""
    exe = exe or _chrome_exe
    if not exe:
        return None, None
    try:
        tmp_dir = tempfile.mkdtemp(prefix="robo_cdp_")
        # Copia perfil real para o CDP: cookies + localStorage + fingerprint do ThreatMetrix
        src_ud = os.path.join(os.path.expanduser("~"),
                              r"AppData\Local\Google\Chrome\User Data")
        if os.path.isdir(src_ud):
            src_def = os.path.join(src_ud, "Default")
            dst_def = os.path.join(tmp_dir, "Default")
            os.makedirs(dst_def, exist_ok=True)

            def _cp(rel_src, rel_dst=None):
                s = os.path.join(src_ud, rel_src)
                d = os.path.join(tmp_dir, rel_dst or rel_src)
                if os.path.exists(s):
                    os.makedirs(os.path.dirname(d), exist_ok=True)
                    try:
                        shutil.copy2(s, d)
                    except Exception:
                        pass

            def _cpdir(rel_src):
                s = os.path.join(src_ud, rel_src)
                d = os.path.join(tmp_dir, rel_src)
                if os.path.isdir(s):
                    try:
                        shutil.copytree(s, d, dirs_exist_ok=True)
                    except Exception:
                        pass

            _cp("Local State")                          # chave AES p/ cookies
            _cp(r"Default\Network\Cookies")             # cookies de sessão
            _cp(r"Default\Preferences")                 # config browser
            _cp(r"Default\Web Data")                    # histórico de formulários
            _cpdir(r"Default\Local Storage")            # localStorage (ThreatMetrix/TSPD)
            _cpdir(r"Default\Session Storage")          # sessionStorage
            _cpdir(r"Default\IndexedDB")                # IndexedDB (fingerprint tokens)
        cmd = [
            exe,
            f"--remote-debugging-port={porta}",
            f"--user-data-dir={tmp_dir}",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-extensions",
            "--disable-sync",
            "--hide-crash-restore-bubble",
            "--no-restore-state",
            "--disable-popup-blocking",
        ]
        # Docker/Railway: root + /dev/shm pequeno exigem estas flags
        if os.name != "nt":
            cmd += ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print(f"  Chrome CDP lançado (porta {porta}, perfil isolado)")
        return proc, tmp_dir
    except Exception as e:
        print(f"  Erro ao lançar Chrome CDP: {e}")
        return None, None


with sync_playwright() as p:
    _chrome_proc = None
    _cdp_tmp_dir = None
    _perfil_tmp  = None
    browser      = None
    context      = None

    # Tenta abordagem CDP: Chrome lançado sem flags de automação Playwright
    # Se não achou Chrome/Chromium no sistema, usa o binário bundled do Playwright
    _cdp_exe = _chrome_exe or p.chromium.executable_path
    if _cdp_exe:
        if not _chrome_exe:
            print(f"  Usando Chromium bundled do Playwright para CDP: {_cdp_exe}")
        _chrome_proc, _cdp_tmp_dir = _launch_chrome_cdp(9222, exe=_cdp_exe)
        if _chrome_proc:
            # Polling: tenta conectar a cada 500ms por até 15s
            browser = None
            for _t in range(30):
                time.sleep(0.5)
                try:
                    browser = p.chromium.connect_over_cdp("http://localhost:9222")
                    break
                except Exception:
                    pass
            if browser:
                ctxs = browser.contexts
                context = ctxs[0] if ctxs else browser.new_context(**_CONTEXT_KWARGS)
                print("  Conectado ao Chrome via CDP (sem marcadores de automação)")
            if not browser:
                e = "porta 9222 não respondeu em 15s"
                print(f"  CDP falhou: {e} — usando perfil mínimo como fallback")
                try:
                    _chrome_proc.terminate()
                except Exception:
                    pass
                _chrome_proc = None
                if _cdp_tmp_dir:
                    shutil.rmtree(_cdp_tmp_dir, ignore_errors=True)
                _cdp_tmp_dir = None

    if context is None:
        _perfil_tmp = _criar_perfil_chrome_minimo() if not WEB_MODE else None

    if context is None and _perfil_tmp and _chrome_exe:
        try:
            context = p.chromium.launch_persistent_context(
                _perfil_tmp,
                executable_path=_chrome_exe,
                headless=HEADLESS,
                slow_mo=30,
                args=_CHROME_ARGS,
                **_CONTEXT_KWARGS,
            )
            print("  Usando perfil Chrome com cookies reais (sem session restore)")
        except Exception as e:
            print(f"  AVISO: perfil minimo falhou ({e}), usando contexto limpo")
            _perfil_tmp = None
            context = None

    if context is None:
        _launch_args = dict(headless=HEADLESS, slow_mo=30, args=_CHROME_ARGS)
        if _chrome_exe:
            _launch_args["executable_path"] = _chrome_exe
        browser = p.chromium.launch(**_launch_args)
        context = browser.new_context(**_CONTEXT_KWARGS)

    _is_cdp_mode = _chrome_proc is not None

    page = context.new_page()
    if _is_cdp_mode:
        # Chrome foi lançado sem --enable-automation: navigator.webdriver já é false.
        # Não injetar scripts (Page.addScriptToEvaluateOnNewDocument é detectável pelo PX).
        print("  Modo CDP: sem injeção de scripts de automação")
    elif _HAS_STEALTH:
        _stealth_sync(page)
    else:
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )

    for i, reg in enumerate(todos, start=1):
        doc   = reg["doc"]
        nome  = reg.get("nome", "")
        tipo  = reg["certidao"]
        label = NOMES_TIPO.get(tipo, tipo)

        nome_exibir = f" | {nome}" if nome else ""
        print(f"\n[{i}/{total}] CNPJ: {doc}{nome_exibir} | {label}")

        # Fecha abas extras de iteracoes anteriores
        for aba in context.pages:
            if aba != page:
                aba.close()

        caminho_temp = os.path.join(PASTA_BASE, f"{doc}_{tipo}.pdf")
        caminho_neg  = os.path.join(PASTA_NEGATIVAS, tipo, f"{doc}.pdf")
        caminho_pos  = os.path.join(PASTA_POSITIVAS, tipo, f"{doc}.pdf")

        if os.path.exists(caminho_neg) or os.path.exists(caminho_pos):
            print(f"  PDF ja existe, pulando...")
            ja_existentes.append(f"{doc} ({tipo})")
            continue

        try:
            fn        = BAIXAR_FN[tipo]
            resultado = fn(page, context, doc, caminho_temp)
            # Pausa anti-rate-limit entre CNPJs do SEFAZ-MT
            if tipo == "estadual_mt" and i < total:
                page.wait_for_timeout(4000)

            if resultado == "ok":
                tipo_cert = detectar_tipo_certidao(caminho_temp)
                destino   = caminho_neg if tipo_cert == "negativa" else caminho_pos
                shutil.move(caminho_temp, destino)
                print(f"  Certidao {tipo_cert} salva: {destino} ({label})")
            elif resultado == "nao_encontrado":
                print(f"  CNPJ nao encontrado na base, pulando...")
                nao_encontrados.append(f"{doc} ({tipo})")
            elif resultado == "com_debitos":
                print(f"  CNPJ com debitos — certidao nao emitida.")
                com_debitos_list.append(f"{doc} ({tipo})")
            else:
                print(f"  Falha: {resultado}")
                erros.append(f"{doc} ({tipo}: {resultado})")

        except PlaywrightTimeout as e:
            print(f"  TIMEOUT: {e}")
            erros.append(f"{doc} ({tipo}: timeout)")
        except Exception as e:
            print(f"  ERRO: {e}")
            erros.append(f"{doc} ({tipo}: erro)")

    context.close()
    try:
        browser.close()
    except Exception:
        pass
    if _chrome_proc:
        try:
            _chrome_proc.terminate()
        except Exception:
            pass
    if _cdp_tmp_dir and os.path.isdir(_cdp_tmp_dir):
        shutil.rmtree(_cdp_tmp_dir, ignore_errors=True)
    if _perfil_tmp and os.path.isdir(_perfil_tmp):
        shutil.rmtree(_perfil_tmp, ignore_errors=True)

sucesso = total - len(erros) - len(nao_encontrados) - len(com_debitos_list) - len(ja_existentes)

linhas = [
    "",
    "--- Processamento finalizado ---",
    f"Sucesso:             {sucesso}/{total}",
    f"Ja existia (pulou):  {len(ja_existentes)}/{total}",
    f"Nao encontrado:      {len(nao_encontrados)}/{total}",
]
if nao_encontrados:
    linhas.append(f"  Docs: {', '.join(nao_encontrados)}")
linhas.append(f"Com debitos:         {len(com_debitos_list)}/{total}")
if com_debitos_list:
    linhas.append(f"  Docs: {', '.join(com_debitos_list)}")
if erros:
    linhas.append(f"Erros:               {len(erros)}/{total}")
    linhas.append(f"  Docs: {', '.join(erros)}")

relatorio = "\n".join(linhas)
print(relatorio)

nome_rel = f"relatorio_geral_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
with open(os.path.join(PASTA_BASE, nome_rel), "w", encoding="utf-8") as f:
    f.write(relatorio)
print(f"\nRelatorio salvo em: {os.path.join(PASTA_BASE, nome_rel)}")

if _vdisplay is not None:
    try:
        _vdisplay.stop()
    except Exception:
        pass
