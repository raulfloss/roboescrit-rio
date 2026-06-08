"""
robo_geral.py — Certidoes Federais, FGTS, Trabalhista e Estadual MT

Uso:
  python robo_geral.py federal
  python robo_geral.py federal,fgts,trabalhista,estadual_mt
  python robo_geral.py todas
"""

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
from datetime import datetime
import pandas as pd
import unicodedata
import base64
import shutil
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

def _ler_empresas():
    """Le CNPJs da aba principal."""
    try:
        df = pd.read_excel(ARQUIVO_EXCEL, dtype=str).fillna("")
        rename = {}
        for col in df.columns:
            for esp in [COLUNA_CNPJ, COLUNA_NOME, COLUNA_CIDADE]:
                if _norm(col) == _norm(esp) and col != esp:
                    rename[col] = esp
        if rename:
            df = df.rename(columns=rename)
        if COLUNA_CNPJ not in df.columns:
            print(f"  AVISO: coluna '{COLUNA_CNPJ}' nao encontrada na aba principal.")
            return []
        df[COLUNA_CNPJ] = df[COLUNA_CNPJ].str.replace(r"[.\-/\s]", "", regex=True).str.strip()
        out = []
        for _, row in df.iterrows():
            cnpj = str(row.get(COLUNA_CNPJ, "")).strip()
            if len(cnpj) == 14 and cnpj.isdigit():
                out.append({
                    "doc":  cnpj,
                    "nome": str(row.get(COLUNA_NOME, "")).strip(),
                })
        return out
    except Exception as e:
        print(f"  ERRO ao ler empresas: {e}")
        return []

def _ler_produtores():
    """Le CPFs da aba 'produtores ativos'."""
    try:
        df = pd.read_excel(ARQUIVO_EXCEL, sheet_name=ABA_PRODUTORES, dtype=str).fillna("")

        # Encontra coluna de CPF/documento
        col_doc = None
        for col in df.columns:
            n = _norm(col)
            if "cpf" in n or _norm(COLUNA_CNPJ) == n:
                col_doc = col
                break
        if col_doc is None:
            # Busca primeira coluna com valores de 11 digitos
            for col in df.columns:
                vals = df[col].str.replace(r"\D", "", regex=True).str.strip()
                if vals.str.len().eq(11).any():
                    col_doc = col
                    break
        if col_doc is None:
            print(f"  AVISO: coluna de CPF nao encontrada na aba '{ABA_PRODUTORES}'.")
            return []

        col_nome = None
        for col in df.columns:
            n = _norm(col)
            if "nome" in n or "razao" in n or _norm(COLUNA_NOME) == n:
                col_nome = col
                break

        df[col_doc] = df[col_doc].str.replace(r"\D", "", regex=True).str.strip()
        out = []
        for _, row in df.iterrows():
            cpf = str(row[col_doc]).strip()
            if len(cpf) == 11 and cpf.isdigit():
                out.append({
                    "doc":  cpf,
                    "nome": str(row.get(col_nome, "")).strip() if col_nome else "",
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

def baixar_federal(page, context, doc, caminho):
    """CND Federal — Receita Federal (servicos.receitafederal.gov.br)."""
    page.goto(URL_FEDERAL, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(2000)

    # SPA pode ter submenu; tenta navegar para CND se houver link
    try:
        link = page.get_by_role("link", name=re.compile(
            r"certid.o de d.bitos|CND|d.bitos relativo", re.I
        )).first
        if link.is_visible(timeout=3000):
            link.click()
            page.wait_for_timeout(1500)
    except Exception:
        pass

    # Campo CNPJ/CPF
    campo = page.locator(
        'input[placeholder*="CNPJ"], input[placeholder*="CPF"], '
        'input[id*="cnpj" i], input[id*="cpf" i], '
        'input[name*="numCpfCnpj" i], input[name*="cnpj" i], '
        'input[type="text"]:visible'
    ).first
    try:
        campo.wait_for(state="visible", timeout=15000)
    except PlaywrightTimeout:
        return "nao_encontrado"
    campo.click()
    campo.fill(formatar_doc(doc))
    page.wait_for_timeout(500)

    # Botao de consulta
    try:
        page.get_by_role("button", name=re.compile(
            r"emitir|gerar|consultar|pesquisar|confirmar", re.I
        )).first.click()
    except Exception:
        page.keyboard.press("Enter")
    page.wait_for_timeout(3000)

    # Captura download direto
    try:
        with page.expect_download(timeout=15000) as dl:
            try:
                page.get_by_role("button", name=re.compile(
                    r"baixar|download|salvar|pdf|imprimir", re.I
                )).first.click()
            except Exception:
                page.get_by_role("link", name=re.compile(
                    r"baixar|download|pdf|certid", re.I
                )).first.click()
        dl.value.save_as(caminho)
        return "ok"
    except Exception:
        pass

    if _capturar_popup_como_pdf(page, context, caminho):
        return "ok"

    return "nao_encontrado"

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

def _resolver_captcha_tst(page):
    """Resolve o captcha de imagem do TST via Anti-Captcha (mesmo padrao do robo.py)."""
    if not ANTI_CAPTCHA_KEY:
        print("  AVISO: ANTI_CAPTCHA_KEY nao configurada — defina a variavel de ambiente.")
        return ""
    try:
        import tempfile
        from anticaptchaofficial.imagecaptcha import imagecaptcha
        img = page.locator('[id="idImgBase64"]')
        img.wait_for(state="visible", timeout=10000)
        img_bytes = img.screenshot(timeout=10000)
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
        campo.press_sequentially(doc_limpo, delay=285)

        captcha_text = _resolver_captcha_tst(page)
        if not captcha_text:
            return "nao_encontrado"

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

# ── Fluxo: SEFAZ-MT ───────────────────────────────────────────────────────────

def baixar_sefaz_mt(page, context, doc, caminho):
    """CND Estadual MT — SEFAZ-MT (sefaz.mt.gov.br)."""
    page.goto(URL_SEFAZ_MT, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(1000)

    eh_cpf = len(re.sub(r"\D", "", doc)) == 11

    # Seleciona tipo (CPF/CNPJ) se existir radio ou select na pagina
    try:
        if eh_cpf:
            page.get_by_role("radio", name=re.compile(r"cpf|f[íi]sica", re.I)).first.check()
        else:
            page.get_by_role("radio", name=re.compile(r"cnpj|j[uú]ridica", re.I)).first.check()
        page.wait_for_timeout(400)
    except Exception:
        pass

    try:
        sel = page.locator('select[name*="tipo" i], select[id*="tipo" i]').first
        if sel.is_visible(timeout=1500):
            sel.select_option(index=(1 if eh_cpf else 2))
            page.wait_for_timeout(300)
    except Exception:
        pass

    # Campo de documento
    campo = page.locator(
        'input[name*="cpfcnpj" i], input[name*="documento" i], '
        'input[id*="cpf" i], input[id*="cnpj" i], '
        'input[placeholder*="CPF"], input[placeholder*="CNPJ"], '
        'input[type="text"]:visible'
    ).first
    try:
        campo.wait_for(state="visible", timeout=15000)
    except PlaywrightTimeout:
        return "nao_encontrado"
    campo.click()
    campo.fill(formatar_doc(doc))
    page.wait_for_timeout(500)

    # Submit
    try:
        page.get_by_role("button", name=re.compile(
            r"emitir|gerar|consultar|pesquisar|ok|enviar", re.I
        )).first.click()
    except Exception:
        try:
            page.locator('input[type="submit"], button[type="submit"]').first.click()
        except Exception:
            page.keyboard.press("Enter")
    page.wait_for_timeout(3000)

    # Captura popup / nova aba
    if _capturar_popup_como_pdf(page, context, caminho):
        return "ok"

    # Download direto
    try:
        with page.expect_download(timeout=15000) as dl:
            try:
                page.get_by_role("button", name=re.compile(
                    r"baixar|download|pdf|imprimir|salvar", re.I
                )).first.click()
            except Exception:
                pass
        dl.value.save_as(caminho)
        return "ok"
    except Exception:
        pass

    # PDF inline (portais Java que renderizam na mesma aba)
    try:
        page.emulate_media(media="print")
        page.pdf(path=caminho, format="A4", print_background=True)
        if os.path.exists(caminho) and os.path.getsize(caminho) > 3000:
            return "ok"
    except Exception:
        pass

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

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=HEADLESS,
        slow_mo=30,
        args=["--disable-blink-features=AutomationControlled"],
    )
    context = browser.new_context(
        accept_downloads=True,
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
    )
    context.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    page = context.new_page()
    if _HAS_STEALTH:
        _stealth_sync(page)

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
    browser.close()

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
