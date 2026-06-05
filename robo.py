from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
from anticaptchaofficial.imagecaptcha import imagecaptcha
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeout
from datetime import datetime
import pandas as pd
import unicodedata
import requests
import tempfile
import base64
import platform
import shutil
import json
import sys
import re
import os

_WINDOWS = platform.system() == "Windows"
if _WINDOWS:
    import msvcrt
    import winsound
    import ctypes

try:
    import pypdf
    _PYPDF_OK = True
except ImportError:
    _PYPDF_OK = False

ANTICAPTCHA_KEY = "0b3f1fe79c286e30486d7739165822d2"

# Localiza a pasta do exe (ou do script) para encontrar o Excel corretamente
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ARQUIVO_EXCEL = os.path.join(BASE_DIR, "cnpjs.xlsx.xlsx")
COLUNA_CNPJ   = "CNPJ (MF) N.º"
COLUNA_CIDADE = "Cidade / UF"
COLUNA_NOME   = "Razão Social"

WEB_MODE = os.environ.get("ROBO_WEB") == "1"

_JOB_ID = os.environ.get("ROBO_JOB_ID", "")
_BASE_CERT = os.path.join(os.path.expanduser("~"), "Documents", "certidões")
PASTA_DOWNLOADS  = os.path.join(_BASE_CERT, _JOB_ID) if _JOB_ID else _BASE_CERT
PASTA_NEGATIVAS  = os.path.join(PASTA_DOWNLOADS, "negativas")
PASTA_POSITIVAS  = os.path.join(PASTA_DOWNLOADS, "positivas")
os.makedirs(PASTA_NEGATIVAS, exist_ok=True)
os.makedirs(PASTA_POSITIVAS, exist_ok=True)

# --- Configuração dos sistemas por cidade (carregada do JSON externo) ---
# "sistema": "betha"  → usa o portal Betha (seleção de estado/município)
# "sistema": "gpsrv"  → usa o portal gp.srv.br (fluxo específico por URL)
_CIDADES_CONFIG_FILE = os.path.join(BASE_DIR, "cidades_config.json")

def _carregar_cidades_config():
    if os.path.exists(_CIDADES_CONFIG_FILE):
        try:
            with open(_CIDADES_CONFIG_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

CIDADES_CONFIG = _carregar_cidades_config()

# Código numérico dos estados no portal Betha
ESTADOS_BETHA = {
    "MT": "25",
    "MS": "24",
}

URL_BETHA = "https://e-gov.betha.com.br/cdweb/03114-559/main.faces"

# -------------------------------------------------------

def detectar_tipo_certidao(caminho):
    """Lê o PDF e retorna 'negativa' ou 'positiva' com base no texto."""
    if not _PYPDF_OK:
        return "negativa"
    try:
        reader = pypdf.PdfReader(caminho)
        texto = " ".join(p.extract_text() or "" for p in reader.pages).upper()
        # "Positiva com Efeitos de Negativa" também é tratada como positiva
        if "CERTID" in texto and "NEGATIVA" in texto:
            return "negativa"
        if "CERTID" in texto and "POSITIVA" in texto:
            return "positiva"
    except Exception:
        pass
    return "negativa"  # padrão se não conseguir ler

def normalizar(texto):
    texto = unicodedata.normalize("NFD", str(texto))
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto.strip().title()

def extrair_cidade_estado(cidade_uf):
    cidade_uf = str(cidade_uf).strip()
    for sep in ["/", "-"]:
        if sep in cidade_uf:
            partes = cidade_uf.split(sep)
            return normalizar(partes[0]), partes[-1].strip().upper()[:2]
    return normalizar(cidade_uf), "MT"

def cnpj_formatado(cnpj):
    c = cnpj.zfill(14)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"

def _gerar_pdf_headless(playwright_inst, pdf_url, cookies_sessao, caminho):
    """Abre browser headless para salvar a certidão como PDF."""
    br = playwright_inst.chromium.launch(headless=True)
    try:
        ctx_pdf = br.new_context()
        ctx_pdf.add_cookies(cookies_sessao)
        pag_pdf = ctx_pdf.new_page()
        pag_pdf.goto(pdf_url, wait_until="networkidle", timeout=30000)
        pag_pdf.emulate_media(media="print")
        pag_pdf.pdf(path=caminho, format="A4", print_background=True)
    finally:
        br.close()

# -------------------------------------------------------

def processar_betha(page, cnpj, cidade, estado, caminho):
    """Fluxo para o portal Betha (Apiacas, Nova Andradina e outros)."""
    codigo_estado   = ESTADOS_BETHA.get(estado)
    municipio_label = f"Prefeitura Municipal de {cidade}"

    if not codigo_estado:
        return "estado_nao_suportado"

    page.goto(URL_BETHA, wait_until="domcontentloaded", timeout=30000)
    page.locator('[id="mainForm:estados"]').wait_for(state="visible", timeout=30000)
    page.locator('[id="mainForm:estados"]').select_option(codigo_estado)
    page.wait_for_timeout(600)

    # Alguns estados (ex: MS) usam campo de busca; outros (ex: MT) usam dropdown
    search = page.get_by_role("textbox", name="Digite para pesquisar")
    try:
        search.wait_for(state="visible", timeout=4000)
        search.click()
        search.fill(cidade)
        page.get_by_role("link", name=f"Prefeitura Municipal de {cidade}").wait_for(state="visible", timeout=10000)
        page.get_by_role("link", name=f"Prefeitura Municipal de {cidade}").click()
    except PlaywrightTimeout:
        # Dropdown padrão (MT)
        try:
            page.locator('[id="mainForm:municipios"]').wait_for(state="visible", timeout=5000)
            page.locator('[id="mainForm:municipios"]').select_option(label=municipio_label)
        except Exception:
            return "prefeitura_nao_encontrada"
    page.wait_for_timeout(400)

    page.get_by_role("link", name="Acessar").click()
    page.get_by_role("link", name="Certidão negativa de contribuinte").wait_for(state="visible", timeout=30000)
    page.get_by_role("link", name="Certidão negativa de contribuinte").click()
    page.get_by_role("link", name="CNPJ").wait_for(state="visible", timeout=30000)
    page.get_by_role("link", name="CNPJ").click()

    campo = page.locator('input[placeholder*="CNPJ"], input[id*="cnpj" i], input[name*="cnpj" i]').first
    campo.wait_for(state="visible", timeout=20000)
    page.wait_for_timeout(800)
    campo.click(force=True)
    page.keyboard.press("Home")
    campo.press_sequentially(cnpj, delay=40)
    page.wait_for_timeout(500)

    page.locator('input[name="mainForm:btCnpj"]').click()

    try:
        page.get_by_role("img", name="Emitir").wait_for(state="visible", timeout=5000)
    except PlaywrightTimeout:
        return "nao_encontrado"

    page.get_by_role("img", name="Emitir").click()
    page.locator('iframe[name^="fancybox-frame"]').wait_for(state="attached", timeout=30000)
    page.wait_for_timeout(1000)

    frame = page.frame_locator('iframe[name^="fancybox-frame"]')
    with page.expect_download(timeout=30000) as dl:
        frame.get_by_role("button", name="Salvar").click()
    dl.value.save_as(caminho)
    return "ok"

# -------------------------------------------------------

def processar_agili(page, context, playwright_inst, cnpj, config, caminho):
    """Fluxo para o portal Agili Blue (Nova Bandeirantes e outros)."""
    page.goto(config["url"], wait_until="domcontentloaded", timeout=30000)

    page.get_by_text("Jurídica").wait_for(state="visible", timeout=20000)
    page.get_by_text("Jurídica").click()

    campo = page.get_by_role("textbox", name="__.___.___/____-__")
    campo.wait_for(state="visible", timeout=15000)
    campo.click()
    campo.fill(cnpj_formatado(cnpj))
    page.wait_for_timeout(1000)

    page.get_by_role("button", name="Imprimir").click()

    page.get_by_role("radio", name="Nova certidão ou boletim").wait_for(state="visible", timeout=15000)
    page.get_by_role("radio", name="Nova certidão ou boletim").check()

    with page.expect_popup(timeout=30000) as popup_info:
        page.get_by_role("button", name="Continuar").click()

    popup = popup_info.value
    popup.wait_for_load_state("load", timeout=30000)
    page.wait_for_timeout(1000)
    pdf_url = popup.url
    cookies_sessao = context.cookies()
    popup.close()

    if not pdf_url or pdf_url in ("about:blank", ":") or not pdf_url.startswith("http"):
        return "nao_encontrado"

    _gerar_pdf_headless(playwright_inst, pdf_url, cookies_sessao, caminho)
    return "ok"

# -------------------------------------------------------

def processar_i7sgp(page, context, playwright_inst, cnpj, config, caminho):
    """Fluxo para o portal i7sgp (Nova Canaã do Norte e outros)."""
    page.goto(config["url"], wait_until="domcontentloaded", timeout=30000)

    page.get_by_role("link", name="Funcionalidades relacionadas ao contribuinte").wait_for(state="visible", timeout=20000)
    page.get_by_role("link", name="Funcionalidades relacionadas ao contribuinte").click()

    page.get_by_role("radio", name="Pessoa Jurídica").wait_for(state="visible", timeout=15000)
    page.get_by_role("radio", name="Pessoa Jurídica").check()

    campo = page.locator('[id="compInformarContribuinte:formNumero:itIdent"]')
    campo.wait_for(state="visible", timeout=15000)
    campo.click()
    campo.fill(cnpj_formatado(cnpj))
    page.wait_for_timeout(500)

    page.get_by_role("button", name="OK").click()
    page.wait_for_timeout(2000)

    try:
        page.get_by_role("link", name="Emitir certidão negativa de d").wait_for(state="visible", timeout=10000)
    except PlaywrightTimeout:
        return "nao_encontrado"

    page.get_by_role("link", name="Emitir certidão negativa de d").click()

    with page.expect_popup(timeout=30000) as popup_info:
        page.get_by_role("button", name="Imprimir Certidão").click()

    popup = popup_info.value
    popup.wait_for_load_state("networkidle", timeout=30000)
    pdf_url = popup.url
    cookies_sessao = context.cookies()
    popup.close()

    if not pdf_url or pdf_url in ("about:blank", ":") or not pdf_url.startswith("http"):
        return "nao_encontrado"

    _gerar_pdf_headless(playwright_inst, pdf_url, cookies_sessao, caminho)
    return "ok"

# -------------------------------------------------------

def processar_gpsrv(page, context, cnpj, config, caminho):
    """Fluxo para o portal gp.srv.br (Alta Floresta, Matupá, Sinop e outros)."""
    tipo_certidao           = config.get("tipo_certidao", "1")
    clicar_emitir           = config.get("clicar_emitir", True)
    clicar_table_finalidade = config.get("clicar_table_finalidade", False)

    page.goto(config["url"], wait_until="domcontentloaded", timeout=30000)

    # Clica em "Emitir Certidão" apenas para portais que precisam (ex: Alta Floresta, Matupá)
    if clicar_emitir:
        page.locator("a").filter(has_text="Emitir Certidão").wait_for(state="visible", timeout=20000)
        page.locator("a").filter(has_text="Emitir Certidão").click()

    select_tipo = page.locator("#W0005W0006vCERTIDAO_TIPO_CERTIDAO_ID")
    select_tipo.wait_for(state="visible", timeout=15000)
    page.wait_for_function(
        "document.querySelector('#W0005W0006vCERTIDAO_TIPO_CERTIDAO_ID')?.options.length > 1",
        timeout=15000
    )
    select_tipo.select_option(tipo_certidao)
    page.evaluate("""
        var el = document.getElementById('W0005W0006vCERTIDAO_TIPO_CERTIDAO_ID');
        el.dispatchEvent(new Event('change', {bubbles: true}));
    """)
    page.wait_for_timeout(800)

    campo_cnpj = page.locator("#W0005W0006vCERTIDAO_CONT_PES_CPF_CNPJ_MASC")
    campo_cnpj.wait_for(state="visible", timeout=15000)
    campo_cnpj.click()
    campo_cnpj.fill(cnpj_formatado(cnpj))

    # Sinop exige clicar na célula TABLE_FINALIDADE antes de selecionar a finalidade
    if clicar_table_finalidade:
        page.locator("#W0005W0006TABLE_FINALIDADE").click()
        page.wait_for_timeout(500)

    page.locator("#W0005W0006vCERTIDAO_FINALIDADE").wait_for(state="visible", timeout=10000)
    page.locator("#W0005W0006vCERTIDAO_FINALIDADE").select_option("2")

    # --- CAPTCHA ---
    campo_captcha = page.locator("#W0005W0006vVALOR_IMAGEM")
    campo_captcha.wait_for(state="visible", timeout=10000)
    campo_captcha.scroll_into_view_if_needed()
    page.wait_for_timeout(500)

    valor_captcha = ""
    try:
        img_bytes = page.locator('img[src*="Kaptcha"]').screenshot(timeout=10000)

        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        tmp.write(img_bytes)
        tmp.close()

        print("  Enviando captcha para anti-captcha...")
        solver = imagecaptcha()
        solver.set_verbose(0)
        solver.set_key(ANTICAPTCHA_KEY)
        solver.set_numeric(1)

        with ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(solver.solve_and_return_solution, tmp.name)
            try:
                valor_captcha = future.result(timeout=60)
            except FutureTimeout:
                raise Exception("Anti-captcha não respondeu em 60s")

        os.unlink(tmp.name)

        if not valor_captcha:
            raise Exception(solver.err_string)
        print(f"  Captcha resolvido: '{valor_captcha}'")
    except Exception as e:
        print(f"  Anti-captcha erro: {e}")

    if not valor_captcha.isdigit() or len(valor_captcha) < 3:
        if WEB_MODE:
            print("  *** Captcha falhou automaticamente — CNPJ será pulado ***")
            return "captcha_falhou"
        print("  *** Falha automática — olhe o navegador e digite o captcha ***")
        valor_captcha = input("  Valor do captcha: ").strip()

    campo_captcha.click()
    campo_captcha.fill(valor_captcha)

    # Registra listeners de popup E download ANTES de clicar Confirmar
    popups_capturados   = []
    downloads_capturados = []
    def capturar_popup(p):
        popups_capturados.append(p)
    def capturar_download(d):
        downloads_capturados.append(d)
    context.on("page", capturar_popup)
    context.on("download", capturar_download)

    page.get_by_role("button", name="Confirmar").click()
    page.wait_for_timeout(4000)

    context.remove_listener("page", capturar_popup)
    context.remove_listener("download", capturar_download)

    # Verifica débitos
    try:
        if page.get_by_text("regularizar os débitos", exact=False).first.is_visible():
            for p in popups_capturados:
                p.close()
            return "com_debitos"
    except Exception:
        pass

    # Verifica se o PDF apareceu em iframe na mesma página
    try:
        iframe = page.locator('iframe[name^="fancybox-frame"]')
        if iframe.is_visible(timeout=3000):
            frame = page.frame_locator('iframe[name^="fancybox-frame"]')
            with page.expect_download(timeout=30000) as dl:
                frame.get_by_role("button", name="Salvar").click()
            dl.value.save_as(caminho)
            return "ok"
    except Exception:
        pass

    # Em headless o PDF vira download direto no contexto
    if downloads_capturados:
        downloads_capturados[0].save_as(caminho)
        for p in popups_capturados:
            p.close()
        return "ok"

    # Usa popup capturado pelo listener
    if not popups_capturados:
        raise PlaywrightTimeout("Nenhum popup abriu após o Confirmar")

    popup = popups_capturados[0]

    # Aguarda URL válida (visível ou headless com redirect)
    try:
        popup.wait_for_url(lambda url: url.startswith("http"), timeout=10000)
    except Exception:
        pass
    page.wait_for_timeout(1000)

    pdf_url = popup.url

    if not pdf_url or pdf_url in ("about:blank", ":") or not pdf_url.startswith("http"):
        popup.close()
        return "nao_encontrado"

    # Extrai o PDF via JavaScript dentro do popup (mais confiável que requests)
    try:
        pdf_base64 = popup.evaluate("""async () => {
            const r = await fetch(window.location.href);
            const blob = await r.blob();
            return new Promise(res => {
                const reader = new FileReader();
                reader.onload = () => res(reader.result.split(',')[1]);
                reader.readAsDataURL(blob);
            });
        }""")
        pdf_bytes = base64.b64decode(pdf_base64)
        if b"%PDF" in pdf_bytes[:10]:
            popup.close()
            with open(caminho, "wb") as f:
                f.write(pdf_bytes)
            return "ok"
    except Exception:
        pass

    popup.close()

    # Fallback: requests com cookies da sessão
    cookies = {c["name"]: c["value"] for c in context.cookies()}
    resposta = requests.get(pdf_url, cookies=cookies, timeout=30)
    if resposta.status_code == 200 and b"%PDF" in resposta.content[:10]:
        with open(caminho, "wb") as f:
            f.write(resposta.content)
        return "ok"

    return "erro_download"

# -------------------------------------------------------
# Modos de execução:
#   python robo.py                              → todas as cidades
#   python robo.py todas                        → todas as cidades
#   python robo.py apiacas                      → só Apiacas
#   python robo.py outras                       → todas menos Apiacas
#   python robo.py cnpj 12345678000195          → CNPJ específico (busca cidade no Excel)
#   python robo.py cnpj 12345678000195 Sinop/MT → CNPJ específico com cidade manual
# -------------------------------------------------------

HEADLESS = False  # padrão: navegador visível

# Menu interativo quando aberto sem argumentos (duplo clique no exe/bat)
if len(sys.argv) == 1:
    print("=" * 45)
    print("     ROBO DE CERTIDOES NEGATIVAS")
    print("=" * 45)
    print()
    print("  1 - Todos os CNPJs")
    print("  2 - Apenas Apiacas")
    print("  3 - Todas as cidades menos Apiacas")
    print("  4 - CNPJ especifico")
    print()
    opcao = input("Escolha uma opcao (1/2/3/4): ").strip()

    if opcao == "1":
        sys.argv = [sys.argv[0], "todas"]
    elif opcao == "2":
        sys.argv = [sys.argv[0], "apiacas"]
    elif opcao == "3":
        sys.argv = [sys.argv[0], "outras"]
    elif opcao == "4":
        cnpj_input = input("Digite o CNPJ: ").strip()
        sys.argv = [sys.argv[0], "cnpj", cnpj_input]
    else:
        print("Opcao invalida. Encerrando.")
        input("Pressione ENTER para fechar...")
        sys.exit(1)


MODO = sys.argv[1].lower() if len(sys.argv) > 1 else "todas"

df = pd.read_excel(ARQUIVO_EXCEL, dtype=str)
df[COLUNA_CNPJ] = df[COLUNA_CNPJ].str.replace(r"[.\-/]", "", regex=True).str.strip()

# Busca a coluna de nome ignorando maiúsculas e acentos
_nome_norm = normalizar(COLUNA_NOME).lower()
COLUNA_NOME_REAL = next(
    (c for c in df.columns if normalizar(str(c)).lower() == _nome_norm),
    None
)

colunas_ler = [COLUNA_CNPJ, COLUNA_CIDADE] + ([COLUNA_NOME_REAL] if COLUNA_NOME_REAL else [])
todos = df[colunas_ler].fillna("").to_dict("records")
todos = [r for r in todos if r[COLUNA_CNPJ].strip()]

if MODO == "cnpj":
    if len(sys.argv) < 3:
        print("Informe o CNPJ: python robo.py cnpj 12345678000195")
        sys.exit(1)

    cnpj_busca = re.sub(r"[.\-/]", "", sys.argv[2]).strip()

    if len(sys.argv) >= 4:
        cidade_manual = sys.argv[3]
        registros = [{COLUNA_CNPJ: cnpj_busca, COLUNA_CIDADE: cidade_manual}]
    else:
        registros = [r for r in todos if r[COLUNA_CNPJ] == cnpj_busca]
        if not registros:
            print(f"CNPJ {cnpj_busca} não encontrado no Excel.")
            print("Informe a cidade manualmente: python robo.py cnpj 12345678000195 \"Sinop/MT\"")
            sys.exit(1)

    print(f"Modo: CNPJ ÚNICO | {registros[0][COLUNA_CNPJ]} | {registros[0][COLUNA_CIDADE]}\n")

elif MODO == "apiacas":
    registros = [r for r in todos if extrair_cidade_estado(r.get(COLUNA_CIDADE, ""))[0] == "Apiacas"]
    print(f"Modo: APIACAS | {len(registros)} CNPJs selecionados.\n")

elif MODO == "outras":
    registros = [r for r in todos if extrair_cidade_estado(r.get(COLUNA_CIDADE, ""))[0] != "Apiacas"]
    print(f"Modo: OUTRAS | {len(registros)} CNPJs selecionados.\n")

elif MODO == "todas":
    registros = todos
    print(f"Modo: TODAS | {len(registros)} CNPJs selecionados.\n")

else:
    print(f"Modo inválido: '{MODO}'.")
    print("Use: todas | apiacas | outras | cnpj <numero> [cidade/UF]")
    sys.exit(1)

if os.environ.get("ROBO_HEADLESS") == "1" or WEB_MODE:
    HEADLESS = True
else:
    print()
    bg = input("Rodar em segundo plano? O navegador fica invisivel e o computador fica livre (S/N): ").strip().upper()
    HEADLESS = (bg == "S")
print()

erros            = []
nao_encontrados  = []
nao_suportados   = []
com_debitos_list = []
ja_existentes    = []

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
    # Remove a marca navigator.webdriver que sites usam para detectar automação
    context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    page    = context.new_page()

    print("  Dica: pressione Q a qualquer momento para parar após o CNPJ atual.\n")

    parar = False
    for i, reg in enumerate(registros, start=1):
        # Verifica se o usuário pressionou Q para parar (apenas Windows, fora do modo web)
        if _WINDOWS and not WEB_MODE:
            while msvcrt.kbhit():
                tecla = msvcrt.getch()
                if tecla in (b'q', b'Q', b'\x1b'):
                    parar = True

        if parar:
            print("\n  *** Parada solicitada — encerrando após este ponto. ***")
            break

        cnpj   = reg[COLUNA_CNPJ]
        nome   = str(reg.get(COLUNA_NOME_REAL, "")).strip() if COLUNA_NOME_REAL else ""
        cidade, estado = extrair_cidade_estado(reg.get(COLUNA_CIDADE, ""))

        cidade_pasta = re.sub(r'[\\/*?:"<>|]', "", cidade)  # remove caracteres inválidos
        os.makedirs(os.path.join(PASTA_NEGATIVAS, cidade_pasta), exist_ok=True)
        os.makedirs(os.path.join(PASTA_POSITIVAS, cidade_pasta), exist_ok=True)

        caminho      = os.path.join(PASTA_DOWNLOADS, f"{cnpj}.pdf")  # temp
        caminho_neg  = os.path.join(PASTA_NEGATIVAS, cidade_pasta, f"{cnpj}.pdf")
        caminho_pos  = os.path.join(PASTA_POSITIVAS, cidade_pasta, f"{cnpj}.pdf")

        nome_exibir = f" | {nome}" if nome else ""
        print(f"\n[{i}/{len(registros)}] CNPJ: {cnpj}{nome_exibir} | Cidade: {cidade} | Estado: {estado}")

        # Fecha abas extras que ficaram abertas na iteração anterior
        for aba in context.pages:
            if aba != page:
                aba.close()

        # Pula se o PDF já foi baixado (em qualquer das subpastas ou temp)
        if os.path.exists(caminho_neg) or os.path.exists(caminho_pos) or os.path.exists(caminho):
            print(f"  PDF já existe, pulando...")
            ja_existentes.append(cnpj)
            continue

        # Cidade sem portal cadastrado — avisa e pula imediatamente
        if cidade not in CIDADES_CONFIG and estado not in ESTADOS_BETHA:
            print(f"  ERRO: Sem link de portal cadastrado para '{cidade}/{estado}' — pulando.")
            nao_suportados.append(f"{cnpj} ({cidade}/{estado} - sem portal)")
            continue

        try:
            if cidade in CIDADES_CONFIG:
                config  = CIDADES_CONFIG[cidade]
                sistema = config["sistema"]
                print(f"  Sistema: {sistema} ({cidade})")

                if sistema == "gpsrv":
                    resultado = processar_gpsrv(page, context, cnpj, config, caminho)
                elif sistema == "agili":
                    resultado = processar_agili(page, context, p, cnpj, config, caminho)
                elif sistema == "i7sgp":
                    resultado = processar_i7sgp(page, context, p, cnpj, config, caminho)
                else:
                    resultado = "sistema_desconhecido"
            else:
                print(f"  Sistema: Betha ({cidade}/{estado})")
                resultado = processar_betha(page, cnpj, cidade, estado, caminho)

            if resultado == "ok":
                tipo    = detectar_tipo_certidao(caminho)
                destino = caminho_neg if tipo == "negativa" else caminho_pos
                shutil.move(caminho, destino)
                print(f"  Certidão {tipo} salva: {destino} ({cidade_pasta})")
            elif resultado == "nao_encontrado":
                print(f"  CNPJ não encontrado na base, pulando...")
                nao_encontrados.append(cnpj)
            elif resultado == "com_debitos":
                print(f"  CNPJ com débitos — certidão não emitida.")
                com_debitos_list.append(cnpj)
            elif resultado in ("prefeitura_nao_encontrada", "estado_nao_suportado", "sistema_desconhecido"):
                print(f"  Portal não suportado ({resultado}), pulando...")
                nao_suportados.append(f"{cnpj} ({resultado})")
            else:
                print(f"  Falha: {resultado}")
                erros.append(f"{cnpj} ({resultado})")

        except PlaywrightTimeout as e:
            print(f"  TIMEOUT: {e}")
            erros.append(f"{cnpj} (timeout)")
        except Exception as e:
            print(f"  ERRO: {e}")
            erros.append(f"{cnpj} (erro)")

    context.close()
    browser.close()

total   = len(registros)
sucesso = total - len(erros) - len(nao_encontrados) - len(nao_suportados) - len(com_debitos_list) - len(ja_existentes)

linhas = [
    "",
    "--- Processamento finalizado ---",
    f"Sucesso:             {sucesso}/{total}",
    f"Já existia (pulou):  {len(ja_existentes)}/{total}",
    f"Não encontrado:      {len(nao_encontrados)}/{total}",
]
if nao_encontrados:
    linhas.append(f"  CNPJs: {', '.join(nao_encontrados)}")

linhas.append(f"Com débitos:         {len(com_debitos_list)}/{total}")
if com_debitos_list:
    linhas.append(f"  CNPJs: {', '.join(com_debitos_list)}")

linhas.append(f"Não suportado:       {len(nao_suportados)}/{total}")
if nao_suportados:
    linhas.append(f"  CNPJs: {', '.join(nao_suportados)}")

if erros:
    linhas.append(f"Erros:               {len(erros)}/{total}")
    linhas.append(f"  CNPJs: {', '.join(erros)}")

relatorio = "\n".join(linhas)
print(relatorio)

# Salva relatório em arquivo junto às certidões
nome_relatorio  = f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
caminho_relatorio = os.path.join(PASTA_DOWNLOADS, nome_relatorio)
with open(caminho_relatorio, "w", encoding="utf-8") as f:
    f.write(relatorio)
print(f"\nRelatório salvo em: {caminho_relatorio}")

# Aviso sonoro e popup apenas no Windows fora do modo web
if _WINDOWS and not WEB_MODE:
    winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
    ctypes.windll.user32.MessageBoxW(
        0,
        f"Processamento finalizado!\n\nSucesso: {sucesso}/{total}\nCom débitos: {len(com_debitos_list)}\nErros: {len(erros)}\n\nRelatório salvo em:\n{caminho_relatorio}",
        "Robô de Certidões — Concluído",
        0x40  # ícone de informação
)
