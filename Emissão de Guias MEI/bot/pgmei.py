"""
Bot de automação do PGMEI — Emissão de Guias DAS para MEI.

Fluxo mapeado pelo Network tab:
  /Identificacao  →  CNPJ + hCaptcha (resolvido via Anti-captcha)  →  302 → /Home/Inicio
  /Home/Inicio    →  clica "Emitir Guia (DAS)"  →  /emissao
  /emissao        →  XHR VerificaRetificacaoAutomatica  →  /gerarDas
  /gerarDas       →  clica "Imprimir"  →  /imprimir
  /imprimir       →  salva como PDF via CDP Page.printToPDF
"""

import asyncio
import base64
import logging
import re
import sys
from datetime import datetime
from pathlib import Path

from playwright.async_api import BrowserContext, Page, async_playwright
from playwright_stealth import Stealth

_stealth = Stealth(
    navigator_languages_override=("pt-BR", "pt"),
    navigator_platform_override="Win32",
)

from config import (
    ANTICAPTCHA_API_KEY,
    BROWSER_PROFILE_DIR,
    DELAY_CNPJS,
    DOWNLOADS_DIR,
    HEADLESS,
    MAX_RETRIES,
    MESES_PT,
    PGMEI_URL,
    PGMEI_URL_INICIO,
    TIMEOUT_MS,
)

logger = logging.getLogger(__name__)

# playwright-stealth aplica patches abrangentes (webdriver, plugins, chrome runtime,
# WebGL, canvas, outerWidth, permissions, mimeTypes e mais) — mais completo que JS manual.


# ---------------------------------------------------------------------------
# Modelo de resultado
# ---------------------------------------------------------------------------

class Resultado:
    def __init__(self, cnpj: str, nome: str, competencia: str):
        self.cnpj        = cnpj
        self.nome        = nome
        self.competencia = competencia
        self.status      = "PENDENTE"
        self.tipo_guia   = ""
        self.arquivo     = ""
        self.debitos     = []   # lista de períodos com vencimento em atraso
        self.observacao  = ""
        self.timestamp   = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    def to_dict(self) -> dict:
        return {
            "CNPJ":               self.cnpj,
            "Nome":               self.nome,
            "Competência":        self.competencia,
            "Status":             self.status,
            "Tipo Guia":          self.tipo_guia,
            "Arquivo":            self.arquivo,
            "Débitos em Atraso":  ", ".join(self.debitos) if self.debitos else "",
            "Observação":         self.observacao,
            "Data/Hora":          self.timestamp,
        }


# ---------------------------------------------------------------------------
# Bot principal
# ---------------------------------------------------------------------------

class PGMEIBot:

    def __init__(self):
        self._context: BrowserContext | None = None
        self._display = None  # pyvirtualdisplay no Linux

    # ------------------------------------------------------------------
    # Ciclo de vida
    # ------------------------------------------------------------------

    async def _iniciar(self, playwright):
        """
        Usa perfil persistente (browser_profile/) para acumular confiança
        no hCaptcha entre sessões. Cada execução deixa cookies/histórico que
        fazem o hCaptcha pontuar o navegador como humano nas próximas vezes.
        """
        # No Linux sem display (Railway), inicia Xvfb para rodar Chrome headful
        if sys.platform != "win32" and not HEADLESS:
            import os as _os
            logger.info(f"Linux detectado | DISPLAY={_os.environ.get('DISPLAY','n/d')}")
            try:
                from pyvirtualdisplay import Display
                if self._display is None:
                    self._display = Display(visible=0, size=(1280, 900))
                    self._display.start()
                    logger.info(f"Xvfb OK | DISPLAY={_os.environ.get('DISPLAY','?')}")
            except Exception as e:
                logger.warning(f"pyvirtualdisplay falhou: {e} — Chrome pode não abrir")

        BROWSER_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        logger.info(f"Perfil do browser: {BROWSER_PROFILE_DIR}")

        # --start-minimized só tem efeito no Windows
        # --no-sandbox e --disable-dev-shm-usage são necessários em containers Linux (Railway)
        extra_args = ["--start-minimized"] if sys.platform == "win32" else [
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu",
        ]

        launch_kwargs = dict(
            headless=HEADLESS,
            args=[
                "--disable-blink-features=AutomationControlled",
                *extra_args,
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-infobars",
            ],
            accept_downloads=True,
            viewport={"width": 1280, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="pt-BR",
            timezone_id="America/Cuiaba",
            extra_http_headers={"Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8"},
        )

        try:
            self._context = await playwright.chromium.launch_persistent_context(
                str(BROWSER_PROFILE_DIR),
                channel="chrome",
                **launch_kwargs,
            )
            logger.info("Usando Chrome real com perfil persistente.")
        except Exception as e:
            logger.warning(f"Chrome real indisponível ({e}); usando Chromium.")
            self._context = await playwright.chromium.launch_persistent_context(
                str(BROWSER_PROFILE_DIR),
                **launch_kwargs,
            )

        self._context.set_default_timeout(TIMEOUT_MS)

    async def _finalizar(self):
        if self._context:
            try:
                await asyncio.wait_for(self._context.close(), timeout=10)
            except (Exception, asyncio.CancelledError):
                pass
            self._context = None
        if self._display is not None:
            try:
                self._display.stop()
            except Exception:
                pass
            self._display = None

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _pasta_download(self, mes: int, ano: int) -> Path:
        pasta = DOWNLOADS_DIR / str(ano) / f"{mes:02d}-{MESES_PT[mes]}"
        pasta.mkdir(parents=True, exist_ok=True)
        return pasta

    def _nome_arquivo(self, cnpj: str, nome: str, mes: int, ano: int, tipo: str) -> str:
        nome_limpo = re.sub(r"[^\w\s-]", "", nome).replace(" ", "_")[:30]
        return f"{cnpj}_{nome_limpo}_{tipo}_{mes:02d}_{ano}.pdf"

    @staticmethod
    def _formatar_cnpj(cnpj: str) -> str:
        c = cnpj.zfill(14)
        return f"{c[0:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:14]}"

    async def _screenshot(self, page: Page, label: str):
        try:
            path = DOWNLOADS_DIR.parent / "logs" / f"debug_{label}_{datetime.now().strftime('%H%M%S')}.png"
            path.parent.mkdir(exist_ok=True)
            await page.screenshot(path=str(path), full_page=True)
            logger.debug(f"Screenshot: {path.name}")
        except Exception:
            pass

    async def _clicar_primeiro_visivel(self, page: Page, seletores: list[str]) -> bool:
        for sel in seletores:
            try:
                el = page.locator(sel).first
                if await el.is_visible(timeout=4000):
                    await el.click()
                    return True
            except Exception:
                continue
        return False

    async def _salvar_pdf_cdp(self, page: Page, destino: Path) -> bool:
        """
        Salva a página atual como PDF usando Chrome DevTools Protocol.
        Funciona com Chrome real em modo headless=False, ao contrário de page.pdf().
        """
        try:
            cdp = await self._context.new_cdp_session(page)
            result = await cdp.send("Page.printToPDF", {
                "printBackground": True,
                "paperWidth":  8.27,   # A4
                "paperHeight": 11.69,
                "marginTop":    0.4,
                "marginBottom": 0.4,
                "marginLeft":   0.4,
                "marginRight":  0.4,
                "scale": 0.9,
            })
            destino.write_bytes(base64.b64decode(result["data"]))
            await cdp.detach()
            logger.info(f"  → PDF salvo via CDP: {destino.name}")
            return True
        except Exception as e:
            logger.warning(f"  → CDP PDF falhou: {e}")
            return False

    async def _injetar_token_hcaptcha(self, page: Page) -> bool:
        """
        Detecta hCaptcha na página, obtém token via Anti-Captcha API e injeta no formulário.
        Necessário em modo headless (Railway) onde não há interação manual possível.
        Retorna True se o token foi injetado com sucesso.
        """
        try:
            sitekey = await page.evaluate("""() => {
                const el = document.querySelector('.h-captcha[data-sitekey], [data-hcaptcha-sitekey]');
                if (el) return el.getAttribute('data-sitekey') || el.getAttribute('data-hcaptcha-sitekey');
                const iframe = document.querySelector('iframe[src*="hcaptcha.com"]');
                if (iframe) {
                    const m = iframe.src.match(/sitekey=([^&]+)/);
                    return m ? decodeURIComponent(m[1]) : null;
                }
                return null;
            }""")

            if not sitekey:
                logger.debug("  → hCaptcha não detectado na página")
                return False

            logger.info(f"  → hCaptcha detectado (sitekey: {sitekey[:20]}...). Enviando ao Anti-Captcha...")

            from anticaptchaofficial.hcaptchaproxyless import hCaptchaProxyless
            solver = hCaptchaProxyless()
            solver.set_verbose(0)
            solver.set_key(ANTICAPTCHA_API_KEY)
            solver.set_website_url(page.url)
            solver.set_website_key(sitekey)

            # solve_and_return_solution é bloqueante — roda em thread para não travar o event loop
            loop = asyncio.get_event_loop()
            token = await loop.run_in_executor(None, solver.solve_and_return_solution)

            if not token or token == 0:
                logger.warning(f"  → Anti-Captcha falhou: {solver.error_code}")
                return False

            logger.info("  → Token hCaptcha recebido. Injetando...")
            await page.evaluate("""(t) => {
                const ta = document.querySelector('textarea[name="h-captcha-response"]');
                if (ta) {
                    Object.getOwnPropertyDescriptor(
                        window.HTMLTextAreaElement.prototype, 'value'
                    ).set.call(ta, t);
                    ta.dispatchEvent(new Event('input', { bubbles: true }));
                    ta.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }""", token)
            await asyncio.sleep(0.5)
            return True

        except Exception as e:
            logger.warning(f"  → Erro ao resolver hCaptcha: {e}")
            return False

    # ------------------------------------------------------------------
    # Passos do fluxo
    # ------------------------------------------------------------------

    async def _passo_inserir_cnpj(self, page: Page, cnpj: str):
        logger.info("  → Acessando /Identificacao...")
        await page.goto(PGMEI_URL, wait_until="domcontentloaded")

        # Aguarda o Angular terminar de renderizar (bootstrapping + route guards).
        # "input" como proxy: quando qualquer input é visível, o Angular já decidiu a rota.
        # Sessão ativa pode fazer o Angular redirecionar para /Home/Inicio após renderizar.
        try:
            await page.wait_for_selector("input, h1, h2", state="visible", timeout=15_000)
        except Exception:
            await asyncio.sleep(3)

        # Verifica DEPOIS do Angular estabilizar (não antes — evita race condition com route guard)
        if "Identificacao" not in page.url:
            logger.info(f"  → Sessão anterior detectada ({page.url}). Limpando cookies...")
            try:
                await self._context.clear_cookies(domain="www8.receita.fazenda.gov.br")
            except TypeError:
                await self._context.clear_cookies()
            await page.goto(PGMEI_URL, wait_until="domcontentloaded")
            try:
                await page.wait_for_selector("input", state="visible", timeout=15_000)
            except Exception:
                await asyncio.sleep(3)

        logger.info(f"  → URL: {page.url}")

        # Angular já renderizou — busca rápida com is_visible() (sem aguardar aparecimento)
        campo = None
        for sel in [
            'input[id*="cnpj" i]',
            'input[name*="cnpj" i]',
            '[formcontrolname*="cnpj" i]',
            'input[placeholder*="CNPJ"]',
            'input[type="tel"]',
            'input[maxlength="18"]',
            'input[maxlength="14"]',
            'input[type="text"]',
            'input',
        ]:
            try:
                el = page.locator(sel).first
                if await el.is_visible():
                    campo = el
                    break
            except Exception:
                continue

        # Fallback: se is_visible() falhou em tudo, aguarda mais um pouco e tenta de novo
        if not campo:
            await asyncio.sleep(4)
            for sel in ['input[type="text"]', 'input[type="tel"]', 'input']:
                try:
                    el = page.locator(sel).first
                    await el.wait_for(state="visible", timeout=10_000)
                    campo = el
                    break
                except Exception:
                    continue

        if not campo:
            await self._screenshot(page, f"sem_campo_{cnpj}")
            raise RuntimeError(f"Campo CNPJ não encontrado (URL: {page.url})")

        # Clica, vai para o início do campo e limpa tudo antes de digitar
        await campo.click()
        await asyncio.sleep(0.3)
        await page.keyboard.press("Home")        # posiciona cursor no início da máscara
        await page.keyboard.press("Control+a")   # seleciona tudo
        await page.keyboard.press("Delete")      # apaga conteúdo/máscara
        await asyncio.sleep(0.2)

        # Digita dígito a dígito — a máscara aplica os separadores automaticamente
        await campo.press_sequentially(cnpj, delay=80)
        await asyncio.sleep(0.5)

        # Se o campo ainda estiver vazio, tenta digitar o CNPJ já formatado
        valor = await campo.input_value()
        if not any(d.isdigit() for d in valor):
            await campo.click()
            await page.keyboard.press("Control+a")
            await page.keyboard.press("Delete")
            await asyncio.sleep(0.2)
            await campo.press_sequentially(self._formatar_cnpj(cnpj), delay=80)
            await asyncio.sleep(0.4)

        logger.info(f"  → Campo preenchido: {await campo.input_value()}")

        # Em headless (Railway): injeta token hCaptcha ANTES de submeter
        if HEADLESS and ANTICAPTCHA_API_KEY:
            await self._injetar_token_hcaptcha(page)

        # Clica Continuar para disparar o hCaptcha
        clicou = await self._clicar_primeiro_visivel(page, [
            'button:has-text("Continuar")',
            'button[type="submit"]',
            'input[type="submit"]',
        ])
        if not clicou:
            await campo.press("Enter")

        # Aguarda para ver se o captcha bloqueou ou redirecionou
        await asyncio.sleep(2.5)

        try:
            conteudo = await page.content()
        except Exception:
            conteudo = ""

        bloqueado = "Comportamento de Rob" in conteudo or "Impedido" in conteudo

        if bloqueado or "Inicio" not in page.url:
            cnpj_fmt = self._formatar_cnpj(cnpj)

            if HEADLESS:
                # Headless (Railway): tenta recuperar automaticamente
                logger.warning(f"  → Captcha/bloqueio detectado (headless). Tentando recuperar...")
                if bloqueado:
                    # Fecha alerta e tenta de novo com token
                    try:
                        await page.click('button.close, [aria-label="Close"], .btn-danger, button:has-text("×")', timeout=3000)
                        await asyncio.sleep(0.5)
                    except Exception:
                        pass
                    await page.goto(PGMEI_URL, wait_until="domcontentloaded")
                    try:
                        await page.wait_for_selector("input", state="visible", timeout=10_000)
                    except Exception:
                        await asyncio.sleep(3)
                    # Preenche CNPJ novamente
                    for sel in ['input[type="tel"]', 'input[maxlength="18"]', 'input[maxlength="14"]', 'input']:
                        try:
                            el = page.locator(sel).first
                            if await el.is_visible():
                                await el.click()
                                await page.keyboard.press("Control+a")
                                await page.keyboard.press("Delete")
                                await el.press_sequentially(cnpj, delay=80)
                                break
                        except Exception:
                            continue
                    if ANTICAPTCHA_API_KEY:
                        await self._injetar_token_hcaptcha(page)
                    await self._clicar_primeiro_visivel(page, [
                        'button:has-text("Continuar")', 'button[type="submit"]',
                    ])
                    await asyncio.sleep(2.5)

                # Aguarda redirect com timeout razoável (60s)
                try:
                    await page.wait_for_url("**/Home/Inicio**", timeout=60_000)
                except Exception:
                    raise RuntimeError(
                        f"hCaptcha não resolvido automaticamente para CNPJ {cnpj_fmt}. "
                        "Verifique o ANTICAPTCHA_API_KEY no config.py ou rode localmente."
                    )
            else:
                # Headful (local): pede resolução manual
                logger.warning("")
                logger.warning("=" * 60)
                logger.warning("  CAPTCHA MANUAL NECESSÁRIO")
                logger.warning(f"  CNPJ: {cnpj_fmt}")
                logger.warning("  No navegador aberto:")
                if bloqueado:
                    logger.warning("    1. Feche o alerta vermelho (X)")
                    logger.warning(f"   2. Digite o CNPJ: {cnpj_fmt}")
                    logger.warning("    3. Clique em Continuar")
                else:
                    logger.warning("    → Clique em Continuar e resolva o captcha")
                logger.warning("  O bot assume automaticamente após você passar.")
                logger.warning("  Aguardando (5 minutos)...")
                logger.warning("=" * 60)
                logger.warning("")
                await page.wait_for_url("**/Home/Inicio**", timeout=300_000)

        logger.info("  → Em /Home/Inicio.")

    async def _passo_navegar_para_das(self, page: Page, mes: int, ano: int) -> list[str]:
        """
        Fluxo exato após /Home/Inicio:
          1. Clica "Emitir Guia de Pagamento (DAS)"
          2. Seleciona Ano-Calendário → clica OK
          3. Acha a linha com vencimento no mês atual → clica "Apurar/Gerar DAS"
        Retorna lista de débitos em atraso detectados na tabela.
        """

        # ── Passo 1: Emitir Guia de Pagamento (DAS) ──────────────────────
        logger.info("  → Clicando em 'Emitir Guia de Pagamento (DAS)'...")
        clicou = await self._clicar_primeiro_visivel(page, [
            'a:has-text("Emitir Guia de Pagamento (DAS)")',
            ':text("Emitir Guia de Pagamento (DAS)")',
            'a:has-text("Emitir Guia de Pagamento")',
            ':text("Emitir Guia de Pagamento")',
            'a:has-text("Emitir Guia")',
            'li:has-text("Emitir Guia")',
        ])
        if not clicou:
            await self._screenshot(page, "sem_emitir_guia")
            raise RuntimeError("Link 'Emitir Guia de Pagamento (DAS)' não encontrado.")
        await page.wait_for_load_state("networkidle")

        # ── Passo 2: Seleciona Ano-Calendário e clica OK ──────────────────
        logger.info(f"  → Selecionando Ano-Calendário {ano}...")
        await self._selecionar_ano_calendario(page, ano)

        # ── Passo 3: Acha a linha pelo vencimento do mês atual ────────────
        logger.info(f"  → Buscando linha com vencimento {mes:02d}/{ano}...")
        return await self._apurar_das_linha(page, mes, ano)

    async def _selecionar_ano_calendario(self, page: Page, ano: int):
        """Seleciona o Ano-Calendário e clica OK."""
        # Tenta dropdown (select) com opção do ano
        try:
            sel = page.locator("select").first
            await sel.wait_for(state="visible", timeout=8000)
            opcoes = await sel.locator("option").all_text_contents()
            for op in opcoes:
                if str(ano) in op:
                    await sel.select_option(label=op)
                    logger.info(f"  → Ano selecionado: {op.strip()}")
                    break
        except Exception:
            pass

        # Clica OK após selecionar o ano
        await self._clicar_primeiro_visivel(page, [
            'button:has-text("OK")',
            'input[value="OK"]',
            'input[value="Ok"]',
            'button:has-text("Confirmar")',
            'button:has-text("Pesquisar")',
        ])
        await page.wait_for_load_state("networkidle")

    async def _verificar_alerta_pgmei(self, page: Page):
        """
        Detecta alertas de erro exibidos pelo PGMEI (banners vermelhos).
        Se encontrar, lança RuntimeError com a mensagem do alerta.
        """
        await asyncio.sleep(0.8)
        try:
            alerta = page.locator(
                ".alert, .alert-danger, .alert-warning, "
                "[class*='alert'], [class*='erro'], [class*='error']"
            ).first
            if await alerta.is_visible(timeout=2000):
                texto = (await alerta.text_content() or "").strip()
                if any(k in texto for k in ["Falha", "limite", "excedido", "Erro", "erro", "23998"]):
                    logger.warning(f"  → Alerta PGMEI: {texto[:120]}")
                    # Tenta fechar o alerta (botão X)
                    try:
                        await alerta.locator("button, .close, [aria-label='Close']").first.click()
                    except Exception:
                        pass
                    raise RuntimeError(f"PGMEI: {texto[:120]}")
        except RuntimeError:
            raise
        except Exception:
            pass

    async def _apurar_das_linha(self, page: Page, mes: int, ano: int) -> list[str]:
        """
        Varre a tabela de competências do PGMEI:
        - Detecta linhas com vencimento em meses anteriores (débitos em atraso)
        - Marca o checkbox da linha cujo vencimento cai no mês/ano atual
        - Retorna lista de períodos com débito (pode ser vazia)

        Usa regex para encontrar datas (DD/MM/YYYY) em qualquer célula da linha,
        evitando dependência de posição fixa de coluna.
        """
        _DATA_RE       = re.compile(r"(\d{2})/(\d{2})/(\d{4})")
        vence_no_mes   = f"/{mes:02d}/{ano}"   # ex: "/06/2026" → bate com "22/06/2026"
        debitos_atraso: list[str] = []
        checkbox_marcado = False

        await page.wait_for_load_state("networkidle")
        rows = await page.locator("table tr").all()

        for row in rows:
            try:
                cells = await row.locator("td").all()
                if len(cells) < 2:
                    continue

                textos  = [(await c.text_content() or "").strip() for c in cells]
                # textos[0] pode ser célula de checkbox vazia; pega o primeiro texto não-vazio
                periodo = next((t for t in textos if t.strip() and not _DATA_RE.search(t)), "")

                # Encontra a primeira data DD/MM/YYYY em qualquer célula da linha (pula período)
                texto_venc = ""
                for txt in textos:
                    if _DATA_RE.search(txt):
                        texto_venc = txt
                        break

                # Fallback: penúltima célula (comportamento anterior)
                if not texto_venc:
                    texto_venc = textos[-2]

                # Detecta débito: data encontrada é anterior ao mês-alvo
                m = _DATA_RE.search(texto_venc)
                if m:
                    v_mes_n = int(m.group(2))
                    v_ano_n = int(m.group(3))
                    if (v_ano_n, v_mes_n) < (ano, mes):
                        debitos_atraso.append(f"{periodo} (venc. {texto_venc})")

                # Marca checkbox do mês atual (busca apenas na coluna de vencimento)
                if vence_no_mes in texto_venc and not checkbox_marcado:
                    cb = row.locator('input[type="checkbox"]').first
                    if await cb.is_visible(timeout=2000):
                        if not await cb.is_checked():
                            await cb.check()
                        logger.info(
                            f"  → Checkbox marcada | Período: {periodo} | "
                            f"Vencimento: {texto_venc}"
                        )
                        checkbox_marcado = True
            except Exception:
                continue

        if debitos_atraso:
            logger.warning(f"  → DÉBITOS EM ATRASO detectados ({len(debitos_atraso)}): "
                           f"{', '.join(debitos_atraso)}")

        # Salva na instância para caso RuntimeError seja capturado acima na pilha
        self._debitos_detectados = debitos_atraso

        if not checkbox_marcado:
            await self._screenshot(page, f"sem_vencimento_{mes:02d}_{ano}")
            raise RuntimeError(
                f"Nenhuma linha com Data de Vencimento em {mes:02d}/{ano} encontrada. "
                "Pode não haver DAS disponível para emissão neste mês."
            )

        # ── 3. Clica "Apurar/Gerar DAS" ─────────────────────────────────────
        clicou = await self._clicar_primeiro_visivel(page, [
            'button:has-text("Apurar/Gerar DAS")',
            'input[value="Apurar/Gerar DAS"]',
            ':text("Apurar/Gerar DAS")',
            'input[value*="Apurar" i]',
        ])

        if not clicou:
            await self._screenshot(page, "sem_botao_apurar_das")
            raise RuntimeError("Botão 'Apurar/Gerar DAS' não encontrado.")

        await page.wait_for_load_state("networkidle")

        # Verifica alertas de erro do PGMEI antes de continuar
        await self._verificar_alerta_pgmei(page)

        # Angular SPA: URL não muda, aguarda o conteúdo da Tela 2 renderizar
        try:
            await page.wait_for_selector(
                ':text("DAS gerados"), button:has-text("Imprimir/Visualizar PDF"), '
                'a:has-text("Imprimir/Visualizar PDF")',
                timeout=15_000,
            )
        except Exception:
            await asyncio.sleep(3)

        logger.info(f"  → Tela 2 pronta: {page.url}")
        return debitos_atraso

    async def _passo_imprimir_e_salvar(
        self, page: Page, cnpj: str, nome: str, mes: int, ano: int, tipo: str
    ) -> Path | None:
        """
        Tela 2 (DAS gerados): clica 'Imprimir/Visualizar PDF'.
        O PGMEI pode:
          A) disparar um download direto de PDF
          B) abrir nova aba com o PDF (que pode ou não disparar download)
        Capturamos o download pelo evento do contexto inteiro.
        """
        pasta   = self._pasta_download(mes, ano)
        destino = pasta / self._nome_arquivo(cnpj, nome, mes, ano, tipo)

        seletores = [
            'button:has-text("Imprimir/Visualizar PDF")',
            'a:has-text("Imprimir/Visualizar PDF")',
            ':text("Imprimir/Visualizar PDF")',
            'button:has-text("Imprimir/Visualizar")',
            'a:has-text("Imprimir/Visualizar")',
            'input[value*="Imprimir" i]',
        ]

        botao = None
        for sel in seletores:
            try:
                el = page.locator(sel).first
                if await el.is_visible(timeout=12_000):
                    botao = el
                    break
            except Exception:
                continue

        if not botao:
            await self._screenshot(page, f"sem_imprimir_{cnpj}")
            raise RuntimeError("Botão 'Imprimir/Visualizar PDF' não encontrado na Tela 2.")

        logger.info("  → Clicando 'Imprimir/Visualizar PDF'...")

        # Guarda quais páginas existem antes do clique
        paginas_antes = set(id(p) for p in self._context.pages)

        # ── Cenário A: download direto ───────────────────────────────────────
        try:
            async with page.expect_download(timeout=12_000) as dl_info:
                await botao.click()
            dl = await dl_info.value
            await dl.save_as(str(destino))
            logger.info(f"  → PDF salvo (download direto): {destino.name}")
            return destino
        except Exception as e:
            logger.debug(f"  → Download direto não capturado ({e}). Checando nova aba...")

        # ── Cenário B: abriu nova aba ────────────────────────────────────────
        await asyncio.sleep(3)
        novas = [p for p in self._context.pages if id(p) not in paginas_antes]

        if novas:
            nova = novas[-1]
            try:
                await nova.wait_for_load_state("networkidle", timeout=15_000)
                logger.info(f"  → Nova aba: {nova.url}")

                # Tenta capturar download na nova aba (ex: PDF inline que o browser baixa)
                try:
                    async with nova.expect_download(timeout=8_000) as dl_info:
                        pass  # download já está em curso
                    dl = await dl_info.value
                    await dl.save_as(str(destino))
                    logger.info(f"  → PDF salvo (download nova aba): {destino.name}")
                    await nova.close()
                    return destino
                except Exception:
                    pass

                # Salva via CDP (imprime a aba do PDF como PDF)
                if await self._salvar_pdf_cdp(nova, destino):
                    await nova.close()
                    return destino

                await nova.close()
            except Exception as e:
                logger.warning(f"  → Erro na nova aba: {e}")

        # ── Cenário C: CDP na página atual (fallback) ────────────────────────
        if await self._salvar_pdf_cdp(page, destino):
            return destino

        logger.error("  → Nenhum método de salvamento funcionou.")
        return None

    # ------------------------------------------------------------------
    # Fluxo alternativo (parcelamento, atraso, 2ª via)
    # ------------------------------------------------------------------

    async def _tentar_alternativas(
        self, page: Page, cnpj: str, nome: str, mes: int, ano: int, resultado: Resultado
    ) -> bool:
        alternativas = [
            ("Parcelamento",  ['a:has-text("Parcelamento")',   ':text("Parcelamento")']),
            ("DAS em Atraso", ['a:has-text("Atraso")',         ':text("Atraso")']),
            ("2a Via",        [':text("2ª via")', ':text("2a via")']),
        ]

        for tipo, seletores in alternativas:
            logger.info(f"  → Tentando alternativa: {tipo}")
            try:
                await page.goto(PGMEI_URL_INICIO, wait_until="networkidle")
                clicou = await self._clicar_primeiro_visivel(page, seletores)
                if not clicou:
                    continue

                await page.wait_for_load_state("networkidle")
                try:
                    await self._selecionar_ano_calendario(page, ano)
                except Exception:
                    pass

                arquivo = await self._passo_imprimir_e_salvar(page, cnpj, nome, mes, ano, tipo)
                if arquivo:
                    resultado.status    = "SUCESSO"
                    resultado.arquivo   = str(arquivo)
                    resultado.tipo_guia = tipo
                    resultado.observacao = f"Guia alternativa: {tipo}"
                    return True

            except Exception as e:
                logger.debug(f"  → Alternativa '{tipo}' falhou: {e}")
                continue

        resultado.status    = "FALHOU"
        resultado.observacao = "Nenhuma guia disponível para este CNPJ/competência."
        return False

    # ------------------------------------------------------------------
    # Processamento de um único CNPJ
    # ------------------------------------------------------------------

    async def _processar_um(self, page: "Page", cnpj: str, nome: str, mes: int, ano: int) -> dict:
        resultado = Resultado(cnpj, nome, f"{mes:02d}/{ano}")
        self._debitos_detectados: list[str] = []

        for tentativa in range(1, MAX_RETRIES + 1):
            # Tentativas 2+: reseta a aba atual (evita abrir nova aba e Chrome aparecer na tela)
            if tentativa > 1:
                try:
                    await asyncio.wait_for(page.goto("about:blank", wait_until="load"), timeout=5000)
                except Exception:
                    pass

            logger.info(f"  [Tentativa {tentativa}/{MAX_RETRIES}]")
            try:
                await self._passo_inserir_cnpj(page, cnpj)
                debitos = await self._passo_navegar_para_das(page, mes, ano)
                arquivo = await self._passo_imprimir_e_salvar(page, cnpj, nome, mes, ano, "DAS")

                if arquivo:
                    resultado.status    = "SUCESSO"
                    resultado.arquivo   = str(arquivo)
                    resultado.tipo_guia = "DAS"
                    resultado.debitos   = debitos
                    break

                # DAS normal não disponível → tenta alternativas
                logger.warning("  → DAS normal indisponível; tentando alternativas...")
                if await self._tentar_alternativas(page, cnpj, nome, mes, ano, resultado):
                    resultado.debitos = debitos
                    break

            except RuntimeError as e:
                resultado.status     = "ERRO"
                resultado.observacao = str(e)
                resultado.debitos    = getattr(self, "_debitos_detectados", [])
                logger.error(f"  → {e}")
                break  # erros de CNPJ/captcha não adianta repetir

            except Exception as e:
                resultado.observacao = str(e)
                logger.warning(f"  → Exceção tentativa {tentativa}: {e}")
                if tentativa == MAX_RETRIES:
                    resultado.status  = "ERRO"
                    resultado.debitos = getattr(self, "_debitos_detectados", [])
                await asyncio.sleep(3)

        return resultado.to_dict()

    # ------------------------------------------------------------------
    # Ponto de entrada público
    # ------------------------------------------------------------------

    async def processar_lote(self, clientes: list[dict], mes: int, ano: int) -> list[dict]:
        self._resultados: list[dict] = []
        resultados = self._resultados
        total      = len(clientes)

        async with async_playwright() as pw:
            try:
                await self._iniciar(pw)
            except Exception as e:
                logger.error(f"FALHA AO INICIAR BROWSER: {e}")
                raise
            # Uma única aba reutilizada para todos os CNPJs — Chrome fica minimizado
            page = await self._context.new_page()
            await _stealth.apply_stealth_async(page)
            try:
                for i, cliente in enumerate(clientes, start=1):
                    cnpj = cliente["cnpj"]
                    nome = cliente["nome"]

                    logger.info(f"\n{'─'*55}")
                    logger.info(f"[{i}/{total}] {nome}  |  CNPJ: {cnpj}")

                    try:
                        resultado = await self._processar_um(page, cnpj, nome, mes, ano)
                    except Exception as exc:
                        # Navegador fechado — reabre e cria nova aba
                        if "closed" in str(exc).lower() or "TargetClosed" in type(exc).__name__:
                            logger.warning("  → Navegador fechado. Reabrindo e retentando...")
                            try:
                                await self._finalizar()
                            except Exception:
                                pass
                            await self._iniciar(pw)
                            page = await self._context.new_page()
                            await _stealth.apply_stealth_async(page)
                            try:
                                resultado = await self._processar_um(page, cnpj, nome, mes, ano)
                            except Exception as exc2:
                                r = Resultado(cnpj, nome, f"{mes:02d}/{ano}")
                                r.status    = "ERRO"
                                r.observacao = f"Browser fechado: {exc2}"
                                resultado   = r.to_dict()
                        else:
                            r = Resultado(cnpj, nome, f"{mes:02d}/{ano}")
                            r.status    = "ERRO"
                            r.observacao = str(exc)
                            resultado   = r.to_dict()

                    resultados.append(resultado)

                    status = resultado["Status"]
                    icone = "OK" if status == "SUCESSO" else "FALHOU"
                    detalhe = resultado["Observação"] or resultado["Arquivo"] or ""
                    logger.info(f"  [{icone}] {status} | {detalhe[:100]}")

                    if i < total:
                        await asyncio.sleep(DELAY_CNPJS)

            finally:
                try:
                    await asyncio.wait_for(page.close(), timeout=5)
                except Exception:
                    pass
                await self._finalizar()

        return resultados
