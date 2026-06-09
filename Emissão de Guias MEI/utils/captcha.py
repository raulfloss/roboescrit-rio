"""
Cliente Anti-captcha para hCaptcha.
Documentação: https://anti-captcha.com/apidoc/task-types/HCaptchaTaskProxyless

Passa userAgent + cookies do nosso browser para que o token gerado
case com a sessão ativa, evitando rejeição server-side.
"""

import asyncio
import logging
import requests

logger = logging.getLogger(__name__)

_BASE = "https://api.anti-captcha.com"


def _post(endpoint: str, payload: dict) -> dict:
    resp = requests.post(f"{_BASE}{endpoint}", json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json()


async def resolver_hcaptcha(
    api_key: str,
    site_key: str,
    page_url: str,
    user_agent: str | None = None,
    cookies: list | None = None,
) -> str:
    """
    Envia o hCaptcha para o Anti-captcha e retorna o token.
    Passa userAgent e cookies do browser para maximizar a chance de aceitação.
    """
    loop = asyncio.get_event_loop()

    task: dict = {
        "type":        "HCaptchaTaskProxyless",
        "websiteURL":  page_url,
        "websiteKey":  site_key,
        "isInvisible": False,
    }

    if user_agent:
        task["userAgent"] = user_agent

    if cookies:
        # Anti-captcha espera string "name=value; name2=value2"
        task["cookies"] = "; ".join(
            f"{c['name']}={c['value']}" for c in cookies if c.get("name")
        )

    logger.info("  → [Anti-captcha] Enviando hCaptcha para resolução...")
    data = await loop.run_in_executor(None, lambda: _post("/createTask", {
        "clientKey": api_key,
        "task":      task,
    }))

    if data.get("errorId"):
        raise RuntimeError(f"Anti-captcha erro ao criar tarefa: {data.get('errorDescription')}")

    task_id = data["taskId"]
    logger.info(f"  → [Anti-captcha] Tarefa {task_id}. Aguardando resolução...")

    await asyncio.sleep(15)
    for tentativa in range(20):
        await asyncio.sleep(5)
        result = await loop.run_in_executor(None, lambda: _post("/getTaskResult", {
            "clientKey": api_key,
            "taskId":    task_id,
        }))

        if result.get("errorId"):
            raise RuntimeError(f"Anti-captcha erro: {result.get('errorDescription')}")

        if result.get("status") == "ready":
            token = result["solution"]["gRecaptchaResponse"]
            logger.info(f"  → [Anti-captcha] Token obtido. (tentativa {tentativa + 1})")
            return token

        logger.debug(f"  → [Anti-captcha] {result.get('status')} ({tentativa + 1}/20)")

    raise RuntimeError("Anti-captcha timeout.")


async def extrair_sitekey(page) -> str | None:
    return await page.evaluate("""
        () => {
            const w = document.querySelector('[data-sitekey]');
            if (w) return w.getAttribute('data-sitekey');
            const fr = document.querySelector('iframe[src*="hcaptcha"]');
            if (fr) { try { return new URL(fr.src).searchParams.get('sitekey'); } catch(e){} }
            return null;
        }
    """)


async def injetar_token(page, token: str):
    """
    Injeta o token nos campos do hCaptcha E dispara os callbacks internos
    para que o widget reporte sucesso ao formulário.
    """
    await page.evaluate("""
        (token) => {
            const set = (el, val) => {
                const s = Object.getOwnPropertyDescriptor(
                    HTMLTextAreaElement.prototype, 'value').set;
                s.call(el, val);
                el.dispatchEvent(new Event('input',  {bubbles: true}));
                el.dispatchEvent(new Event('change', {bubbles: true}));
            };

            // Campos padrão do hCaptcha no documento principal
            ['h-captcha-response', 'g-recaptcha-response'].forEach(nome => {
                const el = document.querySelector('textarea[name="' + nome + '"]');
                if (el) set(el, token);
            });

            // Chama o callback do widget hCaptcha se disponível
            if (window.hcaptcha) {
                try {
                    // API v1: executa o callback registrado
                    const container = document.querySelector('[data-sitekey]');
                    if (container && container.dataset.callback) {
                        window[container.dataset.callback]?.(token);
                    }
                } catch(e) {}

                // Tenta forçar o estado de "solved" no widget interno
                try {
                    Object.values(window.hcaptcha._hcaptchaWidgets || {})
                          .forEach(w => { try { w.setResponse?.(token); } catch(e){} });
                } catch(e) {}
            }
        }
    """, token)
