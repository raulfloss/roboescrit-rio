import os as _os
from pathlib import Path
from datetime import datetime

BASE_DIR = Path(__file__).parent

EXCEL_INPUT         = BASE_DIR / "cnpjs.xlsx.xlsx"
DOWNLOADS_DIR       = BASE_DIR / "downloads"
LOGS_DIR            = BASE_DIR / "logs"
RELATORIOS_DIR      = BASE_DIR / "relatorios"
BROWSER_PROFILE_DIR = BASE_DIR / "browser_profile"  # acumula confiança no hCaptcha entre sessões

# --- Anti-captcha ---
# Coloque aqui a mesma chave que usa no seu outro robô
ANTICAPTCHA_API_KEY = "0b3f1fe79c286e30486d7739165822d2"

# --- Configurações do navegador ---
# Railway não tem display; usa headless automático. Localmente permanece headful (melhor pro captcha).
_on_railway = bool(_os.environ.get("RAILWAY_ENVIRONMENT") or _os.environ.get("RAILWAY_PROJECT_ID"))
HEADLESS        = _on_railway or _os.environ.get("MEI_HEADLESS", "0") == "1"
TIMEOUT_MS      = 40_000  # 40 segundos por ação
MAX_RETRIES     = 3       # tentativas por CNPJ
DELAY_CNPJS     = 3       # segundos de pausa entre CNPJs

# --- URLs (mapeadas pelo Network tab) ---
PGMEI_URL          = "https://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Identificacao"
PGMEI_URL_INICIO   = "https://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/pgmei.app/Home/Inicio"

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março",    4: "Abril",
    5: "Maio",    6: "Junho",     7: "Julho",     8: "Agosto",
    9: "Setembro",10: "Outubro",  11: "Novembro", 12: "Dezembro",
}

def competencia_atual() -> tuple[int, int]:
    now = datetime.now()
    return now.month, now.year
