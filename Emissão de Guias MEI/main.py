"""
Ponto de entrada do robô de emissão de Guias MEI.

Uso:
  python main.py
"""

import asyncio
import logging
import sys
from datetime import datetime

from config import (
    DOWNLOADS_DIR,
    EXCEL_INPUT,
    LOGS_DIR,
    RELATORIOS_DIR,
    competencia_atual,
)
from bot.pgmei import PGMEIBot
from utils.excel_reader import ler_clientes
from utils.relatorio import gerar_relatorio


def configurar_log():
    LOGS_DIR.mkdir(exist_ok=True)
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    arq_log  = LOGS_DIR / f"execucao_{ts}.log"

    fmt = "%(asctime)s | %(levelname)-8s | %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[
            logging.FileHandler(arq_log, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger(__name__)


async def main():
    logger = configurar_log()

    # Verifica se o arquivo de entrada existe
    if not EXCEL_INPUT.exists():
        logger.error(f"Arquivo não encontrado: {EXCEL_INPUT}")
        logger.error("Renomeie sua planilha para 'cnpjs.xlsx.xlsx' ou ajuste EXCEL_INPUT em config.py")
        return

    # Lê clientes MEI
    clientes = ler_clientes(EXCEL_INPUT)
    if not clientes:
        logger.error("Nenhum cliente MEI encontrado na planilha. Verifique a coluna 'Enq.'")
        return

    # Filtro por CNPJ específico
    # Prioridade: argumento CLI → prompt interativo → processar todos
    cnpj_alvo       = sys.argv[1].strip() if len(sys.argv) > 1 else None
    cnpj_alvo_limpo = None

    # Só exibe prompt se stdout também for um terminal (não um pipe do servidor web)
    if not cnpj_alvo and sys.stdin.isatty() and sys.stdout.isatty():
        try:
            resposta = input(
                "\nDigite o CNPJ para processar apenas 1 cliente "
                "(somente números, ou ENTER para processar todos): "
            ).strip()
            if resposta:
                cnpj_alvo = resposta
        except EOFError:
            pass

    if cnpj_alvo:
        cnpj_alvo_limpo = "".join(d for d in cnpj_alvo if d.isdigit())
        clientes = [c for c in clientes if c["cnpj"] == cnpj_alvo_limpo]
        if not clientes:
            logger.error(f"CNPJ '{cnpj_alvo}' não encontrado na planilha.")
            return
        logger.info(f"Modo individual: {clientes[0]['nome']}  |  CNPJ: {cnpj_alvo_limpo}")

    # Competência atual (mês/ano de hoje)
    mes, ano = competencia_atual()
    logger.info(f"Competência alvo: {mes:02d}/{ano}")

    # Garante diretórios de saída
    DOWNLOADS_DIR.mkdir(exist_ok=True)
    RELATORIOS_DIR.mkdir(exist_ok=True)

    # Jonas Durski (índice 0) foi limitado hoje pela Receita — processa por último
    # (aplica apenas quando rodando todos os clientes)
    if not cnpj_alvo:
        clientes = clientes[1:] + clientes[:1]
    logger.info(f"Iniciando processamento de {len(clientes)} clientes...\n")
    bot        = PGMEIBot()
    resultados = []
    try:
        resultados = await bot.processar_lote(clientes, mes, ano)
    except Exception as exc:
        logger.error(f"Erro inesperado durante processamento: {exc}")
        # Resgata resultados já coletados antes do crash
        resultados = getattr(bot, "_resultados", [])
        if resultados:
            logger.warning(f"  → Salvando {len(resultados)} resultado(s) parcial(is)...")

    if not resultados:
        logger.warning("Nenhum resultado para salvar.")
        return

    # Relatório final (sufixo com CNPJ quando modo individual)
    sufixo_cnpj = cnpj_alvo_limpo if cnpj_alvo else None
    relatorio = gerar_relatorio(resultados, RELATORIOS_DIR, sufixo=sufixo_cnpj)

    # Resumo no console
    sucessos    = sum(1 for r in resultados if r["Status"] == "SUCESSO")
    falhas      = len(resultados) - sucessos
    com_debitos = [r for r in resultados if r.get("Débitos em Atraso")]

    logger.info(f"\n{'='*55}")
    logger.info(f"CONCLUÍDO  →  ✓ {sucessos} sucesso(s)   ✗ {falhas} falha(s)")
    logger.info(f"Relatório salvo em: {relatorio}")
    logger.info(f"PDFs em: {DOWNLOADS_DIR}")

    if com_debitos:
        logger.info(f"\n{'─'*55}")
        logger.info(f"⚠  CLIENTES COM DÉBITOS PENDENTES ({len(com_debitos)}):")
        for r in com_debitos:
            logger.info(f"   • {r['Nome']}  |  CNPJ: {r['CNPJ']}")
            logger.info(f"     Débitos: {r['Débitos em Atraso']}")
        logger.info(f"{'─'*55}")
    else:
        logger.info("   Nenhum débito pendente detectado.")

    logger.info("=" * 55)


if __name__ == "__main__":
    asyncio.run(main())
