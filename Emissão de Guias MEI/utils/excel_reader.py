import re
import logging
import openpyxl
from pathlib import Path

logger = logging.getLogger(__name__)


def _limpar_cnpj(valor: str) -> str:
    return re.sub(r"\D", "", str(valor))


def _limpar_nome(razao_social: str) -> str:
    """Remove o prefixo 'XX.XXX.XXX ' que aparece antes do nome na planilha."""
    nome = re.sub(r"^\d{2}\.\d{3}\.\d{3}\s+", "", str(razao_social)).strip()
    return nome or razao_social.strip()


def ler_clientes(caminho: Path) -> list[dict]:
    """
    Lê a planilha e retorna apenas as linhas com Enquadramento = MEI e CNPJ preenchido.

    Colunas esperadas (conforme a planilha atual):
      B → Razão Social    (índice 1)
      D → CNPJ            (índice 3)
      H → Enquadramento   (índice 7)
    """
    wb = openpyxl.load_workbook(caminho)
    ws = wb.active

    clientes = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        cnpj_raw   = row[3]  # coluna D
        nome_raw   = row[1]  # coluna B
        enq        = str(row[7] or "").upper().strip()  # coluna H

        if not cnpj_raw or "MEI" not in enq:
            continue

        cnpj = _limpar_cnpj(cnpj_raw)
        if len(cnpj) != 14:
            logger.warning(f"Linha {i}: CNPJ inválido '{cnpj_raw}', ignorando.")
            continue

        clientes.append({
            "cnpj":  cnpj,
            "nome":  _limpar_nome(str(nome_raw)),
            "linha": i,
        })

    logger.info(f"{len(clientes)} clientes MEI carregados de '{caminho.name}'")
    return clientes
