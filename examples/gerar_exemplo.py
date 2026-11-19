# -*- coding: utf-8 -*-
"""Gera uma árvore FICTÍCIA de relatórios RM (mensal) e RO (semanal) para rodar a automação sem o acervo real.

Cria examples/engenharia/<RODOVIA>/3_Gestao Tecnica de Obra/... com PDFs de 1 página em branco
(nomes inventados, datas de set/2025 a nov/2025). Depois:

    py -3.13 prestacao_contas_relatorios.py --roads RD101 RD103

Requer pypdf (requirements_prestacao_contas.txt).
"""
from pathlib import Path

from pypdf import PdfWriter

AQUI = Path(__file__).resolve().parent
MENSAL = Path("3_Gestao Tecnica de Obra") / "5. Relatório Mensal da Obra"
SEMANAL = (Path("3_Gestao Tecnica de Obra") / "4. Relatório Semanal Obras"
           / "5. Relatório Semanal de Acompanhamento de Obras")

ARQUIVOS = {
    "RD101": {
        MENSAL: ["RM-001.1 - Relatorio Mensal - 09-2025.pdf", "RM-002.1 - Relatorio Mensal - 10-2025.pdf",
                 "RM-003.1 - Relatorio Mensal - 11-2025.pdf"],
        SEMANAL: ["RO-001 - Relatorio Semanal - 05-09-2025.pdf", "RO-005 - Relatorio Semanal - 03-10-2025.pdf",
                  "RO-009 - Relatorio Semanal - 07-11-2025.pdf"],
    },
    "RD103": {
        MENSAL: ["RM-001.1 - Relatorio Mensal - 09-2025.pdf"],
        SEMANAL: ["RO-002 - Relatorio Semanal - 12-09-2025.pdf"],
    },
}


def pdf_em_branco(destino: Path):
    destino.parent.mkdir(parents=True, exist_ok=True)
    w = PdfWriter()
    w.add_blank_page(width=595, height=842)
    with open(destino, "wb") as fh:
        w.write(fh)


if __name__ == "__main__":
    for rodovia, pastas in ARQUIVOS.items():
        for sub, nomes in pastas.items():
            for nome in nomes:
                pdf_em_branco(AQUI / "engenharia" / rodovia / sub / nome)
    print("Exemplo fictício criado em", AQUI / "engenharia")
