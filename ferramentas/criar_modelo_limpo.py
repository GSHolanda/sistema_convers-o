"""Cria modelo/modelo_dominio.xlsx a partir de uma planilha real já preenchida.

Remove os dados da empresa e dos funcionários e o caminho de pasta gravado
pelo Excel, mantendo todo o resto do arquivo (estilos, comentários, fórmulas).

Uso: python ferramentas/criar_modelo_limpo.py planilha_preenchida.xlsx
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from conversor.planilha import Planilha

DESTINO = Path(__file__).resolve().parents[1] / "modelo" / "modelo_dominio.xlsx"


def main() -> None:
    planilha = Planilha(sys.argv[1])
    planilha.limpar_dados()

    workbook = planilha.pacote.xml(planilha.caminho_workbook)
    for elemento in workbook.iter("{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"):
        if any(e.tag.endswith("absPath") for e in elemento.iter()):
            elemento.getparent().remove(elemento)
            break
    planilha.pacote.definir_xml(planilha.caminho_workbook, workbook)

    DESTINO.parent.mkdir(exist_ok=True)
    planilha.salvar(DESTINO)
    print(f"Modelo criado: {DESTINO}")


if __name__ == "__main__":
    main()
