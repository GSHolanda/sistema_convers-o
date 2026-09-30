"""Uso pela linha de comando.

    python -m conversor previa.pdf [outra.pdf ...] [--saida PASTA] [--razao-social "NOME"]

Sem arquivos, abre a janela do programa.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .conversao import caminho_config, converter_pdf
from .mapeamento import Config


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        from .app import main as abrir_janela

        abrir_janela([])
        return 0

    parser = argparse.ArgumentParser(prog="conversor", description="Converte Prévias (PDF) na planilha do Domínio.")
    parser.add_argument("pdfs", nargs="+", type=Path)
    parser.add_argument("--saida", type=Path, help="pasta de destino (padrão: a pasta de cada PDF)")
    parser.add_argument("--config", type=Path, default=None, help="arquivo de configuração")
    parser.add_argument("--razao-social", help="razão social a usar na planilha (substitui a do PDF)")
    args = parser.parse_args(argv)

    config = Config.carregar(args.config or caminho_config())
    erros = 0
    for pdf in args.pdfs:
        try:
            if args.razao_social:
                from .leitor_pdf import ler_previa

                for folha in ler_previa(pdf):
                    config.definir_razao_social(folha.empresa_codigo, args.razao_social)
            resultados = converter_pdf(pdf, config, args.saida)
        except Exception as erro:  # noqa: BLE001
            print(f"ERRO {pdf}: {erro}", file=sys.stderr)
            erros += 1
            continue
        for r in resultados:
            print(f"OK  {r.destino}  ({len(r.linhas)} funcionários)")
            for aviso in r.avisos:
                print(f"    AVISO: {aviso}")
    return 1 if erros else 0


if __name__ == "__main__":
    sys.exit(main())
