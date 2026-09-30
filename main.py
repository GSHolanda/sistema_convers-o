"""Ponto de entrada do executável (PyInstaller).

Sem argumentos (ou com PDFs arrastados sobre o .exe) abre a janela.
Com --converter, converte sem janela:  ConversorPrevia.exe --converter previa.pdf --saida PASTA
"""

import sys


def main() -> int:
    argumentos = sys.argv[1:]
    if argumentos and argumentos[0] == "--converter":
        from conversor.__main__ import main as linha_de_comando

        return linha_de_comando(argumentos[1:])
    from conversor.app import main as janela

    janela(argumentos)
    return 0


if __name__ == "__main__":
    sys.exit(main())
