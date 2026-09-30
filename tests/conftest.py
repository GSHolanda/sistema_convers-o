import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import pdf_exemplo


@pytest.fixture
def previa(tmp_path):
    caminho = tmp_path / "Prévia exemplo.pdf"
    pdf_exemplo.gerar(str(caminho))
    return caminho


@pytest.fixture
def previa_longa(tmp_path):
    caminho = tmp_path / "Prévia longa.pdf"
    pdf_exemplo.gerar_com_quebra_de_pagina(str(caminho))
    return caminho
