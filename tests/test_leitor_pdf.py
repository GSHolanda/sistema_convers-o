from datetime import date
from decimal import Decimal

import pdf_exemplo

from conversor.leitor_pdf import ler_previa, valor_decimal


def test_valor_decimal():
    assert valor_decimal("1.802,14") == Decimal("1802.14")
    assert valor_decimal("0,00") == Decimal(0)
    assert valor_decimal("12.345.678,90") == Decimal("12345678.90")


def test_cabecalho(previa):
    (folha,) = ler_previa(previa)
    assert folha.empresa_codigo == 123
    assert folha.razao_social == "EMPRESA EXEMPLO EIRELI - ME"
    assert folha.inscricao == "12345678000190"
    assert folha.competencia == date(2026, 8, 1)
    assert folha.tipo_calculo == "Mensal"
    assert folha.total_funcionarios == 3


def test_funcionarios_e_eventos(previa):
    (folha,) = ler_previa(previa)
    assert [(f.codigo, f.nome, f.situacao) for f in folha.funcionarios] == [
        (5, "BELTRANA DE SOUZA", "Férias"),
        (20, "CICRANO PEREIRA", "Demitido"),
        (10, "FULANO DE TAL", "Trabalhando"),
    ]
    cicrano = folha.funcionarios[1]
    assert cicrano.demissao == "14/08/2026"
    fulano = folha.funcionarios[2]
    eventos = {(e.codigo, e.tipo): e for e in fulano.eventos}
    assert eventos[(90, 1)].descricao == "Comissões"
    assert eventos[(90, 1)].referencia == "26 Dias"
    assert eventos[(90, 1)].valor == Decimal("1234.56")
    assert eventos[(1950, 3)].referencia == "9,00 %"
    assert eventos[(1307, 1)].valor == Decimal("111.11")
    assert eventos[(880, 1)].descricao == "DSR Acréscimo comissão domingo"
    assert len(fulano.eventos) == 14


def test_resumo_da_empresa_nao_entra_como_evento(previa):
    (folha,) = ler_previa(previa)
    valores = {e.valor for f in folha.funcionarios for e in f.eventos}
    assert Decimal("9999.99") not in valores
    assert Decimal("8888.88") not in valores


def test_conferencias_sem_avisos(previa):
    (folha,) = ler_previa(previa)
    assert folha.avisos() == []


def test_aviso_quando_total_de_funcionarios_nao_bate(tmp_path):
    caminho = tmp_path / "errado.pdf"
    pdf_exemplo.gerar(str(caminho), total_informado=4)
    (folha,) = ler_previa(caminho)
    assert any("4 funcionário" in a for a in folha.avisos())


def test_eventos_continuam_na_pagina_seguinte(previa_longa):
    (folha,) = ler_previa(previa_longa)
    assert len(folha.funcionarios) == 12
    assert all(len(f.eventos) == 14 for f in folha.funcionarios)
    assert folha.avisos() == []
