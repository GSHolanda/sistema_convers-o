from decimal import Decimal

from conversor.leitor_pdf import Evento, Folha, Funcionario
from conversor.mapeamento import Config, montar_linhas, normalizar


def evento(codigo, descricao, valor, tipo=1):
    return Evento(codigo, tipo, descricao, "", Decimal(valor))


def test_normalizar():
    assert normalizar("Prêmios  Produtividade") == "PREMIOS PRODUTIVIDADE"
    assert normalizar("COMPLEMENTO VR E VA") == normalizar("Complemento VR e VA")
    assert normalizar("Horas Extras 50%") == normalizar("Horas Extras 50 %")


def test_evento_por_codigo_e_por_descricao():
    config = Config()
    assert config.coluna_do_evento(evento(90, "Comissões", "1")) == "37"
    assert config.coluna_do_evento(evento(5555, "Comissões", "1")) == "37"
    assert config.coluna_do_evento(evento(1234, "Horas Extras 50%", "1")) == "150"
    assert config.coluna_do_evento(evento(1211, "Desconto de Vale Refeição", "1", 3)) == "9999"
    assert config.coluna_do_evento(evento(1860, "Contribuição Sindical", "1", 3)) == "52"


def test_eventos_ignorados():
    config = Config()
    for codigo, descricao in [
        (91, "DSR S/Comissões"),
        (884, "DSR Prêmio Produtividade"),
        (357, "Media Comissoes/DSR Ferias"),
        (92, "Complemento Salário Normativo"),
        (816, "Vale Transporte (%)"),
        (1950, "INSS"),
    ]:
        assert config.coluna_do_evento(evento(codigo, descricao, "1")) is None, descricao


def test_montar_linhas_ordena_por_codigo_e_soma():
    folha = Folha(123, "X")
    folha.funcionarios = [
        Funcionario(20, "B", eventos=[evento(692, "Desconto de Vale Refeição", "10", 3),
                                      evento(1211, "Desconto de Vale Refeição", "5.5", 3)]),
        Funcionario(3, "A", eventos=[evento(90, "Comissões", "100"), evento(91, "DSR S/Comissões", "20")]),
    ]
    resultado = montar_linhas(folha, Config())
    assert [l.codigo for l in resultado.linhas] == [3, 20]
    assert resultado.linhas[0].valores == {"37": Decimal(100)}
    assert resultado.linhas[1].valores == {"9999": Decimal("15.5")}
    assert resultado.ignorados == {(91, "DSR S/Comissões"): 1}


def test_demitidos_ficam_sem_valores():
    folha = Folha(123, "X")
    folha.funcionarios = [
        Funcionario(1, "A", situacao="Demitido", eventos=[evento(90, "Comissões", "100")]),
        Funcionario(2, "B", demissao="14/08/2026", eventos=[evento(90, "Comissões", "50")]),
        Funcionario(3, "C", situacao="Trabalhando", eventos=[evento(90, "Comissões", "10")]),
    ]
    linhas = montar_linhas(folha, Config()).linhas
    assert [(l.codigo, l.nome, l.valores) for l in linhas] == [
        (1, "A", {}), (2, "B", {}), (3, "C", {"37": Decimal(10)})
    ]
    config = Config({"demitidos_sem_valores": False})
    assert montar_linhas(folha, config).linhas[0].valores == {"37": Decimal(100)}


def test_tipo_de_calculo():
    config = Config()
    assert config.tipo_calculo(Folha(1, "X", tipo_calculo="Mensal")) == 11
    assert config.tipo_calculo(Folha(1, "X", tipo_calculo="Adiantamento")) == 41
    assert config.tipo_calculo(Folha(1, "X", tipo_calculo="Desconhecido")) == 11


def test_razao_social_e_nome_do_arquivo():
    from datetime import date

    config = Config()
    folha = Folha(123, "EMPRESA EIRELI - ME", competencia=date(2026, 7, 1))
    assert config.razao_social(folha) == "EMPRESA EIRELI - ME"
    config.definir_razao_social(123, "EMPRESA LTDA")
    assert config.razao_social(folha) == "EMPRESA LTDA"
    assert config.nome_arquivo(folha) == "dominio planilha - 0123 - 07-2026.xlsx"


def test_config_salva_e_carrega(tmp_path):
    caminho = tmp_path / "cfg.json"
    config = Config.carregar(caminho)
    assert caminho.exists()
    config.definir_razao_social(1, "ABC")
    config.salvar()
    assert Config.carregar(caminho).razao_social(Folha(1, "XYZ")) == "ABC"


def test_config_com_erro(tmp_path):
    import pytest

    caminho = tmp_path / "cfg.json"
    caminho.write_text("{ isso não é json", encoding="utf-8")
    with pytest.raises(ValueError, match="erro na linha"):
        Config.carregar(caminho)
