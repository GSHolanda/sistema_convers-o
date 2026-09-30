"""Gera uma Prévia (Relação de Cálculo) fictícia com o mesmo layout do relatório real.

Usada só nos testes: nomes, códigos e valores são inventados.
"""

from __future__ import annotations

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

LARGURA, ALTURA = A4
FONTE, FONTE_NEGRITO, TAMANHO = "Times-Roman", "Times-Bold", 8.2

EMPRESA = (123, "EMPRESA EXEMPLO EIRELI - ME", "12.345.678/0001-90")
PERIODO = "01/08/2026 a 31/08/2026"

# (código, nome, admissão, demissão, situação, eventos esquerda, eventos direita)
# evento = (código, tipo, descrição, referência, valor)
FUNCIONARIOS = [
    (
        5, "BELTRANA DE SOUZA", "10/01/2024", "", "Férias",
        [
            (357, 1, "Media Comissoes/DSR Ferias", "212:40 hs", "1.500,00"),
            (386, 1, "1/3 Sobre Férias", "33,33 %", "500,00"),
            (687, 1, "Auxílio Vale Transporte", "", "275,00"),
            (710, 5, "Devolução INSS Mês", "", "6,61"),
        ],
        [
            (692, 3, "Desconto de Vale Refeição", "", "1,26"),
            (890, 3, "Desconto Adiantamento", "", "1.800,00"),
            (4002, 4, "Dedução Simplificada - IRRF", "", "158,98"),
        ],
    ),
    (
        20, "CICRANO PEREIRA", "16/10/2025", "14/08/2026", "Demitido",
        [(90, 1, "Comissões", "12 Dias", "665,83")],
        [
            (1211, 3, "Desconto de Vale Refeição", "", "15,12"),
            (1895, 3, "Desconto Líquido Rescisão", "", "650,71"),
        ],
    ),
    (
        10, "FULANO DE TAL", "13/06/2023", "", "Trabalhando",
        [
            (90, 1, "Comissões", "26 Dias", "1.234,56"),
            (91, 1, "DSR S/Comissões", "5 Dias", "237,42"),
            (687, 1, "Auxílio Vale Transporte", "", "300,00"),
            (879, 1, "Acréscimo de comissão domingo", "", "46,33"),
            (880, 1, "DSR Acréscimo comissão domingo", "5 Dias", "8,91"),
            (883, 1, "Prêmios Produtividade", "", "100,00"),
            (884, 1, "DSR Prêmio Produtividade", "", "19,23"),
            (932, 1, "Premiações Mensais", "", "450,10"),
            (1306, 1, "COMPLEMENTO VR E VA", "", "13,26"),
            (1307, 1, "DSR PREMIAÇÕES", "", "111,11"),
        ],
        [
            (692, 3, "Desconto de Vale Refeição", "", "32,76"),
            (1950, 3, "INSS", "9,00 %", "150,00"),
            (1860, 3, "Contribuição Sindical", "", "42,00"),
            (4000, 4, "Dedução Simplificada - IRRF", "", "341,78"),
        ],
    ),
]


def _br(valor: float) -> str:
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _numero(texto: str) -> float:
    return float(texto.replace(".", "").replace(",", "."))


class _Pdf:
    def __init__(self, caminho: str):
        self.c = canvas.Canvas(caminho, pagesize=A4)
        self.pagina = 0
        self.topo = 0.0
        self._nova_pagina()

    def texto(self, x: float, topo: float, texto: str, negrito: bool = False, tamanho: float = TAMANHO) -> None:
        self.c.setFont(FONTE_NEGRITO if negrito else FONTE, tamanho)
        self.c.drawString(x, ALTURA - topo - tamanho, texto)

    def direita(self, x1: float, topo: float, texto: str) -> None:
        self.c.setFont(FONTE, TAMANHO)
        self.c.drawRightString(x1, ALTURA - topo - TAMANHO, texto)

    def _nova_pagina(self) -> None:
        if self.pagina:
            self.c.showPage()
        self.pagina += 1
        codigo, nome, inscricao = EMPRESA
        self.texto(87, 26, "Empresa:", True)
        self.texto(161, 26, f"{codigo:04d} - {nome}")
        self.texto(485, 26, f"03/09/2026 07:56 Pág:{self.pagina:04d}")
        self.texto(87, 38, "Inscrição Federal:", True)
        self.texto(161, 38, inscricao)
        self.texto(122, 83, f"Relação de Cálculo - Período: {PERIODO} - Mensal", True, 12)
        self.texto(12, 102, f"{codigo:04d} - {nome}", True)
        self.topo = 130

    def linha(self, altura: float = 12) -> float:
        if self.topo + altura > ALTURA - 40:
            self._nova_pagina()
        topo = self.topo
        self.topo += altura
        return topo

    def salvar(self) -> None:
        self.c.save()


def _evento(pdf: _Pdf, topo: float, evento, direita: bool) -> None:
    codigo, tipo, descricao, referencia, valor = evento
    # Posições (x) das colunas no relatório real.
    x_codigo, x_tipo, x_descricao, x_referencia, x_valor = (
        (317, 351, 365, 497, 582) if direita else (12, 52, 66, 192, 292)
    )
    pdf.texto(x_codigo, topo, str(codigo))
    pdf.texto(x_tipo, topo, str(tipo))
    pdf.texto(x_descricao, topo, descricao)
    if referencia:
        pdf.texto(x_referencia, topo, referencia)
    pdf.direita(x_valor, topo, valor)


def gerar(caminho: str, funcionarios=FUNCIONARIOS, total_informado: int | None = None) -> None:
    pdf = _Pdf(caminho)
    pdf.texto(12, pdf.linha(16), "Analítico Contratos", True)
    for codigo, nome, adm, dem, situacao, esquerda, direita in funcionarios:
        topo = pdf.linha()
        pdf.texto(12, topo, "Func:", True)
        pdf.texto(51, topo, str(codigo))
        pdf.texto(75, topo, nome)
        pdf.texto(274, topo, f"Adm{adm}")
        pdf.texto(338, topo, "Dem")
        if dem:
            pdf.texto(356, topo, dem)
        pdf.texto(402, topo, "Dep.IR: 00")
        topo = pdf.linha()
        pdf.texto(12, topo, "Cargo: Vendedor(a) C.H.M: 220:00 Salário: 0,00")
        pdf.texto(471, topo, "Situação:")
        pdf.texto(513, topo, situacao)
        pdf.texto(12, pdf.linha(), "Filial: 0001 - CNPJ/CPF: 12.345.678/0001-90")
        for i in range(max(len(esquerda), len(direita))):
            topo = pdf.linha()
            if i < len(esquerda):
                _evento(pdf, topo, esquerda[i], False)
            if i < len(direita):
                _evento(pdf, topo, direita[i], True)
        proventos = sum(_numero(e[4]) for e in esquerda if e[1] == 1)
        descontos = sum(_numero(e[4]) for e in direita if e[1] == 3)
        topo = pdf.linha(16)
        pdf.texto(12, topo, "Proventos:", True)
        pdf.direita(101, topo, _br(proventos))
        pdf.texto(160, topo, "Vantagens:", True)
        pdf.direita(238, topo, "0,00")
        pdf.texto(317, topo, "Descontos:", True)
        pdf.direita(400, topo, _br(descontos))
        pdf.texto(439, topo, "Líquido:", True)
        pdf.direita(550, topo, _br(proventos - descontos))
        pdf.texto(12, pdf.linha(), "Base Impostos", True)
        pdf.texto(12, pdf.linha(), "IRRF 1.000,00 0,00 0,00 0,00")

    codigo, nome, _ = EMPRESA
    pdf.texto(12, pdf.linha(16), f"Total Empresa: {codigo:04d} - {nome}", True)
    pdf.texto(12, pdf.linha(), "Resumo Contrato", True)
    # Linhas de resumo têm o mesmo formato dos eventos e devem ser ignoradas.
    _evento(pdf, pdf.linha(), (90, 1, "Comissões", "1034:00", "9.999,99"), False)
    _evento(pdf, pdf.linha(), (692, 3, "Desconto de Vale Refeição", "0,00", "8.888,88"), True)
    total = len(funcionarios) if total_informado is None else total_informado
    pdf.texto(12, pdf.linha(), f"Dependentes SF 0 Dependentes IR 0 Dedução do IRRF 0,00 Total Funcionários {total}")
    pdf.salvar()


def gerar_com_quebra_de_pagina(caminho: str) -> None:
    """Muitos funcionários, para que eventos continuem na página seguinte."""
    base = FUNCIONARIOS[2]
    funcionarios = [(100 + i, f"PESSOA NUMERO {i:02d}", *base[2:]) for i in range(12)]
    gerar(caminho, funcionarios)
