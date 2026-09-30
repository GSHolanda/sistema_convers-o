"""Leitura do relatório "Relação de Cálculo" (Prévia da folha) em PDF.

O relatório tem layout fixo: cada funcionário começa com uma linha "Func:",
seguida das linhas de eventos em duas colunas (proventos à esquerda,
descontos/informativos à direita) e termina com a linha "Proventos: ...".
A leitura usa a posição de cada palavra na página, o que evita erros de
espaçamento que a extração de texto simples produz (ex.: "11 3,12").
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path

import pdfplumber

# Tolerância vertical (pt) para considerar palavras na mesma linha.
TOLERANCIA_LINHA = 3.0
# Distância mínima (pt) entre a descrição do evento e a referência (dias, horas, %).
ESPACO_REFERENCIA = 6.0

RE_VALOR = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")
RE_EMPRESA = re.compile(r"^Empresa:\s*(\d+)\s*-\s*(.+?)\s+\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}")
RE_INSCRICAO = re.compile(r"Inscri[çc][ãa]o Federal:\s*([\d./-]+)")
RE_PERIODO = re.compile(
    r"Per[íi]odo:\s*(\d{2})/(\d{2})/(\d{4})\s*a\s*\d{2}/\d{2}/\d{4}\s*-\s*(.+)$"
)
RE_FUNC = re.compile(r"^Func:\s*(\d+)\s+(.+?)\s+Adm\s*(\d{2}/\d{2}/\d{4})?\s*Dem\s*(\d{2}/\d{2}/\d{4})?")
RE_SITUACAO = re.compile(r"Situa[çc][ãa]o:\s*(.+)$")
RE_TOTAIS = re.compile(
    r"Proventos:\s*(\S+)\s+Vantagens:\s*(\S+)\s+Descontos:\s*(\S+)\s+L[íi]quido:\s*(\S+)"
)
RE_TOTAL_FUNC = re.compile(r"Total Funcion[áa]rios\s+(\d+)")


def valor_decimal(texto: str) -> Decimal:
    """Converte "1.802,14" em Decimal("1802.14")."""
    return Decimal(texto.replace(".", "").replace(",", "."))


@dataclass
class Evento:
    codigo: int
    tipo: int
    descricao: str
    referencia: str
    valor: Decimal


@dataclass
class Funcionario:
    codigo: int
    nome: str
    admissao: str = ""
    demissao: str = ""
    situacao: str = ""
    eventos: list[Evento] = field(default_factory=list)
    total_proventos: Decimal | None = None
    total_descontos: Decimal | None = None


@dataclass
class Folha:
    """Dados de uma empresa extraídos da Prévia."""

    empresa_codigo: int
    razao_social: str
    inscricao: str = ""
    competencia: date | None = None
    tipo_calculo: str = ""
    funcionarios: list[Funcionario] = field(default_factory=list)
    total_funcionarios: int | None = None
    arquivo: str = ""

    def avisos(self) -> list[str]:
        """Conferências de integridade da leitura."""
        avisos = []
        if self.competencia is None:
            avisos.append("Período/competência não encontrado no PDF.")
        if self.total_funcionarios is not None and self.total_funcionarios != len(self.funcionarios):
            avisos.append(
                f"O PDF informa {self.total_funcionarios} funcionário(s), "
                f"mas foram lidos {len(self.funcionarios)}."
            )
        for f in self.funcionarios:
            proventos = sum((e.valor for e in f.eventos if e.tipo == 1), Decimal(0))
            descontos = sum((e.valor for e in f.eventos if e.tipo == 3), Decimal(0))
            if f.total_proventos is not None and proventos != f.total_proventos:
                avisos.append(
                    f"Funcionário {f.codigo} {f.nome}: soma dos proventos lidos ({proventos}) "
                    f"difere do total do PDF ({f.total_proventos})."
                )
            if f.total_descontos is not None and descontos != f.total_descontos:
                avisos.append(
                    f"Funcionário {f.codigo} {f.nome}: soma dos descontos lidos ({descontos}) "
                    f"difere do total do PDF ({f.total_descontos})."
                )
        return avisos


def _agrupar_linhas(palavras: list[dict]) -> list[list[dict]]:
    linhas: list[list[dict]] = []
    topo_atual = None
    for p in sorted(palavras, key=lambda p: (p["top"], p["x0"])):
        if topo_atual is None or abs(p["top"] - topo_atual) > TOLERANCIA_LINHA:
            linhas.append([])
            topo_atual = p["top"]
        linhas[-1].append(p)
    return [sorted(linha, key=lambda p: p["x0"]) for linha in linhas]


def _texto(palavras: list[dict]) -> str:
    return " ".join(p["text"] for p in palavras)


def _ler_evento(palavras: list[dict]) -> Evento | None:
    """Interpreta metade de uma linha: código, tipo, descrição, [referência], valor."""
    if len(palavras) < 4:
        return None
    codigo, tipo, valor = palavras[0]["text"], palavras[1]["text"], palavras[-1]["text"]
    if not (codigo.isdigit() and len(tipo) == 1 and tipo.isdigit() and RE_VALOR.match(valor)):
        return None
    meio = palavras[2:-1]
    fim_descricao = len(meio)
    for i in range(1, len(meio)):
        if meio[i]["x0"] - meio[i - 1]["x1"] > ESPACO_REFERENCIA:
            fim_descricao = i
            break
    descricao = _texto(meio[:fim_descricao])
    if not descricao:
        return None
    return Evento(
        codigo=int(codigo),
        tipo=int(tipo),
        descricao=descricao,
        referencia=_texto(meio[fim_descricao:]),
        valor=valor_decimal(valor),
    )


def ler_previa(caminho: str | Path) -> list[Folha]:
    """Lê o PDF e devolve uma Folha por empresa encontrada."""
    folhas: dict[int, Folha] = {}
    folha: Folha | None = None
    func: Funcionario | None = None

    with pdfplumber.open(str(caminho)) as pdf:
        for pagina in pdf.pages:
            meio_pagina = pagina.width * 0.51
            palavras = pagina.extract_words(x_tolerance=3, y_tolerance=3)
            for linha in _agrupar_linhas(palavras):
                texto = _texto(linha)

                m = RE_EMPRESA.match(texto)
                if m:
                    codigo = int(m.group(1))
                    if codigo not in folhas:
                        folhas[codigo] = Folha(
                            empresa_codigo=codigo,
                            razao_social=m.group(2).strip(),
                            arquivo=str(caminho),
                        )
                        func = None
                    folha = folhas[codigo]
                    continue
                if folha is None:
                    continue

                m = RE_INSCRICAO.search(texto)
                if m:
                    folha.inscricao = folha.inscricao or re.sub(r"\D", "", m.group(1))
                    continue

                m = RE_PERIODO.search(texto)
                if m:
                    folha.competencia = folha.competencia or date(int(m.group(3)), int(m.group(2)), 1)
                    folha.tipo_calculo = folha.tipo_calculo or m.group(4).strip()
                    continue

                m = RE_FUNC.match(texto)
                if m:
                    codigo = int(m.group(1))
                    existente = next((f for f in folha.funcionarios if f.codigo == codigo), None)
                    if existente is None:
                        existente = Funcionario(codigo=codigo, nome=m.group(2).strip())
                        folha.funcionarios.append(existente)
                    existente.admissao = m.group(3) or existente.admissao
                    existente.demissao = m.group(4) or existente.demissao
                    func = existente
                    continue

                if texto.startswith("Total Empresa:"):
                    func = None
                    continue

                m = RE_TOTAL_FUNC.search(texto)
                if m:
                    folha.total_funcionarios = int(m.group(1))
                    continue

                if func is None:
                    continue

                if texto.startswith("Cargo:"):
                    m = RE_SITUACAO.search(texto)
                    if m:
                        func.situacao = m.group(1).strip()
                    continue

                m = RE_TOTAIS.search(texto)
                if m:
                    try:
                        func.total_proventos = valor_decimal(m.group(1))
                        func.total_descontos = valor_decimal(m.group(3))
                    except ArithmeticError:
                        pass
                    continue

                esquerda = [p for p in linha if p["x0"] < meio_pagina]
                direita = [p for p in linha if p["x0"] >= meio_pagina]
                for metade in (esquerda, direita):
                    evento = _ler_evento(metade)
                    if evento:
                        func.eventos.append(evento)

    return list(folhas.values())
