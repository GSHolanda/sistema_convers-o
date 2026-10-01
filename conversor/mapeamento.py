"""Regras que ligam os eventos da Prévia às colunas da planilha modelo.

Cada coluna da planilha é identificada pelo código da linha 10 do modelo
(37, 150, 299...). Um evento do PDF vai para a coluna se o código do evento
estiver em "eventos" OU se a descrição (sem acentos/maiúsculas) for igual a
uma das "descricoes". Eventos que não se encaixam em nenhuma coluna são
ignorados (DSR, salário, férias, INSS, IRRF etc.).
"""

from __future__ import annotations

import copy
import json
import re
import unicodedata
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from .leitor_pdf import Evento, Folha, Funcionario

# Ao mudar as regras padrão, aumente a versão: configurações antigas recebem as novas regras
# (as razões sociais cadastradas são mantidas).
VERSAO_CONFIG = 2

CONFIG_PADRAO: dict = {
    "versao": VERSAO_CONFIG,
    # Coluna A da planilha. Chave = texto após o período no PDF ("Mensal").
    "tipo_calculo": {
        "Mensal": 11,
        "Adiantamento": 41,
        "Complementar": 42,
        "13º Adiantamento": 51,
        "13º Integral": 52,
    },
    "tipo_calculo_padrao": 11,
    # Funcionários demitidos no período aparecem na planilha, mas sem valores.
    "demitidos_sem_valores": True,
    # Chave = código da coluna na linha 10 do modelo.
    "colunas": {
        "37": {
            "titulo": "Comissão",
            "eventos": [90, 357],
            "descricoes": ["Comissões", "Comissão", "Media Comissoes/DSR Ferias"],
        },
        "150": {
            "titulo": "Horas Extras 50%",
            "eventos": [],
            "descricoes": ["Horas Extras 50%", "Hora Extra 50%", "Horas Extras 50 %", "H.Extras 50%"],
        },
        "299": {"titulo": "Prêmio Destaque", "eventos": [], "descricoes": ["Prêmio Destaque", "Prêmios Destaque"]},
        "301": {"titulo": "Prêmio Mensal", "eventos": [932], "descricoes": ["Premiações Mensais", "Prêmio Mensal"]},
        "303": {
            "titulo": "Acréscimos Domingo",
            "eventos": [879],
            "descricoes": ["Acréscimo de comissão domingo"],
        },
        "306": {
            "titulo": "Prêmio proatividade",
            "eventos": [883],
            "descricoes": ["Prêmios Produtividade", "Prêmio Produtividade", "Prêmio Proatividade"],
        },
        "309": {"titulo": "Prêmio posição", "eventos": [], "descricoes": ["Prêmio Posição", "Premiação Posição"]},
        "312": {"titulo": "Prêmio Ranking", "eventos": [], "descricoes": ["Prêmio Ranking", "Premiação Ranking"]},
        "316": {
            "titulo": "Prêmio desempenho",
            "eventos": [],
            "descricoes": ["Prêmio Desempenho", "Prêmios Desempenho", "Premiação Desempenho"],
        },
        "318": {
            "titulo": "Prêmio mês anterior",
            "eventos": [],
            "descricoes": ["Prêmio Mês Anterior", "Premiação Mês Anterior"],
        },
        "305": {"titulo": "Complemento VR e VA", "eventos": [1306], "descricoes": ["Complemento VR e VA"]},
        "259": {"titulo": "Auxílio Transporte", "eventos": [687], "descricoes": ["Auxílio Vale Transporte"]},
        "260": {"titulo": "Desconto Refeição", "eventos": [692, 1211], "descricoes": ["Desconto de Vale Refeição"]},
        "52": {"titulo": "Sindical", "eventos": [], "descricoes": []},
        "9999": {"titulo": "", "eventos": [], "descricoes": []},
    },
    # Razão social usada na planilha, por código da empresa (substitui a do PDF).
    "razao_social": {},
    # {empresa} = código com 4 dígitos, {mes}/{ano} = competência.
    "nome_arquivo": "dominio planilha - {empresa} - {mes}-{ano}.xlsx",
}


def normalizar(texto: str) -> str:
    """Remove acentos, pontuação e espaços repetidos; tudo em maiúsculas."""
    sem_acento = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    sem_acento = sem_acento.replace("%", " PORCENTO ")
    return re.sub(r"[^A-Z0-9]+", " ", sem_acento.upper()).strip()


class Config:
    def __init__(self, dados: dict | None = None, caminho: Path | None = None):
        self.dados = copy.deepcopy(CONFIG_PADRAO)
        if dados:
            for chave, valor in dados.items():
                self.dados[chave] = valor
        self.caminho = caminho

    @classmethod
    def carregar(cls, caminho: Path) -> Config:
        if caminho.exists():
            try:
                with open(caminho, encoding="utf-8") as f:
                    dados = json.load(f)
            except json.JSONDecodeError as erro:
                raise ValueError(
                    f"O arquivo de configuração tem um erro na linha {erro.lineno}, "
                    f"coluna {erro.colno}:\n{caminho}\n\n({erro.msg})"
                ) from erro
            if dados.get("versao", 1) < VERSAO_CONFIG:
                razoes = dados.get("razao_social", {})
                dados = copy.deepcopy(CONFIG_PADRAO)
                dados["razao_social"] = razoes
                config = cls(dados, caminho)
                try:
                    config.salvar()
                except OSError:
                    pass
                return config
            return cls(dados, caminho)
        config = cls(caminho=caminho)
        try:
            config.salvar()
        except OSError:
            pass
        return config

    def salvar(self) -> None:
        if self.caminho is None:
            return
        self.caminho.parent.mkdir(parents=True, exist_ok=True)
        with open(self.caminho, "w", encoding="utf-8") as f:
            json.dump(self.dados, f, ensure_ascii=False, indent=2)

    def razao_social(self, folha: Folha) -> str:
        return self.dados.get("razao_social", {}).get(str(folha.empresa_codigo), folha.razao_social)

    def tem_razao_social(self, empresa_codigo: int) -> bool:
        return str(empresa_codigo) in self.dados.get("razao_social", {})

    def definir_razao_social(self, empresa_codigo: int, razao: str) -> None:
        self.dados.setdefault("razao_social", {})[str(empresa_codigo)] = razao

    def tipo_calculo(self, folha: Folha) -> int:
        tabela = {normalizar(k): v for k, v in self.dados.get("tipo_calculo", {}).items()}
        return int(tabela.get(normalizar(folha.tipo_calculo), self.dados.get("tipo_calculo_padrao", 11)))

    def nome_arquivo(self, folha: Folha) -> str:
        comp = folha.competencia
        return self.dados.get("nome_arquivo", CONFIG_PADRAO["nome_arquivo"]).format(
            empresa=f"{folha.empresa_codigo:04d}",
            mes=f"{comp.month:02d}" if comp else "00",
            ano=f"{comp.year}" if comp else "0000",
        )

    def coluna_do_evento(self, evento: Evento) -> str | None:
        desc = normalizar(evento.descricao)
        for codigo_coluna, regra in self.dados.get("colunas", {}).items():
            if evento.codigo in regra.get("eventos", []):
                return codigo_coluna
            if desc and desc in (normalizar(d) for d in regra.get("descricoes", [])):
                return codigo_coluna
        return None


@dataclass
class LinhaPlanilha:
    tipo_calculo: int
    codigo: int
    nome: str
    valores: dict[str, Decimal]  # código da coluna -> valor


@dataclass
class Resultado:
    linhas: list[LinhaPlanilha]
    ignorados: dict[tuple[int, str], int]  # (código, descrição) -> ocorrências


def demitido(func: Funcionario) -> bool:
    return bool(func.demissao) or normalizar(func.situacao).startswith("DEMITID")


def montar_linhas(folha: Folha, config: Config) -> Resultado:
    """Converte os funcionários da Prévia em linhas da planilha (ordem de código)."""
    tipo = config.tipo_calculo(folha)
    linhas = []
    ignorados: dict[tuple[int, str], int] = {}
    sem_valores_demitidos = config.dados.get("demitidos_sem_valores", True)
    for func in sorted(folha.funcionarios, key=lambda f: f.codigo):
        valores: dict[str, Decimal] = {}
        if sem_valores_demitidos and demitido(func):
            linhas.append(LinhaPlanilha(tipo, func.codigo, func.nome, valores))
            continue
        for evento in func.eventos:
            coluna = config.coluna_do_evento(evento)
            if coluna is None:
                chave = (evento.codigo, evento.descricao)
                ignorados[chave] = ignorados.get(chave, 0) + 1
                continue
            valores[coluna] = valores.get(coluna, Decimal(0)) + evento.valor
        linhas.append(LinhaPlanilha(tipo, func.codigo, func.nome, valores))
    return Resultado(linhas, ignorados)
