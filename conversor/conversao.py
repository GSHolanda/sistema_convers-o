"""Converte uma Prévia (PDF) na planilha do Domínio."""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass, field
from pathlib import Path

from .leitor_pdf import Folha, ler_previa
from .mapeamento import Config, LinhaPlanilha, montar_linhas
from .planilha import Planilha


def pasta_do_programa() -> Path:
    """Pasta do .exe (ou da raiz do projeto quando roda pelo Python)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def recurso(relativo: str) -> Path:
    """Arquivo embutido no .exe (PyInstaller) ou no projeto."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
    return base / relativo


MODELO_PADRAO = recurso("modelo/modelo_dominio.xlsx")


NOME_CONFIG = "conversor_config.json"


def caminho_config() -> Path:
    """Configuração ao lado do programa; se a pasta não permitir gravar, em %APPDATA%."""
    local = pasta_do_programa() / NOME_CONFIG
    if local.exists() or os.access(local.parent, os.W_OK):
        return local
    base = Path(os.environ.get("APPDATA") or Path.home())
    return base / "ConversorPrevia" / NOME_CONFIG


@dataclass
class ResultadoConversao:
    folha: Folha
    destino: Path
    razao_social: str
    linhas: list[LinhaPlanilha]
    ignorados: dict[tuple[int, str], int]
    avisos: list[str] = field(default_factory=list)


def converter_folha(
    folha: Folha,
    config: Config,
    destino: Path,
    modelo: Path | None = None,
) -> ResultadoConversao:
    modelo = Path(config.dados.get("modelo") or modelo or MODELO_PADRAO)
    resultado = montar_linhas(folha, config)
    planilha = Planilha(modelo)

    avisos = folha.avisos()
    colunas_sem_lugar = {
        codigo
        for linha in resultado.linhas
        for codigo in linha.valores
        if codigo not in planilha.colunas
    }
    for codigo in sorted(colunas_sem_lugar):
        avisos.append(f"A configuração usa a coluna {codigo}, que não existe no modelo; valores não gravados.")

    razao = config.razao_social(folha)
    planilha.preencher(
        empresa_codigo=folha.empresa_codigo,
        razao_social=razao,
        inscricao=folha.inscricao,
        competencia=folha.competencia,
        linhas=resultado.linhas,
    )
    destino.parent.mkdir(parents=True, exist_ok=True)
    planilha.salvar(destino)
    return ResultadoConversao(folha, destino, razao, resultado.linhas, resultado.ignorados, avisos)


def converter_pdf(pdf: Path, config: Config, pasta_destino: Path | None = None) -> list[ResultadoConversao]:
    """Gera uma planilha por empresa encontrada no PDF."""
    folhas = ler_previa(pdf)
    if not folhas:
        raise ValueError("Nenhuma empresa encontrada. O arquivo é uma 'Relação de Cálculo' (Prévia)?")
    pasta = pasta_destino or pdf.parent
    return [converter_folha(f, config, pasta / config.nome_arquivo(f)) for f in folhas]
