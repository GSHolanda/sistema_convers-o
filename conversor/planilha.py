"""Gera a planilha preenchida a partir do arquivo modelo (.xlsx).

O .xlsx é editado diretamente no XML: só os valores das células mudam.
Estilos, larguras, comentários de ajuda, configuração de impressão e
fórmulas do modelo são mantidos exatamente como estão.
"""

from __future__ import annotations

import copy
import posixpath
import re
import zipfile
from datetime import date
from decimal import Decimal
from pathlib import Path

from lxml import etree

from .mapeamento import LinhaPlanilha, normalizar

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": NS_MAIN}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
DECLARACAO_XML = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
TIPO_CALCCHAIN ="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"

RE_REF = re.compile(r"^([A-Z]+)(\d+)$")
RE_INTERVALO = re.compile(r"(\$?[A-Z]+\$?)(\d+):(\$?[A-Z]+\$?)(\d+)")


def _m(tag: str) -> str:
    return f"{{{NS_MAIN}}}{tag}"


def coluna_para_indice(coluna: str) -> int:
    n = 0
    for ch in coluna:
        n = n * 26 + (ord(ch) - 64)
    return n


def dividir_ref(ref: str) -> tuple[str, int]:
    m = RE_REF.match(ref)
    if not m:
        raise ValueError(f"Referência de célula inválida: {ref}")
    return m.group(1), int(m.group(2))


def numero_excel(valor: Decimal | int) -> str:
    if isinstance(valor, int):
        return str(valor)
    texto = format(valor.normalize(), "f")
    return "0" if texto in ("-0", "") else texto


def data_excel(d: date) -> int:
    return (d - date(1899, 12, 30)).days


class ErroModelo(Exception):
    pass


class _Pacote:
    """Conteúdo do .xlsx (zip) em memória."""

    def __init__(self, caminho: Path):
        with zipfile.ZipFile(caminho) as z:
            self.infos = z.infolist()
            self.dados = {i.filename: z.read(i.filename) for i in self.infos}

    def xml(self, nome: str) -> etree._Element:
        return etree.fromstring(self.dados[nome])

    def definir_xml(self, nome: str, raiz: etree._Element) -> None:
        self.dados[nome] = DECLARACAO_XML + etree.tostring(raiz, encoding="UTF-8")

    def remover(self, nome: str) -> None:
        self.dados.pop(nome, None)
        self.infos = [i for i in self.infos if i.filename != nome]

    def salvar(self, destino: Path) -> None:
        with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
            for info in self.infos:
                nova = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                nova.compress_type = zipfile.ZIP_DEFLATED
                nova.external_attr = info.external_attr
                z.writestr(nova, self.dados[info.filename])


def _alvo(base: str, destino: str) -> str:
    if destino.startswith("/"):
        return destino.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base), destino))


def _rels_de(nome: str) -> str:
    pasta, arquivo = posixpath.split(nome)
    return posixpath.join(pasta, "_rels", arquivo + ".rels")


class Planilha:
    """Planilha modelo aberta para preenchimento."""

    def __init__(self, caminho_modelo: str | Path):
        self.pacote = _Pacote(Path(caminho_modelo))
        self.caminho_workbook = "xl/workbook.xml"
        rels_wb = self.pacote.xml(_rels_de(self.caminho_workbook))
        self.workbook = self.pacote.xml(self.caminho_workbook)

        primeira = self.workbook.find("m:sheets/m:sheet", NS)
        if primeira is None:
            raise ErroModelo("O modelo não tem planilhas.")
        rid = primeira.get(f"{{{NS_REL}}}id")
        self.caminho_sst = None
        self.caminho_planilha = None
        for rel in rels_wb:
            destino = _alvo(self.caminho_workbook, rel.get("Target"))
            if rel.get("Id") == rid:
                self.caminho_planilha = destino
            if rel.get("Type", "").endswith("/sharedStrings"):
                self.caminho_sst = destino
        if self.caminho_planilha is None:
            raise ErroModelo("Não foi possível localizar a planilha dentro do modelo.")

        self.planilha = self.pacote.xml(self.caminho_planilha)
        self.sheet_data = self.planilha.find("m:sheetData", NS)
        self.sst_original: list[etree._Element] = []
        if self.caminho_sst and self.caminho_sst in self.pacote.dados:
            self.sst_raiz = self.pacote.xml(self.caminho_sst)
            self.sst_original = list(self.sst_raiz.findall("m:si", NS))
        else:
            self.sst_raiz = None

        # Troca índices de texto compartilhado pelo próprio texto (resolvido ao salvar).
        for c in self.sheet_data.iter(_m("c")):
            if c.get("t") == "s":
                v = c.find("m:v", NS)
                if v is not None and v.text is not None:
                    v.text = self._texto_si(self.sst_original[int(v.text)])

        self._localizar_estrutura()

    # ------------------------------------------------------------------ leitura
    @staticmethod
    def _texto_si(si: etree._Element) -> str:
        return "".join(t.text or "" for t in si.iter(_m("t")))

    def _linhas(self) -> dict[int, etree._Element]:
        return {int(r.get("r")): r for r in self.sheet_data.findall("m:row", NS)}

    def _celula(self, ref: str, criar: bool = False) -> etree._Element | None:
        coluna, numero = dividir_ref(ref)
        linha = self._linhas().get(numero)
        if linha is None:
            if not criar:
                return None
            linha = self._criar_linha(numero)
        for c in linha.findall("m:c", NS):
            if c.get("r") == ref:
                return c
        if not criar:
            return None
        nova = etree.Element(_m("c"), r=ref)
        indice = coluna_para_indice(coluna)
        for c in linha.findall("m:c", NS):
            if coluna_para_indice(dividir_ref(c.get("r"))[0]) > indice:
                c.addprevious(nova)
                break
        else:
            linha.append(nova)
        return nova

    def _criar_linha(self, numero: int) -> etree._Element:
        nova = etree.Element(_m("row"), r=str(numero))
        for linha in self.sheet_data.findall("m:row", NS):
            if int(linha.get("r")) > numero:
                linha.addprevious(nova)
                break
        else:
            self.sheet_data.append(nova)
        return nova

    def _valor(self, c: etree._Element | None) -> str | None:
        if c is None:
            return None
        v = c.find("m:v", NS)
        if v is not None:
            return v.text
        t = c.find("m:is/m:t", NS)
        return t.text if t is not None else None

    def _localizar_estrutura(self) -> None:
        """Encontra cabeçalho, linha de códigos, área de dados e linha de total."""
        self.ref_cabecalho: dict[str, str] = {}
        rotulos = {
            "CODIGO EMPRESA": "empresa",
            "RAZAO SOCIAL": "razao",
            "INSCRICAO": "inscricao",
            "COMPETENCIA": "competencia",
        }
        self.linha_codigos = None
        self.linha_total = None
        for numero, linha in sorted(self._linhas().items()):
            celulas = {dividir_ref(c.get("r"))[0]: c for c in linha.findall("m:c", NS)}
            texto_a = normalizar(self._valor(celulas.get("A")) or "")
            for rotulo, chave in rotulos.items():
                if texto_a.startswith(rotulo) and chave not in self.ref_cabecalho:
                    self.ref_cabecalho[chave] = f"C{numero}"
            if texto_a == "CALCULO" and self.linha_codigos is None:
                self.linha_codigos = numero
            if texto_a == "TOTAL":
                self.linha_total = numero
        if self.linha_codigos is None or self.linha_total is None:
            raise ErroModelo("Modelo inválido: não achei a linha de códigos ('Calculo') ou a linha 'TOTAL'.")
        faltando = set(rotulos.values()) - set(self.ref_cabecalho)
        if faltando:
            raise ErroModelo(f"Modelo inválido: cabeçalho incompleto ({', '.join(sorted(faltando))}).")

        self.primeira_linha = self.linha_codigos + 1
        self.ultima_linha = self.linha_total - 1

        # Código de cada coluna de valores (linha 10 do modelo).
        self.colunas: dict[str, str] = {}
        linha = self._linhas()[self.linha_codigos]
        for c in linha.findall("m:c", NS):
            coluna, _ = dividir_ref(c.get("r"))
            valor = self._valor(c)
            if c.get("t") in (None, "n") and valor not in (None, ""):
                self.colunas[str(int(float(valor)))] = coluna

        # Estilos de uma linha preenchida (primeira linha de dados do modelo).
        modelo = self._linhas().get(self.primeira_linha)
        self.atributos_linha = dict(modelo.attrib) if modelo is not None else {}
        self.estilos_linha = {
            dividir_ref(c.get("r"))[0]: c.get("s")
            for c in (modelo.findall("m:c", NS) if modelo is not None else [])
            if c.get("s") is not None
        }

    @property
    def capacidade(self) -> int:
        return self.ultima_linha - self.primeira_linha + 1

    # ------------------------------------------------------------------ escrita
    @staticmethod
    def _limpar(c: etree._Element) -> None:
        for filho in list(c):
            if filho.tag in (_m("v"), _m("is"), _m("f")):
                c.remove(filho)
        c.attrib.pop("t", None)

    def _escrever_numero(self, ref: str, valor: Decimal | int) -> None:
        c = self._celula(ref, criar=True)
        self._limpar(c)
        etree.SubElement(c, _m("v")).text = numero_excel(valor)

    def _escrever_texto(self, ref: str, texto: str) -> None:
        c = self._celula(ref, criar=True)
        self._limpar(c)
        c.set("t", "s")
        etree.SubElement(c, _m("v")).text = texto

    def _inserir_linhas(self, quantidade: int) -> None:
        """Abre espaço para mais funcionários do que as linhas do modelo."""
        antigo_total = self.linha_total
        for linha in sorted(self._linhas().values(), key=lambda r: -int(r.get("r"))):
            numero = int(linha.get("r"))
            if numero < antigo_total:
                continue
            linha.set("r", str(numero + quantidade))
            for c in linha.findall("m:c", NS):
                coluna, _ = dividir_ref(c.get("r"))
                c.set("r", f"{coluna}{numero + quantidade}")
                f = c.find("m:f", NS)
                if f is not None:
                    if f.text:
                        f.text = RE_INTERVALO.sub(self._ajustar_intervalo(quantidade), f.text)
                    if f.get("ref"):
                        f.set("ref", re.sub(r"\d+", str(numero + quantidade), f.get("ref")))
        self.linha_total += quantidade
        self.ultima_linha += quantidade

        dimensao = self.planilha.find("m:dimension", NS)
        if dimensao is not None and ":" in dimensao.get("ref", ""):
            inicio, fim = dimensao.get("ref").split(":")
            coluna, numero = dividir_ref(fim)
            dimensao.set("ref", f"{inicio}:{coluna}{numero + quantidade}")
        self._remover_calcchain()

    def _ajustar_intervalo(self, quantidade: int):
        ultima = self.ultima_linha

        def trocar(m: re.Match) -> str:
            fim = int(m.group(4))
            if fim == ultima:
                fim += quantidade
            return f"{m.group(1)}{m.group(2)}:{m.group(3)}{fim}"

        return trocar

    def _remover_calcchain(self) -> None:
        caminho_rels = _rels_de(self.caminho_workbook)
        rels = self.pacote.xml(caminho_rels)
        for rel in list(rels):
            if rel.get("Type") == TIPO_CALCCHAIN:
                self.pacote.remover(_alvo(self.caminho_workbook, rel.get("Target")))
                rels.remove(rel)
        self.pacote.definir_xml(caminho_rels, rels)
        tipos = self.pacote.xml("[Content_Types].xml")
        for o in list(tipos):
            if o.get("PartName", "").endswith("calcChain.xml"):
                tipos.remove(o)
        self.pacote.definir_xml("[Content_Types].xml", tipos)

    def preencher(
        self,
        empresa_codigo: int,
        razao_social: str,
        inscricao: str,
        competencia: date | None,
        linhas: list[LinhaPlanilha],
    ) -> None:
        # Cabeçalho
        self._escrever_numero(self.ref_cabecalho["empresa"], empresa_codigo)
        self._escrever_texto(self.ref_cabecalho["razao"], razao_social)
        if inscricao.isdigit():
            self._escrever_numero(self.ref_cabecalho["inscricao"], int(inscricao))
        else:
            self._escrever_texto(self.ref_cabecalho["inscricao"], inscricao)
        if competencia:
            self._escrever_numero(self.ref_cabecalho["competencia"], data_excel(competencia))

        if len(linhas) > self.capacidade:
            self._inserir_linhas(len(linhas) - self.capacidade)
        self._limpar_area_dados()

        # Uma linha por funcionário, com a formatação de linha preenchida do modelo.
        for i, item in enumerate(linhas):
            numero = self.primeira_linha + i
            linha = self._linhas().get(numero)
            if linha is None:
                linha = self._criar_linha(numero)
            for chave, valor in self.atributos_linha.items():
                if chave != "r":
                    linha.set(chave, valor)
            for coluna, estilo in self.estilos_linha.items():
                self._celula(f"{coluna}{numero}", criar=True).set("s", estilo)
            self._escrever_numero(f"A{numero}", item.tipo_calculo)
            self._escrever_numero(f"B{numero}", item.codigo)
            self._escrever_texto(f"C{numero}", item.nome)
            for codigo_coluna, valor in item.valores.items():
                coluna = self.colunas.get(codigo_coluna)
                if coluna and valor != 0:
                    self._escrever_numero(f"{coluna}{numero}", valor)

        self._atualizar_totais(linhas)
        self._abrir_no_topo()

    def _limpar_area_dados(self) -> None:
        """Apaga os valores entre a linha de códigos e o TOTAL (mantém os estilos)."""
        for numero in range(self.primeira_linha, self.ultima_linha + 1):
            linha = self._linhas().get(numero)
            if linha is not None:
                for c in linha.findall("m:c", NS):
                    self._limpar(c)

    def limpar_dados(self) -> None:
        """Deixa o modelo sem dados de empresa nem de funcionários."""
        for ref in self.ref_cabecalho.values():
            c = self._celula(ref)
            if c is not None:
                self._limpar(c)
        self._limpar_area_dados()
        self._atualizar_totais([])
        self._abrir_no_topo()

    def _atualizar_totais(self, linhas: list[LinhaPlanilha]) -> None:
        """Grava o resultado das fórmulas da linha TOTAL (o Excel recalcula ao editar)."""
        linha = self._linhas()[self.linha_total]
        mestres: dict[str, str] = {}
        for c in linha.findall("m:c", NS):
            f = c.find("m:f", NS)
            if f is not None and f.get("t") == "shared" and f.text:
                mestres[f.get("si")] = f.text
        codigo_por_coluna = {col: cod for cod, col in self.colunas.items()}
        for c in linha.findall("m:c", NS):
            f = c.find("m:f", NS)
            if f is None:
                continue
            formula = (f.text or mestres.get(f.get("si"), "")).upper()
            coluna, _ = dividir_ref(c.get("r"))
            if formula.startswith("COUNTA"):
                resultado: Decimal | int = len(linhas)
            elif formula.startswith("SUM") and coluna in codigo_por_coluna:
                codigo = codigo_por_coluna[coluna]
                resultado = sum((l.valores.get(codigo, Decimal(0)) for l in linhas), Decimal(0))
            else:
                continue
            v = c.find("m:v", NS)
            if v is None:
                v = etree.SubElement(c, _m("v"))
            v.text = numero_excel(resultado)

    def _abrir_no_topo(self) -> None:
        for view in self.planilha.findall("m:sheetViews/m:sheetView", NS):
            view.attrib.pop("topLeftCell", None)
            for sel in view.findall("m:selection", NS):
                sel.set("activeCell", "A1")
                sel.set("sqref", "A1")

    # ------------------------------------------------------------------ salvar
    def _gravar_textos(self) -> None:
        originais = {}
        for si in self.sst_original:
            originais.setdefault(self._texto_si(si), si)
        unicos: dict[str, int] = {}
        total = 0
        for c in self.sheet_data.iter(_m("c")):
            if c.get("t") != "s":
                continue
            v = c.find("m:v", NS)
            if v is None:
                continue
            texto = v.text or ""
            if texto not in unicos:
                unicos[texto] = len(unicos)
            v.text = str(unicos[texto])
            total += 1

        if self.sst_raiz is None:
            raise ErroModelo("O modelo não tem tabela de textos (sharedStrings).")
        for si in list(self.sst_raiz):
            self.sst_raiz.remove(si)
        for texto in unicos:
            if texto in originais:
                self.sst_raiz.append(copy.deepcopy(originais[texto]))
            else:
                si = etree.SubElement(self.sst_raiz, _m("si"))
                t = etree.SubElement(si, _m("t"))
                t.text = texto
                if texto != texto.strip():
                    t.set(XML_SPACE, "preserve")
        self.sst_raiz.set("count", str(total))
        self.sst_raiz.set("uniqueCount", str(len(unicos)))

    def salvar(self, destino: str | Path) -> None:
        self._gravar_textos()
        self.pacote.definir_xml(self.caminho_planilha, self.planilha)
        self.pacote.definir_xml(self.caminho_sst, self.sst_raiz)
        self.pacote.salvar(Path(destino))
