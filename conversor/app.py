"""Janela do Conversor de Prévia para Planilha Domínio."""

from __future__ import annotations

import os
import sys
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from . import __version__
from .conversao import caminho_config, converter_folha
from .leitor_pdf import Folha, ler_previa
from .mapeamento import Config

TITULO = "Conversor Prévia → Planilha Domínio"


def formatar_valor(valor) -> str:
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def abrir_arquivo(caminho: Path) -> None:
    if sys.platform.startswith("win"):
        os.startfile(str(caminho))  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        os.system(f'open "{caminho}"')
    else:
        os.system(f'xdg-open "{caminho}" >/dev/null 2>&1 &')


class Aplicacao(ttk.Frame):
    def __init__(self, raiz: tk.Tk, arquivos_iniciais: list[str] | None = None):
        super().__init__(raiz, padding=10)
        self.raiz = raiz
        self.caminho_config = caminho_config()
        try:
            self.config = Config.carregar(self.caminho_config)
        except ValueError as erro:
            messagebox.showerror(TITULO, f"{erro}\n\nUsando a configuração padrão.")
            self.config = Config(caminho=None)
        self.itens: dict[str, tuple[Path, Folha]] = {}

        raiz.title(f"{TITULO}  (v{__version__})")
        raiz.minsize(820, 560)
        self.grid(sticky="nsew")
        raiz.columnconfigure(0, weight=1)
        raiz.rowconfigure(0, weight=1)
        self._montar_tela()

        if arquivos_iniciais:
            self.raiz.after(100, lambda: self.adicionar_arquivos(arquivos_iniciais))

    # ------------------------------------------------------------------ tela
    def _montar_tela(self) -> None:
        self.columnconfigure(0, weight=1)

        ttk.Label(
            self,
            text="1. Adicione as Prévias (Relação de Cálculo) em PDF",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, sticky="w")

        botoes = ttk.Frame(self)
        botoes.grid(row=1, column=0, sticky="w", pady=(4, 4))
        ttk.Button(botoes, text="Adicionar PDFs...", command=self.escolher_pdfs).pack(side="left")
        ttk.Button(botoes, text="Remover selecionado", command=self.remover_selecionados).pack(side="left", padx=6)
        ttk.Button(botoes, text="Limpar lista", command=self.limpar_lista).pack(side="left")

        quadro_lista = ttk.Frame(self)
        quadro_lista.grid(row=2, column=0, sticky="nsew")
        quadro_lista.columnconfigure(0, weight=1)
        quadro_lista.rowconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        colunas = ("arquivo", "empresa", "razao", "competencia", "funcionarios", "status")
        self.lista = ttk.Treeview(quadro_lista, columns=colunas, show="headings", height=7)
        for coluna, titulo, largura, ancora in (
            ("arquivo", "Arquivo", 210, "w"),
            ("empresa", "Empresa", 75, "center"),
            ("razao", "Razão social (planilha)", 210, "w"),
            ("competencia", "Competência", 100, "center"),
            ("funcionarios", "Funcionários", 100, "center"),
            ("status", "Situação", 140, "w"),
        ):
            self.lista.heading(coluna, text=titulo)
            self.lista.column(coluna, width=largura, anchor=ancora, stretch=coluna in ("arquivo", "razao", "status"))
        self.lista.grid(row=0, column=0, sticky="nsew")
        barra = ttk.Scrollbar(quadro_lista, orient="vertical", command=self.lista.yview)
        barra.grid(row=0, column=1, sticky="ns")
        self.lista.configure(yscrollcommand=barra.set)
        self.lista.bind("<Double-1>", self._duplo_clique)
        ttk.Label(
            self,
            text="Dê dois cliques em uma linha para alterar a razão social que vai na planilha.",
            foreground="#555",
        ).grid(row=3, column=0, sticky="w", pady=(2, 8))

        ttk.Label(self, text="2. Onde salvar", font=("Segoe UI", 10, "bold")).grid(row=4, column=0, sticky="w")
        destino = ttk.Frame(self)
        destino.grid(row=5, column=0, sticky="ew", pady=(4, 8))
        destino.columnconfigure(2, weight=1)
        self.modo_destino = tk.StringVar(value="mesma")
        self.pasta_destino = tk.StringVar()
        ttk.Radiobutton(
            destino, text="Na mesma pasta de cada PDF", variable=self.modo_destino, value="mesma"
        ).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Radiobutton(destino, text="Nesta pasta:", variable=self.modo_destino, value="outra").grid(
            row=1, column=0, sticky="w"
        )
        ttk.Entry(destino, textvariable=self.pasta_destino).grid(row=1, column=2, sticky="ew", padx=4)
        ttk.Button(destino, text="Escolher...", command=self.escolher_pasta).grid(row=1, column=3)

        acoes = ttk.Frame(self)
        acoes.grid(row=6, column=0, sticky="ew", pady=(0, 8))
        self.abrir_depois = tk.BooleanVar(value=True)
        self.mostrar_ignorados = tk.BooleanVar(value=False)
        ttk.Checkbutton(acoes, text="Abrir a planilha depois de gerar", variable=self.abrir_depois).pack(side="left")
        ttk.Checkbutton(
            acoes, text="Listar eventos do PDF que não vão para a planilha", variable=self.mostrar_ignorados
        ).pack(side="left", padx=(12, 0))
        ttk.Button(acoes, text="Configuração...", command=self.abrir_configuracao).pack(side="right")
        self.botao_gerar = ttk.Button(acoes, text="3. Gerar planilha(s)", command=self.gerar)
        self.botao_gerar.pack(side="right", padx=6)

        quadro_log = ttk.LabelFrame(self, text="Resultado", padding=4)
        quadro_log.grid(row=7, column=0, sticky="nsew")
        quadro_log.columnconfigure(0, weight=1)
        quadro_log.rowconfigure(0, weight=1)
        self.rowconfigure(7, weight=1)
        self.log = tk.Text(quadro_log, height=10, wrap="word", state="disabled", font=("Consolas", 9))
        self.log.grid(row=0, column=0, sticky="nsew")
        barra_log = ttk.Scrollbar(quadro_log, orient="vertical", command=self.log.yview)
        barra_log.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=barra_log.set)
        self.log.tag_configure("ok", foreground="#1a7f37")
        self.log.tag_configure("aviso", foreground="#b35900")
        self.log.tag_configure("erro", foreground="#c62828")
        self.log.tag_configure("info", foreground="#555555")

    def escrever(self, texto: str, estilo: str | None = None) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", texto + "\n", estilo or ())
        self.log.see("end")
        self.log.configure(state="disabled")
        self.log.update_idletasks()

    # ------------------------------------------------------------------ lista
    def escolher_pdfs(self) -> None:
        arquivos = filedialog.askopenfilenames(
            title="Selecione as Prévias em PDF",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
        )
        if arquivos:
            self.adicionar_arquivos(list(arquivos))

    def _recarregar_config(self) -> bool:
        try:
            self.config = Config.carregar(self.caminho_config)
            return True
        except ValueError as erro:
            messagebox.showerror(TITULO, str(erro))
            return False

    def adicionar_arquivos(self, arquivos: list[str]) -> None:
        if not self._recarregar_config():
            return
        self.raiz.configure(cursor="watch")
        try:
            for arquivo in arquivos:
                caminho = Path(arquivo)
                if any(p == caminho for p, _ in self.itens.values()):
                    continue
                self.escrever(f"Lendo {caminho.name}...", "info")
                try:
                    folhas = ler_previa(caminho)
                except Exception as erro:  # noqa: BLE001 - mostrar qualquer falha ao usuário
                    self.escrever(f"✖ {caminho.name}: não foi possível ler o PDF ({erro}).", "erro")
                    continue
                if not folhas:
                    self.escrever(
                        f"✖ {caminho.name}: nenhuma empresa encontrada. "
                        "O arquivo é uma 'Relação de Cálculo' (Prévia)?",
                        "erro",
                    )
                    continue
                for folha in folhas:
                    self._confirmar_razao_social(folha)
                    item = self.lista.insert("", "end")
                    self.itens[item] = (caminho, folha)
                    self._atualizar_item(item, "Pronto para gerar")
                    for aviso in folha.avisos():
                        self.escrever(f"  ⚠ {aviso}", "aviso")
        finally:
            self.raiz.configure(cursor="")

    def _confirmar_razao_social(self, folha: Folha) -> None:
        """Na primeira vez que uma empresa aparece, confirma a razão social da planilha."""
        if self.config.tem_razao_social(folha.empresa_codigo):
            return
        resposta = simpledialog.askstring(
            "Razão social",
            f"Empresa {folha.empresa_codigo:04d} ainda não cadastrada.\n\n"
            f"Razão social no PDF: {folha.razao_social}\n\n"
            "Confirme ou altere a razão social que vai na planilha\n"
            "(fica guardada para as próximas vezes):",
            initialvalue=folha.razao_social,
            parent=self.raiz,
        )
        razao = (resposta or "").strip() or folha.razao_social
        self.config.definir_razao_social(folha.empresa_codigo, razao)
        self._salvar_config()

    def _salvar_config(self) -> None:
        try:
            self.config.salvar()
        except OSError as erro:
            self.escrever(f"⚠ Não foi possível salvar a configuração ({erro}).", "aviso")

    def _atualizar_item(self, item: str, status: str) -> None:
        caminho, folha = self.itens[item]
        competencia = folha.competencia.strftime("%m/%Y") if folha.competencia else "?"
        self.lista.item(
            item,
            values=(
                caminho.name,
                f"{folha.empresa_codigo:04d}",
                self.config.razao_social(folha),
                competencia,
                len(folha.funcionarios),
                status,
            ),
        )

    def _duplo_clique(self, evento) -> None:
        item = self.lista.identify_row(evento.y)
        if not item or not self._recarregar_config():
            return
        _, folha = self.itens[item]
        resposta = simpledialog.askstring(
            "Razão social",
            f"Razão social da empresa {folha.empresa_codigo:04d} na planilha\n"
            f"(no PDF: {folha.razao_social}):",
            initialvalue=self.config.razao_social(folha),
            parent=self.raiz,
        )
        if resposta and resposta.strip():
            self.config.definir_razao_social(folha.empresa_codigo, resposta.strip())
            self._salvar_config()
            for outro, (_, f) in self.itens.items():
                if f.empresa_codigo == folha.empresa_codigo:
                    self._atualizar_item(outro, self.lista.set(outro, "status"))

    def remover_selecionados(self) -> None:
        for item in self.lista.selection():
            self.lista.delete(item)
            self.itens.pop(item, None)

    def limpar_lista(self) -> None:
        for item in list(self.itens):
            self.lista.delete(item)
        self.itens.clear()

    def escolher_pasta(self) -> None:
        pasta = filedialog.askdirectory(title="Pasta onde salvar as planilhas")
        if pasta:
            self.pasta_destino.set(pasta)
            self.modo_destino.set("outra")

    def abrir_configuracao(self) -> None:
        if not self.caminho_config.exists():
            self._salvar_config()
        messagebox.showinfo(
            "Configuração",
            "A configuração (ligação dos eventos do PDF com as colunas da planilha e razões "
            f"sociais) fica no arquivo:\n\n{self.caminho_config}\n\n"
            "Ele será aberto agora. Depois de alterar, salve o arquivo; "
            "as mudanças valem para os próximos PDFs adicionados.",
        )
        abrir_arquivo(self.caminho_config)

    # ------------------------------------------------------------------ gerar
    def _destino(self, pdf: Path, folha: Folha) -> Path | None:
        if self.modo_destino.get() == "outra":
            pasta = Path(self.pasta_destino.get().strip())
        else:
            pasta = pdf.parent
        destino = pasta / self.config.nome_arquivo(folha)
        if not destino.exists():
            return destino
        resposta = messagebox.askyesnocancel(
            "Arquivo já existe",
            f"Já existe o arquivo:\n{destino}\n\nSubstituir?\n\n"
            "Sim = substituir   Não = salvar com outro nome   Cancelar = pular",
        )
        if resposta is None:
            return None
        if resposta:
            return destino
        n = 2
        while True:
            alternativo = destino.with_name(f"{destino.stem} ({n}){destino.suffix}")
            if not alternativo.exists():
                return alternativo
            n += 1

    def gerar(self) -> None:
        if not self.itens:
            messagebox.showwarning(TITULO, "Adicione pelo menos uma Prévia em PDF.")
            return
        if self.modo_destino.get() == "outra" and not self.pasta_destino.get().strip():
            messagebox.showwarning(TITULO, "Escolha a pasta onde salvar as planilhas.")
            return
        if not self._recarregar_config():
            return
        gerados: list[Path] = []
        self.botao_gerar.configure(state="disabled")
        self.raiz.configure(cursor="watch")
        try:
            for item, (pdf, folha) in list(self.itens.items()):
                destino = None
                try:
                    destino = self._destino(pdf, folha)
                    if destino is None:
                        self._atualizar_item(item, "Pulado")
                        continue
                    resultado = converter_folha(folha, self.config, destino)
                except PermissionError:
                    self._atualizar_item(item, "Erro: arquivo aberto?")
                    nome = destino.name if destino else pdf.name
                    self.escrever(
                        f"✖ Não foi possível salvar {nome}. "
                        "Se ele estiver aberto no Excel, feche-o e tente de novo.",
                        "erro",
                    )
                    continue
                except Exception as erro:  # noqa: BLE001
                    self._atualizar_item(item, "Erro")
                    self.escrever(f"✖ {pdf.name}: {erro}", "erro")
                    self.escrever(traceback.format_exc(), "info")
                    continue

                gerados.append(resultado.destino)
                self._atualizar_item(item, "Gerada")
                self._resumo(resultado)
        finally:
            self.botao_gerar.configure(state="normal")
            self.raiz.configure(cursor="")

        if gerados and self.abrir_depois.get():
            for caminho in gerados[:5]:
                abrir_arquivo(caminho)

    def _resumo(self, resultado) -> None:
        folha = resultado.folha
        comp = folha.competencia.strftime("%m/%Y") if folha.competencia else "?"
        self.escrever(
            f"✔ {resultado.destino.name}  —  empresa {folha.empresa_codigo:04d}, competência {comp}, "
            f"{len(resultado.linhas)} funcionário(s)",
            "ok",
        )
        self.escrever(f"   Salva em: {resultado.destino.parent}", "info")
        totais: dict[str, object] = {}
        for linha in resultado.linhas:
            for codigo, valor in linha.valores.items():
                totais[codigo] = totais.get(codigo, 0) + valor
        colunas = self.config.dados.get("colunas", {})
        partes = [
            f"{regra.get('titulo', codigo)}: {formatar_valor(totais[codigo])}"
            for codigo, regra in colunas.items()
            if codigo in totais
        ]
        if partes:
            self.escrever("   Totais: " + " | ".join(partes), "info")
        for aviso in resultado.avisos:
            self.escrever(f"   ⚠ {aviso}", "aviso")
        if resultado.ignorados and self.mostrar_ignorados.get():
            nomes = ", ".join(f"{cod} {desc}" for cod, desc in sorted(resultado.ignorados))
            self.escrever(f"   Eventos do PDF que não vão para a planilha: {nomes}", "info")


def main(argv: list[str] | None = None) -> None:
    argv = sys.argv[1:] if argv is None else argv
    raiz = tk.Tk()
    if sys.platform.startswith("win"):
        try:
            ttk.Style(raiz).theme_use("vista")
        except tk.TclError:
            pass
    Aplicacao(raiz, [a for a in argv if a.lower().endswith(".pdf")])
    raiz.mainloop()


if __name__ == "__main__":
    main()
