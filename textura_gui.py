#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textura_gui.py — interface para textura_near.py
===============================================

Coloque na MESMA PASTA que textura_near.py e textura_stats.py, e execute:

    python textura_gui.py

Separador 1 · Matriz    carrega o livro, lista folhas, pré-visualiza, mapeia colunas
Separador 2 · Termos    tabela de termos de busca, vazia à partida
Separador 3 · Análise   parâmetros, execução e registo

Requer tkinter (o Python oficial para Windows e macOS já o inclui;
em Linux: sudo apt install python3-tk) e openpyxl.
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import tempfile
import threading
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

AQUI = Path(__file__).resolve().parent
MOTOR = AQUI / "textura_near.py"
N_PREVIA = 40
PISTAS_NO = ("textur", "gestur", "harmon", "rhythm", "timbr")

AJUDA_PADROES = (
    "Padrões separados por vírgula.   *  trunca à direita (uniform* → uniform, "
    "uniformity, uniformly)   ·   *varying trunca à esquerda   ·   "
    "espaços = expressão de várias palavras (not uniform)"
)


class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Mineração de co-ocorrências — NEAR/x")
        self.geometry("1120x820")
        self.minsize(980, 720)

        self.proc: subprocess.Popen | None = None
        self.fila: queue.Queue[str] = queue.Queue()
        self.previa: list[list[str]] = []

        self.v_xlsx = tk.StringVar()
        self.v_folha = tk.StringVar()
        self.v_dim = tk.StringVar(value="nenhum livro carregado")
        self.v_cab = tk.BooleanVar(value=False)
        self.v_col_no = tk.IntVar(value=6)
        self.v_col_ctx = tk.IntVar(value=15)
        self.v_col_src = tk.IntVar(value=12)
        self.v_col_url = tk.IntVar(value=13)

        self.v_etq = tk.StringVar()
        self.v_polo = tk.StringVar(value="—")
        self.v_pads = tk.StringVar()
        self.v_n_termos = tk.StringVar(value="0 termos definidos")

        self.v_saida = tk.StringVar(value=str(Path.home() / "resultado_near.xlsx"))
        self.v_near = tk.IntVar(value=4)
        self.v_banda = tk.IntVar(value=12)
        self.v_lingua = tk.StringVar(value="en")
        self.v_limite = tk.StringVar(value="")
        self.v_fronteira = tk.BooleanVar(value=True)
        self.v_sintaxe = tk.StringVar(value="heuristica")
        self.v_modelo = tk.StringVar(value="en_core_web_sm")
        self.v_consulta = tk.StringVar(value="")
        self.v_estado = tk.StringVar(value="pronto")

        self._constroi()
        self.after(120, self._drena)

    def _constroi(self):
        nb = ttk.Notebook(self, padding=8)
        nb.pack(fill="both", expand=True)
        nb.add(self._aba_matriz(nb), text="  1 · Matriz  ")
        nb.add(self._aba_termos(nb), text="  2 · Termos de busca  ")
        nb.add(self._aba_analise(nb), text="  3 · Análise  ")

    # ======================================================== 1 · MATRIZ ===
    def _aba_matriz(self, pai):
        f = ttk.Frame(pai, padding=10)
        pad = dict(padx=6, pady=4)

        cx = ttk.LabelFrame(f, text="Livro de trabalho", padding=8)
        cx.pack(fill="x")
        cx.columnconfigure(1, weight=1)
        ttk.Label(cx, text="Ficheiro").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(cx, textvariable=self.v_xlsx).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(cx, text="Abrir…", command=self._abre_livro).grid(row=0, column=2, **pad)
        ttk.Label(cx, text="Folha").grid(row=1, column=0, sticky="w", **pad)
        self.cb_folha = ttk.Combobox(cx, textvariable=self.v_folha,
                                     state="readonly", width=34)
        self.cb_folha.grid(row=1, column=1, sticky="w", **pad)
        self.cb_folha.bind("<<ComboboxSelected>>", lambda _e: self._carrega_previa())
        ttk.Checkbutton(cx, text="a folha tem linha de cabeçalho", variable=self.v_cab,
                        command=self._carrega_previa).grid(row=1, column=2,
                                                           sticky="w", **pad)
        ttk.Label(cx, textvariable=self.v_dim, foreground="#555").grid(
            row=2, column=1, sticky="w", padx=6)

        cm = ttk.LabelFrame(f, text="Mapeamento de colunas", padding=8)
        cm.pack(fill="x", pady=(10, 0))
        campos = (("Nó", self.v_col_no), ("Contexto", self.v_col_ctx),
                  ("Fonte", self.v_col_src),
                  ("Hiperligação (0 = nenhuma)", self.v_col_url))
        for k, (rot, var) in enumerate(campos):
            ttk.Label(cm, text=rot).grid(row=0, column=2 * k, sticky="e", **pad)
            ttk.Spinbox(cm, from_=0, to=200, width=5, textvariable=var).grid(
                row=0, column=2 * k + 1, sticky="w", **pad)
        ttk.Button(cm, text="Detectar", command=self._detecta).grid(row=0, column=8, **pad)
        ttk.Label(cm, foreground="#666", wraplength=1000, justify="left",
                  text="O contexto é a única coluna com pontuação e, por isso, a única "
                       "que permite excluir pares separados por fronteira de frase. "
                       "A coluna de hiperligação é preservada: na folha de concordância, "
                       "cada linha fica clicável para o texto de origem."
                  ).grid(row=1, column=0, columnspan=9, sticky="w", padx=6, pady=(2, 0))

        cp = ttk.LabelFrame(f, text=f"Pré-visualização ({N_PREVIA} primeiras linhas)",
                            padding=8)
        cp.pack(fill="both", expand=True, pady=(10, 0))
        self.tree = ttk.Treeview(cp, show="headings", height=13)
        sx = ttk.Scrollbar(cp, orient="horizontal", command=self.tree.xview)
        sy = ttk.Scrollbar(cp, orient="vertical", command=self.tree.yview)
        self.tree.configure(xscrollcommand=sx.set, yscrollcommand=sy.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        cp.rowconfigure(0, weight=1)
        cp.columnconfigure(0, weight=1)
        return f

    # ======================================================== 2 · TERMOS ===
    def _aba_termos(self, pai):
        f = ttk.Frame(pai, padding=10)
        pad = dict(padx=5, pady=4)

        ce = ttk.LabelFrame(f, text="Novo termo", padding=8)
        ce.pack(fill="x")
        ce.columnconfigure(5, weight=1)
        ttk.Label(ce, text="Etiqueta").grid(row=0, column=0, sticky="e", **pad)
        ttk.Entry(ce, textvariable=self.v_etq, width=22).grid(row=0, column=1,
                                                              sticky="w", **pad)
        ttk.Label(ce, text="Pólo").grid(row=0, column=2, sticky="e", **pad)
        ttk.Combobox(ce, textvariable=self.v_polo, width=18, state="readonly",
                     values=["—", "E (estabilidade)", "V (variabilidade)"]).grid(
            row=0, column=3, sticky="w", **pad)
        ttk.Label(ce, text="Padrões").grid(row=1, column=0, sticky="e", **pad)
        ent = ttk.Entry(ce, textvariable=self.v_pads)
        ent.grid(row=1, column=1, columnspan=5, sticky="ew", **pad)
        ent.bind("<Return>", lambda _e: self._adiciona())
        ttk.Label(ce, foreground="#666", wraplength=1000, justify="left",
                  text=AJUDA_PADROES).grid(row=2, column=1, columnspan=5,
                                           sticky="w", padx=5)

        bb = ttk.Frame(ce)
        bb.grid(row=3, column=1, columnspan=5, sticky="w", pady=(6, 0))
        for rot, cmd in (("Adicionar", self._adiciona),
                         ("Actualizar seleccionado", self._actualiza),
                         ("Remover seleccionado", self._remove),
                         ("Limpar tudo", self._limpa)):
            ttk.Button(bb, text=rot, command=cmd).pack(side="left", padx=3)
        ttk.Separator(bb, orient="vertical").pack(side="left", fill="y", padx=10)
        ttk.Button(bb, text="Importar…", command=self._importa).pack(side="left", padx=3)
        ttk.Button(bb, text="Exportar…", command=self._exporta).pack(side="left", padx=3)

        ct = ttk.LabelFrame(f, text="Termos de busca", padding=8)
        ct.pack(fill="both", expand=True, pady=(10, 0))
        cols = ("etiqueta", "polo", "padroes")
        self.tab = ttk.Treeview(ct, columns=cols, show="headings", height=16)
        for c, larg, rot in zip(cols, (200, 140, 640),
                                ("Etiqueta", "Pólo", "Padrões")):
            self.tab.heading(c, text=rot)
            self.tab.column(c, width=larg, anchor="w")
        sy = ttk.Scrollbar(ct, orient="vertical", command=self.tab.yview)
        self.tab.configure(yscrollcommand=sy.set)
        self.tab.pack(side="left", fill="both", expand=True)
        sy.pack(side="right", fill="y")
        self.tab.bind("<Double-1>", self._carrega_linha)
        self.tab.bind("<Delete>", lambda _e: self._remove())

        ttk.Label(f, textvariable=self.v_n_termos, foreground="#444").pack(
            anchor="w", pady=(6, 0))
        return f

    @staticmethod
    def _sigla(texto: str) -> str:
        return texto[0] if texto and texto[0] in "EV" else ""

    def _adiciona(self):
        etq, pads = self.v_etq.get().strip(), self.v_pads.get().strip()
        if not etq or not pads:
            messagebox.showwarning("Termo", "Indique etiqueta e pelo menos um padrão.")
            return
        if any(c in etq for c in " =:,"):
            messagebox.showwarning(
                "Etiqueta", "A etiqueta não pode conter espaços, '=', ':' ou ','.")
            return
        for it in self.tab.get_children():
            if self.tab.item(it, "values")[0] == etq:
                messagebox.showwarning("Etiqueta", f"'{etq}' já existe.")
                return
        self.tab.insert("", "end", values=(etq, self.v_polo.get(), pads))
        self.v_etq.set(""); self.v_pads.set(""); self.v_polo.set("—")
        self._conta()

    def _actualiza(self):
        sel = self.tab.selection()
        if sel:
            self.tab.item(sel[0], values=(self.v_etq.get().strip(),
                                          self.v_polo.get(),
                                          self.v_pads.get().strip()))
            self._conta()

    def _remove(self):
        for it in self.tab.selection():
            self.tab.delete(it)
        self._conta()

    def _limpa(self):
        if self.tab.get_children() and messagebox.askyesno(
                "Limpar", "Remover todos os termos?"):
            self.tab.delete(*self.tab.get_children())
            self._conta()

    def _carrega_linha(self, _e=None):
        sel = self.tab.selection()
        if sel:
            e, p, d = self.tab.item(sel[0], "values")
            self.v_etq.set(e); self.v_polo.set(p); self.v_pads.set(d)

    def _conta(self):
        n = len(self.tab.get_children())
        s = "" if n == 1 else "s"
        self.v_n_termos.set(f"{n} termo{s} definido{s}")

    def _importa(self):
        c = filedialog.askopenfilename(filetypes=[("Texto", "*.txt"), ("Todos", "*.*")])
        if not c:
            return
        for linha in Path(c).read_text(encoding="utf-8").splitlines():
            linha = linha.split("#", 1)[0].strip()
            if not linha or "=" not in linha:
                continue
            esq, pads = linha.split("=", 1)
            if ":" in esq:
                etq, polo = (x.strip() for x in esq.split(":", 1))
                polo = ("E (estabilidade)" if polo.upper().startswith("E")
                        else "V (variabilidade)" if polo.upper().startswith("V")
                        else "—")
            else:
                etq, polo = esq.strip(), "—"
            self.tab.insert("", "end", values=(etq, polo, pads.strip()))
        self._conta()

    def _exporta(self):
        c = filedialog.asksaveasfilename(defaultextension=".txt",
                                         filetypes=[("Texto", "*.txt")])
        if c:
            Path(c).write_text(self._serializa(), encoding="utf-8")
            self._log(f"termos exportados para {c}")

    def _serializa(self) -> str:
        linhas = ["# gerado por textura_gui.py"]
        for it in self.tab.get_children():
            e, p, d = self.tab.item(it, "values")
            s = self._sigla(p)
            linhas.append(f"{e}{' : ' + s if s else ''} = {d}")
        return "\n".join(linhas) + "\n"

    # ======================================================= 3 · ANÁLISE ===
    def _aba_analise(self, pai):
        f = ttk.Frame(pai, padding=10)
        pad = dict(padx=6, pady=4)

        cp = ttk.LabelFrame(f, text="Parâmetros", padding=8)
        cp.pack(fill="x")
        cp.columnconfigure(1, weight=1)

        l1 = ttk.Frame(cp); l1.grid(row=0, column=0, columnspan=6, sticky="w")
        ttk.Label(l1, text="Janela NEAR/").pack(side="left", padx=(6, 2))
        ttk.Spinbox(l1, from_=1, to=25, width=4, textvariable=self.v_near).pack(side="left")
        ttk.Label(l1, text="   Banda de referência").pack(side="left", padx=(14, 2))
        ttk.Spinbox(l1, from_=2, to=60, width=4, textvariable=self.v_banda).pack(side="left")
        ttk.Label(l1, text="   Língua do nó").pack(side="left", padx=(14, 2))
        ttk.Combobox(l1, textvariable=self.v_lingua, width=7, state="readonly",
                     values=["en", "pt", "de", "todas"]).pack(side="left")
        ttk.Label(l1, text="   Limite de linhas").pack(side="left", padx=(14, 2))
        ttk.Entry(l1, textvariable=self.v_limite, width=9).pack(side="left")

        ttk.Checkbutton(
            cp, variable=self.v_fronteira,
            text="Excluir co-ocorrências separadas por fronteira de frase "
                 "(ponto, exclamação, interrogação, ponto-e-vírgula)"
        ).grid(row=1, column=0, columnspan=6, sticky="w", **pad)

        l2 = ttk.Frame(cp); l2.grid(row=2, column=0, columnspan=6, sticky="w")
        ttk.Label(l2, text="Relações sintácticas").pack(side="left", padx=(6, 2))
        ttk.Combobox(l2, textvariable=self.v_sintaxe, width=13, state="readonly",
                     values=["heuristica", "spacy"]).pack(side="left")
        ttk.Label(l2, text="   Modelo").pack(side="left", padx=(12, 2))
        ttk.Entry(l2, textvariable=self.v_modelo, width=20).pack(side="left")
        ttk.Label(l2, foreground="#666",
                  text="   spaCy identifica o governante e separa atribuições "
                       "genuínas de incidentais").pack(side="left")

        ttk.Label(cp, text="Consulta booleana").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(cp, textvariable=self.v_consulta).grid(
            row=3, column=1, columnspan=5, sticky="ew", **pad)
        ttk.Label(cp, foreground="#666",
                  text="usa as etiquetas do separador 2:   "
                       "(uniform OR constant) AND NOT varied"
                  ).grid(row=4, column=1, columnspan=5, sticky="w", padx=6)

        ttk.Label(cp, text="Guardar resultado em").grid(row=5, column=0, sticky="w", **pad)
        ttk.Entry(cp, textvariable=self.v_saida).grid(
            row=5, column=1, columnspan=4, sticky="ew", **pad)
        ttk.Button(cp, text="Escolher…", command=self._escolhe_saida).grid(
            row=5, column=5, **pad)

        ca = ttk.Frame(f); ca.pack(fill="x", pady=(10, 0))
        self.btn_corre = ttk.Button(ca, text="Executar análise", command=self._corre)
        self.btn_corre.pack(side="left", padx=4)
        self.btn_para = ttk.Button(ca, text="Interromper", command=self._para,
                                   state="disabled")
        self.btn_para.pack(side="left", padx=4)
        self.btn_abre = ttk.Button(ca, text="Abrir resultado", command=self._abre_saida,
                                   state="disabled")
        self.btn_abre.pack(side="left", padx=4)
        self.barra = ttk.Progressbar(ca, mode="indeterminate", length=180)
        self.barra.pack(side="left", padx=12)
        ttk.Label(ca, textvariable=self.v_estado, foreground="#444").pack(side="left")

        cl = ttk.LabelFrame(f, text="Registo de execução", padding=8)
        cl.pack(fill="both", expand=True, pady=(10, 0))
        mono = ("Consolas", 9) if os.name == "nt" else ("Menlo", 10)
        self.txt_log = tk.Text(cl, wrap="word", state="disabled",
                               background="#f7f7f5", font=mono)
        sl = ttk.Scrollbar(cl, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=sl.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        sl.pack(side="right", fill="y")
        return f

    # ---------------------------------------------------------- livro/prévia
    def _abre_livro(self):
        c = filedialog.askopenfilename(
            title="Matriz KWIC",
            filetypes=[("Livros Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if not c:
            return
        self.v_xlsx.set(c)
        self.v_saida.set(str(Path(c).with_name("resultado_near.xlsx")))
        try:
            from openpyxl import load_workbook
            wb = load_workbook(c, read_only=True)
            folhas = wb.sheetnames
            wb.close()
        except Exception as exc:                              # noqa: BLE001
            messagebox.showerror("Abrir", f"Não foi possível ler o livro:\n{exc}")
            return
        self.cb_folha["values"] = folhas
        self.v_folha.set(folhas[0])
        self._carrega_previa()

    def _carrega_previa(self):
        cam, folha = self.v_xlsx.get(), self.v_folha.get()
        if not (cam and folha):
            return
        self.v_dim.set("a ler…")
        self.update_idletasks()
        try:
            from openpyxl import load_workbook
            wb = load_workbook(cam, read_only=True)
            ws = wb[folha]
            linhas = []
            for k, row in enumerate(ws.iter_rows(values_only=True)):
                linhas.append(["" if v is None else str(v) for v in row])
                if k + 1 >= N_PREVIA:
                    break
            n_lin, n_col = ws.max_row, ws.max_column
            wb.close()
        except Exception as exc:                              # noqa: BLE001
            messagebox.showerror("Pré-visualização", str(exc))
            return

        self.previa = linhas
        self.v_dim.set(f"{n_lin:,} linhas × {n_col} colunas".replace(",", " "))
        cabec = (linhas[0] if self.v_cab.get() and linhas
                 else [f"col {i+1}" for i in range(n_col)])
        corpo = linhas[1:] if self.v_cab.get() else linhas

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = [f"c{i}" for i in range(n_col)]
        for i in range(n_col):
            rot = cabec[i] if i < len(cabec) else f"col {i+1}"
            self.tree.heading(f"c{i}", text=f"{i+1}· {rot}"[:38])
            self.tree.column(f"c{i}", width=130, stretch=False, anchor="w")
        for r in corpo:
            self.tree.insert("", "end",
                             values=[(r[i] if i < len(r) else "")[:120]
                                     for i in range(n_col)])
        self._detecta()

    def _detecta(self):
        if not self.previa:
            return
        corpo = self.previa[1:] if self.v_cab.get() else self.previa
        if not corpo:
            return
        n_col = max(len(r) for r in corpo)

        def col(i):
            return [r[i] for r in corpo if i < len(r) and r[i]]

        m_no = m_ctx = m_src = m_url = None
        p_no = p_src = p_url = 0.0
        comp = -1.0
        for i in range(n_col):
            vals = col(i)
            if not vals:
                continue
            baixo = [v.lower() for v in vals]
            p = sum(any(v.startswith(x) for x in PISTAS_NO) for v in baixo) / len(baixo)
            if p > p_no:
                m_no, p_no = i, p
            u = sum(v.startswith(("file:", "http")) for v in baixo) / len(baixo)
            if u > p_url:
                m_url, p_url = i, u
            c = sum(len(v) for v in vals) / len(vals)
            if u < .5 and c > comp:
                m_ctx, comp = i, c
            s = sum(("\\" in v or "/" in v) for v in vals) / len(vals)
            if u < .5 and s > p_src and i != m_ctx:
                m_src, p_src = i, s

        if m_no is not None and p_no > .5:
            self.v_col_no.set(m_no + 1)
        if m_ctx is not None:
            self.v_col_ctx.set(m_ctx + 1)
        if m_src is not None and p_src > .5:
            self.v_col_src.set(m_src + 1)
        self.v_col_url.set(m_url + 1 if m_url is not None and p_url > .5 else 0)

    # ------------------------------------------------------------- diversos
    def _escolhe_saida(self):
        c = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                         filetypes=[("Livro Excel", "*.xlsx")])
        if c:
            self.v_saida.set(c)

    def _abre_saida(self):
        p = Path(self.v_saida.get())
        if not p.exists():
            messagebox.showwarning("Abrir", "O ficheiro ainda não existe.")
            return
        if os.name == "nt":
            os.startfile(p)                                   # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(p)])
        else:
            subprocess.Popen(["xdg-open", str(p)])

    # ------------------------------------------------------------ execução
    def _corre(self):
        if self.proc is not None:
            return
        if not MOTOR.exists():
            messagebox.showerror("Motor em falta",
                                 f"textura_near.py não foi encontrado em:\n{AQUI}")
            return
        if not self.v_xlsx.get() or not self.v_folha.get():
            messagebox.showwarning("Matriz", "Carregue o livro no separador 1.")
            return
        if not self.tab.get_children():
            messagebox.showwarning(
                "Termos", "Defina pelo menos um termo de busca no separador 2.")
            return
        if self.v_banda.get() <= self.v_near.get():
            messagebox.showwarning(
                "Parâmetros", "A banda de referência tem de exceder a janela NEAR.")
            return

        tmp = Path(tempfile.gettempdir()) / "_campo_textura.txt"
        tmp.write_text(self._serializa(), encoding="utf-8")

        cmd = [sys.executable, str(MOTOR),
               "--xlsx", self.v_xlsx.get(),
               "--folha", self.v_folha.get(),
               "--near", str(self.v_near.get()),
               "--banda", str(self.v_banda.get()),
               "--lingua", self.v_lingua.get(),
               "--termos", str(tmp),
               "--col-no", str(self.v_col_no.get()),
               "--col-ctx", str(self.v_col_ctx.get()),
               "--col-src", str(self.v_col_src.get()),
               "--col-url", str(self.v_col_url.get()),
               "--sintaxe", self.v_sintaxe.get(),
               "--modelo", self.v_modelo.get(),
               "--saida", self.v_saida.get()]
        if self.v_cab.get():
            cmd += ["--com-cabecalho"]
        if self.v_limite.get().strip().isdigit():
            cmd += ["--limite", self.v_limite.get().strip()]
        if not self.v_fronteira.get():
            cmd += ["--sem-fronteira"]
        if self.v_consulta.get().strip():
            cmd += ["--consulta", self.v_consulta.get().strip()]

        self._log("\n" + "─" * 78)
        self._log(" ".join(f'"{a}"' if " " in a else a for a in cmd) + "\n")
        self.btn_corre.config(state="disabled")
        self.btn_para.config(state="normal")
        self.btn_abre.config(state="disabled")
        self.barra.start(12)
        self.v_estado.set("em execução…")
        threading.Thread(target=self._trabalha, args=(cmd,), daemon=True).start()

    def _trabalha(self, cmd):
        cod = 1
        try:
            self.proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace", bufsize=1)
            for linha in self.proc.stdout:               # type: ignore[union-attr]
                self.fila.put(linha.rstrip())
            cod = self.proc.wait()
        except Exception as exc:                          # noqa: BLE001
            self.fila.put(f"ERRO: {exc}")
        finally:
            self.proc = None
        self.fila.put(f"__FIM__{cod}")

    def _para(self):
        if self.proc is not None:
            self.proc.terminate()
            self._log("interrompido pelo utilizador")

    def _drena(self):
        while True:
            try:
                item = self.fila.get_nowait()
            except queue.Empty:
                break
            if item.startswith("__FIM__"):
                cod = item.removeprefix("__FIM__")
                self.barra.stop()
                self.btn_corre.config(state="normal")
                self.btn_para.config(state="disabled")
                ok = cod == "0"
                self.v_estado.set("concluído" if ok else f"código {cod}")
                if ok:
                    self.btn_abre.config(state="normal")
            else:
                self._log(item)
        self.after(120, self._drena)

    def _log(self, texto: str):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", texto + "\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
