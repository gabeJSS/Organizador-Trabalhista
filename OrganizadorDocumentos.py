"""
ORGANIZADOR DE DOCUMENTOS TRABALHISTAS
Versão 3 do projeto! 

Instale as seguintes dependências:
    pip install PyPDF2 PyMuPDF pandas openpyxl xlrd tkinterdnd2
"""

import os, re, json, shutil, unicodedata, difflib, threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    TkinterDnD = None

import fitz
import PyPDF2
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ══════════════════════════════════════════════════════════════════════════════
# UTILITÁRIOS
# ══════════════════════════════════════════════════════════════════════════════

STOPWORDS = {"DE","DA","DO","DOS","DAS"}

def normalizar_nome(nome):
    nome = unicodedata.normalize("NFD", nome.upper())
    nome = "".join(c for c in nome if unicodedata.category(c) != "Mn")
    nome = re.sub(r"[^A-Z\s]", "", nome)
    return [p for p in nome.split() if p not in STOPWORDS]

def score_match(nome_json, nome_pdf):
    pj, pp = normalizar_nome(nome_json), normalizar_nome(nome_pdf)
    if not pj or not pp: return 0
    s = 0
    if pj[0]  == pp[0]:  s += 20
    if pj[-1] == pp[-1]: s += 50
    for p in pp:
        if p in pj: s += 10
    if difflib.SequenceMatcher(None," ".join(pj)," ".join(pp)).ratio() > 0.8: s += 10
    return s

def encontrar_melhor_match(texto, funcionarios, limite=70):
    cands = []
    for item in funcionarios:
        for nome in item["funcionarios"]:
            s = score_match(nome, texto)
            if s >= limite: cands.append((s, nome, item["condominio"]))
    cands.sort(reverse=True, key=lambda x: x[0])
    return cands[0] if len(cands) == 1 else None

def convert_xls_to_xlsx(xls_path):
    df = pd.read_excel(xls_path, header=None, engine="xlrd")
    new_path = os.path.splitext(xls_path)[0] + ".converted.xlsx"
    df.to_excel(new_path, index=False, header=False, engine="openpyxl")
    return new_path

def create_json_from_excel(excel_path):
    wb = load_workbook(excel_path, data_only=True)
    sheet = wb.active
    condo_data = []
    for row in sheet.iter_rows(min_row=4, min_col=3, max_col=15, values_only=True):
        nome, cond = row[0], row[12]
        if not nome or not cond: continue
        ex = next((c for c in condo_data if c["condominio"] == cond), None)
        if ex: ex["funcionarios"].append(nome)
        else:  condo_data.append({"condominio": cond, "funcionarios": [nome]})
    return condo_data

def copiar_conteudo_pasta(origem, destino):
    for item in os.listdir(origem):
        src = os.path.join(origem, item)
        dst = os.path.join(destino, item)
        if os.path.isdir(src):
            shutil.copytree(src, dst, dirs_exist_ok=True)
        else:
            shutil.copy2(src, dst)

def limpar_caminho_dnd(path):
    path = path.strip()
    if path.startswith("{") and path.endswith("}"): path = path[1:-1]
    return path

def cnpj_no_pdf_fitz(caminho, cnpj):
    try:
        doc = fitz.open(caminho)
        for pg in doc:
            if cnpj in pg.get_text():
                doc.close(); return True
        doc.close()
    except: pass
    return False


# ══════════════════════════════════════════════════════════════════════════════
# WIDGET: DROP ZONE
# ══════════════════════════════════════════════════════════════════════════════

class DropZone(tk.Frame):
    CN = "#f0f4f8"; CH = "#d0e8ff"; COK = "#d4edda"; BOK = "#5aaa6a"; CB = "#b0c4de"

    def __init__(self, parent, label, icon, modo="arquivo", filetypes=None, **kw):
        super().__init__(parent, **kw)
        self.modo = modo; self.filetypes = filetypes or []
        self.var  = tk.StringVar()
        self.var.trace_add("write", self._on_var)
        self._build(label, icon)
        if DND_AVAILABLE: self._reg_dnd()

    def _build(self, label, icon):
        self.config(bg=self.CN, relief="groove", bd=2, cursor="hand2")
        self.columnconfigure(1, weight=1)
        tk.Label(self, text=icon, font=("",16), bg=self.CN, fg="#5577aa"
                ).grid(row=0, column=0, rowspan=2, padx=(8,4), pady=6)
        self._lbl  = tk.Label(self, text=label, font=("",9,"bold"),
                               bg=self.CN, fg="#334466", anchor="w")
        self._lbl.grid(row=0, column=1, sticky="w", pady=(6,0))
        self._dica = tk.Label(self, text="Arraste ou clique para selecionar",
                               font=("",8), bg=self.CN, fg="#8899aa", anchor="w")
        self._dica.grid(row=1, column=1, sticky="w")
        self._ent  = tk.Entry(self, textvariable=self.var, font=("",8),
                               relief="flat", bg=self.CN, fg="#223344",
                               readonlybackground=self.CN, state="readonly")
        self._ent.grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=(0,4))
        for w in (self, self._lbl, self._dica, self._ent):
            w.bind("<Button-1>", self._click)
            w.bind("<Enter>",    lambda e: self._set(self.CH, self.CB) if not self.var.get() else None)
            w.bind("<Leave>",    lambda e: self._set(self.CN, self.CB) if not self.var.get() else None)

    def _reg_dnd(self):
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>",      self._drop)
        self.dnd_bind("<<DragEnter>>", lambda e: self._set(self.CH, self.CB))
        self.dnd_bind("<<DragLeave>>", lambda e: self._set(self.CN, self.CB))

    def _drop(self, event):
        path = limpar_caminho_dnd(event.data)
        self._set(self.CN, self.CB)
        if (self.modo=="pasta" and os.path.isdir(path)) or (self.modo=="arquivo" and os.path.isfile(path)):
            self.var.set(path)

    def _click(self, _=None):
        p = filedialog.askdirectory() if self.modo=="pasta" else filedialog.askopenfilename(filetypes=self.filetypes)
        if p: self.var.set(p)

    def _on_var(self, *_):
        if self.var.get():
            self._set(self.COK, self.BOK)
            self._dica.config(text="OK  " + os.path.basename(self.var.get()))
        else:
            self._set(self.CN, self.CB)
            self._dica.config(text="Arraste ou clique para selecionar")

    def _set(self, bg, bd):
        self.config(bg=bg, highlightbackground=bd, highlightthickness=2, highlightcolor=bd)
        self._lbl.config(bg=bg); self._dica.config(bg=bg)
        self._ent.config(bg=bg, readonlybackground=bg)

    def get(self):   return self.var.get().strip()
    def set(self,v): self.var.set(v)
    def clear(self): self.var.set("")


# ══════════════════════════════════════════════════════════════════════════════
# WIDGET: PAINEL DE PROGRESSO (múltiplos processos, sem distorção)
# ══════════════════════════════════════════════════════════════════════════════

class PainelProgresso(tk.Frame):
    """
    Exibe N barras de progresso alinhadas, uma por processo.
    As barras são criadas dinamicamente e ficam sempre visíveis,
    eliminando o problema de distorção ao alternar entre processos.
    """
    COR = "#ccd8e8"

    PROCESSOS = [
        ("holerite",  "Holerite / Comprovante / Cartão Ponto"),
        ("fgts",      "FGTS"),
        ("nf",        "NF / Boleto / Recibo"),
        ("extrato",   "Extrato Mensal"),
        ("certidoes", "FGTS Gerais / Certidões"),
        ("mescla",    "Mescla de Pastas"),
    ]

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=self.COR, **kw)
        self._bars  = {}
        self._lbls  = {}
        self._build()

    def _build(self):
        tk.Label(self, text="Progresso", font=("",9,"bold"),
                 bg=self.COR, fg="#334466").grid(row=0, column=0, columnspan=2,
                                                  sticky="w", padx=8, pady=(6,2))
        for i, (key, nome) in enumerate(self.PROCESSOS, start=1):
            tk.Label(self, text=nome+":", font=("",8), bg=self.COR,
                     fg="#445566", anchor="e", width=32
                    ).grid(row=i, column=0, sticky="e", padx=(8,4), pady=2)

            bar = ttk.Progressbar(self, mode="determinate", length=340)
            bar.grid(row=i, column=1, sticky="w", padx=(0,8), pady=2)

            lbl = tk.Label(self, text="Aguardando", font=("",8),
                           bg=self.COR, fg="#778899", anchor="w", width=30)
            lbl.grid(row=i, column=2, sticky="w", padx=(4,8), pady=2)

            self._bars[key] = bar
            self._lbls[key] = lbl

    def set(self, key, texto, valor=None, maximo=None):
        bar = self._bars.get(key)
        lbl = self._lbls.get(key)
        if not bar: return
        if maximo is not None: bar["maximum"] = maximo
        if valor  is not None: bar["value"]   = valor
        if lbl: lbl.config(text=texto)
        self.update_idletasks()

    def reset(self, key):
        self.set(key, "Aguardando", 0, 100)


# ══════════════════════════════════════════════════════════════════════════════
# JANELA PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

_Base = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk

class App(_Base):
    BG  = "#eef2f7"
    SEC = "#dde6f0"

    def __init__(self):
        super().__init__()
        self.title("Organizador de Documentos Trabalhistas")
        self.configure(bg=self.BG)
        self.minsize(860, 600)
        self.resizable(True, True)
        self._build()

    # ──────────────────────────────────────────────────────────────────────────
    # LAYOUT PRINCIPAL COM NOTEBOOK (2 ABAS)
    # ──────────────────────────────────────────────────────────────────────────

    def _build(self):
        if not DND_AVAILABLE:
            tk.Label(self, text="  Drag & Drop desativado — instale: pip install tkinterdnd2",
                     bg="#fff3cd", fg="#856404", font=("",8), anchor="w"
                    ).pack(fill="x")

        tk.Label(self, text="Organizador de Documentos Trabalhistas",
                 font=("",14,"bold"), bg=self.BG, fg="#1a2e4a"
                ).pack(anchor="w", padx=16, pady=(10,2))
        tk.Label(self,
                 text="Configure ano/mês/saída e use as abas abaixo. "
                      "Arraste arquivos diretamente nas zonas indicadas.",
                 font=("",9), bg=self.BG, fg="#556"
                ).pack(anchor="w", padx=16, pady=(0,6))

        # Configurações gerais (fora das abas — sempre visível)
        self._secao_config()

        # Notebook com 2 abas
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=6)

        aba_proc   = tk.Frame(nb, bg=self.BG)
        aba_rel    = tk.Frame(nb, bg=self.BG)
        nb.add(aba_proc, text="  Processamento  ")
        nb.add(aba_rel,  text="  Relatório / Auditoria  ")

        # Aba processamento: canvas rolável
        canvas = tk.Canvas(aba_proc, bg=self.BG, highlightthickness=0)
        sb     = ttk.Scrollbar(aba_proc, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._inner = tk.Frame(canvas, bg=self.BG)
        win = canvas.create_window((0,0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
            lambda e: canvas.itemconfig(win, width=e.width))
        self.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        # Seções de processamento
        self._secao_holerite()
        self._secao_fgts_nf()        # FGTS + NF juntos (JSON compartilhado)
        self._secao_extrato()
        self._secao_certidoes()
        self._secao_mescla()

        # Painel de progresso (fixo na parte inferior da aba proc)
        self.progresso = PainelProgresso(aba_proc, relief="sunken", bd=1)
        self.progresso.pack(fill="x", side="bottom", padx=0, pady=0)

        # Aba relatório
        self._build_aba_relatorio(aba_rel)

    # ──────────────────────────────────────────────────────────────────────────
    # HELPERS DE LAYOUT
    # ──────────────────────────────────────────────────────────────────────────

    def _bloco(self, titulo, cor="#1a2e4a"):
        outer = tk.Frame(self._inner, bg=self.BG)
        outer.pack(fill="x", padx=10, pady=(6,2))
        tk.Label(outer, text=titulo, font=("",10,"bold"),
                 bg=self.BG, fg=cor).pack(anchor="w", pady=(0,2))
        inner = tk.Frame(outer, bg=self.SEC, relief="groove", bd=1)
        inner.pack(fill="x")
        return inner

    def _grid2(self, p):
        g = tk.Frame(p, bg=self.SEC)
        g.pack(fill="x", padx=8, pady=4)
        g.columnconfigure(0, weight=1); g.columnconfigure(1, weight=1)
        return g

    def _grid3(self, p):
        g = tk.Frame(p, bg=self.SEC)
        g.pack(fill="x", padx=8, pady=4)
        g.columnconfigure(0,weight=1); g.columnconfigure(1,weight=1); g.columnconfigure(2,weight=1)
        return g

    def _grid4(self, p):
        g = tk.Frame(p, bg=self.SEC)
        g.pack(fill="x", padx=8, pady=4)
        for i in range(4): g.columnconfigure(i, weight=1)
        return g

    def _btn(self, parent, texto, cor, cmd):
        tk.Button(parent, text=texto, bg=cor, fg="white",
                  font=("",10,"bold"), relief="flat", cursor="hand2",
                  pady=5, command=cmd
                 ).pack(fill="x", padx=8, pady=(4,8))

    def _lbl_sec(self, parent, texto, cor="#334"):
        tk.Label(parent, text=texto, bg=self.SEC,
                 font=("",9,"bold"), fg=cor
                ).pack(anchor="w", padx=8, pady=(4,0))

    # ──────────────────────────────────────────────────────────────────────────
    # CONFIGURAÇÕES GERAIS
    # ──────────────────────────────────────────────────────────────────────────

    def _secao_config(self):
        f = tk.Frame(self, bg="#d5e3f0", relief="groove", bd=1)
        f.pack(fill="x", padx=12, pady=(4,2))
        row = tk.Frame(f, bg="#d5e3f0")
        row.pack(fill="x", padx=8, pady=8)

        tk.Label(row, text="Ano:", bg="#d5e3f0", font=("",9,"bold")).pack(side="left", padx=(4,2))
        self.ano_var = tk.StringVar()
        tk.Entry(row, textvariable=self.ano_var, width=6, font=("",10)).pack(side="left", padx=(0,14))

        tk.Label(row, text="Mês:", bg="#d5e3f0", font=("",9,"bold")).pack(side="left", padx=(0,2))
        self.mes_var = tk.StringVar()
        ttk.Combobox(row, textvariable=self.mes_var, state="readonly", width=5,
                     values=["01","02","03","04","05","06","07","08","09","10","11","12"]
                    ).pack(side="left", padx=(0,14))

        tk.Label(row, text="Pasta de saída:", bg="#d5e3f0", font=("",9,"bold")).pack(side="left", padx=(0,2))
        self.saida_var = tk.StringVar()
        tk.Entry(row, textvariable=self.saida_var, width=36, font=("",9)).pack(side="left", padx=(0,4))
        tk.Button(row, text="Selecionar", relief="flat", bg="#a0b8d0", cursor="hand2",
                  font=("",8),
                  command=lambda: self.saida_var.set(
                      filedialog.askdirectory() or self.saida_var.get())
                 ).pack(side="left")

    # ──────────────────────────────────────────────────────────────────────────
    # SEÇÃO: HOLERITE / COMPROVANTE / CARTÃO PONTO
    # ──────────────────────────────────────────────────────────────────────────

    def _secao_holerite(self):
        f = self._bloco("  Holerite  /  Comprovante  /  Cartão Ponto", "#1a3d2e")
        self._lbl_sec(f, "Planilha Excel — exportação do sistema de cartão ponto (obrigatória):")
        self.dz_excel = DropZone(f, "Planilha Excel", "S", filetypes=[("Excel","*.xls *.xlsx")], bg=self.SEC)
        self.dz_excel.pack(fill="x", padx=8, pady=4)

        self._lbl_sec(f, "PDFs — deixe vazio o que não tiver neste mês:")
        g = self._grid3(f)
        self.dz_holerite    = DropZone(g, "Holerite",    "H", filetypes=[("PDF","*.pdf")], bg=self.SEC)
        self.dz_comprovante = DropZone(g, "Comprovante", "C", filetypes=[("PDF","*.pdf")], bg=self.SEC)
        self.dz_cartao      = DropZone(g, "Cartão Ponto","P", filetypes=[("PDF","*.pdf")], bg=self.SEC)
        self.dz_holerite.grid(row=0,column=0,sticky="nsew",padx=4,pady=2)
        self.dz_comprovante.grid(row=0,column=1,sticky="nsew",padx=4,pady=2)
        self.dz_cartao.grid(row=0,column=2,sticky="nsew",padx=4,pady=2)

        self._btn(f, "Processar Holerite / Comprovante / Cartão Ponto", "#2d6a4f",
                  lambda: threading.Thread(target=self._run_holerite, daemon=True).start())

    # ──────────────────────────────────────────────────────────────────────────
    # SEÇÃO: FGTS + NF/BOLETO/RECIBO (JSON de CNPJ compartilhado)
    # ──────────────────────────────────────────────────────────────────────────

    def _secao_fgts_nf(self):
        f = self._bloco("  FGTS  e  Notas Fiscais / Boletos / Recibos", "#12325e")

        # JSON compartilhado
        self._lbl_sec(f, "JSON de CNPJs — usado pelo FGTS e pelo NF/Boleto/Recibo:")
        self.dz_cnpj_json = DropZone(f, "JSON de CNPJs (CNPJs.json)", "J",
                                      filetypes=[("JSON","*.json")], bg=self.SEC)
        self.dz_cnpj_json.pack(fill="x", padx=8, pady=4)

        sep = tk.Frame(f, bg="#b8cde0", height=1)
        sep.pack(fill="x", padx=8, pady=4)

        # FGTS
        self._lbl_sec(f, "FGTS — Relatório:")
        row_fgts = tk.Frame(f, bg=self.SEC)
        row_fgts.pack(fill="x", padx=8, pady=(2,0))
        tk.Label(row_fgts, text="Nome da subpasta FGTS:", bg=self.SEC, font=("",9)).pack(side="left")
        self.fgts_subpasta = tk.StringVar(value="FGTS")
        tk.Entry(row_fgts, textvariable=self.fgts_subpasta, width=16, font=("",9)).pack(side="left", padx=6)

        self.dz_fgts_pdf = DropZone(f, "PDF Relatório FGTS", "F",
                                     filetypes=[("PDF","*.pdf")], bg=self.SEC)
        self.dz_fgts_pdf.pack(fill="x", padx=8, pady=4)

        self._btn(f, "Processar FGTS", "#1d4e89",
                  lambda: threading.Thread(target=self._run_fgts, daemon=True).start())

        sep2 = tk.Frame(f, bg="#b8cde0", height=1)
        sep2.pack(fill="x", padx=8, pady=4)

        # NF / Boleto / Recibo
        self._lbl_sec(f, "NF / Boleto / Recibo — Pasta com os PDFs:")
        self.dz_nf_pasta = DropZone(f, "Pasta com os PDFs", "D", modo="pasta", bg=self.SEC)
        self.dz_nf_pasta.pack(fill="x", padx=8, pady=4)

        self._btn(f, "Processar NF / Boleto / Recibo", "#5a3e8e",
                  lambda: threading.Thread(target=self._run_nf_boleto, daemon=True).start())

    # ──────────────────────────────────────────────────────────────────────────
    # SEÇÃO: EXTRATO MENSAL (do contador)
    # ──────────────────────────────────────────────────────────────────────────

    def _secao_extrato(self):
        f = self._bloco("  Extrato Mensal  (pasta do contador)", "#3d2a00")
        self._lbl_sec(f,
            "Pasta recebida do contador: cada subpasta = nome do cliente, "
            "com 'Extrato Mensal.pdf' dentro.")
        g = self._grid2(f)
        self.dz_extrato_origem = DropZone(g, "Pasta do Contador (origem)", "D",
                                           modo="pasta", bg=self.SEC)
        self.dz_extrato_saida  = DropZone(g, "Pasta de Saída (output)", "D",
                                           modo="pasta", bg=self.SEC)
        self.dz_extrato_origem.grid(row=0,column=0,sticky="nsew",padx=4,pady=(8,4))
        self.dz_extrato_saida.grid( row=0,column=1,sticky="nsew",padx=4,pady=(8,4))
        self._btn(f, "Processar Extrato Mensal", "#7a5200",
                  lambda: threading.Thread(target=self._run_extrato, daemon=True).start())

    # ──────────────────────────────────────────────────────────────────────────
    # SEÇÃO: FGTS GERAIS / CERTIDÕES (distribuir para todas as pastas)
    # ──────────────────────────────────────────────────────────────────────────

    def _secao_certidoes(self):
        f = self._bloco("  FGTS Gerais e Certidões  — distribuir para todos os clientes", "#4a2060")
        self._lbl_sec(f,
            "Selecione as pastas de origem (FGTS, Gerais, Certidões) e a pasta de clientes (output). "
            "O conteúdo será copiado para ANO/MES/[subpasta] de cada cliente.")

        g = self._grid4(f)
        self.dz_cer_fgts    = DropZone(g, "Pasta FGTS\n(Guia + Comprovante)", "F", modo="pasta", bg=self.SEC)
        self.dz_cer_gerais  = DropZone(g, "Pasta Gerais\n(DCTFWeb / GPS)", "G", modo="pasta", bg=self.SEC)
        self.dz_cer_certs   = DropZone(g, "Pasta Certidões\n(CNDs)", "C", modo="pasta", bg=self.SEC)
        self.dz_cer_clientes= DropZone(g, "Pasta de Clientes\n(output)", "D", modo="pasta", bg=self.SEC)
        self.dz_cer_fgts.grid(   row=0,column=0,sticky="nsew",padx=4,pady=(8,4))
        self.dz_cer_gerais.grid( row=0,column=1,sticky="nsew",padx=4,pady=(8,4))
        self.dz_cer_certs.grid(  row=0,column=2,sticky="nsew",padx=4,pady=(8,4))
        self.dz_cer_clientes.grid(row=0,column=3,sticky="nsew",padx=4,pady=(8,4))

        self._btn(f, "Distribuir FGTS Gerais e Certidões para todos os clientes", "#5a3e8e",
                  lambda: threading.Thread(target=self._run_certidoes, daemon=True).start())

    # ──────────────────────────────────────────────────────────────────────────
    # SEÇÃO: MESCLAR PASTAS
    # ──────────────────────────────────────────────────────────────────────────

    def _secao_mescla(self):
        f = self._bloco("  Mesclar Pastas  (juntador.json)", "#5c2e00")
        g = self._grid2(f)
        self.dz_mescla_json  = DropZone(g, "JSON de mescla (juntador.json)", "J",
                                         filetypes=[("JSON","*.json")], bg=self.SEC)
        self.dz_mescla_pasta = DropZone(g, "Pasta base (output)", "D",
                                         modo="pasta", bg=self.SEC)
        self.dz_mescla_json.grid( row=0,column=0,sticky="nsew",padx=4,pady=(8,4))
        self.dz_mescla_pasta.grid(row=0,column=1,sticky="nsew",padx=4,pady=(8,4))
        self._btn(f, "Executar Mescla", "#7a4419",
                  lambda: threading.Thread(target=self._run_mescla, daemon=True).start())

    # ──────────────────────────────────────────────────────────────────────────
    # ABA: RELATÓRIO / AUDITORIA
    # ──────────────────────────────────────────────────────────────────────────

    def _build_aba_relatorio(self, parent):
        f = tk.Frame(parent, bg=self.BG)
        f.pack(fill="both", expand=True, padx=16, pady=12)

        tk.Label(f, text="Gerar Relatório de Auditoria",
                 font=("",12,"bold"), bg=self.BG, fg="#1a2e4a").pack(anchor="w")
        tk.Label(f,
                 text="Aponta condomínios organizados, documentos faltando por condomínio "
                      "e documentos faltando por funcionário. Exporta para Excel.",
                 font=("",9), bg=self.BG, fg="#556", wraplength=700, justify="left"
                ).pack(anchor="w", pady=(2,12))

        # Inputs
        frm = tk.Frame(f, bg=self.SEC, relief="groove", bd=1)
        frm.pack(fill="x", pady=(0,10))

        row1 = tk.Frame(frm, bg=self.SEC)
        row1.pack(fill="x", padx=8, pady=8)
        tk.Label(row1, text="Pasta output (clientes):", bg=self.SEC, font=("",9,"bold")).pack(side="left")
        self.rel_pasta_var = tk.StringVar()
        tk.Entry(row1, textvariable=self.rel_pasta_var, width=46, font=("",9)).pack(side="left", padx=6)
        tk.Button(row1, text="Selecionar", relief="flat", bg="#a0b8d0", cursor="hand2",
                  command=lambda: self.rel_pasta_var.set(
                      filedialog.askdirectory() or self.rel_pasta_var.get())
                 ).pack(side="left")

        row2 = tk.Frame(frm, bg=self.SEC)
        row2.pack(fill="x", padx=8, pady=(0,8))
        tk.Label(row2, text="Ano:", bg=self.SEC, font=("",9,"bold")).pack(side="left")
        self.rel_ano_var = tk.StringVar()
        tk.Entry(row2, textvariable=self.rel_ano_var, width=6, font=("",9)).pack(side="left", padx=(4,14))
        tk.Label(row2, text="Mês:", bg=self.SEC, font=("",9,"bold")).pack(side="left")
        self.rel_mes_var = tk.StringVar()
        ttk.Combobox(row2, textvariable=self.rel_mes_var, state="readonly", width=5,
                     values=["01","02","03","04","05","06","07","08","09","10","11","12"]
                    ).pack(side="left", padx=(4,14))
        tk.Label(row2, text="Salvar relatório em:", bg=self.SEC, font=("",9,"bold")).pack(side="left")
        self.rel_saida_var = tk.StringVar(value="relatorio_auditoria.xlsx")
        tk.Entry(row2, textvariable=self.rel_saida_var, width=30, font=("",9)).pack(side="left", padx=4)
        tk.Button(row2, text="Salvar como...", relief="flat", bg="#a0b8d0", cursor="hand2",
                  command=self._escolher_saida_rel).pack(side="left")

        tk.Button(f, text="Gerar Relatório Excel", bg="#1a3d2e", fg="white",
                  font=("",11,"bold"), relief="flat", cursor="hand2", pady=8,
                  command=lambda: threading.Thread(
                      target=self._run_relatorio, daemon=True).start()
                 ).pack(fill="x", pady=4)

        # Preview na tela
        tk.Label(f, text="Preview:", font=("",9,"bold"), bg=self.BG).pack(anchor="w", pady=(8,2))
        cols = ("Condomínio","Subpasta","Status","Detalhe")
        self._tree = ttk.Treeview(f, columns=cols, show="headings", height=14)
        for c in cols:
            self._tree.heading(c, text=c)
            self._tree.column(c, width=160 if c!="Detalhe" else 280)
        vsb = ttk.Scrollbar(f, orient="vertical", command=self._tree.yview)
        hsb = ttk.Scrollbar(f, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._rel_status = tk.Label(f, text="", font=("",9), bg=self.BG, fg="#334")
        self._rel_status.pack(anchor="w", pady=4)

    def _escolher_saida_rel(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")],
            initialfile="relatorio_auditoria.xlsx")
        if p: self.rel_saida_var.set(p)

    # ══════════════════════════════════════════════════════════════════════════
    # HELPERS COMUNS
    # ══════════════════════════════════════════════════════════════════════════

    def _get_periodo(self):
        ano   = self.ano_var.get().strip()
        mes   = self.mes_var.get().strip().zfill(2)
        saida = self.saida_var.get().strip() or "output"
        if not ano or not mes:
            raise ValueError("Informe o ano e o mês nas Configurações Gerais.")
        return ano, mes, saida

    def _salvar_pagina(self, page, cond, nome, tipo, ano, mes, saida):
        pasta = os.path.join(saida, str(cond), ano, f"{mes}.{ano}", tipo)
        os.makedirs(pasta, exist_ok=True)
        w = PyPDF2.PdfWriter(); w.add_page(page)
        with open(os.path.join(pasta, f"{tipo.replace(' ','_')}_{nome}.pdf"), "wb") as fh:
            w.write(fh)

    # ══════════════════════════════════════════════════════════════════════════
    # LÓGICA: HOLERITE / COMPROVANTE / CARTÃO PONTO
    # ══════════════════════════════════════════════════════════════════════════

    def _run_holerite(self):
        try:
            ano, mes, saida = self._get_periodo()
            excel = self.dz_excel.get()
            if not excel:
                raise ValueError("Selecione a planilha Excel.")

            self.progresso.set("holerite", "Carregando planilha...", 0, 100)

            # Converte XLS se necessário
            if excel.lower().endswith(".xls"):
                self.progresso.set("holerite", "Convertendo XLS→XLSX...")
                excel = convert_xls_to_xlsx(excel)

            funcionarios = create_json_from_excel(excel)

            # 🔥 Mapa nome_upper → (condominio, nome_original)
            mapa_nomes = {}
            for item in funcionarios:
                cond = item["condominio"]
                for nome in item["funcionarios"]:
                    if nome:
                        mapa_nomes[nome.upper()] = (cond, nome)

            arquivos = {
                t: dz.get()
                for t, dz in [
                    ("Holerites", self.dz_holerite),
                    ("Comprovantes", self.dz_comprovante),
                    ("Cartao Ponto", self.dz_cartao),
                ]
                if dz.get()
            }

            if not arquivos:
                raise ValueError("Selecione ao menos um PDF.")

            total = sum(len(PyPDF2.PdfReader(open(p, "rb")).pages)
                        for p in arquivos.values())

            self.progresso.set("holerite", "Processando...", 0, total)
            prog = 0

            for tipo, caminho in arquivos.items():

                with open(caminho, "rb") as fh:
                    reader = PyPDF2.PdfReader(fh)

                    for page in reader.pages:

                        texto = (page.extract_text() or "").upper()
                        encontrado = False

                        # 🔹 BUSCA DIRETA (rápida)
                        for nome_upper, (cond, nome_original) in mapa_nomes.items():
                            if nome_upper in texto:
                                self._salvar_pagina(
                                    page, cond, nome_original,
                                    tipo, ano, mes, saida
                                )
                                encontrado = True
                                break

                        # 🔹 FUZZY MATCH APENAS PARA COMPROVANTES
                        if not encontrado and tipo == "Comprovantes":
                            m = encontrar_melhor_match(texto, funcionarios)
                            if m:
                                _, nome, cond = m
                                self._salvar_pagina(
                                    page, cond, nome,
                                    tipo, ano, mes, saida
                                )
                            else:
                                self._salvar_pagina(
                                    page, "__PENDENTE", "SEM_MATCH",
                                    tipo, ano, mes, saida
                                )

                        prog += 1
                        self.progresso.set(
                            "holerite",
                            f"{tipo} ({prog}/{total})",
                            prog,
                            total
                        )

            self.progresso.set("holerite", "Concluído!", total, total)
            messagebox.showinfo(
                "Sucesso",
                "Holerite / Comprovante / Cartão Ponto\nconcluídos com sucesso!"
            )

        except Exception as e:
            self.progresso.set("holerite", f"Erro: {e}")
            messagebox.showerror("Erro", str(e))

    # ══════════════════════════════════════════════════════════════════════════
    # LÓGICA: FGTS
    # ══════════════════════════════════════════════════════════════════════════

    def _run_fgts(self):
        try:
            ano, mes, saida = self._get_periodo()
            subpasta  = self.fgts_subpasta.get().strip() or "FGTS"
            json_path = self.dz_cnpj_json.get()
            pdf_path  = self.dz_fgts_pdf.get()
            if not json_path or not pdf_path:
                raise ValueError("Selecione o JSON de CNPJs e o PDF do FGTS.")

            with open(json_path,"r",encoding="utf-8") as fh: data = json.load(fh)

            with open(pdf_path,"rb") as fh:
                reader = PyPDF2.PdfReader(fh)
                total_pgs = len(reader.pages)
                self.progresso.set("fgts","Processando...",0,len(data))

                for i,entry in enumerate(data):
                    cnpj, cond = entry["CNPJ"], entry["condominio"]
                    pasta = os.path.join(saida, cond, ano, f"{mes}.{ano}", subpasta)
                    os.makedirs(pasta, exist_ok=True)
                    pgs = [reader.pages[n] for n in range(total_pgs-1)
                           if cnpj in (reader.pages[n].extract_text() or "")]
                    if pgs:
                        w = PyPDF2.PdfWriter()
                        for pg in pgs: w.add_page(pg)
                        with open(os.path.join(pasta,"Relatorio FGTS Mensal.pdf"),"wb") as fw:
                            w.write(fw)
                    self.progresso.set("fgts", f"{i+1}/{len(data)} condomínios", i+1, len(data))

            self.progresso.set("fgts","Concluído!",len(data),len(data))
            messagebox.showinfo("Sucesso","Relatório FGTS organizado com sucesso!")
        except Exception as e:
            self.progresso.set("fgts",f"Erro: {e}")
            messagebox.showerror("Erro",str(e))

    # ══════════════════════════════════════════════════════════════════════════
    # LÓGICA: NF / BOLETO / RECIBO
    # ══════════════════════════════════════════════════════════════════════════

    def _run_nf_boleto(self):
        try:
            ano, mes, saida = self._get_periodo()
            json_path  = self.dz_cnpj_json.get()
            pasta_docs = self.dz_nf_pasta.get()
            if not json_path or not pasta_docs:
                raise ValueError("Selecione o JSON de CNPJs e a pasta com PDFs.")

            with open(json_path,"r",encoding="utf-8") as fh:
                data = json.load(fh)

            # cria mapa CNPJ → condomínio
            mapa_cnpj = {d["CNPJ"]: d["condominio"] for d in data}

            pdfs = [a for a in os.listdir(pasta_docs) if a.lower().endswith(".pdf")]
            self.progresso.set("nf","Processando...",0,len(pdfs))

            for i, arq in enumerate(pdfs):
                cam = os.path.join(pasta_docs, arq)

                try:
                    doc = fitz.open(cam)
                    texto = ""
                    for pg in doc:
                        texto += pg.get_text()
                    doc.close()
                except:
                    continue

                # verifica qual CNPJ está no texto
                for cnpj, cond in mapa_cnpj.items():
                    if cnpj in texto:
                        nl = arq.lower()
                        sub = ("Boletos" if "boleto" in nl
                            else "Notas Fiscais" if "nota" in nl or "nf" in nl
                            else "Recibos" if "recibo" in nl
                            else "Documentos")

                        dst = os.path.join(saida, cond, ano, f"{mes}.{ano}", sub)
                        os.makedirs(dst, exist_ok=True)
                        shutil.move(cam, dst)
                        break

                self.progresso.set("nf", f"{i+1}/{len(pdfs)} PDFs", i+1, len(pdfs))

            self.progresso.set("nf","Concluído!",len(pdfs),len(pdfs))
            messagebox.showinfo("Sucesso","NF / Boleto / Recibo organizados com sucesso!")

        except Exception as e:
            self.progresso.set("nf",f"Erro: {e}")
            messagebox.showerror("Erro",str(e))


    # ══════════════════════════════════════════════════════════════════════════
    # LÓGICA: MESCLAR PASTAS
    # ══════════════════════════════════════════════════════════════════════════

    def _run_mescla(self):
        try:
            json_path  = self.dz_mescla_json.get()
            pasta_base = self.dz_mescla_pasta.get()
            if not json_path or not pasta_base:
                raise ValueError("Selecione o JSON de mescla e a pasta base.")

            with open(json_path,"r",encoding="utf-8") as fh: dados = json.load(fh)
            self.progresso.set("mescla","Mesclando...",0,len(dados))
            deletadas = 0

            for i, item in enumerate(dados):
                destino = os.path.join(pasta_base, item["nome"])
                os.makedirs(destino, exist_ok=True)
                for np_ in item["pastas"]:
                    origem = os.path.join(pasta_base, np_)
                    if os.path.isdir(origem): copiar_conteudo_pasta(origem, destino)
                for np_ in item["pastas"]:
                    cam = os.path.join(pasta_base, np_)
                    if os.path.isdir(cam):
                        try: shutil.rmtree(cam); deletadas += 1
                        except Exception as ex: print(f"Erro ao deletar {cam}: {ex}")
                self.progresso.set("mescla", f"{i+1}/{len(dados)}", i+1, len(dados))

            self.progresso.set("mescla",f"Concluído! {deletadas} pastas removidas.",len(dados),len(dados))
            messagebox.showinfo("Sucesso",f"Mescla concluída!\n{deletadas} pastas removidas.")
        except Exception as e:
            self.progresso.set("mescla",f"Erro: {e}")
            messagebox.showerror("Erro",str(e))

    # ══════════════════════════════════════════════════════════════════════════
    # LÓGICA: RELATÓRIO DE AUDITORIA
    # ══════════════════════════════════════════════════════════════════════════

    # Subpastas esperadas por condomínio (nível do mês)
    SUBPASTAS_COND = ["Boletos","Certidoes","Extrato Mensal","FGTS",
                      "Gerais","Notas Fiscais","Recibos"]
    # Subpastas esperadas por funcionário
    SUBPASTAS_FUNC = ["Holerites","Comprovantes","Cartao Ponto"]

    def _run_relatorio(self):
        try:
            pasta_root = self.rel_pasta_var.get().strip()
            ano        = self.rel_ano_var.get().strip()
            mes        = self.rel_mes_var.get().strip().zfill(2)
            saida      = self.rel_saida_var.get().strip() or "relatorio_auditoria.xlsx"

            if not pasta_root: raise ValueError("Selecione a pasta de clientes.")
            if not ano or not mes: raise ValueError("Informe ano e mês.")

            periodo = f"{mes}.{ano}"
            self._rel_status.config(text="Analisando pastas...")
            self.update_idletasks()

            # Limpa preview
            for row in self._tree.get_children(): self._tree.delete(row)

            linhas = []  # (condominio, subpasta, status, detalhe)

            condominios = sorted([d for d in os.listdir(pasta_root)
                                   if os.path.isdir(os.path.join(pasta_root, d))])

            for cond in condominios:
                pasta_mes = os.path.join(pasta_root, cond, ano, periodo)
                if not os.path.isdir(pasta_mes):
                    linhas.append((cond, "—", "SEM PASTA DO MÊS",
                                   f"Não existe: {pasta_mes}"))
                    continue

                subs_existentes = {d.lower() for d in os.listdir(pasta_mes)
                                   if os.path.isdir(os.path.join(pasta_mes, d))}

                # Verifica subpastas de condomínio
                for sub in self.SUBPASTAS_COND:
                    pasta_sub = os.path.join(pasta_mes, sub)
                    if not os.path.isdir(pasta_sub):
                        linhas.append((cond, sub, "FALTANDO", "Subpasta não encontrada"))
                    else:
                        pdfs = [f for f in os.listdir(pasta_sub) if f.lower().endswith(".pdf")]
                        if not pdfs:
                            linhas.append((cond, sub, "VAZIA", "Nenhum PDF encontrado"))
                        else:
                            linhas.append((cond, sub, "OK", f"{len(pdfs)} arquivo(s)"))

                # Verifica por funcionário (nas subpastas Holerites, Comprovantes, Cartao Ponto)
                nomes_por_sub = {}
                for sub in self.SUBPASTAS_FUNC:
                    pasta_sub = os.path.join(pasta_mes, sub)
                    if os.path.isdir(pasta_sub):
                        arqs = [os.path.splitext(f)[0] for f in os.listdir(pasta_sub)
                                if f.lower().endswith(".pdf")]
                        # Remove prefixo do tipo (ex: "Holerites_NOME" → "NOME")
                        nomes = set()
                        for a in arqs:
                            partes = a.split("_", 1)
                            nomes.add(partes[1] if len(partes) == 2 else a)
                        nomes_por_sub[sub] = nomes

                if nomes_por_sub:
                    todos_nomes = set().union(*nomes_por_sub.values())
                    for nome in sorted(todos_nomes):
                        faltando = [s for s in self.SUBPASTAS_FUNC
                                    if s in nomes_por_sub and nome not in nomes_por_sub[s]]
                        if faltando:
                            linhas.append((cond, "Funcionário", "INCOMPLETO",
                                           f"{nome} — faltando: {', '.join(faltando)}"))
                        else:
                            linhas.append((cond, "Funcionário", "OK", nome))

            # Atualiza preview
            for row in linhas:
                tag = ("ok"       if row[2]=="OK"
                       else "falt" if row[2] in ("FALTANDO","SEM PASTA DO MÊS")
                       else "inc"  if row[2] in ("INCOMPLETO","VAZIA")
                       else "")
                self._tree.insert("", "end", values=row, tags=(tag,))

            self._tree.tag_configure("ok",   background="#e8f5e9")
            self._tree.tag_configure("falt", background="#fdecea")
            self._tree.tag_configure("inc",  background="#fff8e1")

            # Gera Excel
            self._exportar_excel(linhas, saida, ano, mes)

            total_ok   = sum(1 for l in linhas if l[2]=="OK")
            total_prob = sum(1 for l in linhas if l[2]!="OK")
            self._rel_status.config(
                text=f"Concluído — {total_ok} OK | {total_prob} problemas | "
                     f"Exportado: {saida}")
            messagebox.showinfo("Relatório gerado",
                                f"{len(condominios)} condomínios analisados.\n"
                                f"{total_ok} itens OK | {total_prob} problemas.\n"
                                f"Arquivo: {saida}")

        except Exception as e:
            self._rel_status.config(text=f"Erro: {e}")
            messagebox.showerror("Erro",str(e))

    def _exportar_excel(self, linhas, caminho, ano, mes):
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = f"Auditoria {mes}.{ano}"

        # Estilos
        h_fill = PatternFill("solid", fgColor="1A2E4A")
        h_font = Font(color="FFFFFF", bold=True, size=10)
        ok_fill   = PatternFill("solid", fgColor="C8E6C9")
        falt_fill = PatternFill("solid", fgColor="FFCDD2")
        inc_fill  = PatternFill("solid", fgColor="FFF9C4")
        borda = Border(
            left  =Side(style="thin"),right =Side(style="thin"),
            top   =Side(style="thin"),bottom=Side(style="thin"))

        headers = ["Condomínio","Subpasta / Categoria","Status","Detalhe"]
        ws.append(headers)
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col)
            c.fill, c.font, c.alignment = h_fill, h_font, Alignment(horizontal="center")
            c.border = borda

        for row in linhas:
            ws.append(list(row))
            r = ws.max_row
            status = row[2]
            fill = (ok_fill if status=="OK"
                    else falt_fill if status in ("FALTANDO","SEM PASTA DO MÊS")
                    else inc_fill)
            for col in range(1, 5):
                c = ws.cell(row=r, column=col)
                c.fill, c.border = fill, borda
                c.alignment = Alignment(wrap_text=True)

        # Larguras
        for col, w in zip(range(1,5), [38, 22, 18, 50]):
            ws.column_dimensions[get_column_letter(col)].width = w

        # Aba resumo
        ws2 = wb.create_sheet("Resumo")
        condominios_unicos = sorted(set(l[0] for l in linhas))
        ws2.append(["Condomínio","Total Itens","OK","Problemas"])
        for cond in condominios_unicos:
            rows_c = [l for l in linhas if l[0]==cond]
            ok_c   = sum(1 for l in rows_c if l[2]=="OK")
            prob_c = sum(1 for l in rows_c if l[2]!="OK")
            ws2.append([cond, len(rows_c), ok_c, prob_c])
            r = ws2.max_row
            fill = ok_fill if prob_c == 0 else (falt_fill if prob_c > 2 else inc_fill)
            for col in range(1,5):
                ws2.cell(row=r,column=col).fill   = fill
                ws2.cell(row=r,column=col).border = borda

        for col, w in zip(range(1,5),[40,14,10,12]):
            ws2.column_dimensions[get_column_letter(col)].width = w

        # Cabeçalhos da aba resumo
        for col in range(1,5):
            c = ws2.cell(row=1,column=col)
            c.fill, c.font = h_fill, h_font
            c.border, c.alignment = borda, Alignment(horizontal="center")

        wb.save(caminho)


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()
