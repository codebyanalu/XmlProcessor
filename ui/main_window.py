"""
ui/main_window.py
Interface profissional ERP — sidebar 300px, abas NF-e / NFS-e / Dashboard.
Performance: salva em lote, throttle UI, log resumido.
"""

import os, sys, threading, tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from datetime import datetime

import customtkinter as ctk
import pandas as pd

import config.settings as cfg
from extract import extrair_produtos, extrair_servicos
from transform import filtrar_novos, carregar_chaves_existentes
from load import (
    inicializar_sessao, verificar_locks_ativos,
    salvar_produtos_csv, sincronizar_excel_temp,
    sincronizar_com_principal, atualizar_excel_principal,
    limpar_temporarios, total_registros,
)

# ── Paleta ────────────────────────────────────────────────────────────────────
C_PRIM   = "#1a5276"
C_SEC    = "#2980b9"
C_ACENT  = "#e67e22"
C_FUNDO  = "#f0f3f7"
C_FUNDO2 = "#ffffff"
C_SIDE   = "#1c2833"
C_SIDE2  = "#2c3e50"
C_TEXTO  = "#1a252f"
C_TEXTO2 = "#5d6d7e"
C_BORDA  = "#d5d8dc"
C_OK     = "#1e8449"
C_WARN   = "#d68910"
C_ERR    = "#c0392b"
C_INFO   = "#2471a3"
C_CANCEL = "#c0392b"
C_AUTH   = "#1e8449"

FONTE_LOG = ("Consolas", 10)
FONTE_BTN = ("Segoe UI", 10, "bold")
UI_THROTTLE = 10
LOTE_MAX = 500

# ── Detectar tipo de XML ───────────────────────────────────────────────────────
def _detectar_tipo(caminho):
    """Retorna 'nfe' ou 'nfse' baseado no conteúdo do arquivo."""
    try:
        with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
            trecho = f.read(500)
        if any(x in trecho for x in ["CompNFe","NFSe","infNFSe","nNFSe"]):
            return "nfse"
        return "nfe"
    except Exception:
        return "nfe"


# ══════════════════════════════════════════════════════════════════════════════
# JANELA PLANILHA NFS-e (estilo sistema da imagem)
# ══════════════════════════════════════════════════════════════════════════════

class JanelaPlanilhaNFSe(ctk.CTkToplevel):

    _COLUNAS = [
        "#", "Situacao", "Numero_NFSe", "Data_Emissao", "Formato",
        "Nome_Prestador", "CNPJ_Prestador", "Mun_Prestador",
        "Nome_Tomador",   "CNPJ_Tomador",
        "Desc_Servico", "cTribNac",
        "BC_ISS", "Aliq_ISS", "Valor_ISS", "tpRetISSQN",
        "Valor_Bruto", "Valor_Liquido",
        "IBS_vBC", "IBS_pIBSUF", "CBS_pCBS",
        "Valor_PIS", "Valor_COFINS", "Valor_IRRF", "Valor_INSS",
        "Discriminacao", "Arquivo_Origem",
    ]
    _LARG = {
        "#":20, "Situacao":80, "Numero_NFSe":80, "Data_Emissao":130,
        "Formato":110,
        "Nome_Prestador":200, "CNPJ_Prestador":120, "Mun_Prestador":120,
        "Nome_Tomador":180, "CNPJ_Tomador":120,
        "Desc_Servico":220, "cTribNac":80,
        "BC_ISS":100, "Aliq_ISS":70, "Valor_ISS":100, "tpRetISSQN":80,
        "Valor_Bruto":110, "Valor_Liquido":110,
        "IBS_vBC":90, "IBS_pIBSUF":80, "CBS_pCBS":70,
        "Valor_PIS":80,"Valor_COFINS":80,"Valor_IRRF":80,"Valor_INSS":80,
        "Discriminacao":280,"Arquivo_Origem":200,
    }

    def __init__(self, master, df, on_exportar):
        super().__init__(master)
        self.title(f"NFS-e  —  Planilha de Notas de Serviço  ({len(df)} notas)")
        self.geometry("1650x900")
        self.resizable(True, True)
        self.configure(fg_color=C_FUNDO)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # ── Topbar ──
        top = tk.Frame(self, bg=C_PRIM, height=48)
        top.grid(row=0, column=0, sticky="ew")
        top.grid_propagate(False)
        top.grid_columnconfigure(0, weight=1)
        tk.Label(top, text="  GERENCIAMENTO DE NOTAS FISCAIS DE SERVIÇO  —  NFS-e",
                 bg=C_PRIM, fg="white",
                 font=("Segoe UI",12,"bold")).grid(row=0,column=0,sticky="w",padx=12,pady=12)
        tk.Label(top, text=f"{len(df)} notas  |  "
                 f"Total: R$ {df['Valor_Bruto'].apply(lambda x: float(x) if x else 0).sum():,.2f}",
                 bg=C_PRIM, fg="#aed6f1",
                 font=("Segoe UI",9)).grid(row=0,column=1,sticky="e",padx=16)

        # ── Barra de filtros ──
        filtros = tk.Frame(self, bg=C_FUNDO2,
                           highlightbackground=C_BORDA, highlightthickness=1)
        filtros.grid(row=1, column=0, sticky="ew", padx=0, pady=0)
        tk.Label(filtros, text="  Filtrar:", bg=C_FUNDO2, fg=C_TEXTO2,
                 font=("Segoe UI",9,"bold")).pack(side="left", padx=(10,4), pady=8)
        self._filtro_var = tk.StringVar()
        entry = tk.Entry(filtros, textvariable=self._filtro_var,
                         font=("Segoe UI",10), width=35,
                         relief="solid", bd=1)
        entry.pack(side="left", padx=4, pady=6)
        entry.bind("<KeyRelease>", lambda e: self._aplicar_filtro(df))
        tk.Button(filtros, text="Limpar", command=lambda: [self._filtro_var.set(""), self._aplicar_filtro(df)],
                  bg=C_BORDA, fg=C_TEXTO, font=("Segoe UI",9),
                  relief="flat", padx=10, pady=3).pack(side="left", padx=4)

        # totais rápidos
        total_val = df["Valor_Bruto"].apply(lambda x: float(x) if x else 0).sum()
        total_iss = df["Valor_ISS"].apply(lambda x: float(x) if x else 0).sum()
        for txt, val, cor in [
            (f"Total Bruto: R$ {total_val:,.2f}", None, C_PRIM),
            (f"Total ISS:   R$ {total_iss:,.2f}", None, C_ERR),
            (f"Qtd Notas: {len(df)}", None, C_OK),
        ]:
            tk.Label(filtros, text=f"  {txt}  ", bg=C_FUNDO2,
                     fg=cor, font=("Segoe UI",9,"bold")).pack(side="right", padx=6)

        # ── Treeview ──
        style = ttk.Style()
        style.theme_use("default")
        style.configure("NFS.Treeview",
            background=C_FUNDO2, foreground=C_TEXTO,
            rowheight=22, fieldbackground=C_FUNDO2,
            font=("Segoe UI",9))
        style.configure("NFS.Treeview.Heading",
            background=C_PRIM, foreground="white",
            font=("Segoe UI",9,"bold"), relief="flat")
        style.map("NFS.Treeview",
            background=[("selected","#d6eaf8")],
            foreground=[("selected",C_PRIM)])

        tree_frame = tk.Frame(self, bg=C_FUNDO)
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=8, pady=(4,0))
        tree_frame.grid_columnconfigure(0, weight=1)
        tree_frame.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_frame, columns=self._COLUNAS,
                                  show="headings", style="NFS.Treeview",
                                  selectmode="extended")
        for col in self._COLUNAS:
            self.tree.heading(col, text=col, command=lambda c=col: self._ordenar(df,c))
            self.tree.column(col, width=self._LARG.get(col,90), anchor=tk.W, minwidth=30)

        self.tree.tag_configure("auth",   background="#d5f5e3")  # autorizada
        self.tree.tag_configure("cancel", background="#fadbd8")  # cancelada
        self.tree.tag_configure("par",    background="#eaf2fb")
        self.tree.tag_configure("impar",  background=C_FUNDO2)

        sy = ttk.Scrollbar(tree_frame, orient="vertical",   command=self.tree.yview)
        sx = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")

        self._df_original = df
        self._preencher(df)

        # ── Rodapé ──
        rod = tk.Frame(self, bg=C_FUNDO, height=50,
                       highlightbackground=C_BORDA, highlightthickness=1)
        rod.grid(row=3, column=0, sticky="ew")
        rod.grid_propagate(False)
        self._lbl_sel = tk.Label(rod, text="Nenhum item selecionado",
                                 bg=C_FUNDO, fg=C_TEXTO2, font=("Segoe UI",9))
        self._lbl_sel.pack(side="left", padx=12, pady=12)
        self.tree.bind("<<TreeviewSelect>>", self._on_sel)

        for txt, cmd, bg, hv in [
            ("Copiar Seleção", self._copiar, C_PRIM,   C_SEC),
            ("Exportar CSV",   on_exportar,  C_OK,     "#1a7a40"),
            ("Dashboard",      self._dashboard, C_SEC, C_PRIM),
            ("Fechar",         self.destroy,  C_ERR,   "#a93226"),
        ]:
            tk.Button(rod, text=txt, command=cmd,
                      bg=bg, fg="white", font=FONTE_BTN,
                      relief="flat", padx=14, pady=5, cursor="hand2",
                      activebackground=hv, activeforeground="white",
                      bd=0).pack(side="right", padx=6, pady=8)

    def _preencher(self, df):
        self.tree.delete(*self.tree.get_children())
        for i, (_, row) in enumerate(df.iterrows(), 1):
            val_bruto = row.get("Valor_Bruto","0") or "0"
            try:
                val_f = f"R$ {float(val_bruto):,.2f}"
            except Exception:
                val_f = val_bruto

            situacao = "Autorizada"  # NFS-e extraídas estão autorizadas por padrão
            tag = "auth" if i % 2 == 0 else "auth"
            # se houvesse campo situação poderíamos distinguir canceladas

            vals = [i, situacao,
                    row.get("Numero_NFSe",""),
                    str(row.get("Data_Emissao",""))[:19],
                    row.get("Formato",""),
                    str(row.get("Nome_Prestador",""))[:35],
                    row.get("CNPJ_Prestador",""),
                    row.get("Mun_Prestador",""),
                    str(row.get("Nome_Tomador",""))[:30],
                    row.get("CNPJ_Tomador",""),
                    str(row.get("Desc_Servico",""))[:40],
                    row.get("cTribNac",""),
                    row.get("BC_ISS",""),
                    row.get("Aliq_ISS",""),
                    row.get("Valor_ISS",""),
                    row.get("tpRetISSQN",""),
                    val_f,
                    row.get("Valor_Liquido",""),
                    row.get("IBS_vBC",""),
                    row.get("IBS_pIBSUF",""),
                    row.get("CBS_pCBS",""),
                    row.get("Valor_PIS",""),
                    row.get("Valor_COFINS",""),
                    row.get("Valor_IRRF",""),
                    row.get("Valor_INSS",""),
                    str(row.get("Discriminacao",""))[:60],
                    row.get("Arquivo_Origem",""),
                    ]
            tag_row = "auth" if i % 2 == 0 else "par"
            self.tree.insert("", tk.END, tags=(tag_row,), values=vals)

    def _aplicar_filtro(self, df):
        termo = self._filtro_var.get().lower()
        if not termo:
            self._preencher(df)
            return
        mask = df.apply(lambda row: any(termo in str(v).lower() for v in row), axis=1)
        self._preencher(df[mask])

    def _ordenar(self, df, col):
        pass  # poderia implementar sort

    def _on_sel(self, event):
        sel = self.tree.selection()
        if sel:
            vals = [self.tree.item(s,"values") for s in sel]
            total = 0
            for v in vals:
                try:
                    total += float(str(v[16]).replace("R$","").replace(",","").strip())
                except Exception:
                    pass
            self._lbl_sel.configure(
                text=f"{len(sel)} item(ns) selecionado(s)  |  Total: R$ {total:,.2f}")
        else:
            self._lbl_sel.configure(text="Nenhum item selecionado")

    def _copiar(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Info","Nenhum item selecionado.")
            return
        linhas = ["\t".join(str(v) for v in self.tree.item(s,"values")) for s in sel]
        self.clipboard_clear()
        self.clipboard_append("\n".join(linhas))

    def _dashboard(self):
        JanelaDashboard(self, self._df_original)


# ══════════════════════════════════════════════════════════════════════════════
# JANELA DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════

class JanelaDashboard(ctk.CTkToplevel):

    def __init__(self, master, df):
        super().__init__(master)
        self.title("Dashboard  —  Análise NFS-e")
        self.geometry("1100x780")
        self.resizable(True, True)
        self.configure(fg_color=C_FUNDO)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Topbar
        top = tk.Frame(self, bg=C_PRIM, height=48)
        top.grid(row=0, column=0, sticky="ew")
        top.grid_propagate(False)
        top.grid_columnconfigure(0, weight=1)
        tk.Label(top, text="  DASHBOARD  —  NFS-e", bg=C_PRIM, fg="white",
                 font=("Segoe UI",13,"bold")).grid(row=0,column=0,sticky="w",padx=14,pady=12)

        area = tk.Frame(self, bg=C_FUNDO)
        area.grid(row=1, column=0, sticky="nsew", padx=14, pady=12)
        for c in range(3):
            area.grid_columnconfigure(c, weight=1)
        for r in range(4):
            area.grid_rowconfigure(r, weight=1 if r >= 2 else 0)

        # Calcular métricas
        def to_float(s):
            try: return float(str(s).replace(",","").strip()) if s else 0
            except: return 0

        total_bruto  = df["Valor_Bruto"].apply(to_float).sum()
        total_liq    = df["Valor_Liquido"].apply(to_float).sum()
        total_iss    = df["Valor_ISS"].apply(to_float).sum()
        qtd_notas    = len(df)
        ticket_medio = total_bruto / qtd_notas if qtd_notas else 0
        qtd_prest    = df["CNPJ_Prestador"].nunique()

        # ── Cards KPI ──
        def kpi(parent, row, col, titulo, valor, cor_top, sub=""):
            f = tk.Frame(parent, bg=C_FUNDO2,
                         highlightbackground=C_BORDA, highlightthickness=1)
            f.grid(row=row, column=col, padx=6, pady=6, sticky="nsew", ipady=4)
            tk.Frame(f, bg=cor_top, height=5).pack(fill="x")
            tk.Label(f, text=titulo, bg=C_FUNDO2, fg=C_TEXTO2,
                     font=("Segoe UI",9,"bold")).pack(anchor="w",padx=12,pady=(8,0))
            tk.Label(f, text=valor, bg=C_FUNDO2, fg=C_TEXTO,
                     font=("Segoe UI",18,"bold")).pack(anchor="w",padx=12)
            if sub:
                tk.Label(f, text=sub, bg=C_FUNDO2, fg=C_TEXTO2,
                         font=("Segoe UI",8)).pack(anchor="w",padx=12,pady=(0,6))

        kpi(area,0,0,"TOTAL BRUTO",       f"R$ {total_bruto:,.2f}",  C_PRIM,  f"{qtd_notas} notas")
        kpi(area,0,1,"TOTAL LÍQUIDO",     f"R$ {total_liq:,.2f}",   C_OK,    f"Dedução: R$ {total_bruto-total_liq:,.2f}")
        kpi(area,0,2,"TOTAL ISS",         f"R$ {total_iss:,.2f}",   C_ERR,   f"Média por nota: R$ {total_iss/qtd_notas if qtd_notas else 0:,.2f}")
        kpi(area,1,0,"QTD NOTAS",         str(qtd_notas),            C_SEC,   "NFS-e importadas")
        kpi(area,1,1,"TICKET MÉDIO",      f"R$ {ticket_medio:,.2f}", C_ACENT, "Valor médio por nota")
        kpi(area,1,2,"PRESTADORES",       str(qtd_prest),            C_INFO,  "CNPJs distintos")

        # ── Top 10 Prestadores ──
        frame_prest = tk.Frame(area, bg=C_FUNDO2,
                               highlightbackground=C_BORDA, highlightthickness=1)
        frame_prest.grid(row=2, column=0, columnspan=2, padx=6, pady=6, sticky="nsew")
        frame_prest.grid_columnconfigure(0, weight=1)
        frame_prest.grid_rowconfigure(1, weight=1)

        tk.Frame(frame_prest, bg=C_PRIM, height=28).grid(row=0,column=0,sticky="ew")
        tk.Label(frame_prest, text="  TOP PRESTADORES (VALOR)", bg=C_PRIM, fg="white",
                 font=("Segoe UI",9,"bold")).place(x=8,y=6)

        cols_p = ("Prestador","CNPJ","Qtd","Total Bruto","Total ISS")
        tree_p = ttk.Treeview(frame_prest, columns=cols_p, show="headings", height=10)
        for c,w in zip(cols_p,[220,120,50,110,100]):
            tree_p.heading(c,text=c)
            tree_p.column(c,width=w,anchor=tk.W)
        tree_p.tag_configure("par", background="#eaf2fb")

        top10 = (df.groupby(["Nome_Prestador","CNPJ_Prestador"])
                   .agg(Qtd=("Numero_NFSe","count"),
                        Bruto=("Valor_Bruto", lambda x: x.apply(to_float).sum()),
                        ISS=("Valor_ISS",   lambda x: x.apply(to_float).sum()))
                   .sort_values("Bruto", ascending=False)
                   .head(10).reset_index())

        for i, row in top10.iterrows():
            tag = "par" if i%2==0 else ""
            tree_p.insert("","end",tags=(tag,),values=(
                str(row["Nome_Prestador"])[:30],
                row["CNPJ_Prestador"],
                row["Qtd"],
                f"R$ {row['Bruto']:,.2f}",
                f"R$ {row['ISS']:,.2f}",
            ))
        tree_p.grid(row=1,column=0,sticky="nsew",padx=4,pady=(32,4))

        # ── Por Formato ──
        frame_fmt = tk.Frame(area, bg=C_FUNDO2,
                             highlightbackground=C_BORDA, highlightthickness=1)
        frame_fmt.grid(row=2, column=2, padx=6, pady=6, sticky="nsew")
        frame_fmt.grid_columnconfigure(0, weight=1)
        frame_fmt.grid_rowconfigure(1, weight=1)

        tk.Frame(frame_fmt, bg=C_SEC, height=28).grid(row=0,column=0,sticky="ew")
        tk.Label(frame_fmt, text="  POR FORMATO", bg=C_SEC, fg="white",
                 font=("Segoe UI",9,"bold")).place(x=8,y=6)

        cols_f = ("Formato","Qtd","Total")
        tree_f = ttk.Treeview(frame_fmt, columns=cols_f, show="headings", height=5)
        for c,w in zip(cols_f,[130,60,120]):
            tree_f.heading(c,text=c); tree_f.column(c,width=w,anchor=tk.W)

        por_fmt = (df.groupby("Formato")
                     .agg(Qtd=("Numero_NFSe","count"),
                          Total=("Valor_Bruto", lambda x: x.apply(to_float).sum()))
                     .reset_index())
        for i, row in por_fmt.iterrows():
            tree_f.insert("","end",values=(row["Formato"],row["Qtd"],f"R$ {row['Total']:,.2f}"))
        tree_f.grid(row=1,column=0,sticky="nsew",padx=4,pady=(32,4))

        # ── Por Mês ──
        frame_mes = tk.Frame(area, bg=C_FUNDO2,
                             highlightbackground=C_BORDA, highlightthickness=1)
        frame_mes.grid(row=3, column=0, columnspan=3, padx=6, pady=6, sticky="nsew")
        frame_mes.grid_columnconfigure(0, weight=1)
        frame_mes.grid_rowconfigure(1, weight=1)

        tk.Frame(frame_mes, bg=C_OK, height=28).grid(row=0,column=0,sticky="ew")
        tk.Label(frame_mes, text="  EVOLUÇÃO MENSAL", bg=C_OK, fg="white",
                 font=("Segoe UI",9,"bold")).place(x=8,y=6)

        cols_m = ("Mês/Ano","Qtd Notas","Total Bruto","Total ISS","Total Líquido")
        tree_m = ttk.Treeview(frame_mes, columns=cols_m, show="headings", height=6)
        for c,w in zip(cols_m,[100,80,130,110,120]):
            tree_m.heading(c,text=c); tree_m.column(c,width=w,anchor=tk.W)
        tree_m.tag_configure("par",background="#eaf2fb")

        df2 = df.copy()
        df2["_mes"] = df2["Data_Emissao"].apply(
            lambda x: str(x)[:7] if x else "????-??")
        por_mes = (df2.groupby("_mes")
                      .agg(Qtd=("Numero_NFSe","count"),
                           Bruto=("Valor_Bruto",  lambda x: x.apply(to_float).sum()),
                           ISS=("Valor_ISS",      lambda x: x.apply(to_float).sum()),
                           Liq=("Valor_Liquido",  lambda x: x.apply(to_float).sum()))
                      .sort_index().reset_index())
        for i, row in por_mes.iterrows():
            tag = "par" if i%2==0 else ""
            tree_m.insert("","end",tags=(tag,),values=(
                row["_mes"],row["Qtd"],
                f"R$ {row['Bruto']:,.2f}",
                f"R$ {row['ISS']:,.2f}",
                f"R$ {row['Liq']:,.2f}",
            ))
        tree_m.grid(row=1,column=0,sticky="nsew",padx=4,pady=(32,4))


# ══════════════════════════════════════════════════════════════════════════════
# JANELA VISUALIZAÇÃO NF-e
# ══════════════════════════════════════════════════════════════════════════════

class JanelaVisualizacaoNFe(ctk.CTkToplevel):

    _COLUNAS = [
        "#","Chave_NFe","Numero_NFe","Serie_NFe","NatOp","Data_Emissao",
        "Nome_Emitente","UF_Emitente","Nome_Destinatario","UF_Destinatario",
        "Item","cProd","xProd","NCM","CFOP","qCom","vProd",
        "ICMS_CST","ICMS_vBC","ICMS_pICMS","ICMS_vICMS",
        "PIS_CST","PIS_pPIS","PIS_vPIS",
        "COFINS_CST","COFINS_pCOFINS","COFINS_vCOFINS",
        "IBS_CST","IBS_vIBS","CBS_vCBS",
    ]
    _LARG = {
        "#":40,"Chave_NFe":160,"Numero_NFe":70,"Serie_NFe":50,"NatOp":160,
        "Data_Emissao":130,"Nome_Emitente":180,"UF_Emitente":40,
        "Nome_Destinatario":180,"UF_Destinatario":40,
        "Item":40,"cProd":100,"xProd":200,"NCM":80,"CFOP":55,
        "qCom":70,"vProd":90,
        "ICMS_CST":65,"ICMS_vBC":90,"ICMS_pICMS":80,"ICMS_vICMS":90,
        "PIS_CST":60,"PIS_pPIS":75,"PIS_vPIS":90,
        "COFINS_CST":75,"COFINS_pCOFINS":100,"COFINS_vCOFINS":100,
        "IBS_CST":65,"IBS_vIBS":75,"CBS_vCBS":80,
    }

    def __init__(self, master, df, on_exportar, on_copiar):
        super().__init__(master)
        self.title(f"NF-e  —  Produtos  ({len(df)} registros)")
        self.geometry("1600x860")
        self.resizable(True, True)
        self.configure(fg_color=C_FUNDO)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        top = tk.Frame(self, bg=C_PRIM, height=48)
        top.grid(row=0, column=0, sticky="ew")
        top.grid_propagate(False)
        top.grid_columnconfigure(0, weight=1)
        tk.Label(top, text="  NF-e  —  PRODUTOS E IMPOSTOS",
                 bg=C_PRIM, fg="white",
                 font=("Segoe UI",12,"bold")).grid(row=0,column=0,sticky="w",padx=12,pady=12)
        tk.Label(top, text=f"{len(df)} registros",
                 bg=C_PRIM, fg="#aed6f1",
                 font=("Segoe UI",9)).grid(row=0,column=1,sticky="e",padx=12)

        style = ttk.Style()
        style.configure("NFe.Treeview",
            background=C_FUNDO2, foreground=C_TEXTO,
            rowheight=22, fieldbackground=C_FUNDO2, font=("Segoe UI",9))
        style.configure("NFe.Treeview.Heading",
            background=C_PRIM, foreground="white",
            font=("Segoe UI",9,"bold"), relief="flat")
        style.map("NFe.Treeview",
            background=[("selected",C_SEC)],
            foreground=[("selected","white")])

        tf = tk.Frame(self, bg=C_FUNDO)
        tf.grid(row=1, column=0, padx=8, pady=6, sticky="nsew")
        tf.grid_columnconfigure(0, weight=1)
        tf.grid_rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(tf, columns=self._COLUNAS,
                                  show="headings", style="NFe.Treeview")
        for col in self._COLUNAS:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=self._LARG.get(col,90), anchor=tk.W, minwidth=30)
        self.tree.tag_configure("par",  background="#eaf2fb")
        self.tree.tag_configure("impar",background=C_FUNDO2)

        for i, (_, row) in enumerate(df.iterrows(), 1):
            tag = "par" if i%2==0 else "impar"
            vals = [i] + [str(row.get(c,""))[:50] for c in self._COLUNAS[1:]]
            self.tree.insert("","end",tags=(tag,),values=vals)

        sy = ttk.Scrollbar(tf, orient="vertical",   command=self.tree.yview)
        sx = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.tree.grid(row=0,column=0,sticky="nsew")
        sy.grid(row=0,column=1,sticky="ns")
        sx.grid(row=1,column=0,sticky="ew")

        rod = tk.Frame(self, bg=C_FUNDO, height=50,
                       highlightbackground=C_BORDA, highlightthickness=1)
        rod.grid(row=2, column=0, sticky="ew")
        rod.grid_propagate(False)
        for txt,cmd,bg,hv in [
            ("Copiar Seleção", lambda: on_copiar(self.tree), C_PRIM, C_SEC),
            ("Exportar CSV",   on_exportar, C_OK, "#1a7a40"),
            ("Fechar",         self.destroy, C_ERR, "#a93226"),
        ]:
            tk.Button(rod, text=txt, command=cmd, bg=bg, fg="white",
                      font=FONTE_BTN, relief="flat", padx=14, pady=5,
                      cursor="hand2", activebackground=hv,
                      activeforeground="white", bd=0).pack(side="right",padx=6,pady=8)


# ══════════════════════════════════════════════════════════════════════════════
# APLICAÇÃO PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

class AplicacaoLeitorXML:

    def __init__(self):
        ctk.set_appearance_mode("light")
        self.janela = ctk.CTk()
        self.janela.title("GCON/SIAN  —  NF-e / NFS-e")
        self.janela.geometry("1400x900")
        self.janela.minsize(1200,750)
        self.janela.configure(fg_color=C_FUNDO)
        self.janela.protocol("WM_DELETE_WINDOW", self._fechar)
        self.janela.grid_columnconfigure(1, weight=1)
        self.janela.grid_rowconfigure(0, weight=1)

        self.processando = False
        self.arquivos_selecionados: list[str] = []

        locks = verificar_locks_ativos()
        if locks:
            if not messagebox.askyesno("Sessoes Ativas",
                f"Existem {len(locks)} sessao(oes) ativa(s).\nDeseja continuar?"):
                sys.exit(0)

        if not inicializar_sessao():
            messagebox.showerror("Erro","Nao foi possivel inicializar a sessao!")
            return

        self._construir_interface()
        self._log_inicial()

    # ── SIDEBAR ───────────────────────────────────────────────────────────────

    def _construir_interface(self):
        self._construir_sidebar()
        self._construir_area_principal()

    def _construir_sidebar(self):
        sb = tk.Frame(self.janela, bg=C_SIDE, width=300)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)

        logo = tk.Frame(sb, bg=C_PRIM, height=90)
        logo.pack(fill="x")
        logo.pack_propagate(False)
        tk.Label(logo, text="GCON / SIAN", bg=C_PRIM, fg="white",
                 font=("Segoe UI",17,"bold")).pack(pady=(18,2))
        tk.Label(logo, text="NF-e  |  NFS-e", bg=C_PRIM,
                 fg="#aed6f1", font=("Segoe UI",10)).pack()
        tk.Frame(sb, bg=C_ACENT, height=3).pack(fill="x")

        usr = tk.Frame(sb, bg=C_SIDE, pady=14)
        usr.pack(fill="x", padx=16)
        tk.Label(usr, text="USUÁRIO ATIVO", bg=C_SIDE,
                 fg="#7f8c8d", font=("Segoe UI",8,"bold")).pack(anchor="w")
        tk.Label(usr, text=cfg.USUARIO_ID, bg=C_SIDE,
                 fg="white", font=("Segoe UI",12,"bold")).pack(anchor="w")
        tk.Label(usr, text=f"Sessão: {cfg.SESSAO_ID[:14]}…", bg=C_SIDE,
                 fg="#7f8c8d", font=("Segoe UI",9)).pack(anchor="w")
        tk.Frame(sb, bg=C_SIDE2, height=1).pack(fill="x", padx=16)

        def btn(texto, icone, cmd):
            f = tk.Frame(sb, bg=C_SIDE, cursor="hand2")
            f.pack(fill="x")
            lbl = tk.Label(f, text=f"   {icone}   {texto}",
                           bg=C_SIDE, fg="#bdc3c7",
                           font=("Segoe UI",11), anchor="w", padx=10, pady=11)
            lbl.pack(fill="x")
            def on_e(e): lbl.configure(bg=C_SIDE2,fg="white"); f.configure(bg=C_SIDE2)
            def on_l(e): lbl.configure(bg=C_SIDE,fg="#bdc3c7"); f.configure(bg=C_SIDE)
            for w in (lbl,f):
                w.bind("<Enter>",on_e); w.bind("<Leave>",on_l)
                w.bind("<Button-1>", lambda e,c=cmd: c())

        def sec(t):
            tk.Label(sb, text=f"  {t}", bg=C_SIDE, fg="#7f8c8d",
                     font=("Segoe UI",8,"bold")).pack(anchor="w",padx=16,pady=(14,2))

        sec("IMPORTAR NF-e / NFS-e")
        btn("Selecionar 1 XML",       "▶",  self._selecionar_um)
        btn("Selecionar Vários XMLs", "▶▶", self._selecionar_varios)

        tk.Frame(sb, bg=C_SIDE2, height=1).pack(fill="x",padx=16,pady=4)
        sec("VISUALIZAR")
        btn("Ver NF-e (Produtos)",    "≡",  self._visualizar_nfe)
        btn("Ver NFS-e (Planilha)",   "⊟",  self._visualizar_nfse)
        btn("Dashboard NFS-e",        "◉",  self._abrir_dashboard)
        btn("Abrir Excel Local",      "⊞",  self._abrir_excel_local)
        btn("Sincronizar",            "⟳",  self._sincronizar_manual)

        tk.Frame(sb, bg=C_SIDE2, height=1).pack(fill="x",padx=16,pady=4)
        sec("LOG")
        btn("Limpar Log", "✕", self._limpar_log)
        btn("Salvar Log",  "↓", self._salvar_log)

        tk.Frame(sb, bg=C_SIDE).pack(fill="both", expand=True)

        rod = tk.Frame(sb, bg=C_ERR, cursor="hand2")
        rod.pack(fill="x", side="bottom")
        lf = tk.Label(rod, text="   ✕   FECHAR SISTEMA",
                      bg=C_ERR, fg="white",
                      font=("Segoe UI",11,"bold"), pady=16)
        lf.pack(fill="x")
        for w in (rod,lf):
            w.bind("<Button-1>", lambda e: self._fechar())
            w.bind("<Enter>", lambda e: [rod.configure(bg="#a93226"),lf.configure(bg="#a93226")])
            w.bind("<Leave>", lambda e: [rod.configure(bg=C_ERR),lf.configure(bg=C_ERR)])

    # ── ÁREA PRINCIPAL ────────────────────────────────────────────────────────

    def _construir_area_principal(self):
        area = tk.Frame(self.janela, bg=C_FUNDO)
        area.grid(row=0, column=1, sticky="nsew")
        area.grid_columnconfigure(0, weight=1)
        area.grid_rowconfigure(2, weight=1)

        top = tk.Frame(area, bg=C_PRIM, height=52)
        top.grid(row=0, column=0, sticky="ew")
        top.grid_propagate(False)
        top.grid_columnconfigure(0, weight=1)
        tk.Label(top, text="  SISTEMA DE EXTRAÇÃO  —  NF-e / NFS-e  (MULTIUSUÁRIO)",
                 bg=C_PRIM, fg="white",
                 font=("Segoe UI",13,"bold")).grid(row=0,column=0,sticky="w",padx=16,pady=14)
        tk.Label(top, text=datetime.now().strftime("%d/%m/%Y  %H:%M"),
                 bg=C_PRIM, fg="#aed6f1",
                 font=("Segoe UI",9)).grid(row=0,column=1,sticky="e",padx=16)
        tk.Frame(area, bg=C_ACENT, height=3).grid(row=0,column=0,sticky="sew")

        cards = tk.Frame(area, bg=C_FUNDO)
        cards.grid(row=1, column=0, sticky="ew", padx=16, pady=12)
        for i in range(6): cards.grid_columnconfigure(i, weight=1)

        def card(col, titulo, valor, cor):
            f = tk.Frame(cards, bg=C_FUNDO2,
                         highlightbackground=C_BORDA, highlightthickness=1)
            f.grid(row=0, column=col, padx=4, sticky="nsew", ipady=4)
            tk.Frame(f, bg=cor, height=4).pack(fill="x")
            tk.Label(f, text=titulo, bg=C_FUNDO2, fg=C_TEXTO2,
                     font=("Segoe UI",8,"bold")).pack(anchor="w",padx=10,pady=(5,0))
            lbl = tk.Label(f, text=valor, bg=C_FUNDO2, fg=C_TEXTO,
                           font=("Segoe UI",15,"bold"))
            lbl.pack(anchor="w",padx=10,pady=(0,5))
            return lbl

        self._c_status   = card(0,"STATUS",           "PRONTO",                C_OK)
        self._c_arqs     = card(1,"SELECIONADOS",     "0",                     C_SEC)
        self._c_proc     = card(2,"PROCESSADOS",      "0",                     C_INFO)
        self._c_nfe      = card(3,"NF-e PRODUTOS",    str(total_registros()),  C_PRIM)
        self._c_nfse     = card(4,"NFS-e SERVIÇOS",   "0",                     C_ACENT)
        self._c_prog     = card(5,"PROGRESSO",        "0%",                    C_WARN)

        prog = tk.Frame(area, bg=C_FUNDO)
        prog.grid(row=1, column=0, sticky="sew", padx=16, pady=(0,4))
        prog.grid_columnconfigure(0, weight=1)
        self._prog_var = tk.DoubleVar()
        self._prog_bar = ctk.CTkProgressBar(prog, variable=self._prog_var,
                                             height=8, corner_radius=4,
                                             fg_color=C_BORDA, progress_color=C_SEC)
        self._prog_bar.grid(row=0,column=0,sticky="ew",pady=2)
        self._prog_bar.set(0)

        log_outer = tk.Frame(area, bg=C_FUNDO2,
                             highlightbackground=C_BORDA, highlightthickness=1)
        log_outer.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0,12))
        log_outer.grid_columnconfigure(0, weight=1)
        log_outer.grid_rowconfigure(1, weight=1)

        log_hdr = tk.Frame(log_outer, bg=C_PRIM, height=34)
        log_hdr.grid(row=0, column=0, sticky="ew")
        log_hdr.grid_propagate(False)
        tk.Label(log_hdr, text="  LOG DE PROCESSAMENTO",
                 bg=C_PRIM, fg="white",
                 font=("Segoe UI",9,"bold")).pack(side="left",padx=10,pady=7)
        # Badge NF-e / NFS-e
        self._lbl_badge_nfe  = tk.Label(log_hdr, text="  NF-e: 0  ", bg="#1a5276", fg="white",
                                         font=("Segoe UI",8,"bold"))
        self._lbl_badge_nfe.pack(side="right", padx=4, pady=6)
        self._lbl_badge_nfse = tk.Label(log_hdr, text="  NFS-e: 0  ", bg="#e67e22", fg="white",
                                         font=("Segoe UI",8,"bold"))
        self._lbl_badge_nfse.pack(side="right", padx=4, pady=6)

        self.txt_log = scrolledtext.ScrolledText(
            log_outer, wrap=tk.WORD, font=FONTE_LOG,
            bg="#1c2833", fg="#d5d8dc",
            insertbackground="white", relief="flat",
            selectbackground=C_SEC)
        self.txt_log.grid(row=1, column=0, sticky="nsew", padx=1, pady=1)

        for tag, fg_cor, bold in [
            ("success","#2ecc71",True),("error","#e74c3c",True),
            ("warning","#f39c12",True),("info","#5dade2",False),
            ("value","#aab7b8",False),("file","#5dade2",True),
            ("timestamp","#566573",False),("border","#2c3e50",False),
            ("nfe","#5dade2",True),("nfse","#f39c12",True),
        ]:
            self.txt_log.tag_config(tag, foreground=fg_cor,
                font=(FONTE_LOG[0],FONTE_LOG[1],"bold" if bold else "normal"))

    # ── LOG HELPERS ───────────────────────────────────────────────────────────

    def _ts(self): return datetime.now().strftime("%H:%M:%S")

    def log(self, msg, tag="info"):
        self.txt_log.insert(tk.END, f"[{self._ts()}] ","timestamp")
        self.txt_log.insert(tk.END, f"{msg}\n", tag)
        self.txt_log.see(tk.END)
        self.janela.update_idletasks()

    def _divider(self, char="─"):
        self.txt_log.insert(tk.END, f"{char*100}\n","border")
        self.txt_log.see(tk.END)

    def _centered(self, msg, tag="info"):
        pad = max(0,(100-len(msg))//2)
        self.txt_log.insert(tk.END, f"{' '*pad}{msg}\n", tag)
        self.txt_log.see(tk.END)

    def _log_inicial(self):
        self.txt_log.delete(1.0, tk.END)
        self._divider("=")
        self._centered("GCON/SIAN  —  NF-e / NFS-e  —  MULTIUSUARIO")
        self._centered(f"Sessao: {cfg.SESSAO_ID}  |  Usuario: {cfg.USUARIO_ID}")
        self._centered(datetime.now().strftime("%d/%m/%Y  %H:%M:%S"))
        self._divider("=")
        self.log("Sistema iniciado. Pronto para processar NF-e e NFS-e.","success")
        self.log(f"Pasta compartilhada: {cfg.PASTA_BASE}","info")
        self.log(f"NF-e na sessao    : {total_registros()}","info")
        self.log("")

    # ── STATUS / PROGRESSO ────────────────────────────────────────────────────

    def _set_status(self, t, cor=C_OK):
        self._c_status.configure(text=t, fg=cor)

    def _set_progresso(self, atual, total):
        if total:
            pct = atual/total
            self._prog_var.set(pct)
            self._c_prog.configure(text=f"{pct*100:.0f}%")
            self._c_proc.configure(text=str(atual))

    # ── SELEÇÃO ───────────────────────────────────────────────────────────────

    def _selecionar_um(self):
        if self.processando:
            messagebox.showwarning("Aguarde","Processamento em andamento!"); return
        arq = filedialog.askopenfilename(
            title="Selecione um XML",
            filetypes=[("XML","*.xml"),("Todos","*.*")])
        if arq:
            self.arquivos_selecionados = [arq]
            self._c_arqs.configure(text="1")
            threading.Thread(target=self._processar, daemon=True).start()

    def _selecionar_varios(self):
        if self.processando:
            messagebox.showwarning("Aguarde","Processamento em andamento!"); return
        arqs = filedialog.askopenfilenames(
            title="Selecione XMLs",
            filetypes=[("XML","*.xml"),("Todos","*.*")])
        if arqs:
            self.arquivos_selecionados = list(arqs)
            self._c_arqs.configure(text=str(len(arqs)))
            threading.Thread(target=self._processar, daemon=True).start()

    # ── PIPELINE ETL ──────────────────────────────────────────────────────────

    def _processar(self):
        self.processando = True
        self._set_status("PROCESSANDO...", C_WARN)

        total_arqs = len(self.arquivos_selecionados)
        erros, add_nfe, add_nfse = [], 0, 0
        lote_nfe, lote_nfse = [], []
        cnt_nfe, cnt_nfse = 0, 0

        self.log(""); self._divider("=")
        self._centered("INICIO DO PROCESSAMENTO")
        self._divider("=")
        self.log(f"Arquivos: {total_arqs}","info")
        self._set_progresso(0, total_arqs)

        chaves_nfe  = carregar_chaves_existentes(cfg.CSV_TEMP)
        try:
            chaves_nfse = carregar_chaves_existentes(cfg.CSV_NFSE_TEMP)
        except Exception:
            chaves_nfse = set()

        for i, caminho in enumerate(self.arquivos_selecionados, 1):
            nome = os.path.basename(caminho)
            tipo = _detectar_tipo(caminho)

            if tipo == "nfse":
                cnt_nfse += 1
                registros, msg = extrair_servicos(caminho)
                if msg.startswith("ERRO"):
                    erros.append(f"{nome}: {msg[6:]}")
                else:
                    novos, _ = filtrar_novos(registros, chaves_nfse)
                    lote_nfse.extend(novos)
                    add_nfse += len(novos)
                    for p in novos:
                        chaves_nfse.add(
                            f"{p.get('Chave_NFSe','')}_{p.get('Numero_NFSe','')}")
            else:
                cnt_nfe += 1
                registros, msg = extrair_produtos(caminho)
                if msg.startswith("ERRO"):
                    erros.append(f"{nome}: {msg[6:]}")
                else:
                    novos, _ = filtrar_novos(registros, chaves_nfe)
                    lote_nfe.extend(novos)
                    add_nfe += len(novos)
                    for p in novos:
                        chaves_nfe.add(
                            f"{p.get('Chave_NFe','')}_{p.get('Item','')}_{p.get('cProd','')}")

            # Salvar lotes
            if len(lote_nfe) >= LOTE_MAX:
                salvar_produtos_csv(lote_nfe)
                lote_nfe = []
            if len(lote_nfse) >= LOTE_MAX:
                self._salvar_nfse_lote(lote_nfse)
                lote_nfse = []

            # Throttle UI
            if i % UI_THROTTLE == 0 or i == total_arqs:
                self._set_progresso(i, total_arqs)
                tag = "nfse" if tipo=="nfse" else "nfe"
                self.log(f"[{i}/{total_arqs}] [{tipo.upper()}] {nome}  —  "
                         f"{len(registros) if not msg.startswith('ERRO') else 0} registros"
                         + (f"  ERRO: {msg[6:40]}" if msg.startswith("ERRO") else ""), tag)
                self._lbl_badge_nfe.configure(text=f"  NF-e: {cnt_nfe}  ")
                self._lbl_badge_nfse.configure(text=f"  NFS-e: {cnt_nfse}  ")
                self.janela.update_idletasks()

        # Salvar restos
        if lote_nfe:  salvar_produtos_csv(lote_nfe)
        if lote_nfse: self._salvar_nfse_lote(lote_nfse)

        if sincronizar_excel_temp():
            self.log("Excel NF-e sincronizado.","success")

        total_nfe_final = total_registros()
        self._c_nfe.configure(text=str(total_nfe_final))
        self._c_nfse.configure(text=str(add_nfse))

        self.log(""); self._divider("=")
        self._centered("RESUMO")
        self._divider("=")
        self.log(f"Arquivos       : {total_arqs}","info")
        self.log(f"NF-e processados : {cnt_nfe}  |  Adicionados: {add_nfe}","nfe")
        self.log(f"NFS-e processados: {cnt_nfse}  |  Adicionados: {add_nfse}","nfse")
        if erros: self.log(f"Erros: {len(erros)}","error")
        self.log(f"Total NF-e na sessao: {total_nfe_final}","info")
        self._divider("=")

        self._set_status("CONCLUIDO", C_OK)
        messagebox.showinfo("Concluido",
            f"NF-e: {cnt_nfe} arqs, {add_nfe} produtos adicionados\n"
            f"NFS-e: {cnt_nfse} arqs, {add_nfse} notas adicionadas\n"
            f"Erros: {len(erros)}")

        self.processando = False
        self.arquivos_selecionados = []
        self._c_arqs.configure(text="0")
        self._set_progresso(0,1)

    def _salvar_nfse_lote(self, lote):
        """Salva lote de NFS-e no CSV temporário de serviços."""
        import csv
        caminho = cfg.CSV_NFSE_TEMP
        existe = os.path.exists(caminho) and os.path.getsize(caminho) > 0
        try:
            with open(caminho,"a",newline="",encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=cfg.CABECALHO_NFSE)
                if not existe:
                    writer.writeheader()
                for reg in lote:
                    writer.writerow({k: reg.get(k,"") for k in cfg.CABECALHO_NFSE})
        except Exception as e:
            self.log(f"Erro ao salvar NFS-e: {e}","error")

    # ── VISUALIZAÇÕES ─────────────────────────────────────────────────────────

    def _visualizar_nfe(self):
        if not os.path.exists(cfg.EXCEL_TEMP):
            messagebox.showwarning("Aviso","Processe NF-e primeiro."); return
        sincronizar_excel_temp()
        df = pd.read_excel(cfg.EXCEL_TEMP, dtype=str)
        if df.empty:
            messagebox.showinfo("Info","Nenhum dado NF-e."); return
        JanelaVisualizacaoNFe(self.janela, df,
                              on_exportar=self._exportar_csv_nfe,
                              on_copiar=self._copiar_selecao)

    def _visualizar_nfse(self):
        if not os.path.exists(cfg.CSV_NFSE_TEMP):
            messagebox.showwarning("Aviso","Processe NFS-e primeiro."); return
        try:
            df = pd.read_csv(cfg.CSV_NFSE_TEMP, dtype=str, encoding="utf-8",
                             on_bad_lines="skip")
        except Exception as e:
            messagebox.showerror("Erro",f"Erro ao ler NFS-e:\n{e}"); return
        if df.empty:
            messagebox.showinfo("Info","Nenhuma NFS-e encontrada."); return
        JanelaPlanilhaNFSe(self.janela, df,
                           on_exportar=self._exportar_csv_nfse)

    def _abrir_dashboard(self):
        if not os.path.exists(cfg.CSV_NFSE_TEMP):
            messagebox.showwarning("Aviso","Processe NFS-e primeiro."); return
        try:
            df = pd.read_csv(cfg.CSV_NFSE_TEMP, dtype=str, encoding="utf-8",
                             on_bad_lines="skip")
        except Exception as e:
            messagebox.showerror("Erro",f"Erro ao ler NFS-e:\n{e}"); return
        if df.empty:
            messagebox.showinfo("Info","Nenhuma NFS-e encontrada."); return
        JanelaDashboard(self.janela, df)

    def _copiar_selecao(self, tree):
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Info","Nenhum item selecionado."); return
        linhas = ["\t".join(str(v) for v in tree.item(s,"values")) for s in sel]
        self.janela.clipboard_clear()
        self.janela.clipboard_append("\n".join(linhas))
        self.log("Selecao copiada.","success")

    def _exportar_csv_nfe(self):
        if not os.path.exists(cfg.EXCEL_TEMP): return
        arq = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV","*.csv")],
            initialfile=f"nfe_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        if arq:
            pd.read_excel(cfg.EXCEL_TEMP,dtype=str).to_csv(
                arq,index=False,encoding="utf-8-sig")
            self.log(f"NF-e exportada: {arq}","success")
            messagebox.showinfo("Sucesso",f"Exportado!\n{arq}")

    def _exportar_csv_nfse(self):
        if not os.path.exists(cfg.CSV_NFSE_TEMP): return
        arq = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV","*.csv")],
            initialfile=f"nfse_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        if arq:
            pd.read_csv(cfg.CSV_NFSE_TEMP,dtype=str,on_bad_lines="skip").to_csv(
                arq,index=False,encoding="utf-8-sig")
            self.log(f"NFS-e exportada: {arq}","success")
            messagebox.showinfo("Sucesso",f"Exportado!\n{arq}")

    # ── DEMAIS ────────────────────────────────────────────────────────────────

    def _sincronizar_manual(self):
        if self.processando:
            messagebox.showwarning("Aguarde","Processamento em andamento!"); return
        ok, msg = sincronizar_com_principal()
        if ok:
            self.log(f"Sincronizado: {msg}","success")
            ok2,msg2 = atualizar_excel_principal()
            self.log(msg2,"success" if ok2 else "warning")
            messagebox.showinfo("Sincronizacao",f"Concluida!\n{msg}")
        else:
            self.log(f"Erro: {msg}","error")
            messagebox.showerror("Erro",f"Falha:\n{msg}")

    def _fechar(self):
        if self.processando:
            if not messagebox.askyesno("Processamento","Deseja fechar mesmo assim?"): return
        ok,msg = sincronizar_com_principal()
        self.log(msg,"success" if ok else "error")
        if ok:
            ok2,msg2 = atualizar_excel_principal()
            self.log(msg2,"success" if ok2 else "warning")
        limpar_temporarios()
        self.janela.destroy()
        sys.exit(0)

    def _abrir_excel_local(self):
        if not os.path.exists(cfg.EXCEL_TEMP):
            messagebox.showwarning("Aviso","Excel NF-e nao encontrado!"); return
        sincronizar_excel_temp()
        os.startfile(cfg.EXCEL_TEMP)

    def _limpar_log(self):
        self.txt_log.delete(1.0, tk.END)
        self.log("Log limpo.","info")

    def _salvar_log(self):
        arq = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Texto","*.txt")],
            initialfile=f"log_{cfg.USUARIO_ID}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        if arq:
            with open(arq,"w",encoding="utf-8") as f:
                f.write(self.txt_log.get(1.0,tk.END))
            self.log(f"Log salvo: {arq}","success")

    def run(self):
        self.janela.mainloop()
