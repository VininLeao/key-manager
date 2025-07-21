# Gerenciador de Chaves v9.5 - Por Vinícius Leão
# coding: utf-8
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import json
import os
import sqlite3
from datetime import datetime, timedelta
import pyperclip
import shutil
from collections import defaultdict
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import threading
import re
import html
import webbrowser # Para a pré-visualização

# --- Biblioteca para gerar PDF ---
try:
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    from reportlab.lib.utils import ImageReader
    PDF_DISPONIVEL = True
except ImportError:
    PDF_DISPONIVEL = False

# --- Biblioteca para importar de XLS (NOVO) ---
try:
    import pandas as pd
    PANDAS_DISPONIVEL = True
except ImportError:
    PANDAS_DISPONIVEL = False

# --- Constantes ---
DB_NAME = "gerenciador.db"
UNDO_FILE = "gerenciador.db.undo"
REDO_FILE = "gerenciador.db.redo"
BACKUP_DIR = "backups"
PDF_DIR = "pdfs"
EMAIL_CONFIG_FILE = "email_config.json"
APP_VERSION = "9.5" # Versão atualizada com a nova funcionalidade

# --- Utilitários ---
def _sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '', filename)

def _converter_markdown_para_html(texto):
    """Converte uma marcação simples (tipo Markdown) para tags HTML que o ReportLab entende."""
    texto = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', texto)
    texto = re.sub(r'\*(.*?)\*', r'<i>\1</i>', texto)
    texto = re.sub(r'__(.*?)__', r'<u>\1</u>', texto)
    texto = texto.replace('\n', '<br/>')
    return texto

# --- Classe Geradora de PDF (integrada para melhor organização) ---
class GeradorPDF:
    def __init__(self, nome_arquivo):
        self.nome_arquivo = nome_arquivo
        self.story = []
        self._setup_estilos()

    def _setup_estilos(self):
        self.estilos = getSampleStyleSheet()
        self.estilos.add(ParagraphStyle(name='HeaderStyle', fontName='Helvetica-Bold', fontSize=16, alignment=TA_CENTER, spaceAfter=15))
        self.estilos.add(ParagraphStyle(name='InfoLabel', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT, textColor=colors.darkblue))
        self.estilos.add(ParagraphStyle(name='InfoText', fontName='Helvetica', fontSize=10, alignment=TA_LEFT))
        self.estilos.add(ParagraphStyle(name='KeyLabel', fontName='Helvetica-Bold', fontSize=12, alignment=TA_CENTER, spaceBefore=10, spaceAfter=5))
        self.estilos.add(ParagraphStyle(name='KeyStyle', fontName='Courier-Bold', fontSize=16, alignment=TA_CENTER, backColor=colors.HexColor('#f0f0f0'), textColor=colors.black, borderRadius=5, borderWidth=1.5, borderColor=colors.black, padding=12, leading=20))
        self.estilos.add(ParagraphStyle(name='InstructionTitleStyle', fontName='Helvetica-Bold', fontSize=14, alignment=TA_CENTER, spaceBefore=20, spaceAfter=10, textColor=colors.HexColor("#094771")))
        self.estilos.add(ParagraphStyle(name='InstructionBody', fontName='Helvetica', fontSize=10, leading=14, leftIndent=20))
        self.estilos.add(ParagraphStyle(name='FooterStyle', fontName='Helvetica-Oblique', fontSize=9, alignment=TA_CENTER, spaceBefore=30, textColor=colors.grey))

    def adicionar_paragrafo(self, texto, estilo='Normal'):
        texto_seguro = html.escape(texto)
        texto_formatado = _converter_markdown_para_html(texto_seguro)
        p = Paragraph(texto_formatado, self.estilos[estilo])
        self.story.append(p)

    def adicionar_imagem(self, caminho_imagem, largura_cm):
        if not (caminho_imagem and os.path.exists(caminho_imagem)): return
        try:
            img_reader = ImageReader(caminho_imagem)
            iw, ih = img_reader.getSize()
            aspect = ih / float(iw) if iw > 0 else 0
            largura = largura_cm * cm
            altura = (largura * aspect) if aspect > 0 else 0
            img = Image(caminho_imagem, width=largura, height=altura)
            img.hAlign = 'CENTER'
            self.story.append(img)
            self.adicionar_espaco_cm(0.8)
        except Exception as e:
            print(f"Erro ao carregar logo: {e}")

    def adicionar_tabela_info(self, dados, col_widths_cm):
        col_widths = [w * cm for w in col_widths_cm]
        tabela_formatada = []
        for rotulo, texto in dados:
            if texto:
                p_rotulo = Paragraph(f"🔹 <b>{rotulo}</b>", self.estilos['InfoText'])
                p_texto = Paragraph(texto, self.estilos['InfoText'])
                tabela_formatada.append([p_rotulo, p_texto])
        if tabela_formatada:
            tabela = Table(tabela_formatada, colWidths=col_widths)
            tabela.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('LEFTPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 2)]))
            self.story.append(tabela)

    def adicionar_espaco_cm(self, altura_cm):
        self.story.append(Spacer(1, altura_cm * cm))
    
    def adicionar_quebra_pagina(self):
        self.story.append(PageBreak())

    def construir(self):
        try:
            doc = SimpleDocTemplate(self.nome_arquivo, topMargin=0.5*inch, bottomMargin=0.5*inch, leftMargin=0.7*inch, rightMargin=0.7*inch)
            doc.build(self.story)
            return True
        except Exception as e:
            messagebox.showerror("Erro de PDF", f"Não foi possível gerar o arquivo PDF.\nErro: {e}")
            logar_acao(f"FALHA ao gerar PDF. Erro: {e}")
            return False

# --- Funções de Banco de Dados e Utilitárias ---
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS chaves (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        chave TEXT NOT NULL UNIQUE,
        categoria TEXT NOT NULL,
        vendida INTEGER NOT NULL DEFAULT 0,
        comprador TEXT,
        data_venda TEXT,
        ordem_manual INTEGER,
        preco_venda_brl REAL,
        preco_venda_usd REAL,
        canal_venda TEXT
    )
    ''')
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS categorias (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL UNIQUE,
        instrucao_pt TEXT,
        instrucao_en TEXT,
        custo_padrao_brl REAL,
        custo_padrao_usd REAL,
        logo_path TEXT,
        info_licenca_pt TEXT, info_licenca_en TEXT,
        info_idioma_pt TEXT, info_idioma_en TEXT,
        info_entrega_pt TEXT, info_entrega_en TEXT,
        layout_pdf_pt TEXT, layout_pdf_en TEXT,
        instrucao_es TEXT,
        info_licenca_es TEXT,
        info_idioma_es TEXT,
        info_entrega_es TEXT,
        layout_pdf_es TEXT
    )
    ''')
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS canais_venda (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL UNIQUE
    )
    ''')
    conn.commit()
    conn.close()

def _adicionar_coluna_se_nao_existir(cursor, tabela, coluna, tipo):
    cursor.execute(f"PRAGMA table_info({tabela})")
    colunas = [info[1] for info in cursor.fetchall()]
    if coluna not in colunas:
        try:
            cursor.execute(f"ALTER TABLE {tabela} ADD COLUMN {coluna} {tipo}")
            print(f"Coluna '{coluna}' adicionada à tabela '{tabela}' com sucesso.")
            return True
        except Exception as e:
            print(f"Erro ao adicionar coluna '{coluna}': {e}")
            messagebox.showerror("Erro de Banco de Dados", f"Não foi possível atualizar o banco de dados para a nova versão.\nErro: {e}")
            return False
    return True

def verificar_e_migrar_schema():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    sucesso = True
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'chaves', 'ordem_manual', 'INTEGER')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'chaves', 'preco_venda_brl', 'REAL')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'chaves', 'preco_venda_usd', 'REAL')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'chaves', 'canal_venda', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'custo_padrao_brl', 'REAL')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'custo_padrao_usd', 'REAL')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'logo_path', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'info_licenca_pt', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'info_licenca_en', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'info_idioma_pt', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'info_idioma_en', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'info_entrega_pt', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'info_entrega_en', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'layout_pdf_pt', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'layout_pdf_en', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'instrucao_es', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'info_licenca_es', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'info_idioma_es', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'info_entrega_es', 'TEXT')
    sucesso &= _adicionar_coluna_se_nao_existir(cursor, 'categorias', 'layout_pdf_es', 'TEXT')
    cursor.execute("SELECT COUNT(*) FROM chaves WHERE ordem_manual IS NULL")
    if cursor.fetchone()[0] > 0: cursor.execute("UPDATE chaves SET ordem_manual = id WHERE ordem_manual IS NULL")
    conn.commit()
    conn.close()
    if not sucesso: exit()

def migrar_de_json_para_sqlite():
    if not os.path.exists("estoque.json") and not os.path.exists("categorias.json"): return
    conn = sqlite3.connect(DB_NAME); cursor = conn.cursor()
    cursor.execute("SELECT COUNT(id) FROM chaves")
    if cursor.fetchone()[0] > 0 and not (os.path.exists("estoque.json") or os.path.exists("categorias.json")):
        conn.close(); return
    if not messagebox.askyesno("Migração de Dados Detectada", "Arquivos .json antigos foram encontrados. Deseja migrar os dados?"):
        conn.close(); return
    if os.path.exists("estoque.json"):
        try:
            with open("estoque.json", "r", encoding="utf-8") as f: estoque_json = json.load(f)
            chaves = [(item['chave'], item.get('categoria', 'S/C'), 1 if item.get('vendida') else 0, item.get('comprador'), item.get('data_venda')) for item in estoque_json]
            cursor.executemany("INSERT OR IGNORE INTO chaves (chave, categoria, vendida, comprador, data_venda) VALUES (?, ?, ?, ?, ?)", chaves)
            os.rename("estoque.json", "estoque.json.bak")
        except Exception as e: print(f"Erro ao migrar estoque.json: {e}")
    if os.path.exists("categorias.json"):
        try:
            with open("categorias.json", "r", encoding="utf-8") as f: categorias_json = json.load(f)
            cats = [(cat, "", "") for cat in categorias_json] if categorias_json and isinstance(categorias_json[0], str) else [(c['nome'], c.get('inst_pt',''), c.get('inst_en','')) for c in categorias_json]
            cursor.executemany("INSERT OR IGNORE INTO categorias (nome, instrucao_pt, instrucao_en) VALUES (?, ?, ?)", cats)
            os.rename("categorias.json", "categorias.json.bak")
        except Exception as e: print(f"Erro ao migrar categorias.json: {e}")
    cursor.execute("INSERT OR IGNORE INTO categorias (nome) VALUES ('Sem Categoria')")
    conn.commit(); conn.close()
    messagebox.showinfo("Atualização", "Dados migrados com sucesso!")

def logar_acao(acao):
    try:
        with open("log.txt", "a", encoding="utf-8") as log: log.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {acao}\n")
    except IOError: pass

class CustomAskStringDialog(simpledialog.Dialog):
    def __init__(self, parent, title=None, prompt=None, style_colors=None):
        self.prompt = prompt; self.style_colors = style_colors or {}; super().__init__(parent, title)
    def body(self, master):
        self.configure(bg=self.style_colors.get('bg', '#2D2D30')); master.configure(bg=self.style_colors.get('bg', '#2D2D30'))
        ttk.Label(master, text=self.prompt, background=self.style_colors.get('bg', '#2D2D30'), foreground=self.style_colors.get('fg', '#CCCCCC')).pack(pady=(10, 5), padx=10)
        self.entry = ttk.Entry(master, width=40); self.entry.pack(padx=10, pady=5)
        style = ttk.Style(self); style.configure("Dialog.TEntry", fieldbackground=self.style_colors.get('entry_bg', '#3C3C3C'), foreground=self.style_colors.get('text', 'white'), insertbackground=self.style_colors.get('text', 'white'))
        self.entry.configure(style="Dialog.TEntry"); return self.entry
    def buttonbox(self):
        box = ttk.Frame(self, style="TFrame")
        ttk.Button(box, text="OK", width=10, command=self.ok, default=tk.ACTIVE).pack(side=tk.LEFT, padx=5, pady=10)
        ttk.Button(box, text="Cancelar", width=10, command=self.cancel).pack(side=tk.LEFT, padx=5, pady=10)
        self.bind("<Return>", self.ok); self.bind("<Escape>", self.cancel); box.pack()
    def apply(self): self.result = self.entry.get().strip()

# --- Classe Principal ---
class GerenciadorChaves(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(f"Gerenciador de Chaves v{APP_VERSION} - por Vinícius Leão")
        self.state('zoomed'); self.resizable(True, True)
        init_db(); verificar_e_migrar_schema(); migrar_de_json_para_sqlite(); self.migrar_canais_para_tabela()
        self.is_manually_sorted, self.drag_data = True, {"item": None}
        self.configurar_tema_escuro(); self.carregar_dados_do_db()
        self.email_subject_pt = "Seu Pedido de Chave(s) de Ativação"
        self.email_subject_en = "Your Activation Key(s) Order"
        self.email_subject_es = "Su Pedido de Clave(s) de Activación"
        self.criar_menus(); self.criar_widgets()
        self.atualizar_tabela(); self.atualizar_status_bar(); self.atualizar_menus_undo_redo()
        if not PDF_DISPONIVEL: messagebox.showwarning("Biblioteca Faltando", "A biblioteca 'reportlab' não foi encontrada.\nA funcionalidade de gerar PDF estará desativada.\n\nInstale com: pip install reportlab")
        # --- NOVO: Verificação da biblioteca pandas ---
        if not PANDAS_DISPONIVEL: messagebox.showwarning("Biblioteca Faltando", "A biblioteca 'pandas' não foi encontrada.\nA funcionalidade de importar de XLS/XLSX estará desativada.\n\nInstale com: pip install pandas xlrd openpyxl")

    def migrar_canais_para_tabela(self):
        conn = sqlite3.connect(DB_NAME); cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='canais_venda'")
        if cursor.fetchone():
            cursor.execute("SELECT DISTINCT canal_venda FROM chaves WHERE canal_venda IS NOT NULL AND canal_venda != ''")
            canais_existentes = [row[0] for row in cursor.fetchall()]
            if canais_existentes: cursor.executemany("INSERT OR IGNORE INTO canais_venda (nome) VALUES (?)", [(c,) for c in canais_existentes])
            conn.commit()
        conn.close()

    def carregar_dados_do_db(self):
        conn = sqlite3.connect(DB_NAME); conn.row_factory = sqlite3.Row; cursor = conn.cursor()
        self.estoque = [dict(row) for row in cursor.execute("SELECT * FROM chaves").fetchall()]
        for item in self.estoque: item['tree_id'] = f"I{item['id']:08X}"
        self.categorias = [dict(row) for row in cursor.execute("SELECT * FROM categorias").fetchall()]
        conn.close(); self._atualizar_estoque_dict()

    def _get_lista_canais_venda(self):
        conn = sqlite3.connect(DB_NAME); cursor = conn.cursor()
        cursor.execute("SELECT nome FROM canais_venda ORDER BY nome"); nomes = [row[0] for row in cursor.fetchall()]
        conn.close(); return nomes

    def _garantir_canal_venda_existe(self, nome_canal):
        if not nome_canal or not nome_canal.strip(): return
        nome_canal = nome_canal.strip(); conn = sqlite3.connect(DB_NAME); cursor = conn.cursor()
        cursor.execute("INSERT OR IGNORE INTO canais_venda (nome) VALUES (?)", (nome_canal,)); conn.commit(); conn.close()

    def _atualizar_estoque_dict(self):
        self.estoque_dict = {item['chave']: item for item in self.estoque}
        self.tree_id_map = {item['tree_id']: item for item in self.estoque}
        self.categoria_dict = {cat['nome']: cat for cat in self.categorias}

    def salvar_e_atualizar_tudo(self):
        self.carregar_dados_do_db(); self.atualizar_combo_categoria()
        self.atualizar_combo_canal_venda(); self.atualizar_tabela(); self.atualizar_menus_undo_redo()

    def registrar_undo(self):
        if os.path.exists(DB_NAME): shutil.copy2(DB_NAME, UNDO_FILE)
        if os.path.exists(REDO_FILE): os.remove(REDO_FILE)
        self.atualizar_menus_undo_redo()

    def desfazer(self, event=None):
        if not os.path.exists(UNDO_FILE): messagebox.showinfo("Desfazer", "Nenhuma ação para desfazer."); return
        shutil.copy2(DB_NAME, REDO_FILE); shutil.copy2(UNDO_FILE, DB_NAME); os.remove(UNDO_FILE)
        self.salvar_e_atualizar_tudo(); logar_acao("Ação 'desfazer' executada."); messagebox.showinfo("Desfazer", "A última ação foi desfeita.")

    def refazer(self, event=None):
        if not os.path.exists(REDO_FILE): messagebox.showinfo("Refazer", "Nenhuma ação para refazer."); return
        shutil.copy2(DB_NAME, UNDO_FILE); shutil.copy2(REDO_FILE, DB_NAME); os.remove(REDO_FILE)
        self.salvar_e_atualizar_tudo(); logar_acao("Ação 'refazer' executada."); messagebox.showinfo("Refazer", "Ação refeita com sucesso.")

    def atualizar_menus_undo_redo(self):
        self.menu_editar.entryconfig("Desfazer", state="normal" if os.path.exists(UNDO_FILE) else "disabled")
        self.menu_editar.entryconfig("Refazer", state="normal" if os.path.exists(REDO_FILE) else "disabled")

    def configurar_tema_escuro(self):
        style = ttk.Style(self); style.theme_use("clam")
        self.bg_color, self.fg_color, self.entry_bg, self.select_bg, self.text_color = "#2D2D30", "#CCCCCC", "#3C3C3C", "#094771", "white"
        self.menu_style = {"bg": "#333333", "fg": "white", "tearoff": 0, "activebackground": self.select_bg, "activeforeground": "white", "selectcolor": self.fg_color, "borderwidth": 0}
        self.configure(bg=self.bg_color)
        options = ["*Toplevel*background", "*Toplevel*foreground", "*Label*background", "*Label*foreground", "*Frame*background", '*TLabelFrame*background', '*Dialog.msg.width', '*Dialog.Entry.background', '*Dialog.Entry.foreground', '*Dialog.Entry.insertBackground', '*TCombobox*Listbox*background', '*TCombobox*Listbox*foreground', '*TCombobox*Listbox*selectBackground', '*TCombobox*Listbox*selectForeground', '*TNotebook*background', '*TNotebook.Tab*background', '*TNotebook.Tab*foreground']
        values = [self.bg_color, self.fg_color, self.bg_color, self.fg_color, self.bg_color, self.bg_color, 40, self.entry_bg, self.text_color, self.text_color, self.entry_bg, self.text_color, self.select_bg, self.text_color, self.bg_color, self.bg_color, self.fg_color]
        for opt, val in zip(options, values): self.option_add(opt, val)
        style.configure("TNotebook", background=self.bg_color, borderwidth=0)
        style.configure("TNotebook.Tab", background=self.bg_color, foreground=self.fg_color, padding=[10, 5], font=('Segoe UI', 9)); style.map("TNotebook.Tab", background=[("selected", self.select_bg)], expand=[("selected", [1, 1, 1, 1])])
        style.configure("Treeview", background=self.entry_bg, foreground=self.text_color, fieldbackground=self.entry_bg, font=('Segoe UI', 10), rowheight=25)
        style.configure("Treeview.Heading", font=('Segoe UI', 11, 'bold'), background="#333333", foreground=self.fg_color); style.map('Treeview', background=[('selected', self.select_bg)])
        style.configure("TLabel", background=self.bg_color, foreground=self.fg_color); style.configure("TFrame", background=self.bg_color)
        style.configure("TRadiobutton", background=self.bg_color, foreground=self.fg_color, font=('Segoe UI', 10)); style.map("TRadiobutton", background=[('active', self.bg_color)], indicatorcolor=[('selected', self.text_color), ('!selected', '#555555')])
        style.configure("TCheckbutton", background=self.bg_color, foreground=self.fg_color); style.map("TCheckbutton", background=[('active', self.bg_color)], indicatorcolor=[('selected', self.text_color), ('!selected', '#555555')])
        style.configure("TEntry", fieldbackground=self.entry_bg, foreground=self.text_color, insertbackground=self.text_color)
        style.configure("TButton", background="#555555", foreground=self.text_color, font=('Segoe UI', 9)); style.map("TButton", background=[('active', '#666666')])
        style.configure("TCombobox", fieldbackground=self.entry_bg, background=self.entry_bg, foreground=self.text_color, insertbackground=self.text_color, arrowcolor=self.fg_color)
        style.map('TCombobox', fieldbackground=[('readonly', self.entry_bg)], selectbackground=[('readonly', self.entry_bg)], selectforeground=[('readonly', self.text_color)])
        style.configure("TLabelFrame", background=self.bg_color, bordercolor=self.fg_color, relief="groove"); style.configure("TLabelFrame.Label", background=self.bg_color, foreground=self.fg_color, font=('Segoe UI', 9, 'bold'))

    def criar_menus(self):
        menubar = tk.Menu(self, **self.menu_style); self.config(menu=menubar)
        menu_arquivo = tk.Menu(menubar, **self.menu_style)
        # --- NOVO: Comando de importação ---
        menu_arquivo.add_command(label="Importar Chaves de XLS...", command=self.janela_importar_xls, state="normal" if PANDAS_DISPONIVEL else "disabled")
        menu_arquivo.add_command(label="Exportar Estoque", command=self.exportar_estoque); menu_arquivo.add_separator(); menu_arquivo.add_command(label="Sair", command=self.quit)
        self.menu_editar = tk.Menu(menubar, **self.menu_style); self.menu_editar.add_command(label="Desfazer", command=self.desfazer, accelerator="Ctrl+Z"); self.menu_editar.add_command(label="Refazer", command=self.refazer, accelerator="Ctrl+Y"); self.menu_editar.add_separator(); self.menu_editar.add_command(label="Copiar Chave(s)", command=self.copiar_chave_selecionada, accelerator="Ctrl+C"); self.menu_editar.add_command(label="Editar Chave(s)", command=self.acao_editar_selecao, accelerator="F2"); self.menu_editar.add_command(label="Excluir Chave(s)", command=self.excluir_chave_selecionada, accelerator="Delete")
        menu_exibir = tk.Menu(menubar, **self.menu_style); menu_exibir.add_command(label="Atualizar Tabela", command=lambda: self.salvar_e_atualizar_tudo(), accelerator="F5")
        menu_ferramentas = tk.Menu(menubar, **self.menu_style)
        menu_ferramentas.add_command(label="Entregar Chave Única...", command=self.janela_entregar_chave_fluxo_antigo)
        menu_ferramentas.add_command(label="Entregar Várias Chaves...", command=self.janela_entregar_varias_chaves)
        menu_ferramentas.add_separator(); menu_ferramentas.add_command(label="Gerenciar Categorias...", command=self.janela_gerenciar_categorias)
        menu_ferramentas.add_command(label="Gerenciar Canais de Venda...", command=self.janela_gerenciar_canais_venda)
        menu_ferramentas.add_command(label="Dashboard de Vendas...", command=self.janela_dashboard_vendas)
        menu_ferramentas.add_separator(); menu_ferramentas.add_command(label="Configurar Email...", command=self.janela_configurar_email); menu_ferramentas.add_separator()
        menu_ferramentas.add_command(label="Fazer Backup do BD", command=self.fazer_backup_db)
        menu_ajuda = tk.Menu(menubar, **self.menu_style); menu_ajuda.add_command(label=f"Notas da Versão v{APP_VERSION}", command=self.mostrar_notas_atualizacao); menu_ajuda.add_separator(); menu_ajuda.add_command(label="Sobre", command=lambda: messagebox.showinfo("Sobre", f"Gerenciador de Chaves v{APP_VERSION}\n\nDesenvolvido por Vinícius Leão."))
        menubar.add_cascade(label="Arquivo", menu=menu_arquivo); menubar.add_cascade(label="Editar", menu=self.menu_editar); menubar.add_cascade(label="Exibir", menu=menu_exibir); menubar.add_cascade(label="Ferramentas", menu=menu_ferramentas); menubar.add_cascade(label="Ajuda", menu=menu_ajuda)
        self.bind_all("<Control-z>", self.desfazer); self.bind_all("<Control-y>", self.refazer); self.bind_all("<Control-c>", self.copiar_chave_selecionada); self.bind_all("<Delete>", self.excluir_chave_selecionada); self.bind_all("<F5>", lambda e: self.salvar_e_atualizar_tudo()); self.bind_all("<F2>", self.acao_editar_selecao)

    def mostrar_notas_atualizacao(self):
        messagebox.showinfo(f"Notas da Versão v{APP_VERSION}",
        f"v{APP_VERSION} - Importação de Chaves de Planilhas (XLS/XLSX)\n\n"
        "- **NOVO: Importação Direta de Arquivos Excel!** Agora você pode importar chaves diretamente de arquivos .xls e .xlsx, como os que recebe de seus fornecedores.\n\n"
        "- **Como usar:** Vá em `Arquivo` > `Importar Chaves de XLS...`.\n\n"
        "- **Janela de Configuração:** Após selecionar o arquivo, uma janela pedirá para você especificar a coluna (ex: B), a linha de início (ex: 4) e a categoria para as novas chaves.\n\n"
        "- **Validação Automática:** O sistema evita a importação de chaves duplicadas, garantindo a integridade do seu estoque.\n\n"
        "- **Dependência:** Esta função requer a biblioteca 'pandas'. Se não estiver instalada, o programa avisará e a opção de menu ficará desabilitada. (Instale com: pip install pandas xlrd openpyxl)")
    
    def fazer_backup_db(self):
        os.makedirs(BACKUP_DIR, exist_ok=True); nome_backup = f"backup_db_{datetime.now():%Y%m%d_%H%M%S}.db"; caminho_backup = os.path.join(BACKUP_DIR, nome_backup)
        if os.path.exists(DB_NAME): shutil.copy2(DB_NAME, caminho_backup); messagebox.showinfo("Backup", f"Backup criado em:\n{caminho_backup}")
        else: messagebox.showwarning("Backup", "Banco de dados não encontrado.")

    def criar_widgets(self):
        frame_top = ttk.Frame(self); frame_top.pack(fill=tk.X, padx=10, pady=10)
        frame_acoes = ttk.Frame(frame_top); frame_acoes.pack(side=tk.LEFT, fill=tk.Y); ttk.Button(frame_acoes, text="Adicionar Chave", command=self.janela_adicionar_chave).pack(side=tk.LEFT); ttk.Button(frame_acoes, text="Editar Chave(s)", command=self.acao_editar_selecao).pack(side=tk.LEFT, padx=5); ttk.Button(frame_acoes, text="Excluir Chave(s)", command=self.excluir_chave_selecionada).pack(side=tk.LEFT)
        frame_filtros = ttk.Frame(frame_top); frame_filtros.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        self.busca_var = tk.StringVar(); entry_busca = ttk.Entry(frame_filtros, textvariable=self.busca_var); self.status_var = tk.StringVar(value="Todos"); combo_status = ttk.Combobox(frame_filtros, textvariable=self.status_var, state="readonly", values=["Todos", "Disponível", "Vendida"], width=10)
        self.canal_venda_var = tk.StringVar(value="Todos"); self.combo_canal_venda = ttk.Combobox(frame_filtros, textvariable=self.canal_venda_var, state="readonly", width=15)
        self.categoria_var = tk.StringVar(value="Todos"); self.combo_categoria = ttk.Combobox(frame_filtros, textvariable=self.categoria_var, state="readonly", width=20)
        combo_status.pack(side=tk.RIGHT, padx=(5,0)); ttk.Label(frame_filtros, text="Status:").pack(side=tk.RIGHT)
        self.combo_canal_venda.pack(side=tk.RIGHT, padx=(5,0)); ttk.Label(frame_filtros, text="Canal:").pack(side=tk.RIGHT)
        self.combo_categoria.pack(side=tk.RIGHT, padx=(5,0)); ttk.Label(frame_filtros, text="Categoria:").pack(side=tk.RIGHT)
        entry_busca.pack(side=tk.RIGHT, fill=tk.X, expand=True); ttk.Label(frame_filtros, text="Buscar:").pack(side=tk.RIGHT, padx=(10, 2))
        entry_busca.bind("<KeyRelease>", lambda e: self.atualizar_tabela()); self.combo_categoria.bind("<<ComboboxSelected>>", lambda e: self.atualizar_tabela()); self.combo_canal_venda.bind("<<ComboboxSelected>>", lambda e: self.atualizar_tabela()); self.status_var.trace_add("write", lambda *args: self.atualizar_tabela()); self.atualizar_combo_categoria(); self.atualizar_combo_canal_venda()
        ttk.Separator(self, orient='horizontal').pack(fill='x', padx=10, pady=(0, 5))
        frame_tree = ttk.Frame(self); frame_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=0)
        colunas = ("chave", "categoria", "status", "comprador", "canal_venda", "data_venda"); self.tree = ttk.Treeview(frame_tree, columns=colunas, show="headings", selectmode="extended"); yscrollbar = ttk.Scrollbar(frame_tree, orient="vertical", command=self.tree.yview); self.tree.configure(yscrollcommand=yscrollbar.set); yscrollbar.pack(side=tk.RIGHT, fill=tk.Y); self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        headings = {"chave": "Chave", "categoria": "Categoria", "status": "Status", "comprador": "Comprador", "canal_venda": "Canal de Venda", "data_venda": "Data da Venda"}; [self.tree.heading(c, text=t, command=lambda c=c: self.ordenar_por(c)) for c,t in headings.items()]
        col_widths = {"chave": 350, "categoria": 180, "status": 100, "comprador": 150, "canal_venda": 120, "data_venda": 160}; [self.tree.column(c, width=w, anchor=tk.W) for c,w in col_widths.items()]
        self.tree.bind("<Double-1>", self.on_double_click_edit); self.tree.bind("<Button-3>", self.menu_contexto_tree); self.tree.bind("<<TreeviewSelect>>", self.atualizar_status_bar); self.tree.bind("<ButtonPress-1>", self.on_drag_start); self.tree.bind("<B1-Motion>", self.on_drag_motion); self.tree.bind("<ButtonRelease-1>", self.on_drag_end)
        self.status_bar_frame = ttk.Frame(self, style="TFrame"); self.status_bar_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5); self.status_counts_var = tk.StringVar(); ttk.Label(self.status_bar_frame, textvariable=self.status_counts_var).pack(side=tk.LEFT); ttk.Label(self.status_bar_frame, text=f"v{APP_VERSION} - por Vinícius Leão", font=('Segoe UI', 8)).pack(side=tk.RIGHT)

    def on_drag_start(self, e):
        if not self.is_manually_sorted or self.tree.identify_region(e.x, e.y) == "heading": return
        if (item := self.tree.identify_row(e.y)) and len(self.tree.selection()) <= 1: self.drag_data["item"] = item

    def on_drag_motion(self, e):
        if not self.drag_data["item"]: return
        if dest_item := self.tree.identify_row(e.y): self.tree.move(self.drag_data["item"], '', self.tree.index(dest_item))

    def on_drag_end(self, event):
        if not self.drag_data["item"]: return
        self.registrar_undo(); self._update_order_in_db(); self.drag_data["item"] = None

    def _update_order_in_db(self):
        ordered_keys = [(i, self.tree_id_map[iid]['chave']) for i, iid in enumerate(self.tree.get_children()) if iid in self.tree_id_map]
        if not ordered_keys: return
        conn = sqlite3.connect(DB_NAME); cursor = conn.cursor()
        try: cursor.executemany("UPDATE chaves SET ordem_manual = ? WHERE chave = ?", ordered_keys); conn.commit(); logar_acao("Ordem das chaves atualizada.")
        except Exception as e: conn.rollback(); messagebox.showerror("Erro de DB", f"Não foi possível salvar a ordem: {e}")
        finally: conn.close(); self.salvar_e_atualizar_tudo()

    def on_double_click_edit(self, e):
        if len(self.tree.selection()) == 1: self.janela_editar_chave(e)

    def _popup_finalizar_entrega_unica(self, chave_obj):
        popup = tk.Toplevel(self); popup.title("Finalizar Entrega"); popup.geometry("450x700"); popup.resizable(False, False); popup.grab_set(); popup.configure(bg=self.bg_color)
        
        ttk.Label(popup, text=f"Chave: {chave_obj['chave']}", font=('Segoe UI', 10, 'bold')).pack(pady=(10, 5))
        ttk.Label(popup, text=f"Categoria: {chave_obj.get('categoria', 'N/A')}").pack()
        
        frame_form = ttk.Frame(popup, style="TFrame"); frame_form.pack(pady=15, padx=20, fill=tk.X); frame_form.columnconfigure(1, weight=1)
        
        ttk.Label(frame_form, text="Comprador:").grid(row=0, column=0, sticky="w", pady=5); comprador_var = tk.StringVar(); entry_comprador = ttk.Entry(frame_form, textvariable=comprador_var); entry_comprador.grid(row=0, column=1, sticky="ew"); entry_comprador.focus()
        ttk.Label(frame_form, text="Email do Comprador:").grid(row=1, column=0, sticky="w", pady=5); email_comprador_var = tk.StringVar(); entry_email = ttk.Entry(frame_form, textvariable=email_comprador_var); entry_email.grid(row=1, column=1, sticky="ew")
        ttk.Label(frame_form, text="Canal de Venda:").grid(row=2, column=0, sticky="w", pady=5); canal_venda_var = tk.StringVar(); canal_venda_var.set(chave_obj.get('canal_venda') or "")
        combo_canal = ttk.Combobox(frame_form, textvariable=canal_venda_var, values=[''] + self._get_lista_canais_venda()); combo_canal.grid(row=2, column=1, sticky="ew")
        ttk.Label(frame_form, text="Preço Venda (R$):").grid(row=3, column=0, sticky="w", pady=5); preco_brl_var = tk.StringVar(value="0.00"); entry_preco_brl = ttk.Entry(frame_form, textvariable=preco_brl_var); entry_preco_brl.grid(row=3, column=1, sticky="ew")
        ttk.Label(frame_form, text="Preço Venda (US$):").grid(row=4, column=0, sticky="w", pady=5); preco_usd_var = tk.StringVar(value="0.00"); entry_preco_usd = ttk.Entry(frame_form, textvariable=preco_usd_var); entry_preco_usd.grid(row=4, column=1, sticky="ew")
        
        f_opc = ttk.LabelFrame(frame_form, text="Opções de Entrega"); f_opc.grid(row=5, column=0, columnspan=2, pady=(15,0), sticky="ew")
        content_opc = ttk.Frame(f_opc, style="TFrame"); content_opc.pack(fill="both", expand=True, padx=10, pady=10)
        
        acao_entrega_var = tk.StringVar(value="copiar_msg_e_pdf_pt") # Alterado para o novo padrão
        enviar_email_var = tk.BooleanVar(value=False); anexar_pdf_var = tk.BooleanVar(value=False)
        
        ttk.Radiobutton(content_opc, text="Copiar Apenas a Chave", variable=acao_entrega_var, value="copiar_chave").pack(anchor="w", pady=2)
        ttk.Separator(content_opc, orient='horizontal').pack(fill='x', pady=4)
        
        # PT-BR
        ttk.Radiobutton(content_opc, text="Copiar Mensagem (PT-BR)", variable=acao_entrega_var, value="copiar_msg_pt").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Gerar PDF (PT-BR)", variable=acao_entrega_var, value="pdf_pt", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Copiar Mensagem + Gerar PDF (PT-BR)", variable=acao_entrega_var, value="copiar_msg_e_pdf_pt", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Separator(content_opc, orient='horizontal').pack(fill='x', pady=4)

        # EN-US
        ttk.Radiobutton(content_opc, text="Copiar Mensagem (EN-US)", variable=acao_entrega_var, value="copiar_msg_en").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Gerar PDF (EN-US)", variable=acao_entrega_var, value="pdf_en", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Copiar Mensagem + Gerar PDF (EN-US)", variable=acao_entrega_var, value="copiar_msg_e_pdf_en", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Separator(content_opc, orient='horizontal').pack(fill='x', pady=4)

        # ES
        ttk.Radiobutton(content_opc, text="Copiar Mensagem (ES)", variable=acao_entrega_var, value="copiar_msg_es").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Gerar PDF (ES)", variable=acao_entrega_var, value="pdf_es", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Copiar Mensagem + Gerar PDF (ES)", variable=acao_entrega_var, value="copiar_msg_e_pdf_es", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Separator(content_opc, orient='horizontal').pack(fill='x', pady=8)
        
        chk_enviar_email = ttk.Checkbutton(content_opc, text="Enviar email para o comprador", variable=enviar_email_var); chk_enviar_email.pack(anchor="w", pady=2)
        chk_anexar_pdf = ttk.Checkbutton(content_opc, text="Anexar PDF ao email", variable=anexar_pdf_var, state="disabled"); chk_anexar_pdf.pack(anchor="w", pady=2, padx=(20, 0))

        def toggle_anexo(*args):
            acao = acao_entrega_var.get()
            if ("pdf" in acao or "pdf_e_msg" in acao) and enviar_email_var.get(): chk_anexar_pdf.config(state="normal")
            else: chk_anexar_pdf.config(state="disabled"); anexar_pdf_var.set(False)

        acao_entrega_var.trace_add("write", toggle_anexo); enviar_email_var.trace_add("write", toggle_anexo)

        def entregar():
            comprador = comprador_var.get().strip(); canal_venda = canal_venda_var.get().strip() or None
            if not comprador: messagebox.showwarning("Aviso", "Informe o nome do comprador.", parent=popup); return
            self._garantir_canal_venda_existe(canal_venda)
            email_comprador = email_comprador_var.get().strip()
            if enviar_email_var.get() and not email_comprador: messagebox.showwarning("Aviso", "Informe o email.", parent=popup); return
            try: preco_brl, preco_usd = float(preco_brl_var.get().replace(",", ".")), float(preco_usd_var.get().replace(",", "."))
            except ValueError: messagebox.showerror("Erro de Formato", "Preços devem ser números.", parent=popup); return
            
            self.registrar_undo(); data_venda = f"{datetime.now():%Y-%m-%d %H:%M:%S}"
            conn = sqlite3.connect(DB_NAME); conn.execute("UPDATE chaves SET vendida=1, comprador=?, data_venda=?, preco_venda_brl=?, preco_venda_usd=?, canal_venda=? WHERE id=?", (comprador, data_venda, preco_brl, preco_usd, canal_venda, chave_obj['id'])); conn.commit(); conn.close()
            self.salvar_e_atualizar_tudo()
            
            chave_atualizada = self.estoque_dict.get(chave_obj['chave'])
            caminho_pdf_gerado, texto_email = None, None
            acao_selecionada = acao_entrega_var.get()
            
            if acao_selecionada == "copiar_chave":
                pyperclip.copy(chave_atualizada['chave']); messagebox.showinfo("Copiado", "Chave copiada com sucesso!", parent=self)
            
            elif acao_selecionada == "copiar_msg_pt":
                pyperclip.copy(self._construir_mensagem_entrega([chave_atualizada], 'pt_br')); messagebox.showinfo("Copiado", "Mensagem em PT-BR copiada!", parent=self)
            elif acao_selecionada == "copiar_msg_en":
                pyperclip.copy(self._construir_mensagem_entrega([chave_atualizada], 'en_us')); messagebox.showinfo("Copiado", "Mensagem em EN-US copiada!", parent=self)
            elif acao_selecionada == "copiar_msg_es":
                pyperclip.copy(self._construir_mensagem_entrega([chave_atualizada], 'es_es')); messagebox.showinfo("Copiado", "Mensagem em ES copiada!", parent=self)

            elif acao_selecionada.startswith("pdf_"):
                idioma = "en_us" if acao_selecionada == "pdf_en" else "es_es" if acao_selecionada == "pdf_es" else "pt_br"
                caminho_pdf_gerado = self.gerar_pdf_entrega([chave_atualizada], idioma, comprador, email_comprador)
                if caminho_pdf_gerado: messagebox.showinfo("PDF Gerado", f"PDF salvo em:\n{caminho_pdf_gerado}", parent=self)

            elif acao_selecionada.startswith("copiar_msg_e_pdf_"):
                idioma = "en_us" if acao_selecionada == "copiar_msg_e_pdf_en" else "es_es" if acao_selecionada == "copiar_msg_e_pdf_es" else "pt_br"
                pyperclip.copy(self._construir_mensagem_entrega([chave_atualizada], idioma))
                caminho_pdf_gerado = self.gerar_pdf_entrega([chave_atualizada], idioma, comprador, email_comprador)
                if caminho_pdf_gerado:
                    messagebox.showinfo("Sucesso", "Mensagem copiada e PDF gerado com sucesso!", parent=self)
                else:
                    messagebox.showwarning("Sucesso Parcial", "Mensagem copiada, mas houve uma falha ao gerar o PDF.", parent=self)

            if enviar_email_var.get():
                if "en" in acao_selecionada: idioma_email = "en_us"
                elif "es" in acao_selecionada: idioma_email = "es_es"
                else: idioma_email = "pt_br"
                
                texto_email = self._construir_mensagem_entrega([chave_atualizada], idioma_email)
                
                if idioma_email == "en_us": assunto_email = self.email_subject_en
                elif idioma_email == "es_es": assunto_email = self.email_subject_es
                else: assunto_email = self.email_subject_pt

                anexo = caminho_pdf_gerado if anexar_pdf_var.get() else None
                threading.Thread(target=self.enviar_email_com_chave, args=(email_comprador, assunto_email, texto_email, anexo), daemon=True).start()

            logar_acao(f"Chave '{chave_obj['chave']}' entregue para {comprador}")
            popup.destroy()

        frame_botoes = ttk.Frame(popup, style="TFrame"); frame_botoes.pack(pady=10); ttk.Button(frame_botoes, text="Confirmar Entrega", command=entregar).pack(side=tk.LEFT, padx=5); ttk.Button(frame_botoes, text="Cancelar", command=popup.destroy).pack(side=tk.LEFT, padx=5)

    def gerar_pdf_entrega(self, chaves_entregues, idioma, comprador, email_comprador="", preview_mode=False, caminho_salvar_override=None):
        if not PDF_DISPONIVEL: return None
        if caminho_salvar_override: caminho_salvar = caminho_salvar_override
        elif preview_mode:
            os.makedirs(PDF_DIR, exist_ok=True); caminho_salvar = os.path.join(PDF_DIR, "preview_temp.pdf")
        else:
            os.makedirs(PDF_DIR, exist_ok=True); data_hoje = datetime.now().strftime("%Y-%m-%d"); pasta_data = os.path.join(PDF_DIR, data_hoje)
            os.makedirs(pasta_data, exist_ok=True); safe_comprador_name = _sanitize_filename(comprador)
            nome_arquivo = f"Entrega_{safe_comprador_name.replace(' ','_')}_{datetime.now():%Y%m%d%H%M%S}.pdf"; caminho_salvar = os.path.join(pasta_data, nome_arquivo)
        pdf = GeradorPDF(caminho_salvar)
        header_text = "Thank you for your purchase!" if idioma == "en_us" else "¡Gracias por su compra!" if idioma == "es_es" else "Obrigado por sua compra!"
        footer_text = "If you have any questions, please contact us." if idioma == "en_us" else "Cualquier duda o problema, por favor, entre en contacto."
        chaves_por_cat = defaultdict(list)
        for chave in chaves_entregues: chaves_por_cat[chave.get("categoria", "S/C")].append(chave['chave'])
        for i, (cat_nome, chaves_lista) in enumerate(sorted(chaves_por_cat.items())):
            if i > 0: pdf.adicionar_quebra_pagina()
            cat_obj = self.categoria_dict.get(cat_nome)
            pdf.adicionar_imagem(cat_obj.get("logo_path") if cat_obj else None, largura_cm=6.5)
            pdf.adicionar_paragrafo(header_text, estilo='HeaderStyle')
            if cat_obj:
                pt_map = {"Comprador": comprador, "Email": email_comprador, "Produto": cat_nome, "Tipo de licença": cat_obj.get('info_licenca_pt', ''), "Idioma": cat_obj.get('info_idioma_pt', ''), "Entrega": cat_obj.get('info_entrega_pt', '')}
                en_map = {"Buyer": comprador, "Email": email_comprador, "Product": cat_nome, "License type": cat_obj.get('info_licenca_en', ''), "Language": cat_obj.get('info_idioma_en', ''), "Delivery": cat_obj.get('info_entrega_en', '')}
                es_map = {"Comprador": comprador, "Email": email_comprador, "Producto": cat_nome, "Tipo de licencia": cat_obj.get('info_licenca_es', ''), "Idioma": cat_obj.get('info_idioma_es', ''), "Entrega": cat_obj.get('info_entrega_es', '')}
                info_map = en_map if idioma == 'en_us' else es_map if idioma == 'es_es' else pt_map
                info_dados_escapados = {k: html.escape(v) for k, v in info_map.items()}; pdf.adicionar_tabela_info(info_dados_escapados.items(), col_widths_cm=[4.5, 11]); pdf.adicionar_espaco_cm(0.8)
            key_label = 'Your Activation Keys:' if len(chaves_lista) > 1 else 'Your Activation Key:' if idioma == 'en_us' else 'Sus Claves de Activación:' if len(chaves_lista) > 1 else 'Su Clave de Activación:' if idioma == 'es_es' else 'Suas Chaves de Ativação:' if len(chaves_lista) > 1 else 'Sua Chave de Ativação:'
            pdf.adicionar_paragrafo(key_label, estilo='KeyLabel')
            for chave_str in chaves_lista: pdf.adicionar_paragrafo(html.escape(chave_str), estilo='KeyStyle'); pdf.adicionar_espaco_cm(0.2)
            inst_key = 'layout_pdf_en' if idioma == 'en_us' else 'layout_pdf_es' if idioma == 'es_es' else 'layout_pdf_pt'
            instrucao_texto = (cat_obj.get(inst_key) if cat_obj else '').strip()
            if instrucao_texto:
                chaves_formatadas_str = "\n".join(chaves_lista); saudacao_str = "Bom dia" if 5 <= datetime.now().hour < 12 else "Boa tarde" if 12 <= datetime.now().hour < 18 else "Boa noite"
                instrucao_texto = instrucao_texto.replace("{chave_entregue}", chaves_formatadas_str); instrucao_texto = instrucao_texto.replace("{comprador}", comprador); instrucao_texto = instrucao_texto.replace("{saudacao}", saudacao_str)
            if instrucao_texto:
                inst_header = "Activation Instructions" if idioma == "en_us" else "Instrucciones de Activación" if idioma == "es_es" else "Instruções de Ativação"
                pdf.adicionar_paragrafo(inst_header, estilo='InstructionTitleStyle')
                secoes = instrucao_texto.split('[NOVA_PAGINA]')
                for idx, secao in enumerate(secoes):
                    if secao.strip(): pdf.adicionar_paragrafo(secao, estilo='InstructionBody')
                    if idx < len(secoes) - 1: pdf.adicionar_quebra_pagina()
        pdf.adicionar_espaco_cm(1.5); pdf.adicionar_paragrafo(footer_text, estilo='FooterStyle')
        if pdf.construir():
            if not preview_mode and not caminho_salvar_override: logar_acao(f"PDF gerado com sucesso em {caminho_salvar}")
            return caminho_salvar
        return None

    def _construir_mensagem_entrega(self, chaves_entregues, idioma='pt_br'):
        if idioma == 'en_us':
            header = "Thank you for your purchase! Here are your order details:"
            footer = "If you have any questions or issues with activation, please contact us."
        elif idioma == 'es_es':
            header = "¡Gracias por su compra! Siguen los detalles de su pedido:"
            footer = "Cualquier duda o problema con la activación, por favor, póngase en contacto."
        else: # pt_br
            header = "Obrigado por sua compra! Seguem os detalhes do seu pedido:"
            footer = "Qualquer dúvida ou problema com a ativação, por favor, entre em contato."
        partes_chaves, partes_instrucoes = [], []; chaves_por_cat = defaultdict(list)
        for chave in chaves_entregues: chaves_por_cat[chave.get("categoria", "S/C")].append(chave['chave'])
        for cat_nome, chaves in sorted(chaves_por_cat.items()):
            partes_chaves.append(f"**{cat_nome}:**"); partes_chaves.extend(chaves); partes_chaves.append("")
            if cat_obj := self.categoria_dict.get(cat_nome):
                inst_key, inst_header_tpl = ('instrucao_en', "**Instructions for {cat_nome} (EN-US):**") if idioma == 'en_us' else ('instrucao_es', "**Instrucciones para {cat_nome} (ES):**") if idioma == 'es_es' else ('instrucao_pt', "**Instruções para {cat_nome} (PT-BR):**")
                if inst_text := (cat_obj.get(inst_key) or "").strip():
                    partes_instrucoes.extend(["----------", inst_header_tpl.format(cat_nome=cat_nome), inst_text, ""])
        mensagem_final = [header, "", *partes_chaves]
        if partes_instrucoes: mensagem_final.extend(partes_instrucoes)
        mensagem_final.extend([footer]); return "\n".join(mensagem_final)

    def janela_entregar_chave_fluxo_rapido(self):
        if not (sel := self.tree.selection()): return
        if not (chave_obj := self.tree_id_map.get(sel[0])): messagebox.showerror("Erro", "Chave não encontrada."); return
        if chave_obj.get("vendida"): messagebox.showwarning("Aviso", "Esta chave já foi vendida."); return
        self._popup_finalizar_entrega_unica(chave_obj)

    def janela_entregar_chave_fluxo_antigo(self):
        popup = tk.Toplevel(self); popup.title("Entregar Chave Única"); popup.geometry("600x400"); popup.grab_set(); popup.configure(bg=self.bg_color)
        ttk.Label(popup, text="Selecione uma chave disponível:").pack(pady=5, padx=10, anchor="w")
        frame_tree = ttk.Frame(popup, style="TFrame"); frame_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        tree = ttk.Treeview(frame_tree, columns=("chave", "categoria"), show="headings"); tree.heading("chave", text="Chave"); tree.heading("categoria", text="Categoria"); tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(frame_tree, orient="vertical", command=tree.yview); tree.configure(yscrollcommand=scrollbar.set); scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        chaves_disponiveis = [item for item in sorted(self.estoque, key=lambda x: (x.get("vendida", 0), x.get("categoria", ""))) if not item.get("vendida")]
        for item in chaves_disponiveis: tree.insert("", "end", values=(item["chave"], item.get("categoria", "S/C")))
        def prosseguir():
            if not (sel := tree.selection()): messagebox.showwarning("Aviso", "Selecione uma chave.", parent=popup); return
            if chave_obj := self.estoque_dict.get(tree.item(sel[0], "values")[0]):
                popup.destroy(); self._popup_finalizar_entrega_unica(chave_obj)
        tree.bind("<Double-1>", lambda e: prosseguir())
        frame_botoes = ttk.Frame(popup, style="TFrame"); frame_botoes.pack(pady=10); ttk.Button(frame_botoes, text="Prosseguir", command=prosseguir).pack(side=tk.LEFT, padx=5); ttk.Button(frame_botoes, text="Cancelar", command=popup.destroy).pack(side=tk.LEFT, padx=5)

    def menu_contexto_tree(self, event):
        sel = self.tree.selection()
        if not sel and (item := self.tree.identify_row(event.y)): self.tree.selection_set(item); sel = self.tree.selection()
        if sel:
            menu = tk.Menu(self, **self.menu_style)
            if len(sel) == 1 and (chave_obj := self.tree_id_map.get(sel[0])) and not chave_obj.get("vendida"):
                menu.add_command(label="Entregar Chave...", command=self.janela_entregar_chave_fluxo_rapido); menu.add_separator()
            menu.add_command(label=f"Editar Chave{'s' if len(sel) > 1 else ''}...", command=self.acao_editar_selecao)
            menu.add_command(label=f"Copiar Chave{'s' if len(sel) > 1 else ''}", command=self.copiar_chave_selecionada); menu.add_separator()
            menu.add_command(label=f"Excluir Chave{'s' if len(sel) > 1 else ''}", command=self.excluir_chave_selecionada)
            menu.tk_popup(event.x_root, event.y_root)

    def janela_adicionar_chave(self):
        popup = tk.Toplevel(self); popup.title("Adicionar Chaves"); popup.geometry("450x450"); popup.grab_set(); popup.configure(bg=self.bg_color)
        ttk.Label(popup, text="Digite ou cole chaves (1 por linha):").pack(anchor="w", padx=10, pady=(10,0)); texto_chaves = tk.Text(popup, height=10, bg=self.entry_bg, fg=self.text_color, insertbackground=self.text_color, relief="flat", borderwidth=1); texto_chaves.pack(fill=tk.BOTH, expand=True, padx=10, pady=5); texto_chaves.focus()
        frame_cat = ttk.Frame(popup, style="TFrame"); frame_cat.pack(fill=tk.X, padx=10, pady=5); ttk.Label(frame_cat, text="Categoria:").pack(side=tk.LEFT, padx=(0,5)); cat_var = tk.StringVar(); combo_cat = ttk.Combobox(frame_cat, textvariable=cat_var, state="readonly", values=[c['nome'] for c in self.categorias]); (combo_cat.set(self.categorias[0]['nome']) if self.categorias else None); combo_cat.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(frame_cat, text="Nova", command=lambda: nova_cat_func(combo_cat), width=5).pack(side=tk.LEFT, padx=(5,0))
        frame_canal = ttk.Frame(popup, style="TFrame"); frame_canal.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(frame_canal, text="Canal de Venda (Opcional):").pack(side=tk.LEFT, padx=(0,5))
        canal_var = tk.StringVar(); combo_canal = ttk.Combobox(frame_canal, textvariable=canal_var, values=[''] + self._get_lista_canais_venda()); combo_canal.pack(side=tk.LEFT, fill=tk.X, expand=True)
        def nova_cat_func(combo):
            dialog = CustomAskStringDialog(parent=popup, title="Nova Categoria", prompt="Nome da categoria:", style_colors={'bg':self.bg_color, 'fg':self.fg_color, 'entry_bg':self.entry_bg, 'text':self.text_color})
            if nova := dialog.result:
                if any(c['nome'].lower() == nova.lower() for c in self.categorias): messagebox.showwarning("Aviso", "Categoria já existe.", parent=popup); return
                self.registrar_undo(); conn=sqlite3.connect(DB_NAME); conn.execute("INSERT INTO categorias(nome) VALUES (?)", (nova,)); conn.commit(); conn.close()
                self.salvar_e_atualizar_tudo(); combo['values'] = [c['nome'] for c in self.categorias]; combo.set(nova); logar_acao(f"Categoria adicionada: {nova}")
        def adicionar():
            chaves = [c.strip() for c in texto_chaves.get("1.0", tk.END).strip().splitlines() if c.strip()]
            if not chaves: messagebox.showwarning("Aviso", "Nenhuma chave digitada.", parent=popup); return
            self.registrar_undo(); cat_sel = cat_var.get() or "Sem Categoria"; canal_sel = canal_var.get().strip() or None; add_c, dup_c = 0, 0
            if canal_sel: self._garantir_canal_venda_existe(canal_sel)
            conn=sqlite3.connect(DB_NAME); cursor = conn.cursor(); cursor.execute("SELECT MAX(ordem_manual) FROM chaves"); max_o = cursor.fetchone()[0] or 0; to_insert = []
            for i, chave in enumerate(chaves):
                if chave not in self.estoque_dict: to_insert.append((chave, cat_sel, max_o + i + 1, canal_sel)); add_c+=1
                else: dup_c+=1
            if add_c > 0: cursor.executemany("INSERT INTO chaves(chave, categoria, ordem_manual, canal_venda) VALUES(?, ?, ?, ?)", to_insert); conn.commit(); self.salvar_e_atualizar_tudo(); logar_acao(f"{add_c} chaves adicionadas")
            conn.close(); msg = f"{add_c} chave(s) adicionada(s)."; msg+= f"\n{dup_c} duplicada(s) foi(ram) ignorada(s)." if dup_c else ""; messagebox.showinfo("Resultado", msg, parent=popup); popup.destroy()
        frame_b = ttk.Frame(popup, style="TFrame"); frame_b.pack(pady=10); ttk.Button(frame_b, text="Adicionar", command=adicionar).pack(side=tk.LEFT,padx=5); ttk.Button(frame_b, text="Cancelar", command=popup.destroy).pack(side=tk.LEFT,padx=5)

    # --- INÍCIO: NOVAS FUNÇÕES PARA IMPORTAÇÃO DE XLS ---
    def _excel_col_to_int(self, col_str):
        """Converte uma string de coluna do Excel (ex: 'A', 'B', 'AA') para um índice 0."""
        index = 0
        for char in col_str:
            index = index * 26 + (ord(char.upper()) - ord('A') + 1)
        return index - 1

    def janela_importar_xls(self):
        """Abre o diálogo para selecionar um arquivo XLS/XLSX e, em seguida, o popup de configuração."""
        if not PANDAS_DISPONIVEL:
            messagebox.showerror("Função Indisponível", "A biblioteca 'pandas' é necessária para esta função.")
            return

        caminho_arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo XLS ou XLSX",
            filetypes=[("Arquivos Excel", "*.xls *.xlsx"), ("Todos os arquivos", "*.*")]
        )
        if not caminho_arquivo:
            return

        self._popup_configurar_importacao_xls(caminho_arquivo)

    def _popup_configurar_importacao_xls(self, caminho_arquivo):
        """Cria um popup para o usuário configurar os parâmetros de importação do Excel."""
        popup = tk.Toplevel(self)
        popup.title("Configurar Importação de XLS")
        popup.geometry("400x250")
        popup.resizable(False, False)
        popup.grab_set()
        popup.configure(bg=self.bg_color)
        
        mf = ttk.Frame(popup, padding=15, style="TFrame")
        mf.pack(fill=tk.BOTH, expand=True)
        mf.columnconfigure(1, weight=1)

        # Campo para Coluna
        ttk.Label(mf, text="Coluna das Chaves (Letra):").grid(row=0, column=0, sticky="w", pady=5, padx=5)
        col_var = tk.StringVar(value="B")
        col_entry = ttk.Entry(mf, textvariable=col_var, width=10)
        col_entry.grid(row=0, column=1, sticky="w", pady=5, padx=5)
        col_entry.focus()

        # Campo para Linha de Início
        ttk.Label(mf, text="Linha de Início (Número):").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        linha_var = tk.StringVar(value="4")
        linha_entry = ttk.Entry(mf, textvariable=linha_var, width=10)
        linha_entry.grid(row=1, column=1, sticky="w", pady=5, padx=5)

        # Campo para Categoria
        ttk.Label(mf, text="Associar à Categoria:").grid(row=2, column=0, sticky="w", pady=5, padx=5)
        cat_var = tk.StringVar()
        cat_combo = ttk.Combobox(mf, textvariable=cat_var, state="readonly", values=[c['nome'] for c in self.categorias])
        if self.categorias: cat_combo.set(self.categorias[0]['nome'])
        cat_combo.grid(row=2, column=1, sticky="ew", pady=5, padx=5)

        def processar_importacao():
            col_letra = col_var.get().strip().upper()
            linha_inicio_str = linha_var.get().strip()
            categoria_sel = cat_var.get()

            if not col_letra or not linha_inicio_str or not categoria_sel:
                messagebox.showerror("Erro de Validação", "Todos os campos são obrigatórios.", parent=popup)
                return
            
            try:
                linha_inicio = int(linha_inicio_str)
                if linha_inicio < 1: raise ValueError()
            except ValueError:
                messagebox.showerror("Erro de Validação", "A linha de início deve ser um número positivo.", parent=popup)
                return

            try:
                col_index = self._excel_col_to_int(col_letra)
                # Lê o arquivo sem tratar a primeira linha como cabeçalho
                df = pd.read_excel(caminho_arquivo, header=None, sheet_name=0)
                
                # Seleciona a coluna pelo índice e a partir da linha de início, removendo valores nulos
                chaves_a_importar = df.iloc[linha_inicio - 1:, col_index].dropna().astype(str).tolist()

                if not chaves_a_importar:
                    messagebox.showwarning("Nenhum Dado", "Nenhuma chave foi encontrada na coluna e linha especificadas.", parent=popup)
                    return
                
                self.registrar_undo()
                add_c, dup_c = 0, 0
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                cursor.execute("SELECT MAX(ordem_manual) FROM chaves")
                max_ordem = (cursor.fetchone()[0] or 0)
                
                to_insert = []
                for i, chave in enumerate(chaves_a_importar):
                    chave_limpa = chave.strip()
                    if chave_limpa and chave_limpa not in self.estoque_dict:
                        to_insert.append((chave_limpa, categoria_sel, max_ordem + i + 1, None))
                        add_c += 1
                    elif chave_limpa:
                        dup_c += 1
                
                if add_c > 0:
                    cursor.executemany("INSERT INTO chaves(chave, categoria, ordem_manual, canal_venda) VALUES(?, ?, ?, ?)", to_insert)
                    conn.commit()
                    self.salvar_e_atualizar_tudo()
                    logar_acao(f"{add_c} chaves importadas do arquivo {os.path.basename(caminho_arquivo)}")
                
                conn.close()
                msg_final = f"{add_c} chave(s) nova(s) importada(s) com sucesso!"
                if dup_c > 0:
                    msg_final += f"\n{dup_c} chave(s) duplicada(s) foi(ram) ignorada(s)."
                
                messagebox.showinfo("Importação Concluída", msg_final, parent=self)
                popup.destroy()

            except FileNotFoundError:
                messagebox.showerror("Erro de Arquivo", f"O arquivo não foi encontrado:\n{caminho_arquivo}", parent=popup)
            except Exception as e:
                messagebox.showerror("Erro na Leitura", f"Ocorreu um erro ao processar o arquivo Excel.\n\nVerifique se a coluna '{col_letra}' existe e se o arquivo não está corrompido.\n\nDetalhes do erro: {e}", parent=popup)

        botoes_f = ttk.Frame(mf, style="TFrame")
        botoes_f.grid(row=3, column=0, columnspan=2, pady=20)
        ttk.Button(botoes_f, text="Importar", command=processar_importacao, style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(botoes_f, text="Cancelar", command=popup.destroy).pack(side=tk.LEFT, padx=5)

    # --- FIM: NOVAS FUNÇÕES PARA IMPORTAÇÃO DE XLS ---

    def janela_entregar_varias_chaves(self):
        popup = tk.Toplevel(self); popup.title("Entregar Várias Chaves"); popup.geometry("800x850"); popup.grab_set(); popup.configure(bg=self.bg_color)
        frame_info = ttk.Frame(popup, style="TFrame"); frame_info.pack(fill=tk.X, padx=10, pady=5); ttk.Label(frame_info, text="Selecione as chaves:").pack(side=tk.LEFT); self.contador_sel_var = tk.StringVar(value="0 selecionadas"); ttk.Label(frame_info, textvariable=self.contador_sel_var, font=('Segoe UI', 9, 'italic')).pack(side=tk.RIGHT)
        frame_tree = ttk.Frame(popup, style="TFrame"); frame_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5); tree = ttk.Treeview(frame_tree, columns=("chave", "categoria"), show="headings", selectmode="extended")
        tree.heading("chave", text="Chave", anchor=tk.CENTER); tree.heading("categoria", text="Categoria", anchor=tk.CENTER)
        tree.column("chave", width=450, anchor=tk.CENTER); tree.column("categoria", width=250, anchor=tk.CENTER)
        scrollbar = ttk.Scrollbar(frame_tree, orient="vertical", command=tree.yview); tree.configure(yscrollcommand=scrollbar.set); tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True); scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        def upd_count(e=None): self.contador_sel_var.set(f"{len(tree.selection())} selecionadas")
        tree.bind("<<TreeviewSelect>>", upd_count)
        chaves_disponiveis = [item for item in sorted(self.estoque, key=lambda x: (x.get("vendida",0), x.get("categoria",""))) if not item.get("vendida")]
        for item in chaves_disponiveis: tree.insert("", "end", iid=item['tree_id'], values=(item["chave"], item.get("categoria", "S/C")))
        frame_form = ttk.Frame(popup, style="TFrame"); frame_form.pack(fill=tk.X, padx=10, pady=10); frame_form.columnconfigure(1, weight=1); frame_form.columnconfigure(3, weight=1)
        ttk.Label(frame_form, text="Comprador:").grid(row=0, column=0, sticky="w", pady=2, padx=(0,5)); comprador_var = tk.StringVar(); entry_comprador = ttk.Entry(frame_form, textvariable=comprador_var); entry_comprador.grid(row=0, column=1, sticky="ew"); entry_comprador.focus()
        ttk.Label(frame_form, text="Email do Comprador:").grid(row=0, column=2, sticky="w", pady=2, padx=(10,5)); email_comprador_var = tk.StringVar(); ttk.Entry(frame_form, textvariable=email_comprador_var).grid(row=0, column=3, sticky="ew")
        ttk.Label(frame_form, text="Canal de Venda:").grid(row=1, column=0, sticky="w", pady=5, padx=(0,5)); canal_venda_var = tk.StringVar()
        ttk.Combobox(frame_form, textvariable=canal_venda_var, values=[''] + self._get_lista_canais_venda()).grid(row=1, column=1, columnspan=3, sticky="ew")
        ttk.Label(frame_form, text="Preço Unit.(R$):").grid(row=2, column=0, sticky="w", pady=5, padx=(0,5)); preco_brl_var = tk.StringVar(value="0.00"); ttk.Entry(frame_form, textvariable=preco_brl_var).grid(row=2, column=1, sticky="ew")
        ttk.Label(frame_form, text="Preço Unit.(US$):").grid(row=2, column=2, sticky="w", pady=5, padx=(10,5)); preco_usd_var = tk.StringVar(value="0.00"); ttk.Entry(frame_form, textvariable=preco_usd_var).grid(row=2, column=3, sticky="ew")
        
        f_opc = ttk.LabelFrame(frame_form, text="Opções de Entrega"); f_opc.grid(row=3, column=0, columnspan=4, pady=10, sticky='ew')
        content_opc = ttk.Frame(f_opc, style="TFrame"); content_opc.pack(fill="both", expand=True, padx=10, pady=10)
        acao_entrega_var = tk.StringVar(value="copiar_msg_e_pdf_pt") # Alterado para o novo padrão
        enviar_email_var = tk.BooleanVar(value=False); anexar_pdf_var = tk.BooleanVar(value=False)
        
        ttk.Radiobutton(content_opc, text="Copiar Apenas as Chaves", variable=acao_entrega_var, value="copiar_chave").pack(anchor="w", pady=2)
        ttk.Separator(content_opc, orient='horizontal').pack(fill='x', pady=4)
        # PT-BR
        ttk.Radiobutton(content_opc, text="Copiar Mensagem (PT-BR)", variable=acao_entrega_var, value="copiar_msg_pt").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Gerar PDF (PT-BR)", variable=acao_entrega_var, value="pdf_pt", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Copiar Mensagem + Gerar PDF (PT-BR)", variable=acao_entrega_var, value="copiar_msg_e_pdf_pt", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Separator(content_opc, orient='horizontal').pack(fill='x', pady=4)
        # EN-US
        ttk.Radiobutton(content_opc, text="Copiar Mensagem (EN-US)", variable=acao_entrega_var, value="copiar_msg_en").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Gerar PDF (EN-US)", variable=acao_entrega_var, value="pdf_en", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Copiar Mensagem + Gerar PDF (EN-US)", variable=acao_entrega_var, value="copiar_msg_e_pdf_en", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Separator(content_opc, orient='horizontal').pack(fill='x', pady=4)
        # ES
        ttk.Radiobutton(content_opc, text="Copiar Mensagem (ES)", variable=acao_entrega_var, value="copiar_msg_es").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Gerar PDF (ES)", variable=acao_entrega_var, value="pdf_es", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Radiobutton(content_opc, text="Copiar Mensagem + Gerar PDF (ES)", variable=acao_entrega_var, value="copiar_msg_e_pdf_es", state="normal" if PDF_DISPONIVEL else "disabled").pack(anchor="w", pady=2)
        ttk.Separator(content_opc, orient='horizontal').pack(fill='x', pady=8)

        chk_enviar_email = ttk.Checkbutton(content_opc, text="Enviar email para o comprador", variable=enviar_email_var); chk_enviar_email.pack(anchor="w", pady=2)
        chk_anexar_pdf = ttk.Checkbutton(content_opc, text="Anexar PDF ao email", variable=anexar_pdf_var, state="disabled"); chk_anexar_pdf.pack(anchor="w", pady=2, padx=(20, 0))

        def toggle_anexo(*args):
            acao = acao_entrega_var.get()
            if ("pdf" in acao or "pdf_e_msg" in acao) and enviar_email_var.get(): chk_anexar_pdf.config(state="normal")
            else: chk_anexar_pdf.config(state="disabled"); anexar_pdf_var.set(False)
        acao_entrega_var.trace_add("write", toggle_anexo); enviar_email_var.trace_add("write", toggle_anexo)

        def entregar():
            sel_ids = tree.selection(); comprador = comprador_var.get().strip()
            if not sel_ids or not comprador: messagebox.showwarning("Aviso", "Selecione chaves e informe o comprador.", parent=popup); return
            email_comprador = email_comprador_var.get().strip(); canal_venda = canal_venda_var.get().strip() or None
            if enviar_email_var.get() and not email_comprador: messagebox.showwarning("Aviso", "Informe o email.", parent=popup); return
            if canal_venda: self._garantir_canal_venda_existe(canal_venda)
            try: preco_brl, preco_usd = float(preco_brl_var.get().replace(",", ".")), float(preco_usd_var.get().replace(",", "."))
            except ValueError: messagebox.showerror("Erro de Formato", "Preços devem ser números.", parent=popup); return
            if not messagebox.askyesno("Confirmar Entrega", f"Entregar {len(sel_ids)} chaves para '{comprador}'?", parent=popup): return
            
            self.registrar_undo(); data_venda = f"{datetime.now():%Y-%m-%d %H:%M:%S}"; entregues_obj, para_update = [], []
            for sel_id in sel_ids:
                if item := self.tree_id_map.get(sel_id): 
                    para_update.append((comprador, data_venda, preco_brl, preco_usd, canal_venda, item['id']))
                    item.update({'vendida':1, 'comprador':comprador, 'data_venda':data_venda, 'preco_venda_brl': preco_brl, 'preco_venda_usd': preco_usd, 'canal_venda': canal_venda})
                    entregues_obj.append(item)
            
            conn=sqlite3.connect(DB_NAME); conn.executemany("UPDATE chaves SET vendida=1, comprador=?, data_venda=?, preco_venda_brl=?, preco_venda_usd=?, canal_venda=? WHERE id=?", para_update); conn.commit(); conn.close()
            self.salvar_e_atualizar_tudo()
            
            caminho_pdf_gerado = None; acao_selecionada = acao_entrega_var.get()
            
            if acao_selecionada == "copiar_chave":
                pyperclip.copy("\n".join([c['chave'] for c in entregues_obj])); messagebox.showinfo("Copiado", f"{len(entregues_obj)} Chaves copiadas!", parent=self)

            elif acao_selecionada == "copiar_msg_pt":
                pyperclip.copy(self._construir_mensagem_entrega(entregues_obj, 'pt_br')); messagebox.showinfo("Copiado", "Mensagem em PT-BR copiada!", parent=self)
            elif acao_selecionada == "copiar_msg_en":
                pyperclip.copy(self._construir_mensagem_entrega(entregues_obj, 'en_us')); messagebox.showinfo("Copiado", "Mensagem em EN-US copiada!", parent=self)
            elif acao_selecionada == "copiar_msg_es":
                pyperclip.copy(self._construir_mensagem_entrega(entregues_obj, 'es_es')); messagebox.showinfo("Copiado", "Mensagem em ES copiada!", parent=self)
            
            elif acao_selecionada.startswith("pdf_"):
                idioma = "en_us" if acao_selecionada == "pdf_en" else "es_es" if acao_selecionada == "pdf_es" else "pt_br"
                caminho_pdf_gerado = self.gerar_pdf_entrega(entregues_obj, idioma, comprador, email_comprador)
                if caminho_pdf_gerado: messagebox.showinfo("PDF Gerado", f"PDF salvo em:\n{caminho_pdf_gerado}", parent=self)
            
            elif acao_selecionada.startswith("copiar_msg_e_pdf_"):
                idioma = "en_us" if acao_selecionada == "copiar_msg_e_pdf_en" else "es_es" if acao_selecionada == "copiar_msg_e_pdf_es" else "pt_br"
                pyperclip.copy(self._construir_mensagem_entrega(entregues_obj, idioma))
                caminho_pdf_gerado = self.gerar_pdf_entrega(entregues_obj, idioma, comprador, email_comprador)
                if caminho_pdf_gerado:
                    messagebox.showinfo("Sucesso", "Mensagem copiada e PDF gerado com sucesso!", parent=self)
                else:
                    messagebox.showwarning("Sucesso Parcial", "Mensagem copiada, mas houve uma falha ao gerar o PDF.", parent=self)

            if enviar_email_var.get():
                if "en" in acao_selecionada: idioma_email = "en_us"
                elif "es" in acao_selecionada: idioma_email = "es_es"
                else: idioma_email = "pt_br"

                texto_email = self._construir_mensagem_entrega(entregues_obj, idioma_email)

                if idioma_email == "en_us": assunto_email = self.email_subject_en
                elif idioma_email == "es_es": assunto_email = self.email_subject_es
                else: assunto_email = self.email_subject_pt

                anexo = caminho_pdf_gerado if anexar_pdf_var.get() else None
                threading.Thread(target=self.enviar_email_com_chave, args=(email_comprador, assunto_email, texto_email, anexo), daemon=True).start()
                
            logar_acao(f"{len(entregues_obj)} chaves entregues para {comprador}"); popup.destroy()

        f_botoes=ttk.Frame(popup, style="TFrame"); f_botoes.pack(pady=10); ttk.Button(f_botoes, text="Confirmar Entrega", command=entregar).pack(side=tk.LEFT,padx=5); ttk.Button(f_botoes, text="Cancelar", command=popup.destroy).pack(side=tk.LEFT,padx=5)


    def janela_gerenciar_categorias(self):
        popup = tk.Toplevel(self); popup.title("Gerenciar Categorias"); popup.geometry("950x850"); popup.grab_set(); popup.configure(bg=self.bg_color)
        main_frame = ttk.Frame(popup, style="TFrame", padding=10); main_frame.pack(fill=tk.BOTH, expand=True); main_frame.grid_columnconfigure(1, weight=1); main_frame.grid_rowconfigure(0, weight=1)
        list_frame = ttk.Frame(main_frame, style="TFrame"); list_frame.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(0, 10))
        listbox = tk.Listbox(list_frame, bg=self.entry_bg, fg=self.text_color, selectbackground=self.select_bg, relief="flat", borderwidth=0, highlightthickness=0, exportselection=False, width=25); listbox.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        list_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview); listbox.config(yscrollcommand=list_scroll.set); list_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        notebook = ttk.Notebook(main_frame, style="TNotebook"); notebook.grid(row=0, column=1, sticky="nsew")
        f_inst = ttk.Frame(notebook, style="TFrame", padding=10); notebook.add(f_inst, text=" Instruções (Email) ")
        f_layout = ttk.Frame(notebook, style="TFrame", padding=10); notebook.add(f_layout, text=" Texto Detalhado (PDF) ")
        f_pdf = ttk.Frame(notebook, style="TFrame", padding=10); notebook.add(f_pdf, text=" Layout do PDF ")
        f_custos = ttk.Frame(notebook, style="TFrame", padding=10); notebook.add(f_custos, text=" Custos ")
        f_inst.columnconfigure(0, weight=1); f_inst.rowconfigure(1, weight=1); f_inst.rowconfigure(3, weight=1); f_inst.rowconfigure(5, weight=1)
        ttk.Label(f_inst, text="Instruções para o corpo do Email (PT-BR):").grid(row=0, column=0, sticky="w"); text_pt = tk.Text(f_inst, height=5, bg=self.entry_bg, fg=self.text_color, insertbackground=self.text_color, relief="flat", borderwidth=1, wrap="word"); text_pt.grid(row=1, column=0, sticky="nsew", pady=(2,10))
        ttk.Label(f_inst, text="Instruções para o corpo do Email (EN-US):").grid(row=2, column=0, sticky="w"); text_en = tk.Text(f_inst, height=5, bg=self.entry_bg, fg=self.text_color, insertbackground=self.text_color, relief="flat", borderwidth=1, wrap="word"); text_en.grid(row=3, column=0, sticky="nsew", pady=(2,10))
        ttk.Label(f_inst, text="Instrucciones para el cuerpo del Email (ES):").grid(row=4, column=0, sticky="w"); text_es = tk.Text(f_inst, height=5, bg=self.entry_bg, fg=self.text_color, insertbackground=self.text_color, relief="flat", borderwidth=1, wrap="word"); text_es.grid(row=5, column=0, sticky="nsew", pady=(2,0))
        f_layout.columnconfigure(0, weight=1); f_layout.rowconfigure(2, weight=1); f_layout.rowconfigure(5, weight=1); f_layout.rowconfigure(8, weight=1)
        def aplicar_tag(widget_texto, tag):
            try:
                inicio, fim = widget_texto.index("sel.first"), widget_texto.index("sel.last")
                texto_selecionado = widget_texto.get(inicio, fim); widget_texto.delete(inicio, fim); widget_texto.insert(inicio, f"{tag}{texto_selecionado}{tag}")
            except tk.TclError: messagebox.showinfo("Aviso", "Selecione um texto para aplicar a formatação.", parent=popup)
        ttk.Label(f_layout, text="Texto detalhado para o PDF (PT-BR):").grid(row=0, column=0, sticky="w")
        toolbar_pt = ttk.Frame(f_layout, style="TFrame"); toolbar_pt.grid(row=1, column=0, sticky="ew", pady=(2,0))
        layout_pt = tk.Text(f_layout, height=5, bg=self.entry_bg, fg=self.text_color, insertbackground=self.text_color, relief="flat", borderwidth=1, wrap="word", undo=True); layout_pt.grid(row=2, column=0, sticky="nsew", pady=(2,10))
        ttk.Label(f_layout, text="Texto detalhado para o PDF (EN-US):").grid(row=3, column=0, sticky="w")
        toolbar_en = ttk.Frame(f_layout, style="TFrame"); toolbar_en.grid(row=4, column=0, sticky="ew", pady=(2,0))
        layout_en = tk.Text(f_layout, height=5, bg=self.entry_bg, fg=self.text_color, insertbackground=self.text_color, relief="flat", borderwidth=1, wrap="word", undo=True); layout_en.grid(row=5, column=0, sticky="nsew", pady=(2,10))
        ttk.Label(f_layout, text="Texto detallado para el PDF (ES):").grid(row=6, column=0, sticky="w")
        toolbar_es = ttk.Frame(f_layout, style="TFrame"); toolbar_es.grid(row=7, column=0, sticky="ew", pady=(2,0))
        layout_es = tk.Text(f_layout, height=5, bg=self.entry_bg, fg=self.text_color, insertbackground=self.text_color, relief="flat", borderwidth=1, wrap="word", undo=True); layout_es.grid(row=8, column=0, sticky="nsew", pady=(2,0))
        for toolbar, widget_texto in [(toolbar_pt, layout_pt), (toolbar_en, layout_en), (toolbar_es, layout_es)]:
            ttk.Button(toolbar, text=" B ", width=3, command=lambda w=widget_texto: aplicar_tag(w, "**")).pack(side=tk.LEFT)
            ttk.Button(toolbar, text=" I ", width=3, command=lambda w=widget_texto: aplicar_tag(w, "*")).pack(side=tk.LEFT)
            ttk.Button(toolbar, text=" U ", width=3, command=lambda w=widget_texto: aplicar_tag(w, "__")).pack(side=tk.LEFT)
            ttk.Label(toolbar, text="|").pack(side=tk.LEFT, padx=5)
            ttk.Button(toolbar, text="[NOVA_PAGINA]", command=lambda w=widget_texto: w.insert(tk.INSERT, "[NOVA_PAGINA]")).pack(side=tk.LEFT)
        f_pdf.columnconfigure(1, weight=1); logo_path_var = tk.StringVar()
        def browse_logo():
            path = filedialog.askopenfilename(title="Selecionar Logo", filetypes=[("Imagens", "*.png *.jpg *.jpeg *.gif"), ("Todos", "*.*")], parent=popup)
            if path: logo_path_var.set(path)
        ttk.Label(f_pdf, text="Logo do Produto:").grid(row=0, column=0, sticky="w", padx=5, pady=5); ttk.Entry(f_pdf, textvariable=logo_path_var).grid(row=0, column=1, sticky="ew", padx=5, pady=5); ttk.Button(f_pdf, text="Procurar...", command=browse_logo).grid(row=0, column=2, padx=5, pady=5)
        pdf_details_pt = ttk.LabelFrame(f_pdf, text=" Detalhes do Produto (PT-BR) "); pdf_details_pt.grid(row=1, column=0, columnspan=3, sticky="ew", pady=5)
        pdf_details_pt.columnconfigure(1, weight=1); lic_pt_var=tk.StringVar(); idiom_pt_var=tk.StringVar(); entr_pt_var=tk.StringVar()
        ttk.Label(pdf_details_pt, text="Tipo de Licença:").grid(row=0,column=0,sticky="w",padx=5,pady=3); ttk.Entry(pdf_details_pt,textvariable=lic_pt_var).grid(row=0,column=1,sticky="ew",padx=5,pady=3)
        ttk.Label(pdf_details_pt, text="Idioma:").grid(row=1,column=0,sticky="w",padx=5,pady=3); ttk.Entry(pdf_details_pt,textvariable=idiom_pt_var).grid(row=1,column=1,sticky="ew",padx=5,pady=3)
        ttk.Label(pdf_details_pt, text="Entrega:").grid(row=2,column=0,sticky="w",padx=5,pady=3); ttk.Entry(pdf_details_pt,textvariable=entr_pt_var).grid(row=2,column=1,sticky="ew",padx=5,pady=3)
        pdf_details_en = ttk.LabelFrame(f_pdf, text=" Product Details (EN-US) "); pdf_details_en.grid(row=2, column=0, columnspan=3, sticky="ew", pady=5)
        pdf_details_en.columnconfigure(1, weight=1); lic_en_var=tk.StringVar(); idiom_en_var=tk.StringVar(); entr_en_var=tk.StringVar()
        ttk.Label(pdf_details_en, text="License Type:").grid(row=0,column=0,sticky="w",padx=5,pady=3); ttk.Entry(pdf_details_en,textvariable=lic_en_var).grid(row=0,column=1,sticky="ew",padx=5,pady=3)
        ttk.Label(pdf_details_en, text="Language:").grid(row=1,column=0,sticky="w",padx=5,pady=3); ttk.Entry(pdf_details_en,textvariable=idiom_en_var).grid(row=1,column=1,sticky="ew",padx=5,pady=3)
        ttk.Label(pdf_details_en, text="Delivery:").grid(row=2,column=0,sticky="w",padx=5,pady=3); ttk.Entry(pdf_details_en,textvariable=entr_en_var).grid(row=2,column=1,sticky="ew",padx=5,pady=3)
        pdf_details_es = ttk.LabelFrame(f_pdf, text=" Detalles del Producto (ES) "); pdf_details_es.grid(row=3, column=0, columnspan=3, sticky="ew", pady=5)
        pdf_details_es.columnconfigure(1, weight=1); lic_es_var=tk.StringVar(); idiom_es_var=tk.StringVar(); entr_es_var=tk.StringVar()
        ttk.Label(pdf_details_es, text="Tipo de Licencia:").grid(row=0,column=0,sticky="w",padx=5,pady=3); ttk.Entry(pdf_details_es,textvariable=lic_es_var).grid(row=0,column=1,sticky="ew",padx=5,pady=3)
        ttk.Label(pdf_details_es, text="Idioma:").grid(row=1,column=0,sticky="w",padx=5,pady=3); ttk.Entry(pdf_details_es,textvariable=idiom_es_var).grid(row=1,column=1,sticky="ew",padx=5,pady=3)
        ttk.Label(pdf_details_es, text="Entrega:").grid(row=2,column=0,sticky="w",padx=5,pady=3); ttk.Entry(pdf_details_es,textvariable=entr_es_var).grid(row=2,column=1,sticky="ew",padx=5,pady=3)
        f_custos.columnconfigure(1, weight=1); f_custos.columnconfigure(3, weight=1); custo_brl_var = tk.StringVar(); custo_usd_var = tk.StringVar()
        ttk.Label(f_custos, text="Custo Padrão (R$):").grid(row=0, column=0, padx=5, pady=5, sticky="w"); ttk.Entry(f_custos, textvariable=custo_brl_var).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Label(f_custos, text="Custo Padrão (US$):").grid(row=0, column=2, padx=5, pady=5, sticky="w"); ttk.Entry(f_custos, textvariable=custo_usd_var).grid(row=0, column=3, sticky="ew", padx=5, pady=5)
        def load_cat_details(e=None):
            if not (sel_idx := listbox.curselection()): return
            cat_nome = listbox.get(sel_idx[0]); cat_obj = self.categoria_dict.get(cat_nome)
            text_pt.delete("1.0", tk.END); text_en.delete("1.0", tk.END); text_es.delete("1.0", tk.END)
            layout_pt.delete("1.0", tk.END); layout_en.delete("1.0", tk.END); layout_es.delete("1.0", tk.END)
            for var in [custo_brl_var, custo_usd_var, logo_path_var, lic_pt_var, idiom_pt_var, entr_pt_var, lic_en_var, idiom_en_var, entr_en_var, lic_es_var, idiom_es_var, entr_es_var]: var.set("")
            custo_brl_var.set("0.00"); custo_usd_var.set("0.00")
            if cat_obj:
                text_pt.insert("1.0", cat_obj.get("instrucao_pt","")); text_en.insert("1.0", cat_obj.get("instrucao_en","")); text_es.insert("1.0", cat_obj.get("instrucao_es",""))
                layout_pt.insert("1.0", cat_obj.get("layout_pdf_pt", "")); layout_en.insert("1.0", cat_obj.get("layout_pdf_en", "")); layout_es.insert("1.0", cat_obj.get("layout_pdf_es", ""))
                custo_brl_var.set(f"{cat_obj.get('custo_padrao_brl') or 0.0:.2f}"); custo_usd_var.set(f"{cat_obj.get('custo_padrao_usd') or 0.0:.2f}")
                logo_path_var.set(cat_obj.get("logo_path", "")); lic_pt_var.set(cat_obj.get("info_licenca_pt", "")); idiom_pt_var.set(cat_obj.get("info_idioma_pt", "")); entr_pt_var.set(cat_obj.get("info_entrega_pt", ""))
                lic_en_var.set(cat_obj.get("info_licenca_en", "")); idiom_en_var.set(cat_obj.get("info_idioma_en", "")); entr_en_var.set(cat_obj.get("info_entrega_en", ""))
                lic_es_var.set(cat_obj.get("info_licenca_es", "")); idiom_es_var.set(cat_obj.get("info_idioma_es", "")); entr_es_var.set(cat_obj.get("info_entrega_es", ""))
        listbox.bind("<<ListboxSelect>>", load_cat_details)
        def fill_lb(): listbox.delete(0,tk.END); [listbox.insert(tk.END, c['nome']) for c in sorted(self.categorias, key=lambda x: x['nome'])]
        def save_cat():
            if not (sel_idx := listbox.curselection()): messagebox.showwarning("Aviso", "Selecione uma categoria.", parent=popup); return
            cat_nome = listbox.get(sel_idx[0])
            try: custo_brl, custo_usd = float(custo_brl_var.get().replace(",",".")), float(custo_usd_var.get().replace(",","."))
            except ValueError: messagebox.showerror("Erro de Formato", "Custos devem ser números.", parent=popup); return
            self.registrar_undo(); conn = sqlite3.connect(DB_NAME)
            dados = (text_pt.get("1.0",tk.END).strip(),text_en.get("1.0",tk.END).strip(),text_es.get("1.0",tk.END).strip(),custo_brl,custo_usd,logo_path_var.get().strip(),lic_pt_var.get().strip(),lic_en_var.get().strip(),lic_es_var.get().strip(),idiom_pt_var.get().strip(),idiom_en_var.get().strip(),idiom_es_var.get().strip(),entr_pt_var.get().strip(),entr_en_var.get().strip(),entr_es_var.get().strip(),layout_pt.get("1.0",tk.END).strip(),layout_en.get("1.0",tk.END).strip(),layout_es.get("1.0",tk.END).strip(),cat_nome)
            query = "UPDATE categorias SET instrucao_pt=?,instrucao_en=?,instrucao_es=?,custo_padrao_brl=?,custo_padrao_usd=?,logo_path=?,info_licenca_pt=?,info_licenca_en=?,info_licenca_es=?,info_idioma_pt=?,info_idioma_en=?,info_idioma_es=?,info_entrega_pt=?,info_entrega_en=?,info_entrega_es=?,layout_pdf_pt=?,layout_pdf_en=?,layout_pdf_es=? WHERE nome=?"
            conn.execute(query, dados); conn.commit(); conn.close(); self.salvar_e_atualizar_tudo(); messagebox.showinfo("Sucesso", f"Dados de '{cat_nome}' salvos.", parent=popup)
        def previsualizar_pdf_selecionado():
            if not (sel_idx := listbox.curselection()): messagebox.showwarning("Aviso", "Selecione uma categoria.", parent=popup); return
            cat_nome = listbox.get(sel_idx[0]); idioma_foco = 'pt_br'; focused_widget = popup.focus_get()
            if focused_widget == layout_en: idioma_foco = 'en_us'
            elif focused_widget == layout_es: idioma_foco = 'es_es'
            chave_dummy = {'chave': 'XXXX-XXXX-XXXX-XXXX', 'categoria': cat_nome}
            if idioma_foco == 'pt_br': chave_dummy['layout_pdf_pt'] = layout_pt.get("1.0", tk.END).strip()
            elif idioma_foco == 'en_us': chave_dummy['layout_pdf_en'] = layout_en.get("1.0", tk.END).strip()
            else: chave_dummy['layout_pdf_es'] = layout_es.get("1.0", tk.END).strip()
            cat_obj_preview = chave_dummy.copy(); cat_obj_preview.update({"logo_path": logo_path_var.get(),"info_licenca_pt": lic_pt_var.get(),"info_idioma_pt": idiom_pt_var.get(),"info_entrega_pt": entr_pt_var.get(),"info_licenca_en": lic_en_var.get(),"info_idioma_en": idiom_en_var.get(),"info_entrega_en": entr_en_var.get(),"info_licenca_es": lic_es_var.get(),"info_idioma_es": idiom_es_var.get(),"info_entrega_es": entr_es_var.get(),})
            self.categoria_dict[cat_nome] = cat_obj_preview
            caminho_preview = self.gerar_pdf_entrega(chaves_entregues=[chave_dummy],idioma=idioma_foco,comprador="Comprador de Teste",email_comprador="teste@email.com",preview_mode=True)
            if caminho_preview: webbrowser.open_new(f'file://{os.path.realpath(caminho_preview)}')
            self.carregar_dados_do_db()
        btn_frame = ttk.Frame(main_frame, style="TFrame"); btn_frame.grid(row=1, column=1, sticky="sew", pady=(10,0))
        def add_cat(cb):
            d = CustomAskStringDialog(parent=popup, title="Nova Categoria", prompt="Nome:", style_colors={'bg': self.bg_color, 'fg': self.fg_color, 'entry_bg': self.entry_bg, 'text': self.text_color})
            if nova := d.result:
                if any(c['nome'].lower() == nova.lower() for c in self.categorias): messagebox.showwarning("Aviso", "Categoria já existe.", parent=popup); return
                self.registrar_undo(); conn = sqlite3.connect(DB_NAME); conn.execute("INSERT INTO categorias(nome,custo_padrao_brl,custo_padrao_usd) VALUES(?,0.0,0.0)", (nova,)); conn.commit(); conn.close()
                self.salvar_e_atualizar_tudo(); cb(); logar_acao(f"Categoria adicionada: {nova}")
        def del_cat(l, cb):
            if not (s := l.curselection()): messagebox.showwarning("Aviso", "Selecione uma categoria.", parent=popup); return
            nc = l.get(s[0])
            if nc == "Sem Categoria": messagebox.showerror("Erro", "'Sem Categoria' não pode ser excluída.", parent=popup); return
            if messagebox.askyesno("Excluir Categoria", f"Deseja excluir '{nc}'?", parent=popup, icon='warning'):
                self.registrar_undo(); conn = sqlite3.connect(DB_NAME); c = conn.cursor()
                c.execute("UPDATE chaves SET categoria='Sem Categoria' WHERE categoria=?", (nc,)); c.execute("DELETE FROM categorias WHERE nome=?", (nc,)); conn.commit(); conn.close()
                self.salvar_e_atualizar_tudo(); cb(); logar_acao(f"Categoria excluída: {nc}")
        ttk.Button(btn_frame, text="Nova", command=lambda: add_cat(fill_lb)).pack(side=tk.LEFT, padx=(0,5))
        ttk.Button(btn_frame, text="Excluir", command=lambda: del_cat(listbox, fill_lb)).pack(side=tk.LEFT, padx=(0,5))
        ttk.Button(btn_frame, text="Salvar Alterações", command=save_cat, style="Accent.TButton").pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="Pré-visualizar PDF", command=previsualizar_pdf_selecionado).pack(side=tk.RIGHT, padx=5)
        fill_lb()
    
    def janela_gerenciar_canais_venda(self):
        popup = tk.Toplevel(self); popup.title("Gerenciar Canais de Venda"); popup.geometry("500x400"); popup.grab_set(); popup.configure(bg=self.bg_color)
        mf = ttk.Frame(popup, style="TFrame", padding=10); mf.pack(fill=tk.BOTH, expand=True); mf.rowconfigure(1, weight=1); mf.columnconfigure(0, weight=1)
        ttk.Label(mf, text="Canais de Venda Existentes:", font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 5))
        list_frame = ttk.Frame(mf, style="TFrame"); list_frame.grid(row=1, column=0, columnspan=2, sticky='nsew')
        lb = tk.Listbox(list_frame, bg=self.entry_bg, fg=self.text_color, selectbackground=self.select_bg, relief="flat", borderwidth=0, highlightthickness=0, exportselection=False); lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=lb.yview); lb.configure(yscrollcommand=scrollbar.set); scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        def fill_lb(): lb.delete(0, tk.END); [lb.insert(tk.END, canal) for canal in self._get_lista_canais_venda()]
        def adicionar_canal():
            novo_nome = simpledialog.askstring("Adicionar Canal", "Digite o nome do novo canal:", parent=popup)
            if novo_nome and (nome_limpo := novo_nome.strip()):
                if nome_limpo in self._get_lista_canais_venda(): messagebox.showerror("Erro", f"O canal '{nome_limpo}' já existe.", parent=popup); return
                self._garantir_canal_venda_existe(nome_limpo); self.salvar_e_atualizar_tudo(); logar_acao(f"Canal '{nome_limpo}' adicionado."); fill_lb()
        def renomear_canal():
            if not (sel := lb.curselection()): messagebox.showwarning("Aviso", "Selecione um canal para renomear.", parent=popup); return
            canal_antigo = lb.get(sel[0])
            novo_nome = simpledialog.askstring("Renomear Canal", f"Digite o novo nome para '{canal_antigo}':", parent=popup)
            if novo_nome and (nome_limpo := novo_nome.strip()):
                if nome_limpo == canal_antigo: return
                if nome_limpo in self._get_lista_canais_venda(): messagebox.showerror("Erro", f"O canal '{nome_limpo}' já existe.", parent=popup); return
                self.registrar_undo(); conn = sqlite3.connect(DB_NAME)
                conn.execute("UPDATE canais_venda SET nome=? WHERE nome=?", (nome_limpo, canal_antigo)); conn.execute("UPDATE chaves SET canal_venda=? WHERE canal_venda=?", (nome_limpo, canal_antigo)); conn.commit(); conn.close()
                self.salvar_e_atualizar_tudo(); logar_acao(f"Canal '{canal_antigo}' renomeado para '{nome_limpo}'"); fill_lb()
        def excluir_canal():
            if not (sel := lb.curselection()): messagebox.showwarning("Aviso", "Selecione um canal para excluir.", parent=popup); return
            canal = lb.get(sel[0])
            if messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja remover o canal '{canal}'?\nIsso o removerá de todas as chaves associadas.", icon='warning', parent=popup):
                self.registrar_undo(); conn = sqlite3.connect(DB_NAME)
                conn.execute("DELETE FROM canais_venda WHERE nome=?", (canal,)); conn.execute("UPDATE chaves SET canal_venda=NULL WHERE canal_venda=?", (canal,)); conn.commit(); conn.close()
                self.salvar_e_atualizar_tudo(); logar_acao(f"Canal '{canal}' excluído"); fill_lb()
        btn_frame = ttk.Frame(mf, style="TFrame"); btn_frame.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        ttk.Button(btn_frame, text="Adicionar", command=adicionar_canal).pack(side=tk.LEFT, padx=5); ttk.Button(btn_frame, text="Renomear", command=renomear_canal).pack(side=tk.LEFT, padx=5); ttk.Button(btn_frame, text="Excluir", command=excluir_canal).pack(side=tk.LEFT, padx=5); ttk.Button(btn_frame, text="Fechar", command=popup.destroy).pack(side=tk.RIGHT, padx=5)
        fill_lb()

    def excluir_chave_selecionada(self, event=None):
        if not (sel := self.tree.selection()): messagebox.showwarning("Excluir", "Selecione chaves."); return
        if not messagebox.askyesno("Confirmar", f"Excluir permanentemente as {len(sel)} chaves?", icon='warning'): return
        ids=[self.tree_id_map[i]['id'] for i in sel if i in self.tree_id_map]
        if not ids: messagebox.showerror("Erro","Chaves não encontradas."); return
        self.registrar_undo(); conn=sqlite3.connect(DB_NAME); conn.execute(f"DELETE FROM chaves WHERE id IN ({','.join('?'*len(ids))})", ids); conn.commit(); conn.close()
        self.salvar_e_atualizar_tudo(); logar_acao(f"{len(ids)} chaves excluídas."); messagebox.showinfo("Excluído",f"{len(ids)} chaves excluídas.")

    def exportar_estoque(self):
        if not (caminho := filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv"), ("All", "*.*")])): return
        try:
            conn=sqlite3.connect(DB_NAME); cursor=conn.cursor()
            cursor.execute("SELECT chave, categoria, vendida, comprador, canal_venda, data_venda, preco_venda_brl, preco_venda_usd FROM chaves ORDER BY ordem_manual ASC")
            with open(caminho,"w",encoding="utf-8",newline='') as f:
                import csv
                w=csv.writer(f); w.writerow(["Chave","Categoria","Status","Comprador", "Canal de Venda", "Data","PrecoBRL","PrecoUSD"])
                for row in cursor.fetchall(): r=list(row); r[2]="Vendida" if r[2]==1 else "Disponível"; w.writerow(r)
            conn.close(); messagebox.showinfo("Exportar", "Estoque exportado com sucesso.")
        except Exception as e: messagebox.showerror("Erro", f"Erro ao exportar:\n{e}")

    def atualizar_combo_categoria(self):
        nomes = sorted([cat['nome'] for cat in self.categorias]); self.combo_categoria['values'] = ["Todos"] + nomes
        if self.categoria_var.get() not in self.combo_categoria['values']: self.categoria_var.set("Todos")

    def atualizar_combo_canal_venda(self):
        nomes = self._get_lista_canais_venda(); self.combo_canal_venda['values'] = ["Todos", "Nenhum"] + nomes
        if self.canal_venda_var.get() not in self.combo_canal_venda['values']: self.canal_venda_var.set("Todos")

    def atualizar_tabela(self, event=None):
        self.is_manually_sorted=True; sel_previa = self.tree.selection(); self.tree.delete(*self.tree.get_children()); self.tree_id_map={i['tree_id']:i for i in self.estoque}
        self.tree.tag_configure("vendida",background="#4a2e2e",foreground="#f09090"); self.tree.tag_configure("disponivel",background="#2e4d2e",foreground="#a0eea0")
        busca,cat_f,stat_f,canal_f = self.busca_var.get().lower(),self.categoria_var.get(),self.status_var.get(),self.canal_venda_var.get()
        filtrada=self.estoque
        if busca: filtrada=[i for i in filtrada if busca in i['chave'].lower() or busca in i.get('categoria','').lower() or busca in (i.get('comprador')or'').lower() or busca in (i.get('canal_venda')or'').lower()]
        if cat_f != "Todos": filtrada=[i for i in filtrada if i.get("categoria")==cat_f]
        if canal_f == "Todos": pass
        elif canal_f == "Nenhum": filtrada = [i for i in filtrada if not i.get("canal_venda")]
        else: filtrada = [i for i in filtrada if i.get("canal_venda") == canal_f]
        if stat_f != "Todos": filtrada=[i for i in filtrada if i.get("vendida",0)==(1 if stat_f=="Vendida" else 0)]
        filtrada.sort(key=lambda x:x.get('ordem_manual',x.get('id')))
        for item in filtrada:
            tag="vendida" if item.get("vendida",0) else "disponivel"
            valores = (item["chave"], item.get("categoria","S/C"), "Vendida" if item.get("vendida") else "Disponível", item.get("comprador") or "", item.get("canal_venda") or "", item.get("data_venda") or "")
            self.tree.insert("",tk.END,iid=item['tree_id'],values=valores,tags=(tag,))
        try: self.tree.selection_set([i for i in sel_previa if self.tree.exists(i)])
        except tk.TclError: pass
        self.atualizar_status_bar()

    def atualizar_status_bar(self, event=None):
        texto = f"Total: {len(self.estoque)} | Mostrando: {len(self.tree.get_children())} | Selecionadas: {len(self.tree.selection())}"; self.status_counts_var.set(texto)

    def ordenar_por(self, col):
        self.is_manually_sorted = False; rev = getattr(self,"ord_rev",False) if getattr(self,"last_col",None)==col else False
        lista=[(self.tree.item(i,'values'), i) for i in self.tree.get_children()]; col_idx = self.tree["columns"].index(col)
        lista.sort(key=lambda x:str(x[0][col_idx] or "").lower(), reverse=rev)
        for i,(v,iid) in enumerate(lista): self.tree.move(iid, '', i)
        self.last_col=col; self.ord_rev = not rev

    def copiar_chave_selecionada(self, event=None):
        if event is not None and isinstance(self.focus_get(), (tk.Text, ttk.Entry, tk.Listbox)): return
        if not (sel := self.tree.selection()):
            if event is None: messagebox.showwarning("Copiar", "Selecione uma ou mais chaves na tabela.")
            return
        pyperclip.copy("\n".join([self.tree.item(i,"values")[0] for i in sel]))
        if event is None: messagebox.showinfo("Copiado", f"{len(sel)} chave(s) copiada(s).")

    def acao_editar_selecao(self, event=None):
        sel=self.tree.selection()
        if len(sel) == 0: messagebox.showwarning("Editar","Selecione uma chave."); return
        elif len(sel)==1: self.janela_editar_chave()
        else: self.janela_editar_varias_chaves()

    def janela_editar_varias_chaves(self):
        sel, num_chaves = self.tree.selection(), len(self.tree.selection())
        popup = tk.Toplevel(self); popup.title("Edição em Massa"); popup.geometry("450x350"); popup.grab_set(); popup.resizable(False,False); popup.configure(bg=self.bg_color)
        mf = ttk.Frame(popup,padding=15, style="TFrame"); mf.pack(fill=tk.BOTH, expand=True)
        ttk.Label(mf,text=f"Editando {num_chaves} chaves",font=('Segoe UI',12,'bold')).pack(pady=(0,20))
        fc = ttk.Frame(mf, style="TFrame"); fc.pack(fill=tk.X,pady=5); alt_cat=tk.BooleanVar(); ttk.Checkbutton(fc,text="Alterar Categoria:",variable=alt_cat,style="TCheckbutton").pack(side=tk.LEFT); cat_var=tk.StringVar(); c_cat=ttk.Combobox(fc,textvariable=cat_var,state="readonly",values=[c['nome']for c in self.categorias]); (c_cat.set(self.categorias[0]['nome']) if self.categorias else None); c_cat.pack(side=tk.LEFT,fill=tk.X,expand=True,padx=(5,0))
        f_canal = ttk.Frame(mf, style="TFrame"); f_canal.pack(fill=tk.X, pady=5)
        alt_canal = tk.BooleanVar(); ttk.Checkbutton(f_canal, text="Alterar Canal Venda:", variable=alt_canal, style="TCheckbutton").pack(side=tk.LEFT)
        canal_var = tk.StringVar(); c_canal = ttk.Combobox(f_canal, textvariable=canal_var, values=[''] + self._get_lista_canais_venda()); c_canal.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5,0))
        fs=ttk.Frame(mf, style="TFrame"); fs.pack(fill=tk.X,pady=5); alt_stat=tk.BooleanVar(); ttk.Checkbutton(fs,text="Alterar Status:",variable=alt_stat,style="TCheckbutton").pack(side=tk.LEFT); stat_var=tk.StringVar(value="Disponível"); c_stat=ttk.Combobox(fs,textvariable=stat_var,state="readonly",values=["Disponível","Vendida"]); c_stat.pack(side=tk.LEFT,fill=tk.X,expand=True,padx=(5,0))
        def salvar_massa():
            if not alt_cat.get() and not alt_stat.get() and not alt_canal.get(): messagebox.showwarning("Aviso","Nenhuma alteração selecionada.",parent=popup); return
            ids_editar=[self.tree_id_map[i]['id'] for i in sel if i in self.tree_id_map]; canal_selecionado = canal_var.get().strip() or None
            if not ids_editar: return
            if canal_selecionado: self._garantir_canal_venda_existe(canal_selecionado)
            self.registrar_undo(); campos_upd, params = [],[]
            if alt_cat.get(): campos_upd.append("categoria=?"); params.append(cat_var.get())
            if alt_canal.get(): campos_upd.append("canal_venda=?"); params.append(canal_selecionado)
            if alt_stat.get():
                vendida=1 if stat_var.get()=="Vendida" else 0; campos_upd.append("vendida=?"); params.append(vendida)
                if not vendida: campos_upd.extend(["comprador=NULL","data_venda=NULL","preco_venda_brl=NULL","preco_venda_usd=NULL"])
            placeh = ','.join(['?']*len(ids_editar)); query=f"UPDATE chaves SET {', '.join(campos_upd)} WHERE id IN ({placeh})"; params.extend(ids_editar)
            conn=sqlite3.connect(DB_NAME); conn.execute(query,params); conn.commit(); conn.close()
            self.salvar_e_atualizar_tudo(); logar_acao(f"Edição em massa em {num_chaves} chaves."); messagebox.showinfo("Sucesso","Chaves atualizadas.",parent=self); popup.destroy()
        fb=ttk.Frame(popup, style="TFrame"); fb.pack(side=tk.BOTTOM, pady=15); ttk.Button(fb,text="Salvar",command=salvar_massa).pack(side=tk.LEFT,padx=5); ttk.Button(fb,text="Cancelar",command=popup.destroy).pack(side=tk.LEFT,padx=5)

    def janela_editar_chave(self, event=None):
        if not(sel_id:=self.tree.selection()): messagebox.showwarning("Editar","Selecione uma chave."); return
        if not(chave_obj:=self.tree_id_map.get(sel_id[0])): messagebox.showerror("Erro", "Chave não encontrada."); return
        popup=tk.Toplevel(self); popup.title("Editar Chave"); popup.geometry("400x580"); popup.grab_set(); popup.configure(bg=self.bg_color)
        mf=ttk.Frame(popup,padding=10, style="TFrame"); mf.pack(fill=tk.BOTH,expand=True)
        ttk.Label(mf, text="Chave:").pack(anchor="w", pady=(5,0)); entry_chave = tk.Text(mf, height=3, bg=self.entry_bg, fg=self.text_color, insertbackground=self.text_color, relief="flat", borderwidth=1); entry_chave.pack(fill=tk.X,pady=2); entry_chave.insert("1.0", chave_obj["chave"])
        ttk.Label(mf,text="Categoria:").pack(anchor="w", pady=(5,0)); cat_var=tk.StringVar(value=chave_obj.get("categoria","S/C")); ttk.Combobox(mf,textvariable=cat_var,state="readonly",values=[c['nome'] for c in self.categorias]).pack(fill=tk.X, pady=2)
        ttk.Label(mf, text="Canal de Venda:").pack(anchor="w", pady=(5,0)); canal_var = tk.StringVar(value=chave_obj.get("canal_venda") or "")
        ttk.Combobox(mf,textvariable=canal_var, values=[''] + self._get_lista_canais_venda()).pack(fill=tk.X, pady=2)
        ttk.Label(mf,text="Status:").pack(anchor="w",pady=(5,0)); status_var=tk.StringVar(value="Vendida" if chave_obj.get("vendida") else "Disponível"); ttk.Combobox(mf,textvariable=status_var,state="readonly",values=["Disponível", "Vendida"]).pack(fill=tk.X, pady=2)
        ttk.Label(mf,text="Comprador:").pack(anchor="w",pady=(5,0)); entry_comp=ttk.Entry(mf); entry_comp.pack(fill=tk.X,pady=2); entry_comp.insert(0,chave_obj.get("comprador") or "")
        ttk.Label(mf,text="Data Venda(YYYY-MM-DD HH:MM:SS):").pack(anchor="w",pady=(5,0)); entry_data=ttk.Entry(mf); entry_data.pack(fill=tk.X,pady=2); entry_data.insert(0,chave_obj.get("data_venda") or "")
        ttk.Label(mf,text="Preço Venda(R$):").pack(anchor="w",pady=(5,0)); entry_brl=ttk.Entry(mf); entry_brl.pack(fill=tk.X,pady=2); entry_brl.insert(0,f"{chave_obj.get('preco_venda_brl') or 0.0:.2f}")
        ttk.Label(mf,text="Preço Venda(US$):").pack(anchor="w",pady=(5,0)); entry_usd=ttk.Entry(mf); entry_usd.pack(fill=tk.X,pady=2); entry_usd.insert(0,f"{chave_obj.get('preco_venda_usd') or 0.0:.2f}")
        def salvar():
            if not(nova_chave := entry_chave.get("1.0",tk.END).strip()): messagebox.showwarning("Aviso", "Chave não pode ser vazia.", parent=popup); return
            try: preco_brl,preco_usd = float(entry_brl.get().replace(",","")),float(entry_usd.get().replace(",",""))
            except ValueError: messagebox.showerror("Erro", "Preços devem ser numéricos.", parent=popup); return
            self.registrar_undo(); vendida=1 if status_var.get()=="Vendida" else 0; canal_venda = canal_var.get().strip() or None; comprador=entry_comp.get().strip() if vendida else None; data_venda=entry_data.get().strip() if vendida else None
            if canal_venda: self._garantir_canal_venda_existe(canal_venda)
            if vendida and comprador and not data_venda: data_venda=f"{datetime.now():%Y-%m-%d %H:%M:%S}"
            if not vendida: comprador,data_venda,preco_brl,preco_usd=None,None,None,None
            conn=sqlite3.connect(DB_NAME); conn.execute("UPDATE chaves SET chave=?,categoria=?,vendida=?,comprador=?,data_venda=?,preco_venda_brl=?,preco_venda_usd=?,canal_venda=? WHERE id=?",(nova_chave,cat_var.get(),vendida,comprador,data_venda,preco_brl,preco_usd,canal_venda,chave_obj['id'])); conn.commit(); conn.close()
            self.salvar_e_atualizar_tudo(); logar_acao(f"Chave ID {chave_obj['id']} editada."); messagebox.showinfo("Sucesso","Chave atualizada.",parent=self); popup.destroy()
        fb=ttk.Frame(mf, style="TFrame"); fb.pack(pady=20); ttk.Button(fb,text="Salvar",command=salvar).pack(side=tk.LEFT,padx=5); ttk.Button(fb,text="Cancelar",command=popup.destroy).pack(side=tk.LEFT,padx=5)
    
    def obter_cotacao_dolar(self, cotacao_var):
        try:
            api_url = "https://economia.awesomeapi.com.br/last/USD-BRL"; response = requests.get(api_url, timeout=5)
            response.raise_for_status(); data = response.json(); cotacao = float(data['USDBRL']['bid']); cotacao_var.set(f"{cotacao:.2f}")
        except requests.exceptions.RequestException as e: messagebox.showwarning("Erro de Rede", f"Não foi possível buscar a cotação do dólar.\nVerifique sua conexão ou a API.\nErro: {e}", parent=self)
        except Exception as e: messagebox.showerror("Erro Inesperado", f"Ocorreu um erro ao processar a cotação.\n{e}", parent=self)

    def janela_dashboard_vendas(self):
        popup = tk.Toplevel(self); popup.title("Dashboard de Vendas"); popup.geometry("1000x600"); popup.grab_set(); popup.configure(bg=self.bg_color)
        mf = ttk.Frame(popup, padding=15, style="TFrame"); mf.pack(fill=tk.BOTH, expand=True); mf.rowconfigure(2,weight=1); mf.columnconfigure(0,weight=1)
        filtro_f = ttk.LabelFrame(mf,text=" Filtros "); filtro_f.grid(row=0,column=0,sticky="ew", pady=(0,10))
        content_filtro = ttk.Frame(filtro_f, style="TFrame"); content_filtro.pack(fill="x", expand=True, padx=5, pady=5)
        ttk.Label(content_filtro, text="Período:").pack(side=tk.LEFT,padx=(10,5),pady=10)
        periodo_var = tk.StringVar(value="Últimos 30 dias")
        combo_periodo = ttk.Combobox(content_filtro, textvariable=periodo_var, state="readonly", width=18, values=["Hoje", "Ontem", "Últimos 7 dias", "Últimos 30 dias", "Este Mês", "Mês Passado", "Personalizado"]); combo_periodo.pack(side=tk.LEFT, pady=10)
        ttk.Label(content_filtro, text="Início:").pack(side=tk.LEFT,padx=(15,5),pady=10); e_data_ini=ttk.Entry(content_filtro,width=12); e_data_ini.pack(side=tk.LEFT, pady=10)
        ttk.Label(content_filtro, text="Fim:").pack(side=tk.LEFT,padx=(15,5), pady=10); e_data_fim=ttk.Entry(content_filtro, width=12); e_data_fim.pack(side=tk.LEFT,pady=10)
        def _set_date_from_preset(event=None):
            periodo = periodo_var.get(); hoje = datetime.now(); d_ini, d_fim = hoje, hoje
            if periodo == "Personalizado": e_data_ini.config(state='normal'); e_data_fim.config(state='normal'); return
            if periodo == "Hoje": d_ini = d_fim = hoje
            elif periodo == "Ontem": d_ini = d_fim = hoje - timedelta(days=1)
            elif periodo == "Últimos 7 dias": d_fim = hoje; d_ini = hoje - timedelta(days=6)
            elif periodo == "Últimos 30 dias": d_fim = hoje; d_ini = hoje - timedelta(days=29)
            elif periodo == "Este Mês": d_fim = hoje; d_ini = hoje.replace(day=1)
            elif periodo == "Mês Passado": primeiro_dia_mes_atual = hoje.replace(day=1); d_fim = primeiro_dia_mes_atual - timedelta(days=1); d_ini = d_fim.replace(day=1)
            e_data_ini.config(state='normal'); e_data_fim.config(state='normal'); e_data_ini.delete(0, tk.END); e_data_ini.insert(0, d_ini.strftime("%Y-%m-%d")); e_data_fim.delete(0, tk.END); e_data_fim.insert(0, d_fim.strftime("%Y-%m-%d")); e_data_ini.config(state='readonly'); e_data_fim.config(state='readonly'); gerar_relatorio()
        combo_periodo.bind("<<ComboboxSelected>>", _set_date_from_preset)
        ttk.Button(content_filtro, text="Gerar Relatório", command=lambda: gerar_relatorio()).pack(side=tk.RIGHT, padx=(10,5), pady=10)
        ttk.Button(content_filtro, text="Atualizar Cotação", command=lambda: self.obter_cotacao_dolar(cotacao_var), width=18).pack(side=tk.RIGHT, padx=(5,5), pady=10)
        e_cotacao = ttk.Entry(content_filtro,textvariable=(cotacao_var:=tk.StringVar(value="5.00")), width=8); e_cotacao.pack(side=tk.RIGHT, pady=10)
        ttk.Label(content_filtro, text="Cotação Dólar(R$):").pack(side=tk.RIGHT,padx=(15,2),pady=10)
        resumo_f = ttk.LabelFrame(mf,text=" Resumo do Período "); resumo_f.grid(row=1,column=0,sticky="ew",pady=10)
        content_resumo = ttk.Frame(resumo_f, style="TFrame"); content_resumo.pack(fill="x", expand=True, padx=5, pady=5)
        for i in range(3): content_resumo.columnconfigure(i, weight=1)
        tot_vendas, rec_tot, custo_tot, lucro_tot = tk.StringVar(value="Vendas: 0"), tk.StringVar(value="Receita TOTAL: R$ 0,00 / US$ 0,00"), tk.StringVar(value="Custo TOTAL: R$ 0,00 / US$ 0,00"), tk.StringVar(value="LUCRO TOTAL: R$ 0,00 / US$ 0,00")
        ttk.Label(content_resumo,textvariable=rec_tot,font=('Segoe UI',10)).grid(row=0,column=0,sticky="w",padx=10,pady=5)
        ttk.Label(content_resumo,textvariable=custo_tot,font=('Segoe UI',10)).grid(row=0,column=1,sticky="w",padx=10,pady=5)
        ttk.Label(content_resumo,textvariable=lucro_tot,foreground="#90ee90",font=('Segoe UI',12,'bold')).grid(row=1,column=0,columnspan=2,sticky="w",padx=10,pady=5)
        ttk.Label(content_resumo,textvariable=tot_vendas,font=('Segoe UI',11,'bold')).grid(row=0,rowspan=2,column=2,sticky="e",padx=20)
        detalhes_f = ttk.LabelFrame(mf,text=" Detalhes por Categoria (Valores em R$) "); detalhes_f.grid(row=2,column=0,sticky="nsew",pady=10)
        content_detalhes = ttk.Frame(detalhes_f, style="TFrame"); content_detalhes.pack(fill="both", expand=True); content_detalhes.rowconfigure(0, weight=1); content_detalhes.columnconfigure(0, weight=1)
        tree = ttk.Treeview(content_detalhes,columns=("cat","qtd","rec","custo","lucro","lucro_medio"),show="headings"); tree.grid(row=0,column=0,sticky="nsew")
        ys = ttk.Scrollbar(content_detalhes, orient='vertical', command=tree.yview); tree.configure(yscrollcommand=ys.set); ys.grid(row=0, column=1, sticky='ns')
        headings = {"cat":"Categoria", "qtd":"Qtd Vendida", "rec":"Receita Total", "custo":"Custo Total", "lucro":"Lucro Total", "lucro_medio":"Lucro Médio/Venda"}
        for col,txt in headings.items(): tree.heading(col, text=txt, anchor=tk.CENTER)
        widths = {"cat":200, "qtd":100, "rec":130, "custo":130, "lucro":130, "lucro_medio":150}
        for col,w in widths.items(): tree.column(col,width=w,anchor=tk.CENTER)
        def format_brl(val): return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        def format_usd(val,cotacao): return f"US$ {val/cotacao if cotacao > 0 else 0:.2f}"
        def gerar_relatorio():
            if not e_data_ini.get() or not e_data_fim.get(): return
            try: cotacao=float(cotacao_var.get().replace(",","."))
            except ValueError: messagebox.showerror("Erro", "Cotação inválida.",parent=popup); return
            d_ini, d_fim = e_data_ini.get(), e_data_fim.get()
            try: dt_fim_query=(datetime.strptime(d_fim,"%Y-%m-%d")+timedelta(days=1)).strftime("%Y-%m-%d"); dt_ini_query = datetime.strptime(d_ini,"%Y-%m-%d").strftime("%Y-%m-%d")
            except ValueError: messagebox.showerror("Erro","Formato de data inválido (Use AAAA-MM-DD).",parent=popup); return
            conn=sqlite3.connect(DB_NAME); conn.row_factory=sqlite3.Row; c=conn.cursor()
            c.execute("SELECT c.categoria,c.preco_venda_brl,c.preco_venda_usd,cat.custo_padrao_brl,cat.custo_padrao_usd FROM chaves AS c LEFT JOIN categorias AS cat ON c.categoria=cat.nome WHERE c.vendida=1 AND c.data_venda>=? AND c.data_venda<?", (dt_ini_query, dt_fim_query))
            vendas = c.fetchall(); conn.close()
            tot_rec, tot_custo = 0.0, 0.0; stats = defaultdict(lambda:{"qtd":0, "rec":0, "custo":0})
            for v in vendas:
                rec_conv = (v["preco_venda_brl"] or 0) + ((v["preco_venda_usd"] or 0) * cotacao); custo_conv = (v["custo_padrao_brl"] or 0) + ((v["custo_padrao_usd"] or 0) * cotacao)
                tot_rec += rec_conv; tot_custo += custo_conv; cat_nome = v["categoria"] if v["categoria"] else "Sem Categoria"
                stats[cat_nome]['qtd'] += 1; stats[cat_nome]['rec'] += rec_conv; stats[cat_nome]['custo'] += custo_conv
            tot_lucro = tot_rec - tot_custo
            tot_vendas.set(f"Vendas: {len(vendas)}"); rec_tot.set(f"Receita TOTAL: {format_brl(tot_rec)} / {format_usd(tot_rec, cotacao)}"); custo_tot.set(f"Custo TOTAL: {format_brl(tot_custo)} / {format_usd(tot_custo, cotacao)}"); lucro_tot.set(f"LUCRO TOTAL: {format_brl(tot_lucro)} / {format_usd(tot_lucro, cotacao)}")
            tree.delete(*tree.get_children())
            for cat, data in sorted(stats.items()):
                lucro = data['rec'] - data['custo']; lucro_m = lucro / data['qtd'] if data['qtd'] else 0
                tree.insert("","end", values=(cat, data['qtd'], format_brl(data['rec']), format_brl(data['custo']), format_brl(lucro), format_brl(lucro_m)))
        self.obter_cotacao_dolar(cotacao_var); popup.after(150, _set_date_from_preset)

    def janela_configurar_email(self):
        popup = tk.Toplevel(self); popup.title("Configurações de Email"); popup.geometry("500x300"); popup.grab_set(); popup.resizable(False, False); popup.configure(bg=self.bg_color)
        mf = ttk.Frame(popup, padding=15, style="TFrame"); mf.pack(fill=tk.BOTH, expand=True); mf.columnconfigure(1, weight=1)
        config = self.carregar_config_email()
        email_var = tk.StringVar(value=config.get("email", "")); senha_var = tk.StringVar(value=config.get("senha", "")); servidor_var = tk.StringVar(value=config.get("servidor", "smtp.gmail.com")); porta_var = tk.StringVar(value=config.get("porta", "587"))
        ttk.Label(mf, text="Email do Remetente:").grid(row=0, column=0, sticky="w", pady=5, padx=5); ttk.Entry(mf, textvariable=email_var).grid(row=0, column=1, sticky="ew", pady=5, padx=5)
        ttk.Label(mf, text="Senha/App Password:").grid(row=1, column=0, sticky="w", pady=5, padx=5); ttk.Entry(mf, textvariable=senha_var, show="*").grid(row=1, column=1, sticky="ew", pady=5, padx=5)
        ttk.Label(mf, text="Servidor SMTP:").grid(row=2, column=0, sticky="w", pady=5, padx=5); ttk.Entry(mf, textvariable=servidor_var).grid(row=2, column=1, sticky="ew", pady=5, padx=5)
        ttk.Label(mf, text="Porta SMTP:").grid(row=3, column=0, sticky="w", pady=5, padx=5); ttk.Entry(mf, textvariable=porta_var).grid(row=3, column=1, sticky="ew", pady=5, padx=5)
        ttk.Label(mf, text="Atenção: Use 'Senhas de App' para Gmail, Outlook, etc.", font=('Segoe UI', 8, 'italic'), foreground="yellow").grid(row=4, column=0, columnspan=2, pady=(10,0))
        def salvar_config():
            nova_config = {"email": email_var.get().strip(), "senha": senha_var.get().strip(), "servidor": servidor_var.get().strip(), "porta": porta_var.get().strip()}
            with open(EMAIL_CONFIG_FILE, "w") as f: json.dump(nova_config, f, indent=4)
            messagebox.showinfo("Sucesso", "Configurações de email salvas!", parent=popup); popup.destroy()
        botoes_f = ttk.Frame(popup, style="TFrame"); botoes_f.pack(pady=10)
        ttk.Button(botoes_f, text="Salvar", command=salvar_config).pack(side=tk.LEFT, padx=5); ttk.Button(botoes_f, text="Cancelar", command=popup.destroy).pack(side=tk.LEFT, padx=5)

    def carregar_config_email(self):
        try:
            with open(EMAIL_CONFIG_FILE, "r") as f: return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError): return {}

    def enviar_email_com_chave(self, destinatario, assunto, corpo, caminho_anexo=None):
        config = self.carregar_config_email()
        if not all(k in config and config[k] for k in ["email", "senha", "servidor", "porta"]):
            logar_acao("ERRO: Tentativa de enviar email sem configuração completa."); messagebox.showwarning("Email não Configurado", "As configurações de email estão incompletas.\n\nVá em Ferramentas > Configurar Email... para ajustá-las."); return
        try:
            msg = MIMEMultipart(); msg['From'] = config['email']; msg['To'] = destinatario; msg['Subject'] = assunto
            corpo_html = corpo.replace('\n', '<br>'); corpo_html = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', corpo_html); msg.attach(MIMEText(corpo_html, 'html', 'utf-8'))
            if caminho_anexo and os.path.exists(caminho_anexo):
                with open(caminho_anexo, "rb") as anexo_file: part = MIMEApplication(anexo_file.read(), Name=os.path.basename(caminho_anexo))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_anexo)}"'
                msg.attach(part); logar_acao(f"Anexando PDF: {caminho_anexo}")
            server = smtplib.SMTP(config['servidor'], int(config['porta'])); server.starttls(); server.login(config['email'], config['senha']); server.sendmail(config['email'], destinatario, msg.as_string()); server.quit()
            logar_acao(f"Email enviado com sucesso para {destinatario}"); messagebox.showinfo("Email Enviado", f"Email enviado com sucesso para {destinatario}.")
        except Exception as e: logar_acao(f"FALHA ao enviar email para {destinatario}. Erro: {e}"); messagebox.showerror("Erro de Email", f"Não foi possível enviar o email.\n\nVerifique suas configurações, conexão e senha de app.\n\nErro: {e}")

if __name__ == "__main__":
    app = GerenciadorChaves()
    s = ttk.Style()
    s.configure("Accent.TButton", background="#094771", font=('Segoe UI', 9, 'bold'))
    s.map("Accent.TButton", background=[('active', '#0a588a')])
    app.mainloop()