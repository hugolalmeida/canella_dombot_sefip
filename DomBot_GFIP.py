"""
DomBot GFIP
===========
Automação RPA para exportação do relatório GFIP no sistema Domínio Folha.

Caminho: Relatórios → Informativos → (seta) → (seta) → Enter
Preenchimento: Competência de/até → Código de Recolhimento → Caminho do arquivo → OK
O Domínio salva o TXT direto no caminho informado (nome de arquivo fixo, não editável),
por isso cada empresa é salva em uma subpasta própria dentro da pasta de destino.

Autor: Hugo L. Almeida
Versão: 1.0
"""

import argparse
import ctypes
import json
import logging
import os
import subprocess
import sys
import threading
import time
import traceback
from datetime import datetime
from typing import Optional, Tuple

from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto import findwindows, timings
import win32gui
import win32con
import win32process
from ctypes import wintypes


def _hwnd_com_foco_teclado() -> int:
    """
    Retorna o hwnd do controle que realmente possui o foco de teclado no
    momento, via GetGUIThreadInfo (funciona entre processos, diferente de
    GetFocus() puro que só enxerga a thread chamadora). Usado para confirmar
    que send_keys vai atingir o controle esperado antes de digitar texto
    livre — evita que letras "vazem" como atalhos de tela quando o foco
    real está em outro lugar (ex: fora de qualquer campo de texto).
    """
    hwnd_fg = win32gui.GetForegroundWindow()
    if not hwnd_fg:
        return 0
    tid, _ = win32process.GetWindowThreadProcessId(hwnd_fg)

    class GUITHREADINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("flags", wintypes.DWORD),
            ("hwndActive", wintypes.HWND),
            ("hwndFocus", wintypes.HWND),
            ("hwndCapture", wintypes.HWND),
            ("hwndMenuOwner", wintypes.HWND),
            ("hwndMoveSize", wintypes.HWND),
            ("hwndCaret", wintypes.HWND),
            ("rcCaret", wintypes.RECT),
        ]

    info = GUITHREADINFO()
    info.cbSize = ctypes.sizeof(GUITHREADINFO)
    ok = ctypes.windll.user32.GetGUIThreadInfo(tid, ctypes.byref(info))
    if not ok:
        return 0
    return int(info.hwndFocus) if info.hwndFocus else 0


# Imports de GUI — só disponíveis quando rodando com ambiente desktop completo
try:
    import customtkinter as ctk
    import pandas as pd
    import tkinter.messagebox as messagebox
    from PIL import Image, ImageDraw
    _GUI_DISPONIVEL = True
except ImportError:
    _GUI_DISPONIVEL = False


# ── Log handler para a GUI ──────────────────────────────────────────────────

class GUILogHandler(logging.Handler):
    def __init__(self, gui):
        super().__init__()
        self.gui = gui

    def emit(self, record):
        msg = self.format(record)
        self.gui.window.after(0, lambda: self.gui.adicionar_log(msg, record.levelno))


# ── Interface gráfica ────────────────────────────────────────────────────────

class AutomacaoGUI:
    CORES = {
        'sucesso':     '#2ECC71',
        'erro':        '#E74C3C',
        'aviso':       '#F39C12',
        'info':        '#3498DB',
        'texto':       '#ECF0F1',
        'fundo_card':  '#2C3E50',
        'fundo_escuro':'#1A252F',
        'destaque':    '#1ABC9C',
        'processando': '#9B59B6',
    }

    def __init__(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("green")

        self.window = ctk.CTk()
        self.window.title("DomBot - GFIP v1.0")
        self.window.geometry("860x580")
        self.window.minsize(800, 520)
        self.window.protocol("WM_DELETE_WINDOW", self.ao_fechar)

        self.executando = False
        self.pausa_solicitada = False
        self.thread_automacao = None
        self.df_carregado = None

        self.stats = {'processados': 0, 'sucesso': 0, 'erros': 0, 'tempo_inicio': None}

        self.logs_dir = os.path.join(os.path.dirname(__file__), "logs")
        os.makedirs(self.logs_dir, exist_ok=True)
        self.setup_file_logging()

        # Variáveis de UI
        self.arquivo_excel    = ctk.StringVar()
        self.linha_inicial    = ctk.StringVar(value="2")
        self.comp_inicial     = ctk.StringVar(value=datetime.now().strftime("%m/%Y"))
        self.comp_final       = ctk.StringVar(value=datetime.now().strftime("%m/%Y"))
        self.dir_saida        = ctk.StringVar()
        self.status_var       = ctk.StringVar(value="Aguardando início...")

        self.total_linhas      = 0
        self.linhas_processadas = 0
        self.linhas_com_erro   = 0

        self.logger = logging.getLogger('DomBotGFIP')
        self.logger.setLevel(logging.INFO)
        self.logger.handlers = []
        handler = GUILogHandler(self)
        handler.setFormatter(logging.Formatter('%(message)s'))
        self.logger.addHandler(handler)

        self.criar_interface()

    # ── Logging em arquivo ──────────────────────────────────────────────────

    def setup_file_logging(self):
        data = datetime.now().strftime("%Y-%m-%d")
        fmt  = logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')

        self.success_logger = logging.getLogger('GFIPSuccess')
        self.success_logger.setLevel(logging.INFO)
        if not self.success_logger.handlers:
            h = logging.FileHandler(os.path.join(self.logs_dir, f'gfip_success_{data}.log'), encoding='utf-8')
            h.setFormatter(fmt)
            self.success_logger.addHandler(h)

        self.error_logger = logging.getLogger('GFIPError')
        self.error_logger.setLevel(logging.ERROR)
        if not self.error_logger.handlers:
            h = logging.FileHandler(os.path.join(self.logs_dir, f'gfip_error_{data}.log'), encoding='utf-8')
            h.setFormatter(fmt)
            self.error_logger.addHandler(h)

    # ── Interface ────────────────────────────────────────────────────────────

    def criar_interface(self):
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(0, weight=1)
        main = ctk.CTkFrame(self.window, fg_color="transparent")
        main.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(3, weight=1)

        self.criar_header(main)
        self.criar_painel_config(main)
        self.criar_painel_stats(main)
        self.criar_area_logs(main)

    def criar_header(self, parent):
        frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        frame.grid_columnconfigure(1, weight=1)

        logo_frame = ctk.CTkFrame(frame, fg_color=self.CORES['destaque'], width=44, height=44, corner_radius=22)
        logo_frame.grid(row=0, column=0, padx=10, pady=8)
        logo_frame.grid_propagate(False)
        ctk.CTkLabel(logo_frame, text="📄", font=ctk.CTkFont(size=18)).place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(frame, text="DomBot — GFIP",
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=self.CORES['texto']).grid(row=0, column=1, sticky="w", padx=5)

        self.status_frame = ctk.CTkFrame(frame, fg_color="transparent")
        self.status_frame.grid(row=0, column=2, padx=10)
        self.status_indicator = ctk.CTkFrame(self.status_frame, fg_color="#7F8C8D", width=10, height=10, corner_radius=5)
        self.status_indicator.pack(side="left", padx=(0, 6))
        ctk.CTkLabel(self.status_frame, textvariable=self.status_var,
                     font=ctk.CTkFont(size=11), text_color="#95A5A6").pack(side="left")

    def criar_painel_config(self, parent):
        frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        frame.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        frame.grid_columnconfigure(0, weight=1)

        # Linha 1 — Excel + Linha inicial + Botões
        r1 = ctk.CTkFrame(frame, fg_color="transparent")
        r1.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        r1.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(r1, text="📊", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=(0, 4))
        ctk.CTkEntry(r1, textvariable=self.arquivo_excel,
                     placeholder_text="Planilha Excel com lista de empresas...",
                     height=32, font=ctk.CTkFont(size=11)).grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ctk.CTkButton(r1, text="Procurar", command=self.selecionar_arquivo,
                      width=80, height=32, font=ctk.CTkFont(size=11),
                      fg_color=self.CORES['info'], hover_color="#2980B9").grid(row=0, column=2, padx=(0, 14))

        ctk.CTkLabel(r1, text="Linha:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(row=0, column=3, padx=(0, 3))
        ctk.CTkEntry(r1, textvariable=self.linha_inicial, width=50, height=32,
                     font=ctk.CTkFont(size=11), justify="center").grid(row=0, column=4, padx=(0, 14))

        self.btn_iniciar = ctk.CTkButton(r1, text="▶ Iniciar", command=self.iniciar_thread,
                                         width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
                                         fg_color=self.CORES['sucesso'], hover_color="#27AE60")
        self.btn_iniciar.grid(row=0, column=5, padx=3)

        self.btn_pausar = ctk.CTkButton(r1, text="⏸ Pausar", command=self.pausar,
                                         width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
                                         fg_color=self.CORES['aviso'], hover_color="#E67E22", state="disabled")
        self.btn_pausar.grid(row=0, column=6, padx=3)

        self.btn_parar = ctk.CTkButton(r1, text="⏹ Parar", command=self.parar,
                                        width=90, height=32, font=ctk.CTkFont(size=11, weight="bold"),
                                        fg_color=self.CORES['erro'], hover_color="#C0392B", state="disabled")
        self.btn_parar.grid(row=0, column=7, padx=(3, 0))

        # Linha 2 — Competências + Diretório de saída
        r2 = ctk.CTkFrame(frame, fg_color="transparent")
        r2.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
        r2.grid_columnconfigure(5, weight=1)

        ctk.CTkLabel(r2, text="📅", font=ctk.CTkFont(size=14)).grid(row=0, column=0, padx=(0, 4))
        ctk.CTkLabel(r2, text="Competência de:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(row=0, column=1, padx=(0, 4))
        ctk.CTkEntry(r2, textvariable=self.comp_inicial, width=80, height=32,
                     font=ctk.CTkFont(size=11), justify="center",
                     placeholder_text="MM/AAAA").grid(row=0, column=2, padx=(0, 6))
        ctk.CTkLabel(r2, text="até:", font=ctk.CTkFont(size=11), text_color="#BDC3C7").grid(row=0, column=3, padx=(0, 4))
        ctk.CTkEntry(r2, textvariable=self.comp_final, width=80, height=32,
                     font=ctk.CTkFont(size=11), justify="center",
                     placeholder_text="MM/AAAA").grid(row=0, column=4, padx=(0, 14))

        ctk.CTkLabel(r2, text="📁", font=ctk.CTkFont(size=14)).grid(row=0, column=5, padx=(0, 4))
        ctk.CTkEntry(r2, textvariable=self.dir_saida,
                     placeholder_text="Pasta raiz de destino (uma subpasta por empresa)...",
                     height=32, font=ctk.CTkFont(size=11)).grid(row=0, column=6, sticky="ew", padx=(0, 6))
        r2.grid_columnconfigure(6, weight=1)
        ctk.CTkButton(r2, text="Selecionar", command=self.selecionar_pasta,
                      width=80, height=32, font=ctk.CTkFont(size=11),
                      fg_color=self.CORES['info'], hover_color="#2980B9").grid(row=0, column=7)

    def criar_painel_stats(self, parent):
        frame = ctk.CTkFrame(parent, fg_color=self.CORES['fundo_card'], corner_radius=8)
        frame.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        for i in range(4):
            frame.grid_columnconfigure(i, weight=1)

        self._stat_card(frame, 0, "📋", "Total",   "total_label",   "0")
        self._stat_card(frame, 1, "✅", "Sucesso",  "sucesso_label", "0", self.CORES['sucesso'])
        self._stat_card(frame, 2, "❌", "Erros",    "erros_label",   "0", self.CORES['erro'])
        self._stat_card(frame, 3, "🏢", "Empresa",  "empresa_label", "-", self.CORES['info'])

        prog_frame = ctk.CTkFrame(frame, fg_color="transparent")
        prog_frame.grid(row=1, column=0, columnspan=4, sticky="ew", padx=10, pady=(2, 8))
        prog_frame.grid_columnconfigure(0, weight=1)
        self.progress_bar = ctk.CTkProgressBar(prog_frame, height=6, corner_radius=3,
                                                progress_color=self.CORES['destaque'])
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_bar.set(0)
        self.progress_label = ctk.CTkLabel(prog_frame, text="0%",
                                            font=ctk.CTkFont(size=10), text_color="#95A5A6")
        self.progress_label.grid(row=0, column=1, padx=(8, 0))

    def _stat_card(self, parent, col, icon, titulo, attr, val, cor=None):
        card = ctk.CTkFrame(parent, fg_color="transparent")
        card.grid(row=0, column=col, padx=5, pady=8)
        ctk.CTkLabel(card, text=f"{icon} {titulo}", font=ctk.CTkFont(size=10), text_color="#7F8C8D").pack()
        lbl = ctk.CTkLabel(card, text=val, font=ctk.CTkFont(size=14, weight="bold"),
                            text_color=cor if cor else self.CORES['texto'])
        lbl.pack()
        setattr(self, attr, lbl)

    def criar_area_logs(self, parent):
        tab = ctk.CTkTabview(parent, fg_color=self.CORES['fundo_card'],
                              segmented_button_fg_color=self.CORES['fundo_escuro'],
                              segmented_button_selected_color=self.CORES['destaque'],
                              corner_radius=8)
        tab.grid(row=3, column=0, sticky="nsew")

        t_log  = tab.add("📜 Logs")
        t_prev = tab.add("📄 Preview")

        # Aba Logs
        t_log.grid_columnconfigure(0, weight=1)
        t_log.grid_rowconfigure(0, weight=1)
        log_cnt = ctk.CTkFrame(t_log, fg_color="transparent")
        log_cnt.grid(row=0, column=0, sticky="nsew", padx=3, pady=3)
        log_cnt.grid_columnconfigure(0, weight=1)
        log_cnt.grid_rowconfigure(0, weight=1)
        self.log_text = ctk.CTkTextbox(log_cnt, font=ctk.CTkFont(family="Consolas", size=11),
                                        fg_color=self.CORES['fundo_escuro'], corner_radius=6)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        self.log_text._textbox.tag_config("sucesso",    foreground=self.CORES['sucesso'])
        self.log_text._textbox.tag_config("erro",       foreground=self.CORES['erro'])
        self.log_text._textbox.tag_config("aviso",      foreground=self.CORES['aviso'])
        self.log_text._textbox.tag_config("info",       foreground=self.CORES['info'])
        self.log_text._textbox.tag_config("processando",foreground=self.CORES['processando'])

        btn_log = ctk.CTkFrame(log_cnt, fg_color="transparent")
        btn_log.grid(row=1, column=0, sticky="ew", pady=(5, 0))
        ctk.CTkButton(btn_log, text="🗑 Limpar", command=lambda: self.log_text.delete("1.0","end"),
                      width=90, height=26, font=ctk.CTkFont(size=10),
                      fg_color="#34495E", hover_color="#2C3E50").pack(side="left")
        ctk.CTkButton(btn_log, text="💾 Exportar", command=self.exportar_logs,
                      width=90, height=26, font=ctk.CTkFont(size=10),
                      fg_color="#34495E", hover_color="#2C3E50").pack(side="left", padx=8)

        # Aba Preview
        t_prev.grid_columnconfigure(0, weight=1)
        t_prev.grid_rowconfigure(1, weight=1)
        self.preview_info = ctk.CTkLabel(t_prev, text="Nenhum arquivo carregado",
                                          font=ctk.CTkFont(size=11), text_color="#95A5A6")
        self.preview_info.grid(row=0, column=0, sticky="w", padx=6, pady=3)
        self.preview_text = ctk.CTkTextbox(t_prev, font=ctk.CTkFont(family="Consolas", size=10),
                                            fg_color=self.CORES['fundo_escuro'], corner_radius=6)
        self.preview_text.grid(row=1, column=0, sticky="nsew", padx=3, pady=(0, 3))

    # ── Ações da UI ──────────────────────────────────────────────────────────

    def selecionar_arquivo(self):
        path = ctk.filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls")], title="Planilha de empresas")
        if path:
            self.arquivo_excel.set(path)
            self.adicionar_log(f"Arquivo: {os.path.basename(path)}", logging.INFO, "info")
            self.carregar_preview()

    def selecionar_pasta(self):
        script = (
            "Add-Type -AssemblyName System.Windows.Forms; "
            "$f = New-Object System.Windows.Forms.FolderBrowserDialog; "
            "$f.Description = 'Pasta raiz de destino dos TXT (uma subpasta por empresa)'; "
            "if ($f.ShowDialog() -eq 'OK') { $f.SelectedPath } else { '' }"
        )
        result = subprocess.run(["powershell", "-Command", script],
                                 capture_output=True, text=True, timeout=120)
        pasta = result.stdout.strip()
        if pasta:
            self.dir_saida.set(pasta)
            self.adicionar_log(f"Pasta de saída: {pasta}", logging.INFO, "info")

    def carregar_preview(self):
        try:
            df = pd.read_excel(self.arquivo_excel.get())
            self.df_carregado = df
            self.total_label.configure(text=str(len(df)))
            self.preview_info.configure(text=f"📄 {os.path.basename(self.arquivo_excel.get())} | {len(df)} linhas")
            self.preview_text.delete("1.0", "end")
            cols = list(df.columns[:5])
            header = " | ".join(f"{c:^15}" for c in cols)
            self.preview_text.insert("end", f"{'─'*len(header)}\n{header}\n{'─'*len(header)}\n")
            for _, row in df.head(40).iterrows():
                self.preview_text.insert("end", " | ".join(f"{str(v)[:15]:^15}" for v in list(row)[:5]) + "\n")
            colunas_ok = {'Nº', 'Empresa', 'Cod_Recolhimento'}.issubset(set(df.columns))
            tag = "sucesso" if colunas_ok else "aviso"
            self.adicionar_log(
                f"Preview: {len(df)} empresas" + ("" if colunas_ok else " — verificar colunas Nº, Empresa e Cod_Recolhimento"),
                logging.INFO, tag)
        except Exception as e:
            self.adicionar_log(f"Erro no preview: {e}", logging.ERROR, "erro")

    def exportar_logs(self):
        try:
            path = ctk.filedialog.asksaveasfilename(
                defaultextension=".txt", filetypes=[("Text", "*.txt")],
                initialfilename=f"gfip_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            if path:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get("1.0", "end"))
                self.adicionar_log(f"Log exportado: {path}", logging.INFO, "sucesso")
        except Exception as e:
            self.adicionar_log(f"Erro ao exportar: {e}", logging.ERROR, "erro")

    def adicionar_log(self, msg, level=logging.INFO, tag=None):
        try:
            ts = datetime.now().strftime('%H:%M:%S')
            if tag is None:
                tag = "erro" if level >= logging.ERROR else "aviso" if level >= logging.WARNING else "info"
            icons = {"sucesso":"✅","erro":"❌","aviso":"⚠️","info":"ℹ️","processando":"⏳"}
            self.log_text.insert("end", f"[{ts}] {icons.get(tag,'•')} ", tag)
            self.log_text.insert("end", f"{msg}\n", tag)
            self.log_text.see("end")
            self.window.update_idletasks()
        except Exception:
            pass

    def atualizar_progresso(self, atual, total):
        pct = atual / total if total > 0 else 0
        self.progress_bar.set(pct)
        self.progress_label.configure(text=f"{pct*100:.1f}%")
        self.status_var.set(f"Processando: {atual}/{total}")
        self.window.update_idletasks()

    def atualizar_status_indicator(self, status):
        cores = {'aguardando':'#7F8C8D','executando':self.CORES['sucesso'],
                 'pausado':self.CORES['aviso'],'erro':self.CORES['erro'],'concluido':self.CORES['info']}
        self.status_indicator.configure(fg_color=cores.get(status,'#7F8C8D'))

    def atualizar_tempo(self):
        if self.stats['tempo_inicio']:
            el = datetime.now() - self.stats['tempo_inicio']
            h, r = divmod(int(el.total_seconds()), 3600)
            m, s = divmod(r, 60)
            if self.executando:
                self.window.after(1000, self.atualizar_tempo)

    # ── Controles de execução ────────────────────────────────────────────────

    def validar_entrada(self) -> Tuple[bool, str]:
        if not self.arquivo_excel.get() or not os.path.exists(self.arquivo_excel.get()):
            return False, "Selecione um arquivo Excel válido"
        try:
            li = int(self.linha_inicial.get())
            if li < 1:
                return False, "Linha inicial inválida"
        except ValueError:
            return False, "Linha inicial deve ser um número"
        if not self.comp_inicial.get() or not self.comp_final.get():
            return False, "Informe as competências"
        if not self.dir_saida.get():
            return False, "Selecione a pasta de destino dos TXT"
        return True, "OK"

    def iniciar_thread(self):
        if self.executando:
            return
        ok, msg = self.validar_entrada()
        if not ok:
            messagebox.showerror("Validação", msg)
            return

        self.linhas_processadas = 0
        self.linhas_com_erro    = 0
        self.sucesso_label.configure(text="0")
        self.erros_label.configure(text="0")
        self.stats['tempo_inicio'] = datetime.now()

        self.thread_automacao = threading.Thread(target=self.executar_automacao, daemon=True)
        self.thread_automacao.start()

        self.btn_iniciar.configure(state="disabled")
        self.btn_pausar.configure(state="normal")
        self.btn_parar.configure(state="normal")
        self.atualizar_status_indicator('executando')
        self.atualizar_tempo()

    def pausar(self):
        if self.executando:
            self.pausa_solicitada = not self.pausa_solicitada
            if self.pausa_solicitada:
                self.btn_pausar.configure(text="▶ Retomar")
                self.status_var.set("Pausado")
                self.atualizar_status_indicator('pausado')
                self.adicionar_log("Pausado", logging.INFO, "aviso")
            else:
                self.btn_pausar.configure(text="⏸ Pausar")
                self.status_var.set("Em execução...")
                self.atualizar_status_indicator('executando')
                self.adicionar_log("Retomado", logging.INFO, "info")

    def parar(self):
        if self.executando:
            self.executando = False
            self.pausa_solicitada = False
            self.adicionar_log("Parando...", logging.INFO, "aviso")
            self.status_var.set("Interrompendo...")
            self.atualizar_status_indicator('erro')

    def ao_fechar(self):
        if self.executando:
            if messagebox.askyesno("Sair", "Automação em execução. Deseja sair?"):
                self.executando = False
                self.window.after(1000, self.window.destroy)
        else:
            self.window.destroy()

    # ── Loop principal de automação ─────────────────────────────────────────

    def executar_automacao(self):
        try:
            self.executando = True
            self.adicionar_log("Iniciando automação GFIP...", logging.INFO, "processando")
            self.status_var.set("Em execução...")

            df = pd.read_excel(self.arquivo_excel.get())
            inicio_idx = int(self.linha_inicial.get()) - 2
            df_proc = df.iloc[inicio_idx:]
            self.total_linhas = len(df_proc)
            self.total_label.configure(text=str(self.total_linhas))
            self.adicionar_log(f"{self.total_linhas} empresa(s) para processar", logging.INFO, "info")

            bot = GFIPAutomation(self.logger, self)
            if not bot.connect_to_dominio():
                self.adicionar_log("Não foi possível conectar ao Domínio Folha", logging.ERROR, "erro")
                return

            for idx, (_, row) in enumerate(df_proc.iterrows()):
                if not self.executando:
                    self.adicionar_log("Interrompido pelo usuário", logging.INFO, "aviso")
                    break
                while self.pausa_solicitada and self.executando:
                    time.sleep(0.5)
                if not self.executando:
                    break

                self.atualizar_progresso(idx + 1, self.total_linhas)
                empresa_num = str(int(row['Nº']))
                empresa_nome = str(row.get('Empresa', empresa_num))
                cod_recolhimento = str(row.get('Cod_Recolhimento', '')).strip()
                self.empresa_label.configure(text=empresa_num[:20])
                self.adicionar_log(f"Empresa {empresa_num} — {empresa_nome}", logging.INFO, "processando")

                try:
                    ok = bot.processar_empresa(
                        empresa_num=empresa_num,
                        empresa_nome=empresa_nome,
                        comp_inicial=self.comp_inicial.get(),
                        comp_final=self.comp_final.get(),
                        cod_recolhimento=cod_recolhimento,
                        dir_saida=self.dir_saida.get(),
                    )
                    if ok:
                        self.linhas_processadas += 1
                        self.sucesso_label.configure(text=str(self.linhas_processadas))
                        self.success_logger.info(f"Empresa {empresa_num} - {empresa_nome} - OK")
                        self.adicionar_log(f"✅ Empresa {empresa_num} exportada", logging.INFO, "sucesso")
                    else:
                        self.linhas_com_erro += 1
                        self.erros_label.configure(text=str(self.linhas_com_erro))
                        self.error_logger.error(f"Empresa {empresa_num} - {empresa_nome} - ERRO")
                        self.adicionar_log(f"Falha na empresa {empresa_num}", logging.ERROR, "erro")
                        time.sleep(2)
                except Exception as e:
                    self.linhas_com_erro += 1
                    self.erros_label.configure(text=str(self.linhas_com_erro))
                    self.error_logger.error(f"Empresa {empresa_num} - exceção: {e}")
                    self.adicionar_log(f"Erro empresa {empresa_num}: {e}", logging.ERROR, "erro")

            if self.executando:
                self.status_var.set("Concluído")
                self.progress_bar.set(1.0)
                self.progress_label.configure(text="100%")
                self.atualizar_status_indicator('concluido')
                self.adicionar_log(
                    f"Concluído — {self.linhas_processadas} OK / {self.linhas_com_erro} erros",
                    logging.INFO, "sucesso")

        except Exception as e:
            self.adicionar_log(f"Erro crítico: {e}", logging.ERROR, "erro")
            self.atualizar_status_indicator('erro')
        finally:
            self.executando = False
            self.pausa_solicitada = False
            self.btn_iniciar.configure(state="normal")
            self.btn_pausar.configure(state="disabled", text="⏸ Pausar")
            self.btn_parar.configure(state="disabled")

    def executar(self):
        self.window.mainloop()


# ── Engine de automação do Domínio ──────────────────────────────────────────

class GFIPAutomation:
    """Navega no Domínio Folha e exporta o relatório GFIP por empresa."""

    def __init__(self, logger, gui):
        timings.Timings.window_find_timeout = 20
        self.app          = None
        self.main_window  = None
        self.logger       = logger
        self.gui          = gui
        self.empresa_atual = None

    def log(self, msg):
        self.logger.info(msg)

    def should_stop(self) -> bool:
        return not self.gui.executando

    def check_pause(self):
        while self.gui.pausa_solicitada and self.gui.executando:
            time.sleep(0.5)

    def smart_sleep(self, seconds: float) -> bool:
        interval = 0.15
        elapsed  = 0.0
        while elapsed < seconds:
            if self.should_stop():
                return False
            self.check_pause()
            t = min(interval, seconds - elapsed)
            time.sleep(t)
            elapsed += t
        return True

    def wait_for(self, fn, timeout=30, poll=0.2, desc="") -> bool:
        start = time.time()
        while time.time() - start < timeout:
            if self.should_stop():
                return False
            self.check_pause()
            try:
                if fn():
                    if desc:
                        self.log(f"{desc} OK ({time.time()-start:.1f}s)")
                    return True
            except Exception:
                pass
            time.sleep(poll)
        if desc:
            self.log(f"{desc} — timeout após {timeout}s")
        return False

    def _force_focus(self, hwnd: int):
        try:
            ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)
            ctypes.windll.user32.keybd_event(0x12, 0, 2, 0)
            # SW_RESTORE só se minimizado — evita redimensionar janela maximizada
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
        except Exception:
            pass

    def _set_focus_cross_process(self, hwnd: int) -> bool:
        """
        Dá foco de teclado a um controle de outro processo (ex: o Domínio).
        win32gui.SetFocus falha com ERROR_ACCESS_DENIED (5) quando chamado
        pela thread do Python diretamente, pois SetFocus só é permitido para
        a thread dona da fila de mensagens da janela alvo. A solução padrão
        do Win32 é anexar temporariamente a thread do Python à thread de
        input do processo alvo via AttachThreadInput, chamar SetFocus, e
        desanexar em seguida.
        """
        user32 = ctypes.windll.user32
        hwnd_fg = win32gui.GetForegroundWindow()
        tid_alvo, _ = win32process.GetWindowThreadProcessId(hwnd_fg) if hwnd_fg else (0, 0)
        tid_atual = ctypes.windll.kernel32.GetCurrentThreadId()

        anexado = False
        try:
            if tid_alvo and tid_alvo != tid_atual:
                anexado = bool(user32.AttachThreadInput(tid_atual, tid_alvo, True))
            win32gui.SetFocus(hwnd)
            return True
        except Exception as e:
            self.log(f"⚠️ SetFocus cross-process falhou: {e}")
            return False
        finally:
            if anexado:
                try:
                    user32.AttachThreadInput(tid_atual, tid_alvo, False)
                except Exception:
                    pass

    def _is_alive(self) -> bool:
        if not self.app or not self.main_window:
            return False
        try:
            return win32gui.IsWindow(self.main_window.handle)
        except Exception:
            return False

    def _find_window(self, title_contains: str) -> int:
        """Retorna hwnd de janela visível cujo título contém o texto."""
        result = [0]
        def cb(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                try:
                    if title_contains.lower() in win32gui.GetWindowText(hwnd).lower():
                        result[0] = hwnd
                        return False
                except Exception:
                    pass
            return True
        win32gui.EnumWindows(cb, None)
        return result[0]

    def _get_combobox_text_by_ctrl_id(self, parent_hwnd: int, ctrl_id: int) -> str:
        """
        Lê o texto atualmente exibido em um ComboBox filho identificado por
        ctrl_id, usando WM_GETTEXT via SendMessageW (ctypes). WM_GETTEXT é
        uma das mensagens com marshalling cross-process garantido pelo
        Windows (ao contrário de CB_GETLBTEXT, cujo buffer de saída não é
        thunked entre processos e retornava lixo/vazio na tentativa
        anterior via win32gui.GetWindowText/SendMessage).
        """
        WM_GETTEXT     = 0x000D
        WM_GETTEXTLENGTH = 0x000E
        user32 = ctypes.windll.user32

        h = win32gui.GetWindow(parent_hwnd, 5)  # GW_CHILD
        while h:
            try:
                if win32gui.GetDlgCtrlID(h) == ctrl_id and win32gui.GetClassName(h) == "ComboBox":
                    length = user32.SendMessageW(h, WM_GETTEXTLENGTH, 0, 0)
                    buf = ctypes.create_unicode_buffer(length + 1)
                    user32.SendMessageW(h, WM_GETTEXT, length + 1, ctypes.byref(buf))
                    return buf.value or ""
            except Exception:
                pass
            h = win32gui.GetWindow(h, 2)  # GW_HWNDNEXT
        return ""

    # ── Conexão ──────────────────────────────────────────────────────────────

    def connect_to_dominio(self) -> bool:
        try:
            self.log("🔎 Localizando Domínio Folha...")

            # Busca hwnd pelo título via win32gui (sem pywinauto ainda)
            hwnd = None
            def _cb(h, _):
                nonlocal hwnd
                if hwnd:
                    return
                if not win32gui.IsWindowVisible(h):
                    return
                t = win32gui.GetWindowText(h)
                if "omínio" in t or "OMÍNIO" in t:
                    hwnd = h
            win32gui.EnumWindows(_cb, None)

            if not hwnd:
                self.log("❌ Domínio Folha não encontrado. Abra o sistema antes de iniciar.")
                return False

            titulo = win32gui.GetWindowText(hwnd)
            self.log(f"✅ Janela encontrada: '{titulo}' (hwnd={hwnd})")

            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(1)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5)

            # Conecta usando win32 sem mapear a árvore inteira (timeout curto)
            timings.Timings.window_find_timeout = 5
            self.app = Application(backend="win32").connect(
                handle=hwnd, timeout=10
            )
            self.main_window = self.app.window(handle=hwnd)
            self.log("✅ Conectado ao Domínio Folha")
            return True
        except Exception as e:
            self.log(f"❌ Erro ao conectar: {type(e).__name__}: {e}")
            return False

    # ── Troca de empresa ─────────────────────────────────────────────────────

    def trocar_empresa(self, empresa_num: str) -> bool:
        try:
            if empresa_num == self.empresa_atual:
                self.log(f"🏢 Empresa {empresa_num} já ativa")
                return True

            self.log(f"🔄 Trocando para empresa {empresa_num} (F8)...")

            # Se a janela "Troca de empresas" já estiver aberta (de uma empresa anterior),
            # fecha antes de reabrir — evita que os códigos se concatenem no campo de busca.
            try:
                troca_existente = self.main_window.child_window(title="Troca de empresas", class_name="FNWND3190")
                if troca_existente.exists() and troca_existente.is_visible():
                    self.log("⚠️ 'Troca de empresas' já aberta — fechando antes de reabrir")
                    troca_existente.set_focus()
                    send_keys('{ESC}')
                    time.sleep(0.5)
            except Exception:
                pass

            self.main_window.set_focus()
            if not self.smart_sleep(0.3):
                return False
            send_keys('{F8}')
            if not self.smart_sleep(2):
                return False

            # Aguarda janela "Troca de empresas"
            troca = None
            for _ in range(10):
                if self.should_stop():
                    return False
                try:
                    troca = self.main_window.child_window(title="Troca de empresas", class_name="FNWND3190")
                    if troca.exists():
                        break
                except Exception:
                    pass
                if not self.smart_sleep(0.8):
                    return False

            if not troca or not troca.exists():
                # Pode haver popup bloqueando o F8 — fecha e tenta uma vez mais
                self.log("⚠️ Janela 'Troca de empresas' não apareceu — verificando popups...")
                for _ in range(6):
                    self._tratar_erros_dominio()
                    self._checar_sem_dados()
                    time.sleep(0.3)
                # Segunda tentativa com F8
                self.log("🔄 Segunda tentativa F8 após fechar popups...")
                try:
                    self.main_window.set_focus()
                except Exception:
                    pass
                time.sleep(0.3)
                send_keys('{F8}')
                time.sleep(2)
                troca = None
                for _ in range(10):
                    if self.should_stop():
                        return False
                    try:
                        troca = self.main_window.child_window(title="Troca de empresas", class_name="FNWND3190")
                        if troca.exists():
                            break
                    except Exception:
                        pass
                    time.sleep(0.8)
                if not troca or not troca.exists():
                    self.log("❌ Janela 'Troca de empresas' não apareceu após retry — enviando ESCs para limpar")
                    self._enviar_escs(8)
                    return False

            troca_hwnd = troca.wrapper_object().handle
            self._force_focus(troca_hwnd)
            if not self.smart_sleep(0.3):
                return False

            edit_hwnd = win32gui.FindWindowEx(troca_hwnd, 0, "Edit", None)
            if edit_hwnd:
                edit = self.app.window(handle=edit_hwnd)
                edit.set_focus()
                time.sleep(0.2)
                edit.type_keys('^a', with_spaces=False)
                edit.type_keys(empresa_num, with_spaces=False)
                time.sleep(0.3)
                edit.type_keys('{ENTER}', with_spaces=False)
            else:
                troca.wrapper_object().set_focus()
                send_keys('^a')
                send_keys(empresa_num)
                if not self.smart_sleep(0.3):
                    return False
                send_keys('{ENTER}')

            if not self.smart_sleep(2):
                return False

            # Aguarda a janela "Troca de empresas" fechar de fato — sinal de
            # que o Domínio concluiu o carregamento da empresa. Só depois
            # disso é seguro checar/fechar o aviso de vencimento e navegar
            # pelo menu; fazer isso cedo demais faz o menu ser processado
            # ainda com a troca em andamento, e o aviso "rouba" o Enter.
            self.log("⏳ Aguardando o Domínio concluir a troca de empresa...")
            troca_fechou = False
            for _ in range(20):
                if self.should_stop():
                    return False
                try:
                    troca_check = self.main_window.child_window(title="Troca de empresas", class_name="FNWND3190")
                    if not troca_check.exists():
                        troca_fechou = True
                        break
                except Exception:
                    troca_fechou = True
                    break
                if not self.smart_sleep(0.5):
                    return False

            if not troca_fechou:
                self.log("⚠️ Janela 'Troca de empresas' ainda visível após espera — prosseguindo com cautela")

            # Tempo extra para o aviso de vencimento (se houver) terminar de
            # renderizar antes de tentarmos fechá-lo/navegar pelo menu.
            if not self.smart_sleep(2):
                return False
            self._garantir_aviso_vencimento_fechado()
            self.empresa_atual = empresa_num
            self.log(f"✅ Empresa {empresa_num} ativa")
            return True

        except Exception as e:
            self.log(f"❌ Erro na troca de empresa: {e} — enviando ESCs para limpar")
            self._enviar_escs(8)
            return False

    def _fechar_avisos_vencimento(self):
        """Fecha popup de Avisos de Vencimento — busca como filho e como top-level."""
        fechou = False

        # 1. Tenta como filho da janela principal (título real observado:
        # "Relatório de Avisos de Vencimentos" — por isso regex, não exato)
        try:
            av = self.main_window.child_window(title_re=".*[Vv]enciment.*", class_name="FNWND3190")
            if av.exists() and av.is_visible():
                self.log("🔔 Fechando Avisos de Vencimento (filho)")
                av.set_focus()
                send_keys('{ESC}')
                time.sleep(0.5)
                fechou = True
        except Exception:
            pass

        if fechou:
            return

        # 2. Tenta como janela top-level via win32gui
        def _cb(hwnd, _):
            nonlocal fechou
            if fechou:
                return
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                titulo = win32gui.GetWindowText(hwnd)
                classe = win32gui.GetClassName(hwnd)
                if "Aviso" in titulo and "FNWND" in classe:
                    self.log(f"🔔 Fechando Avisos de Vencimento (top-level: '{titulo}')")
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.3)
                    send_keys('{ESC}')
                    time.sleep(0.5)
                    fechou = True
            except Exception:
                pass

        win32gui.EnumWindows(_cb, None)

        # 3. Fallback: qualquer FNWND visível que não seja a janela principal
        if not fechou:
            def _cb2(hwnd, _):
                nonlocal fechou
                if fechou:
                    return
                try:
                    if hwnd == self.main_window.handle:
                        return
                    if not win32gui.IsWindowVisible(hwnd):
                        return
                    classe = win32gui.GetClassName(hwnd)
                    titulo = win32gui.GetWindowText(hwnd)
                    if "FNWND" in classe and titulo:
                        self.log(f"🔔 Fechando popup: '{titulo}' [{classe}]")
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.3)
                        send_keys('{ESC}')
                        time.sleep(0.5)
                        fechou = True
                except Exception:
                    pass
            win32gui.EnumWindows(_cb2, None)

        if not fechou:
            self.log("ℹ️ Nenhum popup de aviso encontrado")

    def _fechar_popup_erro_gravacao(self) -> bool:
        """
        Fecha o popup de título 'Erro' que pode aparecer durante o processamento.
        Janela: class=#32770, título='Erro'. Botão OK tem ctrl_id=2.
        Retorna True se fechou o popup.
        """
        BM_CLICK = 0x00F5
        fechou = [False]

        def _cb(hwnd, _):
            if fechou[0]:
                return False
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != "#32770":
                return True
            titulo = win32gui.GetWindowText(hwnd)
            titulo_lower = titulo.strip().lower()
            if titulo_lower != "erro" and not ("erro" in titulo_lower and len(titulo_lower) < 20):
                return True
            b = win32gui.GetWindow(hwnd, 5)  # GW_CHILD
            while b:
                try:
                    if win32gui.GetDlgCtrlID(b) == 2 and win32gui.GetClassName(b) == "Button":
                        self.log(f"⚠️ Popup 'Erro' detectado (titulo={repr(titulo)}) — fechando")
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.2)
                        win32gui.SendMessage(b, BM_CLICK, 0, 0)
                        fechou[0] = True
                        return False
                except Exception:
                    pass
                b = win32gui.GetWindow(b, 2)  # GW_HWNDNEXT
            return True

        win32gui.EnumWindows(_cb, None)
        return fechou[0]

    # ── Navegação: Relatórios → Informativos → (seta) → (seta) → Enter ──────

    def _aviso_vencimento_aberto(self) -> bool:
        """Verifica (sem fechar) se a janela/aba 'Avisos de Vencimento(s)' está aberta.
        Usa busca por substring pois o título observado no sistema real é
        'Relatório de Avisos de Vencimentos' (plural, com prefixo), não o
        texto exato 'Avisos de Vencimento'.

        O aviso abre EMBUTIDO como aba/filho da janela principal (mesmo
        padrão do relatório GFIP e do antigo 'Extrator da DIRF'), não como
        janela top-level — por isso EnumWindows sozinho não o enxerga.
        Verificamos: 1) o título da própria main_window, 2) filhos diretos
        FNWND3190 da main_window, 3) fallback top-level via EnumWindows.
        """
        try:
            titulo_main = self.main_window.window_text() or ""
            if "venciment" in titulo_main.lower():
                return True
        except Exception:
            pass

        try:
            av = self.main_window.child_window(title_re=".*[Vv]enciment.*", class_name="FNWND3190")
            if av.exists() and av.is_visible():
                return True
        except Exception:
            pass

        encontrado = [False]
        def _cb(hwnd, _):
            if encontrado[0]:
                return
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                if "venciment" in win32gui.GetWindowText(hwnd).lower():
                    encontrado[0] = True
            except Exception:
                pass
        try:
            win32gui.EnumWindows(_cb, None)
        except Exception:
            pass
        return encontrado[0]

    def _garantir_aviso_vencimento_fechado(self, tentativas: int = 6) -> bool:
        """Fecha repetidamente o aviso de vencimento (aparece com atraso após
        a troca de empresa) até confirmar que não está mais aberto. Deve ser
        chamado logo antes de qualquer navegação de menu, pois se essa janela
        ainda estiver aberta um ENTER subsequente a maximiza em vez de
        confirmar o menu."""
        for _ in range(tentativas):
            if not self._aviso_vencimento_aberto():
                return True
            self.log("🔔 'Avisos de Vencimento' ainda aberto — fechando antes de navegar")
            self._fechar_avisos_vencimento()
            time.sleep(0.4)
        return not self._aviso_vencimento_aberto()

    def abrir_extrator_gfip(self) -> bool:
        try:
            self.log("📂 Abrindo relatório GFIP...")
            if not self._garantir_aviso_vencimento_fechado():
                self.log("⚠️ Não foi possível confirmar o fechamento do 'Avisos de Vencimento' — prosseguindo com cautela")
            self.main_window.set_focus()
            if not self.smart_sleep(0.4):
                return False

            send_keys('%r')           # ALT+R — menu Relatórios
            if not self.smart_sleep(0.5):
                return False
            send_keys('i')            # I — Informativos
            if not self.smart_sleep(0.4):
                return False
            send_keys('{RIGHT}')      # → abre submenu Informativos
            if not self.smart_sleep(0.4):
                return False
            send_keys('{RIGHT}')      # → entra no item/submenu seguinte
            if not self.smart_sleep(0.4):
                return False

            # Última checagem: se o aviso "roubou" o foco durante a navegação
            # do menu, o ENTER a seguir iria apenas maximizá-lo em vez de
            # confirmar a seleção do GFIP.
            if self._aviso_vencimento_aberto():
                self.log("⚠️ 'Avisos de Vencimento' reapareceu durante a navegação — fechando e reiniciando menu")
                self._fechar_avisos_vencimento()
                time.sleep(0.4)
                self.main_window.set_focus()
                if not self.smart_sleep(0.4):
                    return False
                send_keys('%r')
                if not self.smart_sleep(0.5):
                    return False
                send_keys('i')
                if not self.smart_sleep(0.4):
                    return False
                send_keys('{RIGHT}')
                if not self.smart_sleep(0.4):
                    return False
                send_keys('{RIGHT}')
                if not self.smart_sleep(0.4):
                    return False

            send_keys('{ENTER}')      # ENTER — abre a janela do GFIP
            if not self.smart_sleep(1.5):
                return False

            # Aguarda janela do GFIP (filho FNWND3190 ou top-level #32770)
            for _ in range(40):
                if self.should_stop():
                    return False

                # Detecta bloqueio por Honorários — empresa deve ser ignorada
                if self._checar_bloqueio_honorarios():
                    return None  # sentinel: skip empresa sem contar como erro

                try:
                    extrator = self.main_window.child_window(title_re=".*GFIP.*", class_name="FNWND3190")
                    if extrator.exists():
                        self.log("✅ Janela do GFIP aberta")
                        return True
                except Exception:
                    pass

                hwnd_ext = self._find_window("GFIP")
                if hwnd_ext:
                    self.log("✅ Janela do GFIP encontrada (top-level)")
                    return True

                if not self.smart_sleep(0.7):
                    return False

            if self._checar_bloqueio_honorarios():
                return None

            self.log("❌ Janela do GFIP não apareceu — enviando ESCs para limpar")
            self._enviar_escs(8)
            return False

        except Exception as e:
            self.log(f"❌ Erro ao abrir GFIP: {e} — enviando ESCs para limpar")
            self._enviar_escs(8)
            return False

    # ── Preenche a janela do GFIP e clica OK ─────────────────────────────────

    def _achar_filho_por_ctrl_id(self, parent_hwnd: int, ctrl_id: int, classe: str = None) -> int:
        """Retorna hwnd do filho direto com o ctrl_id (e classe, se informada) dados."""
        h = win32gui.GetWindow(parent_hwnd, 5)  # GW_CHILD
        while h:
            try:
                if win32gui.GetDlgCtrlID(h) == ctrl_id:
                    if classe is None or win32gui.GetClassName(h) == classe:
                        return h
            except Exception:
                pass
            h = win32gui.GetWindow(h, 2)  # GW_HWNDNEXT
        return 0

    def _marcar_fgts_em_atraso(self, extrator_hwnd: int, data_atraso: str) -> bool:
        """
        Marca o combo "Indicador recolhimento do FGTS" (ctrl_id=1034) como
        "Em Atraso" e preenche o campo "Data" (ctrl_id=1027, PBEDIT190) com
        `data_atraso` (DD/MM/AAAA). Mapeado via gfip_tab_spy.py.

        O combo 1034 é um ComboBox nativo do Win32 (não VCL) — usamos
        CB_GETCOUNT/CB_GETLBTEXT via SendMessageW para achar o índice do
        item "Em Atraso" e CB_SETCURSEL para selecioná-lo, técnica mais
        confiável que ciclar por tecla (usado no combo 1024, que é custom).
        """
        CB_GETCOUNT   = 0x0146
        CB_GETLBTEXT  = 0x0148
        CB_GETLBTEXTLEN = 0x0149
        CB_SETCURSEL  = 0x014E
        CB_GETCURSEL  = 0x0147
        user32 = ctypes.windll.user32

        combo_hwnd = self._achar_filho_por_ctrl_id(extrator_hwnd, 1034, "ComboBox")
        if not combo_hwnd:
            self.log("❌ Combo 'Indicador recolhimento do FGTS' (ctrl_id=1034) não encontrado")
            return False

        count = user32.SendMessageW(combo_hwnd, CB_GETCOUNT, 0, 0)
        indice_atraso = -1
        for i in range(count):
            length = user32.SendMessageW(combo_hwnd, CB_GETLBTEXTLEN, i, 0)
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.SendMessageW(combo_hwnd, CB_GETLBTEXT, i, ctypes.byref(buf))
            if "atraso" in (buf.value or "").strip().lower():
                indice_atraso = i
                break

        if indice_atraso < 0:
            self.log(f"❌ Item 'Em Atraso' não encontrado no combo 1034 ({count} itens)")
            return False

        user32.SendMessageW(combo_hwnd, CB_SETCURSEL, indice_atraso, 0)
        # ComboBox nativo não notifica o app sozinho via SendMessage direto;
        # o Domínio só reage à mudança (habilitando o campo Data) ao receber
        # WM_COMMAND/CBN_SELCHANGE do combo para a janela pai.
        CBN_SELCHANGE = 1
        parent_id = win32gui.GetDlgCtrlID(combo_hwnd)
        wparam = (CBN_SELCHANGE << 16) | parent_id
        win32gui.SendMessage(extrator_hwnd, win32con.WM_COMMAND, wparam, combo_hwnd)
        time.sleep(0.3)
        self.log("📝 Indicador recolhimento do FGTS: Em Atraso")

        data_hwnd = self._achar_filho_por_ctrl_id(extrator_hwnd, 1027, "PBEDIT190")
        if not data_hwnd:
            self.log("❌ Campo Data (ctrl_id=1027) não encontrado após marcar Em Atraso")
            return False

        self._set_focus_cross_process(data_hwnd)
        time.sleep(0.2)
        send_keys('^a')
        time.sleep(0.1)
        send_keys(data_atraso, with_spaces=False)
        time.sleep(0.2)
        self.log(f"📝 Data do FGTS em atraso: {data_atraso}")
        return True

    def preencher_gfip(self, comp_inicial: str, comp_final: str,
                       cod_recolhimento: str, caminho_completo: str,
                       data_atraso: str = None) -> bool:
        """
        Preenche competência, código de recolhimento e caminho de destino,
        então confirma (OK). A janela abre com foco no campo "Competência"
        (único — não há "de"/"até" nesta tela; comp_final é aceito por
        compatibilidade com o restante do fluxo mas não é usado aqui).
        Ordem de TAB confirmada na tela real:
        Competência → Código de Recolhimento → (5x TAB) → Caminho → OK.

        `data_atraso`: se informado (formato DD/MM/AAAA), marca o combo
        "Indicador recolhimento do FGTS" (ctrl_id=1034) como "Em Atraso" e
        preenche o campo "Data" correspondente (ctrl_id=1027) com esse valor.
        Usado no fluxo de recálculo de FGTS, que sempre envolve recolhimento
        em atraso. Se None, o indicador permanece no padrão "No Prazo" da
        tela (comportamento normal do GFIP). O indicador de Previdência
        Social (2º par de campos) não é alterado — fora do escopo atual.

        O Domínio grava o TXT diretamente no caminho informado, com nome de
        arquivo fixo definido pelo próprio sistema — por isso o caminho deve
        ser uma pasta exclusiva da empresa (ver salvar_gfip / processar_empresa).
        """
        try:
            extrator_hwnd = 0
            try:
                ext = self.main_window.child_window(title_re=".*GFIP.*", class_name="FNWND3190")
                if ext.exists():
                    extrator_hwnd = ext.wrapper_object().handle
            except Exception:
                pass
            if not extrator_hwnd:
                extrator_hwnd = self._find_window("GFIP")
            if not extrator_hwnd:
                self.log("❌ Janela do GFIP não encontrada — enviando ESCs para limpar")
                self._enviar_escs(8)
                return False

            self._force_focus(extrator_hwnd)
            time.sleep(0.5)
            self.log(f"📝 Preenchendo competência: {comp_inicial}")

            # Campo "Competência" (único, não há "de"/"até" nesta tela) já
            # vem focado ao abrir a janela.
            send_keys('^a')
            time.sleep(0.1)
            send_keys(comp_inicial, with_spaces=False)
            time.sleep(0.3)

            # TAB — Código de Recolhimento (combo, ex: "115 - Fgts + Inss").
            # O combo já abre com "115" selecionado por padrão. Este combo
            # NÃO é navegado por F4/setas: com foco nele, apertar a tecla
            # "1" repetidamente cicla entre os itens que começam com "1"
            # (115 -> 150 -> 155 -> 115 -> ...), sem precisar abrir dropdown
            # nem confirmar com Enter — o valor exibido já fica setado.
            send_keys('{TAB}')
            time.sleep(0.2)
            cod_recolhimento_norm = (cod_recolhimento or "").strip()
            if not cod_recolhimento_norm or cod_recolhimento_norm.startswith("115"):
                self.log("📝 Código de Recolhimento: 115 (padrão da tela, não alterado)")
            else:
                self.log(f"📝 Código de Recolhimento: selecionando {cod_recolhimento_norm} via ciclo de tecla")
                achou = False
                item_atual = self._get_combobox_text_by_ctrl_id(extrator_hwnd, 1024)
                self.log(f"   valor inicial do combo: '{item_atual.strip()}'")
                for i in range(6):  # só 3 itens no ciclo (115/150/155); 6 dá folga
                    if item_atual.strip().startswith(cod_recolhimento_norm):
                        achou = True
                        break
                    send_keys('1', with_spaces=False)
                    time.sleep(0.5)
                    item_atual = self._get_combobox_text_by_ctrl_id(extrator_hwnd, 1024)
                    self.log(f"   após aperto {i+1}: '{item_atual.strip()}'")
                if achou:
                    self.log(f"📝 Código de Recolhimento: {item_atual.strip()}")
                else:
                    self.log(f"❌ Item '{cod_recolhimento_norm}' não encontrado ciclando o combo de Recolhimento após 6 tentativas — abortando linha (valor atual: '{item_atual.strip()}')")
                    return False

            # Ordem de TAB confirmada via gfip_tab_spy.py (ctrl_id de cada
            # controle focado, a partir de Competência ctrl_id=1037):
            #   1024 ComboBox (Recolhimento) -> 1047 ComboBox (Característica)
            #   -> 1050 ComboBox -> 1001 PBEDIT190 (Responsável)
            #   -> 2 Button "&Arquivo" (radio) -> 1014 Button "..."
            #   -> 1025 Edit (campo de caminho real)

            # TAB — Característica (não alterado)
            send_keys('{TAB}')
            time.sleep(0.15)
            # TAB — próximo combo (Tipo Folha/Complemento/Modalidade — não alterado)
            send_keys('{TAB}')
            time.sleep(0.15)

            # TAB — Responsável (ctrl_id=1001, preenchido com "1")
            send_keys('{TAB}')
            time.sleep(0.2)
            send_keys('^a')
            time.sleep(0.1)
            send_keys('1', with_spaces=False)
            time.sleep(0.3)
            self.log("📝 Responsável: 1")

            # Campo de caminho real (ctrl_id=1025, Edit). Em vez de confiar
            # cegamente em mais TABs (se algum passo anterior desviar, uma
            # letra do caminho digitada fora de um campo de texto é
            # interpretada como atalho de tela — ex: "S" de "Santos" abrindo
            # "Seleção..."), localizamos e focamos esse Edit diretamente
            # pelo ctrl_id via win32gui, o que é imune a desvios de TAB.
            caminho_hwnd = 0
            h = win32gui.GetWindow(extrator_hwnd, 5)  # GW_CHILD
            while h:
                try:
                    if win32gui.GetDlgCtrlID(h) == 1025 and win32gui.GetClassName(h) == "Edit":
                        caminho_hwnd = h
                        break
                except Exception:
                    pass
                h = win32gui.GetWindow(h, 2)  # GW_HWNDNEXT

            if not caminho_hwnd:
                self.log("❌ Campo de caminho (ctrl_id=1025) não encontrado — abortando linha")
                return False

            # win32gui.SetFocus só funciona quando chamado pela thread dona
            # da fila de mensagens da janela alvo — daí o "Acesso negado"
            # (ERROR_ACCESS_DENIED) ao tentar focar um controle de outro
            # processo diretamente. É preciso anexar a thread do Python à
            # thread de input do Domínio via AttachThreadInput antes de
            # chamar SetFocus, e desanexar logo em seguida.
            self._set_focus_cross_process(caminho_hwnd)
            time.sleep(0.2)

            # Confirma que o foco realmente chegou no campo certo antes de
            # digitar — evita repetir o bug de letras vazando como atalho.
            foco_ok = False
            for _ in range(10):
                foco_atual = _hwnd_com_foco_teclado()
                if foco_atual == caminho_hwnd:
                    foco_ok = True
                    break
                self._set_focus_cross_process(caminho_hwnd)
                time.sleep(0.2)

            if not foco_ok:
                self.log("❌ Não foi possível focar o campo de caminho — abortando linha para evitar digitar teclas soltas")
                return False

            # Este campo já vem preenchido com um caminho padrão do Domínio;
            # Ctrl+A não seleciona/limpa esse Edit de forma confiável, então
            # limpamos explicitamente via WM_SETTEXT (substituição direta,
            # sem depender de seleção por teclado) e então digitamos.
            win32gui.SendMessage(caminho_hwnd, win32con.WM_SETTEXT, 0, "")
            time.sleep(0.1)
            send_keys(caminho_completo, with_spaces=True)
            time.sleep(0.3)
            self.log(f"📁 Caminho de destino: {caminho_completo}")

            # Indicador recolhimento do FGTS "Em Atraso" + Data (fluxo de
            # recálculo de FGTS). Localizado por ctrl_id via
            # gfip_tab_spy.py: 1034 = ComboBox indicador, 1027 = PBEDIT190 Data.
            if data_atraso:
                if not self._marcar_fgts_em_atraso(extrator_hwnd, data_atraso):
                    self.log("❌ Não foi possível marcar 'Em Atraso' — abortando linha")
                    return False

            # Confirma (OK). Botão real inspecionado: AutomationId="1028" (ctrl_id),
            # AccessKey="Alt+o", class="Button". ENTER sozinho não aciona esse botão
            # pois o foco pode estar no campo de caminho, não no botão — por isso
            # localizamos o botão pelo ctrl_id e clicamos via PostMessage (mesma
            # técnica que funciona no DIRF), com fallback Alt+O e ENTER.
            self.log("▶ Confirmando (OK)...")
            BM_CLICK = 0x00F5
            ok_hwnd = 0
            h = win32gui.GetWindow(extrator_hwnd, 5)  # GW_CHILD
            while h:
                try:
                    if win32gui.GetDlgCtrlID(h) == 1028 and win32gui.GetClassName(h) == "Button":
                        ok_hwnd = h
                        break
                except Exception:
                    pass
                h = win32gui.GetWindow(h, 2)  # GW_HWNDNEXT

            if ok_hwnd:
                win32gui.PostMessage(ok_hwnd, BM_CLICK, 0, 0)
                self.log("✅ OK clicado (PostMessage, ctrl_id=1028)")
            else:
                self.log("⚠️ Botão OK não encontrado por ID, usando Alt+O")
                send_keys('%o')
                time.sleep(0.3)
                send_keys('{ENTER}')

            # O popup de confirmação ("GFIP gerada com sucesso.") demora
            # cerca de 10s para aparecer após o clique em OK — aguarda antes
            # de começar a checar, em vez de gastar ciclos do loop achando
            # "nenhum popup" enquanto o Domínio ainda está processando.
            self.log("⏳ Aguardando o Domínio processar (popup de confirmação demora ~10s)...")
            time.sleep(10)

            # Aguarda até 10s fechando avisos informativos que possam surgir.
            for _ in range(20):
                sem_dados = self._checar_sem_dados()
                if sem_dados:
                    self.log("⚠️ 'Sem dados para emitir' — empresa sem movimento no período")
                    return None  # sinaliza "sem dados", não é erro
                self._fechar_avisos_vencimento()
                resultado_popup = self._tratar_erros_dominio()
                if resultado_popup == "gfip_sucesso":
                    # Popup de sucesso já fechado — não há motivo para
                    # continuar o loop de polling até o timeout.
                    self.log("✅ GFIP gerado com sucesso — encerrando espera")
                    return True
                time.sleep(0.5)

            return True

        except Exception as e:
            self.log(f"❌ Erro ao preencher GFIP: {e} — enviando ESCs para limpar")
            self._enviar_escs(8)
            return False

    def _verificar_dialogo_salvamento_imediato(self) -> int:
        """
        Verifica sem esperar se algum diálogo de confirmação (ex: substituir
        arquivo existente) está aberto. Aceita qualquer #32770 visível que
        não seja o popup de erro.
        Retorna hwnd ou 0.
        """
        result = [0]
        def cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != "#32770":
                return True
            titulo = win32gui.GetWindowText(hwnd).lower().strip()
            if "erro" in titulo and len(titulo) < 20:
                return True
            if any(p in titulo for p in ("salvar", "save", "exportar", "substitu", "confirmar")):
                result[0] = hwnd
                return False
            return True
        win32gui.EnumWindows(cb, None)
        return result[0]

    def _confirmar_substituicao_arquivo(self) -> bool:
        """
        Confirma o popup 'Confirmar Salvar Como'/substituição que o sistema
        pode abrir quando o arquivo já existe na pasta de destino.
        Janela: class=#32770. Botão 'Sim' tem ctrl_id=6 (IDYES).
        Retorna True se confirmou a substituição.
        """
        BM_CLICK = 0x00F5
        confirmou = [False]

        def _cb(hwnd, _):
            if confirmou[0]:
                return False
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != "#32770":
                return True
            titulo = win32gui.GetWindowText(hwnd).strip().lower()
            corpo = []
            c = win32gui.GetWindow(hwnd, 5)  # GW_CHILD
            while c:
                try:
                    if win32gui.GetClassName(c) == "Static":
                        corpo.append((win32gui.GetWindowText(c) or "").lower())
                except Exception:
                    pass
                c = win32gui.GetWindow(c, 2)  # GW_HWNDNEXT
            corpo_txt = " ".join(corpo)
            eh_substituir = any(k in corpo_txt for k in ("substitu", "sobrescrev", "já existe", "ja existe"))
            if not eh_substituir and titulo not in ("salvar", "confirmar salvar como"):
                return True
            b = win32gui.GetWindow(hwnd, 5)  # GW_CHILD
            while b:
                try:
                    if win32gui.GetDlgCtrlID(b) == 6 and win32gui.GetClassName(b) == "Button":  # IDYES = Sim
                        self.log(f"♻️ Popup 'substituir arquivo' detectado (titulo={repr(titulo)}) — confirmando (Sim)")
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.2)
                        win32gui.SendMessage(b, BM_CLICK, 0, 0)
                        confirmou[0] = True
                        return False
                except Exception:
                    pass
                b = win32gui.GetWindow(b, 2)  # GW_HWNDNEXT
            return True

        for _ in range(10):
            win32gui.EnumWindows(_cb, None)
            if confirmou[0]:
                break
            time.sleep(0.3)
        if not confirmou[0]:
            try:
                send_keys('%s')
            except Exception:
                pass
        return confirmou[0]

    # ── Salva o arquivo TXT do GFIP na subpasta da empresa ──────────────────

    def salvar_gfip(self, empresa_num: str, empresa_nome: str, dir_saida: str,
                    comp_inicial: str, comp_final: str, cod_recolhimento: str,
                    data_atraso: str = None) -> bool:
        """
        Abre a janela do GFIP, preenche os campos e confirma. O Domínio grava
        o arquivo diretamente na pasta informada, com nome de arquivo fixo
        (não editável pelo usuário) — por isso cada empresa usa sua própria
        subpasta dentro de dir_saida, no formato:
            {dir_saida}/{empresa_num}/

        `data_atraso`: ver preencher_gfip — marca "Em Atraso" no indicador do
        FGTS com essa data, para o fluxo de recálculo de FGTS.
        """
        try:
            pasta_empresa = os.path.join(dir_saida, empresa_num)
            os.makedirs(pasta_empresa, exist_ok=True)

            resultado = self.preencher_gfip(comp_inicial, comp_final, cod_recolhimento,
                                            pasta_empresa, data_atraso=data_atraso)
            if resultado is None:
                return None  # sem dados
            if not resultado:
                return False

            # Pode surgir confirmação de substituição se já existir arquivo na pasta
            self._confirmar_substituicao_arquivo()

            # preencher_gfip já confirmou "GFIP gerada com sucesso" — o
            # arquivo já foi escrito em disco nesse ponto. A janela do GFIP
            # e a de "Seleção de Empregados" (aberta pelo botão "Seleção...")
            # continuam abertas na tela e NÃO fecham sozinhas, então em vez
            # de esperar passivamente (o que só desperdiça até 30s), fecha
            # ambas ativamente agora para liberar a próxima empresa.
            self.log("🔻 Fechando janela de Seleção e GFIP após gravação bem-sucedida...")
            fechou = False
            for _ in range(10):
                if self.should_stop():
                    return False

                fechou_algo = False

                def _cb_selecao(hwnd, _):
                    nonlocal fechou_algo
                    try:
                        if not win32gui.IsWindowVisible(hwnd):
                            return
                        t = win32gui.GetWindowText(hwnd).lower()
                        if "seleção" in t or "selecao" in t:
                            win32gui.SetForegroundWindow(hwnd)
                            time.sleep(0.15)
                            send_keys('{ESC}')
                            time.sleep(0.3)
                            fechou_algo = True
                    except Exception:
                        pass
                try:
                    win32gui.EnumWindows(_cb_selecao, None)
                except Exception:
                    pass

                self._fechar_avisos_vencimento()
                self._tratar_erros_dominio()
                self._checar_sem_dados()

                try:
                    ext = self.main_window.child_window(title_re=".*GFIP.*", class_name="FNWND3190")
                    if ext.exists() and ext.is_visible():
                        ext.set_focus()
                        time.sleep(0.1)
                        send_keys('{ESC}')
                        time.sleep(0.3)
                        fechou_algo = True
                    elif not ext.exists():
                        fechou = True
                        break
                except Exception:
                    fechou = True
                    break

                if not fechou_algo:
                    time.sleep(0.3)

            if not fechou:
                self.log("⚠️ Janela do GFIP ainda aberta após tentativas de fechamento — enviando ESCs finais")
                self._enviar_escs(6)

            if any(os.scandir(pasta_empresa)):
                self.log(f"✅ Arquivo GFIP salvo em: {pasta_empresa}")
                return True
            else:
                self.log(f"❌ Nenhum arquivo encontrado em {pasta_empresa} após gravação")
                return False

        except Exception as e:
            self.log(f"❌ Erro ao salvar GFIP: {e} — enviando ESCs para limpar")
            self._enviar_escs(8)
            return False

    def _fechar_todas_abas(self):
        """
        Fecha todas as abas/janelas do Domínio abertas além da janela principal.
        """
        hwnd_main = self.main_window.handle

        # 0. Fecha popups de aviso/erro antes de tudo
        for _ in range(3):
            self._tratar_erros_dominio()
            self._checar_sem_dados()
            time.sleep(0.15)

        # 1. Fecha a janela "Seleção de Empregados" se ainda estiver aberta —
        # ela pode ficar por trás da janela do GFIP (título contém "Seleção"),
        # então é fechada explicitamente por título em vez de depender só dos
        # ESCs genéricos da janela principal, que nem sempre alcançam quem
        # está atrás.
        for _t in range(5):
            fechou_selecao = False
            def _cb_selecao(hwnd, _):
                nonlocal fechou_selecao
                if fechou_selecao:
                    return
                try:
                    if not win32gui.IsWindowVisible(hwnd):
                        return
                    if "seleção" in win32gui.GetWindowText(hwnd).lower() or "selecao" in win32gui.GetWindowText(hwnd).lower():
                        self.log(f"🔻 Fechando janela 'Seleção' ainda aberta: '{win32gui.GetWindowText(hwnd)}'")
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.2)
                        send_keys('{ESC}')
                        time.sleep(0.4)
                        fechou_selecao = True
                except Exception:
                    pass
            try:
                win32gui.EnumWindows(_cb_selecao, None)
            except Exception:
                pass
            if not fechou_selecao:
                break

        # 2. Fecha a janela do GFIP explicitamente via pywinauto (mais confiável)
        for _t in range(5):
            try:
                ext = self.main_window.child_window(title_re=".*GFIP.*", class_name="FNWND3190")
                if ext.exists() and ext.is_visible():
                    self.log("🔻 Fechando janela do GFIP via pywinauto")
                    ext.set_focus()
                    time.sleep(0.1)
                    send_keys('{ESC}')
                    time.sleep(0.4)
                    continue
            except Exception:
                pass
            break

        # 2. ESC incondicional na janela principal para fechar relatório embutido e outras abas
        try:
            win32gui.SetForegroundWindow(hwnd_main)
            time.sleep(0.2)
        except Exception:
            pass
        for _ in range(10):
            send_keys('{ESC}')
            time.sleep(0.12)
        time.sleep(0.3)

        # 3. Loop: fecha abas embutidas restantes (título com "[") e janelas FNWND flutuantes
        for _i in range(15):
            titulo = ""
            try:
                titulo = self.main_window.window_text() or ""
            except Exception:
                break

            if "[" in titulo:
                self.log(f"🔻 Fechando aba embutida: {titulo[:60]}")
                try:
                    win32gui.SetForegroundWindow(hwnd_main)
                    time.sleep(0.15)
                    for _ in range(6):
                        send_keys('{ESC}')
                        time.sleep(0.12)
                except Exception:
                    pass
                self._tratar_erros_dominio()
                time.sleep(0.4)
                continue

            # Fecha janelas FNWND flutuantes (Troca de empresas, etc.)
            janela_flutuante = [None]
            def _cb(hwnd, _):
                if hwnd == hwnd_main:
                    return True
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                if "FNWND" in win32gui.GetClassName(hwnd):
                    janela_flutuante[0] = (hwnd, win32gui.GetWindowText(hwnd) or "")
                    return False
                return True
            try:
                win32gui.EnumWindows(_cb, None)
            except Exception:
                pass

            if janela_flutuante[0]:
                hwnd_f, titulo_f = janela_flutuante[0]
                self.log(f"🔻 Fechando janela flutuante: '{titulo_f}'")
                try:
                    win32gui.SetForegroundWindow(hwnd_f)
                    time.sleep(0.15)
                    send_keys('{ESC}')
                except Exception:
                    pass
                time.sleep(0.4)
                continue

            break  # nada mais aberto

        # Fecha qualquer popup de aviso/erro que esteja bloqueando a UI
        for _ in range(5):
            self._tratar_erros_dominio()
            self._checar_sem_dados()
            time.sleep(0.2)

        # Garante foco de volta na janela principal
        try:
            win32gui.SetForegroundWindow(hwnd_main)
        except Exception:
            pass

    def _checar_bloqueio_honorarios(self) -> bool:
        """
        Detecta e fecha o popup 'O acesso a essa empresa foi bloqueado pelo módulo Honorários.'
        Retorna True se detectado (empresa deve ser ignorada, não é erro).
        """
        BM_CLICK = 0x00F5
        encontrado = [False]
        PALAVRAS_BLOQUEIO = ["honorários", "honorarios", "bloqueado pelo módulo", "bloqueado pelo modulo"]

        def _ler_textos(hwnd) -> str:
            partes = []
            try: partes.append(win32gui.GetWindowText(hwnd))
            except: pass
            def _child(ch, _):
                try:
                    t = win32gui.GetWindowText(ch)
                    if t: partes.append(t)
                except: pass
                return True
            try: win32gui.EnumChildWindows(hwnd, _child, None)
            except: pass
            return " ".join(partes).lower()

        def _fechar_dialog(hwnd):
            try:
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.2)
            except: pass
            try:
                ok_hwnd = win32gui.FindWindowEx(hwnd, 0, "Button", "OK")
                if not ok_hwnd: ok_hwnd = win32gui.GetDlgItem(hwnd, 1)
                if ok_hwnd:
                    win32gui.SendMessage(ok_hwnd, BM_CLICK, 0, 0)
                    return
            except: pass
            try:
                dlg = self.app.window(handle=hwnd)
                ok_btn = dlg.child_window(title="OK", class_name="Button")
                if ok_btn.exists():
                    ok_btn.click_input()
                    return
            except: pass
            try:
                send_keys('{ENTER}')
                return
            except: pass
            try: send_keys('{ESC}')
            except: pass

        def _verificar(hwnd):
            if not hwnd: return False
            texto = _ler_textos(hwnd)
            if any(kw in texto for kw in PALAVRAS_BLOQUEIO):
                self.log(f"⚠️ Empresa bloqueada pelo módulo Honorários — será ignorada")
                _fechar_dialog(hwnd)
                return True
            return False

        # 1. Busca direta por título "Aviso"
        hwnd_aviso = win32gui.FindWindow(None, "Aviso")
        if hwnd_aviso and _verificar(hwnd_aviso):
            return True

        # 2. EnumWindows fallback
        def cb(hwnd, _):
            cls = win32gui.GetClassName(hwnd)
            if cls not in ("#32770",) and "aviso" not in win32gui.GetWindowText(hwnd).lower():
                return True
            if _verificar(hwnd):
                encontrado[0] = True
                return False
            return True
        if not encontrado[0]:
            try: win32gui.EnumWindows(cb, None)
            except: pass
        return encontrado[0]

    def _enviar_escs(self, n: int = 5, intervalo: float = 0.15):
        """Envia N teclas ESC para fechar qualquer janela/aviso bloqueante."""
        try:
            win32gui.SetForegroundWindow(self.main_window.handle)
            time.sleep(0.2)
        except Exception:
            pass
        for _ in range(n):
            try:
                send_keys('{ESC}')
            except Exception:
                pass
            time.sleep(intervalo)

    def _tratar_erros_dominio(self):
        """
        Fecha diálogos de erro/aviso do Domínio.
        Estratégia: BM_CLICK no botão OK filho do diálogo. Fallback:
        SetForegroundWindow + ENTER.

        Retorna:
          "gfip_sucesso" — fechou o popup "GFIP gerada com sucesso." (o
                           chamador pode concluir a etapa imediatamente,
                           sem esperar o resto do loop de polling)
          True           — fechou algum outro popup de erro/aviso
          False          — nenhum popup encontrado
        """
        BM_CLICK    = 0x00F5
        # "gfip" incluído pois o popup de confirmação final ("GFIP gerada
        # com sucesso.") usa esse título, não "Aviso"/"Informação" — sem
        # isso o popup nunca era reconhecido e o bot ficava preso até o
        # timeout do loop de espera achando "Nenhum popup de aviso encontrado".
        error_titles = {"erro", "aviso", "atenção", "informação", "alerta", "gfip"}

        found_hwnd  = [None]
        found_title = [None]

        def cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != "#32770":
                return True
            t = win32gui.GetWindowText(hwnd).strip().lower()
            if any(e in t for e in error_titles):
                found_hwnd[0]  = hwnd
                found_title[0] = t
                return False  # para na primeira janela
            return True

        try:
            win32gui.EnumWindows(cb, None)
        except Exception:
            pass

        hwnd = found_hwnd[0]
        if hwnd is None:
            return False

        titulo = found_title[0]

        corpo_textos = []
        def _enum_child(child, _):
            try:
                corpo_textos.append(win32gui.GetWindowText(child).lower())
            except Exception:
                pass
            return True
        try:
            win32gui.EnumChildWindows(hwnd, _enum_child, None)
        except Exception:
            pass
        corpo = " ".join(corpo_textos)

        self.log(f"🔔 Popup '{titulo}' detectado — corpo: {corpo[:120]}")

        # NÃO fechar aqui se for "sem dados para emitir" — _checar_sem_dados trata isso
        if "sem dados" in corpo:
            self.log("🔔 Popup 'sem dados' — delegando para _checar_sem_dados, não fechando aqui")
            return False

        eh_sucesso_gfip = "gerada com sucesso" in corpo

        # Método 1: pywinauto click_input
        try:
            dlg = self.app.window(handle=hwnd)
            ok_btn = dlg.child_window(title="OK", class_name="Button")
            if ok_btn.exists():
                ok_btn.click_input()
                self.log(f"🔔 Popup '{titulo}' fechado via click_input()")
                return "gfip_sucesso" if eh_sucesso_gfip else True
        except Exception:
            pass

        # Método 2: PostMessage BM_CLICK no botão filho
        try:
            ok_hwnd = win32gui.FindWindowEx(hwnd, 0, "Button", "OK")
            if not ok_hwnd:
                ok_hwnd = win32gui.GetDlgItem(hwnd, 1)
            if ok_hwnd:
                win32gui.PostMessage(ok_hwnd, BM_CLICK, 0, 0)
                self.log(f"🔔 Popup '{titulo}' fechado via PostMessage BM_CLICK")
                return "gfip_sucesso" if eh_sucesso_gfip else True
        except Exception:
            pass

        # Método 3: WM_CLOSE
        try:
            win32gui.PostMessage(hwnd, 0x0010, 0, 0)
            self.log(f"🔔 Popup '{titulo}' fechado via WM_CLOSE")
            return "gfip_sucesso" if eh_sucesso_gfip else True
        except Exception:
            pass

        # Fallback: foco + ENTER
        try:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.1)
            send_keys('{ENTER}')
            self.log(f"🔔 Popup '{titulo}' fechado via ENTER")
            return "gfip_sucesso" if eh_sucesso_gfip else True
        except Exception:
            pass

        return False

    def _checar_sem_dados(self) -> bool:
        """
        Verifica se surgiu o popup 'Sem dados para emitir!'.
        Lê o título e o corpo de todos os diálogos visíveis (#32770 e outros).
        Se detectado, fecha o popup e retorna True.
        """
        BM_CLICK   = 0x00F5
        encontrado = [False]
        PALAVRAS_SEM_DADOS = ["sem dados", "sem dados para emitir", "no data"]

        def _ler_textos_dialog(hwnd) -> str:
            partes = []
            try:
                partes.append(win32gui.GetWindowText(hwnd))
            except Exception:
                pass
            def _child(ch, _):
                try:
                    t = win32gui.GetWindowText(ch)
                    if t:
                        partes.append(t)
                except Exception:
                    pass
                return True
            try:
                win32gui.EnumChildWindows(hwnd, _child, None)
            except Exception:
                pass
            return " ".join(partes).lower()

        def _fechar_dialog(hwnd):
            try:
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.25)
            except Exception:
                pass

            try:
                ok_hwnd = win32gui.FindWindowEx(hwnd, 0, "Button", "OK")
                if not ok_hwnd:
                    ok_hwnd = win32gui.GetDlgItem(hwnd, 1)
                if ok_hwnd:
                    win32gui.SendMessage(ok_hwnd, BM_CLICK, 0, 0)
                    self.log("🔔 Popup 'Sem dados' fechado via SendMessage BM_CLICK")
                    time.sleep(0.2)
                    return
            except Exception as e:
                self.log(f"⚠️ SendMessage BM_CLICK falhou: {e}")

            try:
                dlg = self.app.window(handle=hwnd)
                ok_btn = dlg.child_window(title="OK", class_name="Button")
                if ok_btn.exists():
                    ok_btn.click_input()
                    self.log("🔔 Popup 'Sem dados' fechado via click_input()")
                    time.sleep(0.2)
                    return
            except Exception as e:
                self.log(f"⚠️ click_input falhou: {e}")

            try:
                ok_hwnd = win32gui.FindWindowEx(hwnd, 0, "Button", "OK")
                if not ok_hwnd:
                    ok_hwnd = win32gui.GetDlgItem(hwnd, 1)
                if ok_hwnd:
                    win32gui.PostMessage(ok_hwnd, BM_CLICK, 0, 0)
                    self.log("🔔 Popup 'Sem dados' fechado via PostMessage BM_CLICK")
                    time.sleep(0.2)
                    return
            except Exception:
                pass

            try:
                send_keys('{ENTER}')
                self.log("🔔 Popup 'Sem dados' fechado via ENTER")
                time.sleep(0.2)
                return
            except Exception:
                pass

            try:
                send_keys('{ESC}')
                self.log("🔔 Popup 'Sem dados' fechado via ESC")
                time.sleep(0.2)
            except Exception:
                pass

        def _verificar_hwnd(hwnd):
            if not hwnd:
                return False
            texto = _ler_textos_dialog(hwnd)
            if any(kw in texto for kw in PALAVRAS_SEM_DADOS):
                cls = win32gui.GetClassName(hwnd)
                self.log(f"⚠️ Popup 'Sem dados para emitir' detectado (class={cls}, texto='{texto[:80]}') — fechando")
                _fechar_dialog(hwnd)
                return True
            return False

        hwnd_aviso = win32gui.FindWindow(None, "Aviso")
        if hwnd_aviso and _verificar_hwnd(hwnd_aviso):
            return True

        def cb(hwnd, _):
            cls = win32gui.GetClassName(hwnd)
            if cls not in ("#32770",) and "aviso" not in win32gui.GetWindowText(hwnd).lower():
                return True
            if _verificar_hwnd(hwnd):
                encontrado[0] = True
                return False
            return True

        if not encontrado[0]:
            try:
                win32gui.EnumWindows(cb, None)
            except Exception:
                pass
        return encontrado[0]

    def cleanup(self):
        """Fecha janelas abertas e volta ao estado limpo."""
        try:
            self._tratar_erros_dominio()
            self._checar_sem_dados()
        except Exception:
            pass
        try:
            self._force_focus(self.main_window.handle)
            for _ in range(10):
                send_keys('{ESC}')
                time.sleep(0.2)
        except Exception:
            pass

    # ── Ponto de entrada por empresa ─────────────────────────────────────────

    def processar_empresa(self, empresa_num: str, empresa_nome: str,
                          comp_inicial: str, comp_final: str,
                          cod_recolhimento: str, dir_saida: str,
                          data_atraso: str = None):
        """
        Retorna:
          True         — GFIP gerado e TXT salvo
          'sem_dados'  — empresa sem movimento no período (não é erro)
          'honorarios' — empresa bloqueada pelo módulo Honorários (não é erro)
          False        — falha real

        `data_atraso`: ver preencher_gfip — se informado, marca "Em Atraso"
        no indicador do FGTS (fluxo de recálculo de FGTS).
        """
        try:
            # Reconecta se necessário
            if not self._is_alive():
                if not self.connect_to_dominio():
                    return False

            hwnd = self.main_window.handle
            try:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    time.sleep(0.5)
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass
            time.sleep(0.2)

            # Limpa janelas/abas remanescentes de empresa anterior que falhou
            self._fechar_todas_abas()
            for _ in range(4):
                self._tratar_erros_dominio()
                self._checar_sem_dados()
                time.sleep(0.2)

            # 1. Troca empresa
            if not self.trocar_empresa(empresa_num):
                return False

            if self.should_stop():
                return False

            # 2. Abre janela do GFIP
            resultado_abrir = self.abrir_extrator_gfip()
            if resultado_abrir is None:
                self.log(f"⏭️ {empresa_nome} ignorada: bloqueada pelo módulo Honorários")
                self._fechar_todas_abas()
                return 'honorarios'
            if not resultado_abrir:
                self.cleanup()
                return False

            if self.should_stop():
                return False

            # 3. Preenche campos, confirma e salva o TXT direto na subpasta da empresa
            resultado = self.salvar_gfip(empresa_num, empresa_nome, dir_saida,
                                         comp_inicial, comp_final, cod_recolhimento,
                                         data_atraso=data_atraso)
            if resultado is None:
                self.log(f"⏭️ {empresa_nome} ignorada: sem dados para emitir no período")
                time.sleep(0.5)
                self._fechar_todas_abas()
                return 'sem_dados'
            if not resultado:
                self.log("🔻 Erro ao salvar — fechando janelas antes de prosseguir")
                self._fechar_todas_abas()
                return False

            self._fechar_todas_abas()
            return True

        except Exception as e:
            self.log(f"❌ Exceção em processar_empresa {empresa_num}: {e}\n{traceback.format_exc()}")
            self.cleanup()
            return False


# ── Planilha de exemplo ──────────────────────────────────────────────────────

def criar_planilha_exemplo():
    """Cria uma planilha Excel de exemplo na pasta do script."""
    caminho = os.path.join(os.path.dirname(__file__), "empresas_gfip_exemplo.xlsx")
    if not os.path.exists(caminho):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Empresas"
        ws.append(["Nº", "Empresa", "Cod_Recolhimento"])
        ws.append([1028, "SABOR E LENHA PIZZARIA LTDA", "115"])
        ws.append([1042, "EMPRESA EXEMPLO LTDA", "115"])
        wb.save(caminho)
        print(f"Planilha de exemplo criada: {caminho}")


# ── CLI mode (chamado pelo servidor FastAPI via subprocess) ─────────────────

def _emit(tipo: str, msg: str, empresa_num=None):
    """Imprime linha JSON no stdout para o servidor capturar via WebSocket."""
    payload = {"tipo": tipo, "msg": msg, "ts": datetime.now().isoformat()}
    if empresa_num is not None:
        payload["empresa"] = empresa_num
    print(json.dumps(payload, ensure_ascii=True), flush=True)


class _CLILogger:
    """Logger que emite para stdout em formato JSON."""
    def __init__(self, empresa_num):
        self.empresa_num = empresa_num

    def _log(self, level_name: str, msg: str):
        tipo = {"DEBUG": "debug", "INFO": "info", "WARNING": "aviso",
                "ERROR": "erro", "CRITICAL": "erro"}.get(level_name, "info")
        _emit(tipo, str(msg), self.empresa_num)

    def debug(self, msg, *args, **kwargs):
        self._log("DEBUG", msg)

    def info(self, msg, *args, **kwargs):
        self._log("INFO", msg)

    def warning(self, msg, *args, **kwargs):
        self._log("WARNING", msg)

    def error(self, msg, *args, **kwargs):
        self._log("ERROR", msg)

    def critical(self, msg, *args, **kwargs):
        self._log("CRITICAL", msg)


class _CLIProgressAdapter:
    """Adapta chamadas da GUI para saída JSON no CLI."""
    def adicionar_log(self, msg, level=logging.INFO, cor=None):
        tipo = "erro" if level >= logging.ERROR else ("aviso" if level >= logging.WARNING else "info")
        _emit(tipo, str(msg))

    @property
    def executando(self):
        return True

    @property
    def pausa_solicitada(self):
        return False

    def should_stop(self):
        return False


def run_cli(args):
    """Executa automação GFIP sem GUI, emitindo logs JSON para stdout."""
    empresa_num = args.empresa
    empresa_nome = args.empresa_nome or str(empresa_num)
    comp_inicial = args.comp_inicial
    comp_final = args.comp_final
    cod_recolhimento = args.cod_recolhimento or ""
    dir_saida = args.dir_saida or os.path.join(os.path.dirname(__file__), "results")
    data_atraso = args.data_atraso or None

    os.makedirs(dir_saida, exist_ok=True)

    logger = _CLILogger(empresa_num)
    gui_adapter = _CLIProgressAdapter()

    _emit("inicio", f"Iniciando GFIP para empresa {empresa_num} ({empresa_nome})", empresa_num)

    bot = GFIPAutomation(logger, gui_adapter)

    if not bot.connect_to_dominio():
        _emit("erro", "Não foi possível conectar ao Domínio Folha", empresa_num)
        sys.exit(1)

    ok = bot.processar_empresa(str(empresa_num), empresa_nome, comp_inicial, comp_final,
                               cod_recolhimento, dir_saida, data_atraso=data_atraso)

    if ok:
        _emit("sucesso", f"Empresa {empresa_num} concluída com sucesso", empresa_num)
        sys.exit(0)
    else:
        _emit("erro", f"Empresa {empresa_num} falhou", empresa_num)
        sys.exit(1)


# ── Worker mode (bot consulta servidor, executa e reporta) ──────────────────

def _post(url: str, data: dict, servidor: str):
    import urllib.request
    body = json.dumps(data, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        f"{servidor}{url}",
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            return json.loads(r.read())
    except Exception as e:
        print(f"[worker] POST {url} falhou: {e}")
        return None


def _get(url: str, servidor: str):
    import urllib.request
    try:
        with urllib.request.urlopen(f"{servidor}{url}", timeout=10) as r:
            return json.loads(r.read())
    except Exception as e:
        print(f"[worker] GET {url} falhou: {e}")
        return None


class _WorkerLogger:
    """Logger que envia linhas ao servidor via HTTP POST."""
    def __init__(self, execucao_id: int, servidor: str):
        self.execucao_id = execucao_id
        self.servidor = servidor

    def _log(self, nivel: str, msg: str):
        print(f"[{nivel}] {msg}")
        _post(f"/api/bot/{self.execucao_id}/log", {
            "nivel": nivel,
            "mensagem": str(msg),
            "timestamp": datetime.now().isoformat(),
        }, self.servidor)

    def debug(self, msg, *args, **kwargs):   self._log("info", msg)
    def info(self, msg, *args, **kwargs):    self._log("info", msg)
    def warning(self, msg, *args, **kwargs): self._log("aviso", msg)
    def error(self, msg, *args, **kwargs):   self._log("erro", msg)
    def critical(self, msg, *args, **kwargs):self._log("erro", msg)


class _WorkerProgressAdapter:
    executando = True
    pausa_solicitada = False
    def should_stop(self): return False
    def adicionar_log(self, msg, level=None, cor=None): print(msg)


def run_worker(servidor: str, intervalo: int):
    """Loop: consulta pendentes, executa um por vez, reporta resultado."""
    print(f"[worker] Iniciando. Servidor: {servidor} | Intervalo: {intervalo}s")

    r = _post("/api/bot/resetar-travados", {}, servidor)
    if r:
        print(f"[worker] Execuções travadas resetadas: {r.get('resetados', 0)}")

    bot = None
    gui_adapter = _WorkerProgressAdapter()

    while True:
        try:
            pendentes = _get("/api/bot/pendentes", servidor) or []
            if pendentes:
                job = pendentes[0]
                execucao_id  = job["execucao_id"]
                empresa_num  = job["empresa_num"]
                empresa_nome = job["empresa_nome"]
                comp_inicial = job["comp_inicial"]
                comp_final   = job["comp_final"]
                cod_recolhimento = job.get("cod_recolhimento", "")
                dir_saida    = job["dir_saida"]

                os.makedirs(dir_saida, exist_ok=True)
                print(f"[worker] Processando execucao_id={execucao_id} empresa={empresa_num}")

                _post(f"/api/bot/{execucao_id}/iniciar", {}, servidor)

                logger = _WorkerLogger(execucao_id, servidor)

                if bot is None or not bot._is_alive():
                    bot = GFIPAutomation(logger, gui_adapter)
                    if not bot.connect_to_dominio():
                        msg_erro = "Domínio Folha não encontrado. Encerrando worker."
                        print(f"[worker] {msg_erro}")
                        _post(f"/api/bot/{execucao_id}/finalizar",
                              {"sucesso": False, "mensagem": msg_erro}, servidor)
                        sys.exit(1)

                resultado = bot.processar_empresa(str(empresa_num), empresa_nome,
                                                  comp_inicial, comp_final,
                                                  cod_recolhimento, dir_saida)
                if resultado == 'sem_dados':
                    _post(f"/api/bot/{execucao_id}/finalizar",
                          {"sucesso": True, "mensagem": "Sem dados para emitir"}, servidor)
                elif resultado == 'honorarios':
                    _post(f"/api/bot/{execucao_id}/finalizar",
                          {"sucesso": True, "mensagem": "Bloqueada pelo módulo Honorários"}, servidor)
                else:
                    _post(f"/api/bot/{execucao_id}/finalizar",
                          {"sucesso": bool(resultado), "mensagem": None}, servidor)
            else:
                print(f"[worker] Nenhuma pendência. Aguardando {intervalo}s...")

        except Exception as e:
            print(f"[worker] Erro no loop: {e}")
            traceback.print_exc()

        time.sleep(intervalo)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="DomBot GFIP")
    parser.add_argument("--worker", action="store_true", help="Modo worker: consulta servidor e executa pendentes")
    parser.add_argument("--servidor", default="http://localhost:8000", help="URL base do servidor (modo worker)")
    parser.add_argument("--intervalo", type=int, default=10, help="Segundos entre consultas ao servidor (modo worker)")
    parser.add_argument("--empresa", type=int, help="Número da empresa (modo CLI)")
    parser.add_argument("--empresa-nome", dest="empresa_nome", help="Nome da empresa (modo CLI)")
    parser.add_argument("--comp-inicial", dest="comp_inicial", help="Competência inicial MM/AAAA (modo CLI)")
    parser.add_argument("--comp-final", dest="comp_final", help="Competência final MM/AAAA (modo CLI)")
    parser.add_argument("--cod-recolhimento", dest="cod_recolhimento", help="Código de Recolhimento (modo CLI)")
    parser.add_argument("--dir-saida", dest="dir_saida", help="Pasta raiz de saída dos TXT (modo CLI)")
    parser.add_argument("--data-atraso", dest="data_atraso",
                        help="Se informado (DD/MM/AAAA), marca 'Em Atraso' no indicador do FGTS "
                             "com essa data (fluxo de recálculo de FGTS). Padrão: não altera (No Prazo).")
    args = parser.parse_args()

    if args.worker:
        run_worker(args.servidor, args.intervalo)
    elif args.empresa:
        run_cli(args)
    else:
        if not _GUI_DISPONIVEL:
            print("Erro: pacotes de GUI (customtkinter, pandas, Pillow) não instalados neste ambiente.")
            sys.exit(1)
        criar_planilha_exemplo()
        try:
            app = AutomacaoGUI()
            app.executar()
        except Exception as e:
            print(f"Erro crítico: {e}")
            traceback.print_exc()


if __name__ == "__main__":
    main()
