#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ==============================================================================
# Copyright (c) 2026 Filipe Estevão
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
# ==============================================================================

"""Interface gráfica para geração dos relatórios de DRX.

Atalho recomendado:
    pythonw.exe "C:\\...\\interface_drx.py"
"""

import contextlib
import os
import queue
import re
import subprocess
import sys
import threading
import traceback
from pathlib import Path

if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

BASE = Path(__file__).resolve().parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk
except Exception as exc:
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0, "Tkinter não está disponível.\n\n" + str(exc),
            "Gerador de Relatório DRX", 0x10)
    except Exception:
        print("Tkinter não está disponível:", exc, file=sys.stderr)
    raise SystemExit(1)

try:
    import gerar_relatorio_drx as drx
    ERRO_MOTOR = None
except Exception as exc:
    drx = None
    ERRO_MOTOR = exc

FONTES_RAIO_X = ("Cu", "Co", "Cr", "Mo", "Ag")
SEM_HDF5 = "— sem HDF5 —"


def _ativar_dpi():
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        return
    except Exception:
        pass
    try:
        import ctypes
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _abrir_caminho(caminho):
    caminho = str(Path(caminho))
    if sys.platform == "win32":
        os.startfile(caminho)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", caminho])
    else:
        subprocess.Popen(["xdg-open", caminho])


class EscritorFila:
    encoding = "utf-8"
    errors = "replace"

    def __init__(self, fila):
        self.fila = fila

    def write(self, texto):
        if texto:
            self.fila.put(("log", texto))
        return len(texto)

    def flush(self):
        pass

    def isatty(self):
        return False


class AplicativoDRX:
    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de Relatório DRX")
        self.root.geometry("900x840")
        self.root.minsize(800, 760)
        self.root.protocol("WM_DELETE_WINDOW", self._fechar)
        self.root.report_callback_exception = self._report_callback_exception
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        self.info_path = BASE / "Info.txt"
        self.info = {}
        self.info_original_analista = ""
        self.info_original_comentarios = ""
        self.pasta = None
        self.xlsx = []
        self.hdf5s = []
        self.ultimo_pdf = None
        self.fonte_hdf5 = None
        self.executando = False
        self.fila = queue.Queue()

        self._configurar_estilo()
        self._criar_variaveis()
        self._construir_ui()
        self._carregar_info()
        self._aplicar_estado()
        self._centralizar()
        self.root.after(100, self._processar_fila)

    def _configurar_estilo(self):
        estilo = ttk.Style(self.root)
        try:
            estilo.theme_use("vista")
        except tk.TclError:
            pass
        estilo.configure("TLabel", font=("Segoe UI", 9))
        estilo.configure("Titulo.TLabel", font=("Segoe UI", 17, "bold"))
        estilo.configure("Subtitulo.TLabel", font=("Segoe UI", 9), foreground="#555555")
        estilo.configure("Dica.TLabel", font=("Segoe UI", 9), foreground="#666666")
        estilo.configure("Valor.TLabel", font=("Segoe UI", 9))
        estilo.configure("Ok.TLabel", font=("Segoe UI", 9), foreground="#2E6B3A")
        estilo.configure("Aviso.TLabel", font=("Segoe UI", 9), foreground="#A15C00")
        estilo.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        estilo.configure("Gerar.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))

    def _criar_variaveis(self):
        self.pasta_var = tk.StringVar()
        self.hdf5_var = tk.StringVar()
        self.equip_var = tk.StringVar(value="—")
        self.serial_var = tk.StringVar(value="—")
        self.software_var = tk.StringVar(value="—")
        self.fonte_var = tk.StringVar()
        self.fonte_status_var = tk.StringVar(
            value="Informe ou deixe em branco para inferir das planilhas.")
        self.analista_var = tk.StringVar()
        self.info_status_var = tk.StringVar()
        self.resumo_pasta_var = tk.StringVar(value="Nenhuma pasta selecionada.")
        self.destino_var = tk.StringVar(value="O PDF será salvo na pasta selecionada.")
        self.offline_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Selecione uma pasta para começar.")

    def _construir_ui(self):
        cabecalho = ttk.Frame(self.root, padding=(20, 14, 20, 4))
        cabecalho.grid(row=0, column=0, sticky="ew")
        cabecalho.columnconfigure(0, weight=1)
        ttk.Label(cabecalho, text="Gerador de Relatório DRX",
                  style="Titulo.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(cabecalho,
                  text="Refinamento de Rietveld a partir das exportações do XRDanalysis PRO",
                  style="Subtitulo.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 0))

        corpo = ttk.Frame(self.root, padding=(20, 8, 20, 4))
        corpo.grid(row=1, column=0, sticky="nsew")
        corpo.columnconfigure(0, weight=1)
        corpo.rowconfigure(3, weight=1, minsize=120)

        info = ttk.LabelFrame(corpo, text="Informações do laboratório",
                              padding=(14, 10, 14, 10))
        info.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        info.columnconfigure(1, weight=1)

        ttk.Label(info, text="Equipamento").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=1)
        ttk.Label(info, textvariable=self.equip_var,
                  style="Valor.TLabel").grid(row=0, column=1, sticky="w", pady=1)
        ttk.Label(info, text="S/N").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=1)
        ttk.Label(info, textvariable=self.serial_var,
                  style="Valor.TLabel").grid(row=1, column=1, sticky="w", pady=1)
        ttk.Label(info, text="Software").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=1)
        ttk.Label(info, textvariable=self.software_var,
                  style="Valor.TLabel").grid(row=2, column=1, sticky="w", pady=1)
        ttk.Label(info, text="Analista").grid(row=3, column=0, sticky="w",
                                              padx=(0, 10), pady=(6, 0))
        self.analista_entry = ttk.Entry(info, textvariable=self.analista_var)
        self.analista_entry.grid(row=3, column=1, sticky="ew", pady=(6, 0))
        ttk.Label(info, text="Comentários").grid(row=4, column=0, sticky="nw",
                                                 padx=(0, 10), pady=(8, 0))
        self.comentarios_text = scrolledtext.ScrolledText(
            info, height=3, wrap="word", font=("Segoe UI", 9), undo=True)
        self.comentarios_text.grid(row=4, column=1, sticky="nsew", pady=(8, 0))

        rodape_info = ttk.Frame(info)
        rodape_info.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        rodape_info.columnconfigure(0, weight=1)
        self.info_status_label = ttk.Label(
            rodape_info, textvariable=self.info_status_var,
            style="Dica.TLabel", wraplength=640)
        self.info_status_label.grid(row=0, column=0, sticky="w")
        self.btn_recarregar_info = ttk.Button(
            rodape_info, text="Recarregar Info.txt", command=self._recarregar_info)
        self.btn_recarregar_info.grid(row=0, column=1, sticky="e")

        dados = ttk.LabelFrame(corpo, text="Medição", padding=(14, 10, 14, 10))
        dados.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        dados.columnconfigure(1, weight=1)

        ttk.Label(dados, text="Pasta").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.pasta_entry = ttk.Entry(dados, textvariable=self.pasta_var, state="readonly")
        self.pasta_entry.grid(row=0, column=1, sticky="ew")
        self.btn_pasta = ttk.Button(dados, text="Selecionar pasta...",
                                    command=self._selecionar_pasta)
        self.btn_pasta.grid(row=0, column=2, padx=(8, 0))
        ttk.Label(dados, textvariable=self.resumo_pasta_var,
                  style="Dica.TLabel").grid(row=1, column=1, columnspan=2,
                                            sticky="w", pady=(3, 0))

        ttk.Label(dados, text="HDF5").grid(row=2, column=0, sticky="w",
                                           padx=(0, 10), pady=(8, 0))
        self.hdf5_combo = ttk.Combobox(dados, textvariable=self.hdf5_var, state="disabled")
        self.hdf5_combo.grid(row=2, column=1, sticky="ew", pady=(8, 0))
        self.hdf5_combo.bind("<<ComboboxSelected>>",
                             lambda _evento: self._hdf5_alterado())
        self.btn_atualizar = ttk.Button(dados, text="Atualizar",
                                        command=self._analisar_pasta)
        self.btn_atualizar.grid(row=2, column=2, padx=(8, 0), pady=(8, 0))

        ttk.Label(dados, text="Fonte de raios X").grid(
            row=3, column=0, sticky="w", padx=(0, 10), pady=(8, 0))
        self.fonte_combo = ttk.Combobox(
            dados, textvariable=self.fonte_var, values=FONTES_RAIO_X,
            state="readonly", width=6)
        self.fonte_combo.grid(row=3, column=1, sticky="w", pady=(8, 0))
        ttk.Label(dados, textvariable=self.fonte_status_var,
                  style="Dica.TLabel", wraplength=520).grid(
            row=3, column=2, sticky="w", padx=(8, 0), pady=(8, 0))

        ttk.Label(dados, textvariable=self.destino_var, style="Dica.TLabel",
                  wraplength=760).grid(row=4, column=0, columnspan=3,
                                       sticky="w", pady=(8, 0))
        self.chk_offline = ttk.Checkbutton(
            dados, text="Não consultar o COD (offline)", variable=self.offline_var)
        self.chk_offline.grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Label(dados, text="Sem internet? Marque para usar apenas o cache local.",
                  style="Dica.TLabel").grid(row=6, column=0, columnspan=3, sticky="w")

        acoes = ttk.Frame(corpo)
        acoes.grid(row=2, column=0, sticky="ew")
        acoes.columnconfigure(2, weight=1)
        self.btn_gerar = ttk.Button(acoes, text="Gerar relatório",
                                    style="Gerar.TButton", command=self._gerar_relatorio)
        self.btn_gerar.grid(row=0, column=0, sticky="w")
        self.progresso = ttk.Progressbar(acoes, mode="indeterminate", length=220)
        self.progresso.grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Label(acoes, textvariable=self.status_var,
                  style="Dica.TLabel").grid(row=0, column=2, sticky="e")

        log_frame = ttk.LabelFrame(corpo, text="Andamento", padding=(10, 6, 10, 8))
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=8, wrap="word", state="disabled", font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky="nsew")

        rodape = ttk.Frame(self.root, padding=(20, 6, 20, 12))
        rodape.grid(row=2, column=0, sticky="ew")
        rodape.columnconfigure(0, weight=1)
        ttk.Label(rodape, text="Info.txt e .fonts são lidos da pasta do programa.",
                  style="Dica.TLabel").grid(row=0, column=0, sticky="w")
        self.btn_abrir_pdf = ttk.Button(rodape, text="Abrir PDF",
                                        command=self._abrir_pdf, state="disabled")
        self.btn_abrir_pdf.grid(row=0, column=1, padx=(0, 8))
        self.btn_abrir_pasta = ttk.Button(rodape, text="Abrir pasta",
                                          command=self._abrir_pasta, state="disabled")
        self.btn_abrir_pasta.grid(row=0, column=2)

    def _centralizar(self):
        self.root.update_idletasks()
        largura = 900
        altura = 780
        x = max(0, (self.root.winfo_screenwidth() - largura) // 2)
        y = max(0, (self.root.winfo_screenheight() - altura) // 2)
        self.root.geometry(f"{largura}x{altura}+{x}+{y}")

    def _carregar_info(self, confirmar=False):
        if confirmar and self._info_alterado():
            if not messagebox.askyesno(
                    "Recarregar Info.txt",
                    "Descartar as alterações feitas no analista e nos comentários?",
                    parent=self.root):
                return
        erro_info = None
        try:
            self.info = drx.ler_info(self.info_path)
        except Exception as exc:
            self.info = {}
            erro_info = exc
        self.equip_var.set(self.info.get("instrumento") or "—")
        self.serial_var.set(self.info.get("serial") or "—")
        software = " ".join(x for x in (self.info.get("software"),
                                        self.info.get("versao")) if x)
        self.software_var.set(software or "—")
        self.fonte_var.set(self.info.get("fonte") or "")
        self.analista_var.set(self.info.get("analista") or drx.ANALISTA)
        self._definir_comentarios("\n\n".join(self.info.get("textos") or []))
        if erro_info is not None:
            self.info_status_var.set(
                f"Erro ao ler Info.txt: {erro_info} — usando valores padrão.")
            self.info_status_label.configure(style="Aviso.TLabel")
        elif self.info_path.exists():
            self.info_status_var.set(f"Info.txt: {self.info_path}")
            self.info_status_label.configure(style="Ok.TLabel")
        else:
            self.info_status_var.set(
                f"Info.txt não encontrado: {self.info_path} — usando valores padrão.")
            self.info_status_label.configure(style="Aviso.TLabel")
        self.info_original_analista = self.analista_var.get().strip()
        self.info_original_comentarios = self._texto_comentarios()
        self._atualizar_fonte()

    def _definir_comentarios(self, texto):
        estado = self.comentarios_text.cget("state")
        self.comentarios_text.configure(state="normal")
        self.comentarios_text.delete("1.0", "end")
        if texto:
            self.comentarios_text.insert("1.0", texto)
        self.comentarios_text.configure(state=estado)

    def _texto_comentarios(self):
        return self.comentarios_text.get("1.0", "end-1c").strip()

    def _comentarios(self):
        texto = self._texto_comentarios()
        if not texto:
            return []
        return [re.sub(r"\s+", " ", p).strip()
                for p in re.split(r"\n\s*\n", texto) if p.strip()]

    def _info_alterado(self):
        return (self.analista_var.get().strip() != self.info_original_analista
                or self._texto_comentarios() != self.info_original_comentarios)

    def _recarregar_info(self):
        if self.executando:
            return
        self._carregar_info(confirmar=True)

    def _selecionar_pasta(self):
        inicial = self.pasta if self.pasta and self.pasta.is_dir() else BASE
        pasta = filedialog.askdirectory(
            parent=self.root, title="Selecione a pasta com os arquivos da medição",
            initialdir=str(inicial), mustexist=True)
        if not pasta:
            return
        self.pasta_var.set(pasta)
        self._analisar_pasta()

    def _limpar_pasta(self, mensagem):
        self.pasta = None
        self.xlsx = []
        self.hdf5s = []
        self.hdf5_var.set("")
        self.hdf5_combo.configure(values=[], state="disabled")
        self.fonte_hdf5 = None
        info_fonte = drx.normalizar_fonte(self.info.get("fonte"))
        self.fonte_var.set(info_fonte or "")
        self.fonte_status_var.set(
            "Padrão do Info.txt" if info_fonte else
            "Informe ou deixe em branco para inferir das planilhas.")
        self.resumo_pasta_var.set(mensagem)
        self.destino_var.set("O PDF será salvo na pasta selecionada.")
        self.ultimo_pdf = None
        self._aplicar_estado()

    def _analisar_pasta(self):
        texto = self.pasta_var.get().strip().strip('"')
        if not texto:
            self._limpar_pasta("Nenhuma pasta selecionada.")
            self.status_var.set("Selecione uma pasta para começar.")
            return
        pasta = Path(texto)
        if not pasta.is_dir():
            self._limpar_pasta("Pasta inválida.")
            self.status_var.set("Selecione uma pasta válida.")
            return
        self.pasta = pasta
        self.xlsx = [p for p in sorted(pasta.glob("*.xlsx"))
                     if not p.name.startswith("~$")]
        self.hdf5s = sorted(pasta.glob("*.hdf5")) + sorted(pasta.glob("*.h5"))
        nomes = [p.name for p in self.hdf5s]
        valores = ([SEM_HDF5] + nomes) if nomes else [SEM_HDF5]
        self.hdf5_combo.configure(values=valores)
        if len(nomes) == 1:
            self.hdf5_var.set(nomes[0])
        elif nomes:
            if self.hdf5_var.get() not in valores:
                self.hdf5_var.set("")
        else:
            self.hdf5_var.set(SEM_HDF5)
        n_xlsx = len(self.xlsx)
        resumo = f"{n_xlsx} planilha{'s' if n_xlsx != 1 else ''} .xlsx"
        if self.hdf5s:
            resumo += f" • {len(self.hdf5s)} HDF5"
            if len(self.hdf5s) > 1:
                resumo += " — confirme o arquivo correto"
        else:
            resumo += " • sem HDF5 (somente planilhas)"
        self.resumo_pasta_var.set(resumo)
        self.destino_var.set(f"O PDF será salvo em: {pasta}")
        self.ultimo_pdf = None
        self._atualizar_fonte()
        if self._pode_gerar():
            self.status_var.set("Pasta pronta para gerar.")
        elif self.xlsx and self.hdf5s and not self.hdf5_var.get().strip():
            self.status_var.set("Escolha o HDF5 ou 'sem HDF5'.")
        else:
            self.status_var.set("Selecione uma pasta com planilhas .xlsx.")

    def _pode_gerar(self):
        if not (self.pasta and self.pasta.is_dir() and self.xlsx):
            return False
        if not self.hdf5s:
            return True
        nome = self.hdf5_var.get().strip()
        if not nome:
            return False
        return nome == SEM_HDF5 or self._hdf5_selecionado() is not None

    def _hdf5_selecionado(self):
        if not self.pasta:
            return None
        nome = self.hdf5_var.get().strip()
        if not nome or nome == SEM_HDF5:
            return None
        caminho = self.pasta / nome
        return caminho if caminho.exists() else None

    def _atualizar_fonte(self):
        h5 = self._hdf5_selecionado()
        fonte = None
        if h5 is not None:
            try:
                fonte = drx.normalizar_fonte(
                    drx.ler_hdf5(h5).get("Source target material"))
            except (Exception, SystemExit) as exc:
                self._log(f"\nAviso: não foi possível ler a fonte do HDF5: {exc}\n")
        self.fonte_hdf5 = fonte
        if fonte:
            self.fonte_var.set(fonte)
            self.fonte_status_var.set("Detectada no HDF5")
        else:
            info_fonte = drx.normalizar_fonte(self.info.get("fonte"))
            atual = self.fonte_var.get().strip()
            if not atual and info_fonte:
                self.fonte_var.set(info_fonte)
                atual = info_fonte
            if atual:
                self.fonte_status_var.set(
                    "Padrão do Info.txt" if atual == info_fonte
                    else "Selecionada para este relatório")
            else:
                self.fonte_status_var.set(
                    "Informe ou deixe em branco para inferir das planilhas.")
        self._aplicar_estado()

    def _hdf5_alterado(self):
        self.ultimo_pdf = None
        if self.pasta:
            self.destino_var.set(f"O PDF será salvo em: {self.pasta}")
        self._atualizar_fonte()
        if self._pode_gerar():
            self.status_var.set("Pasta pronta para gerar.")
        elif self.xlsx and self.hdf5s and not self.hdf5_var.get().strip():
            self.status_var.set("Escolha o HDF5 ou 'sem HDF5'.")

    def _aplicar_estado(self):
        ativo = not self.executando
        estado = "normal" if ativo else "disabled"
        self.pasta_entry.configure(state="readonly" if ativo else "disabled")
        self.btn_pasta.configure(state=estado)
        self.btn_atualizar.configure(state=estado)
        self.btn_recarregar_info.configure(state=estado)
        self.analista_entry.configure(state=estado)
        self.comentarios_text.configure(state=estado)
        self.chk_offline.configure(state=estado)
        if ativo and self.hdf5s:
            self.hdf5_combo.configure(state="readonly")
        else:
            self.hdf5_combo.configure(state="disabled")
        if ativo and not self.fonte_hdf5:
            self.fonte_combo.configure(state="readonly")
        else:
            self.fonte_combo.configure(state="disabled")
        self.btn_gerar.configure(
            state="normal" if ativo and self._pode_gerar() else "disabled")
        self.btn_abrir_pdf.configure(
            state="normal" if ativo and self.ultimo_pdf and self.ultimo_pdf.exists()
            else "disabled")
        self.btn_abrir_pasta.configure(
            state="normal" if ativo and self.pasta and self.pasta.is_dir()
            else "disabled")

    def _gerar_relatorio(self):
        if self.executando:
            return
        if not self.pasta or not self.pasta.is_dir():
            messagebox.showwarning("Pasta", "Selecione a pasta da medição.",
                                   parent=self.root)
            return
        if not self.xlsx:
            messagebox.showwarning(
                "Planilhas", "A pasta selecionada não contém planilhas .xlsx.",
                parent=self.root)
            return
        if self.hdf5s and not self.hdf5_var.get().strip():
            messagebox.showwarning(
                "HDF5", "Escolha um arquivo HDF5 ou a opção 'sem HDF5'.",
                parent=self.root)
            return
        hdf5 = self._hdf5_selecionado()
        if (self.hdf5s and self.hdf5_var.get().strip() != SEM_HDF5
                and hdf5 is None):
            messagebox.showwarning(
                "HDF5", "O arquivo HDF5 selecionado não foi encontrado. "
                        "Atualize a pasta ou escolha outro arquivo.",
                parent=self.root)
            return
        existentes = sorted(self.pasta.glob("Relatorio_DRX_*.pdf"))
        if existentes:
            nomes = "\n".join(p.name for p in existentes[:5])
            if len(existentes) > 5:
                nomes += "\n..."
            if not messagebox.askyesno(
                    "Substituir PDF",
                    "Já existe relatório nesta pasta:\n\n" + nomes +
                    "\n\nContinuar e substituir?",
                    parent=self.root):
                return
        analista = self.analista_var.get().strip()
        comentarios = self._comentarios()
        fonte = self.fonte_var.get().strip() or None
        offline = self.offline_var.get()
        self.ultimo_pdf = None
        self.destino_var.set(f"Gerando em: {self.pasta}")
        self.executando = True
        self.status_var.set("Gerando relatório...")
        self.progresso.start(12)
        self._log_limpar()
        self._log("Iniciando geração...\n")
        if hdf5 is None:
            self._log("HDF5: não fornecido — usando somente planilhas.\n")
        self._aplicar_estado()
        try:
            threading.Thread(target=self._executar,
                             args=(self.pasta, hdf5, fonte, analista, comentarios, offline),
                             daemon=True).start()
        except Exception:
            self.executando = False
            self.progresso.stop()
            self.destino_var.set(f"O PDF será salvo em: {self.pasta}")
            self._aplicar_estado()
            self.status_var.set("Erro ao iniciar a geração.")
            self._mostrar_erro(traceback.format_exc())

    def _executar(self, pasta, hdf5, fonte, analista, comentarios, offline):
        escritor = EscritorFila(self.fila)
        try:
            with contextlib.redirect_stdout(escritor), contextlib.redirect_stderr(escritor):
                pdf = drx.executar_geracao(
                    pasta=pasta, hdf5=str(hdf5) if hdf5 is not None else None,
                    info=str(self.info_path), offline=offline, analista=analista,
                    comentarios=comentarios, fonte=fonte,
                    sem_hdf5=hdf5 is None)
            self.fila.put(("ok", str(pdf)))
        except SystemExit as exc:
            codigo = exc.code
            if isinstance(codigo, str) and codigo.strip():
                mensagem = codigo.strip()
            elif codigo not in (None, 0):
                mensagem = f"Erro ao gerar o relatório (código {codigo})."
            else:
                mensagem = "A geração foi interrompida."
            self.fila.put(("erro", mensagem))
        except Exception:
            self.fila.put(("erro", traceback.format_exc()))
        finally:
            self.fila.put(("fim", None))

    def _processar_fila(self):
        erro = None
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "log":
                    self._log(dado)
                elif tipo == "ok":
                    self.ultimo_pdf = Path(dado)
                    self.status_var.set(f"Relatório gerado: {self.ultimo_pdf.name}")
                    self.destino_var.set(f"PDF gerado em: {self.ultimo_pdf}")
                elif tipo == "erro":
                    erro = dado
                    self.status_var.set("Erro ao gerar o relatório.")
                    self._log(f"\nERRO:\n{dado}\n")
                elif tipo == "fim":
                    self._finalizar_geracao()
        except queue.Empty:
            pass
        if erro is not None:
            self._mostrar_erro(erro)
        self.root.after(100, self._processar_fila)

    def _finalizar_geracao(self):
        self.executando = False
        self.progresso.stop()
        self._aplicar_estado()
        if self.ultimo_pdf and self.ultimo_pdf.exists():
            self.status_var.set(f"Pronto: {self.ultimo_pdf.name}")
        elif self.pasta:
            self.destino_var.set(f"O PDF será salvo em: {self.pasta}")

    def _mostrar_erro(self, texto, titulo="Erro ao gerar relatório"):
        texto = str(texto).strip()
        if len(texto) > 1800:
            texto = texto[:1800] + "\n\n(Detalhes completos no log.)"
        messagebox.showerror(titulo, texto, parent=self.root)

    def _report_callback_exception(self, exc, valor, tb):
        texto = "".join(traceback.format_exception(exc, valor, tb))
        self._log(f"\nERRO:\n{texto}\n")
        self._mostrar_erro(texto, "Erro na interface")

    def _log_limpar(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _log(self, texto):
        if not texto:
            return
        self.log_text.configure(state="normal")
        self.log_text.insert("end", texto)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _abrir_pdf(self):
        if not self.ultimo_pdf or not self.ultimo_pdf.exists():
            messagebox.showwarning("PDF", "Gere o relatório primeiro.",
                                   parent=self.root)
            return
        try:
            _abrir_caminho(self.ultimo_pdf)
        except Exception as exc:
            messagebox.showerror("Não foi possível abrir o PDF", str(exc),
                                 parent=self.root)

    def _abrir_pasta(self):
        if not self.pasta or not self.pasta.is_dir():
            return
        try:
            _abrir_caminho(self.pasta)
        except Exception as exc:
            messagebox.showerror("Não foi possível abrir a pasta", str(exc),
                                 parent=self.root)

    def _fechar(self):
        if self.executando and not messagebox.askyesno(
                "Sair", "A geração está em andamento. Deseja fechar mesmo assim?",
                parent=self.root):
            return
        self.root.destroy()


def main():
    _ativar_dpi()
    try:
        root = tk.Tk()
    except Exception as exc:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                0, "Não foi possível abrir a interface.\n\n" + str(exc),
                "Gerador de Relatório DRX", 0x10)
        except Exception:
            print("Não foi possível abrir a interface:", exc, file=sys.stderr)
        return
    root.withdraw()
    root.title("Gerador de Relatório DRX")
    if drx is None:
        messagebox.showerror(
            "Erro",
            "Não foi possível carregar gerar_relatorio_drx.py.\n\n" + str(ERRO_MOTOR),
            parent=root)
        root.destroy()
        return
    try:
        app = AplicativoDRX(root)
        root.aplicativo = app
    except Exception:
        messagebox.showerror("Erro na interface", traceback.format_exc(), parent=root)
        root.destroy()
        return
    root.deiconify()
    root.mainloop()


if __name__ == "__main__":
    main()
