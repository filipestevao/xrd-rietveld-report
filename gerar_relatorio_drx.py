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

"""
gerar_relatorio_drx.py
======================
Gera um relatório PDF de análise de DRX (refinamento de Rietveld) a partir das
exportações do XRDanalysis PRO (Anton Paar).

COMO USAR
---------
1. Coloque numa mesma pasta (ex.: "Export"):
     * as planilhas .xlsx exportadas (o NOME DO ARQUIVO não importa; cada uma é
       reconhecida pelo nome da aba: Quantitative - Phase List, PeakList,
       GlobalInstrument, RValue, Relative Mass, PrimaryChart, Differential Chart
       e, opcionalmente, Chebyshev);
     * o arquivo .hdf5 da medição                                [recomendado];
     * um arquivo Info.txt                                       [opcional].
   Outros arquivos na pasta são ignorados (inclusive PDFs e relatórios já
   gerados). O .hdf5 fornece os metadados experimentais; sem ele, o relatório
   é gerado apenas com as planilhas e o Info.txt.
2. Execute:   python gerar_relatorio_drx.py Export
   (opções:  -o Relatorio.pdf   --hdf5 medicao.hdf5   --info Info.txt
             --fonte Cu   --sem-hdf5   --offline)

INTERFACE GRÁFICA
-----------------
   pythonw interface_drx.py
   (a interface lê o Info.txt da pasta do programa, permite editar o nome do
    analista e os comentários apenas para o relatório atual e salva o PDF na
    pasta da medição selecionada)

INFO.TXT (texto simples, não precisa saber Python)
--------------------------------------------------
   XRDynamic 500 - S/N 12345678        <- linha 1: equipamento e nº de série
   XRDanalysis PRO 1.2.5.3658          <- linha 2: software de análise e versão
   Fonte: Cu                           <- (opcional) Cu, Co, Cr, Mo ou Ag
   (o S/N do instrumento e o software de análise não existem no .hdf5;
    esses dados só aparecem se o Info.txt existir)
   Analista: Seu Nome                  <- (opcional) nome para o rodapé

   Qualquer texto escrito depois de uma linha em branco vira "Comentários"
   no relatório. Separe parágrafos com linha em branco.

DEPENDÊNCIAS
------------
   pip install numpy pandas openpyxl matplotlib reportlab pillow
   HDF5 (opcional): pip install h5py
FONTES: Roboto-Light/Regular/Medium.ttf numa pasta "fonts" ou ".fonts" ao lado
do script, na pasta dos dados ou instaladas no sistema. Sem elas usa-se DejaVu Sans.
"""
import argparse
import base64
import datetime as dt
import io
import json
import logging
import os
import re
import sys
import unicodedata
import urllib.request
from pathlib import Path
from xml.sax.saxutils import escape

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.gridspec import GridSpec, GridSpecFromSubplotSpec
from matplotlib.lines import Line2D

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (BaseDocTemplate, CondPageBreak, Frame, Image,
                                KeepTogether, PageBreak, PageTemplate,
                                Paragraph, Spacer, Table, TableStyle)

# silencia avisos de peso de fonte do matplotlib (não são erros)
logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)

_ParagraphBase = Paragraph


class Paragraph(_ParagraphBase):
    """Paragraph que troca χ (idêntico ao 'x' na Roboto) por um χ distinto."""

    def __init__(self, text, style=None, *a, **k):
        text = str(text).replace("χ", '<font name="RB-Chi">χ</font>')
        super().__init__(text, style, *a, **k)


# =============================================================================
# CONFIGURAÇÃO  (valores padrão; o Info.txt tem prioridade quando existir)
# =============================================================================
ANALISTA = ""  # rodapé
TITULO_RELATORIO = "Relatório de análise de DRX"

# Opcional: dados que as exportações não trazem e que o COD não resolveu.
# Chave = fórmula (qualquer ordem). Exemplo:
#   INFO_FASES = {"LaBO3": {"grupo_espacial": "Pnma (n.º 62)", "Z": 4}}
INFO_FASES = {}

# Apêndice: omite reflexões com I/I1 abaixo deste valor (%). Use 0 para listar todas.
LIMITE_APENDICE = 0.5
# =============================================================================

PALETA_FASES = ["#CC0000", "#3F6EA6", "#2E8B57", "#B8860B", "#7B5EA7", "#808080"]
COR = dict(vermelho="#CC0000", cinza="#666666", cinza_claro="#B3B3B3",
           linha="#E6E6E6", fundo="#F5F5F5", azul="#3F6EA6", preto="#000000",
           obs="#2B2B2B")
NBSP = "\u00a0"
MINUS = "\u2212"
COD_URL_CIF = "https://www.crystallography.net/cod/{}.cif"
COD_URL_PAG = "https://www.crystallography.net/cod/{}.html"

_TAB_ELEM = """
H 1.008 2.20;He 4.0026 -;Li 6.94 0.98;Be 9.0122 1.57;B 10.81 2.04;C 12.011 2.55;
N 14.007 3.04;O 15.999 3.44;F 18.998 3.98;Ne 20.180 -;Na 22.990 0.93;Mg 24.305 1.31;
Al 26.982 1.61;Si 28.085 1.90;P 30.974 2.19;S 32.06 2.58;Cl 35.45 3.16;Ar 39.948 -;
K 39.098 0.82;Ca 40.078 1.00;Sc 44.956 1.36;Ti 47.867 1.54;V 50.942 1.63;Cr 51.996 1.66;
Mn 54.938 1.55;Fe 55.845 1.83;Co 58.933 1.88;Ni 58.693 1.91;Cu 63.546 1.90;Zn 65.38 1.65;
Ga 69.723 1.81;Ge 72.630 2.01;As 74.922 2.18;Se 78.971 2.55;Br 79.904 2.96;Kr 83.798 -;
Rb 85.468 0.82;Sr 87.62 0.95;Y 88.906 1.22;Zr 91.224 1.33;Nb 92.906 1.6;Mo 95.95 2.16;
Ru 101.07 2.2;Rh 102.91 2.28;Pd 106.42 2.20;Ag 107.87 1.93;Cd 112.41 1.69;In 114.82 1.78;
Sn 118.71 1.96;Sb 121.76 2.05;Te 127.60 2.1;I 126.90 2.66;Xe 131.29 -;Cs 132.91 0.79;
Ba 137.33 0.89;La 138.91 1.10;Ce 140.12 1.12;Pr 140.91 1.13;Nd 144.24 1.14;Sm 150.36 1.17;
Eu 151.96 1.2;Gd 157.25 1.2;Tb 158.93 1.1;Dy 162.50 1.22;Ho 164.93 1.23;Er 167.26 1.24;
Tm 168.93 1.25;Yb 173.05 1.1;Lu 174.97 1.27;Hf 178.49 1.3;Ta 180.95 1.5;W 183.84 2.36;
Re 186.21 1.9;Os 190.23 2.2;Ir 192.22 2.20;Pt 195.08 2.28;Au 196.97 2.54;Hg 200.59 2.00;
Tl 204.38 1.62;Pb 207.2 2.33;Bi 208.98 2.02;Th 232.04 1.3;U 238.03 1.38
"""
MASSA, ELETRONEG = {}, {}
for _it in _TAB_ELEM.replace("\n", "").split(";"):
    _p = _it.split()
    if len(_p) == 3:
        MASSA[_p[0]] = float(_p[1])
        if _p[2] != "-":
            ELETRONEG[_p[0]] = float(_p[2])

KALFA1 = {"Cu": 1.540593, "Co": 1.788965, "Mo": 0.709319, "Cr": 2.289760,
          "Fe": 1.936041, "Ag": 0.559421}

FONTES_RAIO_X = ("Cu", "Co", "Cr", "Mo", "Ag")
TOL_FONTE = 0.01  # Å


def normalizar_fonte(valor):
    """Extrai o símbolo de uma fonte de raios X suportada."""
    if not valor:
        return None
    m = re.search(r"(Cu|Co|Cr|Mo|Ag)(?=\s*K|\b)", str(valor), re.I)
    if not m:
        return None
    s = m.group(1)
    return s[0].upper() + s[1:].lower()


def inferir_fonte_por_lambda(lam):
    """Inferência conservadora da fonte a partir do λ calculado do PeakList."""
    if lam is None or not np.isfinite(lam):
        return None
    fonte = min(FONTES_RAIO_X, key=lambda f: abs(KALFA1[f] - lam))
    return fonte if abs(KALFA1[fonte] - lam) <= TOL_FONTE else None


TRAD = {
    "coupled scan": "Varredura acoplada (θ–2θ)",
    "fixed tube x-ray": "Tubo de raios X fixo",
    "continuous (default, high dynamic range)": "Contínuo (alta faixa dinâmica)",
    "continuous (high count rate)": "Contínuo (alta taxa de contagem)",
    "bragg-brentano (monochromator)": "Bragg-Brentano (com monocromador)",
    "bragg-brentano (no optic)": "Bragg-Brentano (sem óptica adicional)",
    "full detector region": "Região completa do detector",
    "fixed": "Fixa", "optimized": "Otimizada", "automatic": "Automática",
    "linear (1d)": "linear 1D",
    "finished": "Concluída", "completed": "Concluída",
    "running": "Em andamento", "aborted": "Abortada", "stopped": "Interrompida",
}

TRAD_ESTAGIOS = {
    "sample spinner": "Estágio de rotação de amostra (spinner)",
    "sample spinning": "Estágio de rotação de amostra (spinner)",
    "rotationary sample holder": "Porta-amostra rotativo",
    "rotary sample holder": "Porta-amostra rotativo",
    "sample holder": "Porta-amostra",
    "flat sample holder": "Porta-amostra plano",
    "flat plate sample holder": "Porta-amostra de placa plana",
    "capillary sample holder": "Porta-amostra capilar",
    "capillary spinner": "Estágio capilar com rotação",
    "zero background sample holder": "Porta-amostra de fundo zero",
}

ABAS_OBRIGATORIAS = {
    "Quantitative - Phase List": "fases", "PeakList": "picos",
    "GlobalInstrument": "instr", "RValue": "rvalues",
    "Relative Mass": "massa", "PrimaryChart": "padrao",
    "Differential Chart": "dif",
}
ABAS_OPCIONAIS = {"Chebyshev": "fundo"}


# =============================================================================
# Formatação (pt-BR)
# =============================================================================
def num(x, nd=2):
    """Número pt-BR: vírgula decimal, NBSP como separador de milhar."""
    if x is None or (isinstance(x, float) and not np.isfinite(x)):
        return "–"
    s = f"{x:,.{nd}f}".replace(",", "§").replace(".", ",").replace("§", NBSP)
    return s.replace("-", MINUS)


def sci(x, nd=1):
    if x == 0:
        return "0"
    exp = int(np.floor(np.log10(abs(x))))
    man = x / 10 ** exp
    return f"{num(man, nd)}×10<super>{str(exp).replace('-', MINUS)}</super>"


def esc(s):
    return escape(str(s))


def numero_em(texto, padrao=None):
    m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", str(texto))
    return float(m.group()) if m else padrao


def data_pt(s):
    """'10/13/2026 15:05:22 UTC -03:00' (MM/DD/AAAA) ou ISO -> datetime."""
    s = str(s)
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})", s)
    if m:
        return dt.datetime(*map(int, m.groups()))
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})", s)
    if not m:
        return None
    a, b, ano, hh, mm_, ss = map(int, m.groups())
    mes, dia = (a, b) if a <= 12 else (b, a)
    return dt.datetime(ano, mes, dia, hh, mm_, ss)


def duracao_pt(s):
    m = re.match(r"(\d+):(\d{2}):(\d{2})", str(s).strip())
    if not m:
        return str(s)
    h, mi, se = map(int, m.groups())
    return (f"{h} h {mi:02d} min {se:02d} s" if h else f"{mi} min {se:02d} s")


def tr(v):
    return TRAD.get(str(v).strip().lower(), str(v))


def tr_estagio(v):
    chave = " ".join(str(v).split()).lower()
    return TRAD_ESTAGIOS.get(chave, str(v))


def so_ascii(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9_-]+", "_", s).strip("_")


def norm(s):
    return re.sub(r"[^a-z0-9]", "", so_ascii(str(s)).lower())


def formula_livre_html(texto):
    """'ZrO2 + Al2O3' -> subscritos nos tokens que parecem fórmula."""
    def tok(m):
        t = m.group(0)
        if re.match(r"^[A-Z][A-Za-z0-9()]*\d[A-Za-z0-9()]*$", t):
            return re.sub(r"(?<=[A-Za-z)])(\d+)", r"<sub>\1</sub>", esc(t))
        return esc(t)
    return re.sub(r"\S+", tok, str(texto))


# =============================================================================
# Química
# =============================================================================
def parse_formula(f):
    f = re.sub(r"[^A-Za-z0-9.]", " ", str(f))
    cont = {}
    for el, n in re.findall(r"([A-Z][a-z]?)\s*(\d*\.?\d*)", f):
        cont[el] = cont.get(el, 0) + (float(n) if n else 1.0)
    return cont


def canon(f):
    return "".join(f"{e}{c:g}" for e, c in sorted(parse_formula(f).items()))


def formula_bonita(f):
    cont = parse_formula(f)
    if not cont:
        return str(f), esc(f)
    els = list(cont)
    if all(e in ELETRONEG for e in els):
        els.sort(key=lambda e: ELETRONEG[e])
    txt = "".join(e + (f"{cont[e]:g}" if cont[e] != 1 else "") for e in els)
    html = "".join(e + (f"<sub>{cont[e]:g}</sub>" if cont[e] != 1 else "")
                   for e in els)
    return txt, html


def massa_molar(f):
    cont = parse_formula(f)
    if not cont or any(e not in MASSA for e in cont):
        return None
    return sum(MASSA[e] * n for e, n in cont.items())


def sistema_celula(a, b, c, al, be, ga, tol=1e-3):
    eq = lambda x, y: abs(x - y) < tol * max(abs(x), abs(y), 1)
    a90 = [eq(x, 90) for x in (al, be, ga)]
    if all(a90):
        if eq(a, b) and eq(b, c):
            return "cúbica"
        if eq(a, b) or eq(b, c) or eq(a, c):
            return "tetragonal"
        return "ortorrômbica"
    if eq(al, 90) and eq(be, 90) and eq(ga, 120) and eq(a, b):
        return "hexagonal"
    if eq(a, b) and eq(b, c) and eq(al, be) and eq(be, ga):
        return "romboédrica"
    if sum(a90) == 2:
        return "monoclínica"
    return "triclínica"


def grupo_espacial_html(hm):
    """'P 1 21/c 1' -> P2<sub>1</sub>/c ;  'P 32 2 1' -> P3<sub>2</sub>21"""
    if not hm:
        return None
    hm = str(hm).strip()
    hexa = hm.endswith(":H") or " :H" in hm
    hm = hm.replace(":H", "").strip()
    t = hm.split()
    if len(t) == 4 and t[1] == "1" and t[3] == "1":
        t = [t[0], t[2]]
    s = "".join(t)
    s = re.sub(r"(?<![0-9])([2346])([1-5])(?=[/a-z\d]|$)",
               lambda m: (f"{m.group(1)}<sub>{m.group(2)}</sub>"
                          if int(m.group(2)) < int(m.group(1)) else m.group(0)), s)
    return s + (" (eixos hex.)" if hexa else "")


# =============================================================================
# Leitura: HDF5, Info.txt
# =============================================================================
def _casa(nome_dados, candidato):
    a, b = norm(nome_dados), norm(candidato)
    return bool(a and b and (a in b or b in a))


def achar_hdf5(pasta, nome_dados):
    arqs = sorted(list(Path(pasta).glob("*.hdf5")) + list(Path(pasta).glob("*.h5")))
    for p in arqs:
        if _casa(nome_dados, p.stem):
            return p
    return arqs[0] if len(arqs) == 1 else None


def ler_hdf5(caminho):
    """Metadados do .hdf5 da medição (recomendado)."""
    meta = {}
    try:
        import h5py
    except ImportError:
        sys.exit("Erro: h5py não instalado. Instale com: pip install h5py")
    try:
        f = h5py.File(caminho, "r")
    except Exception as e:
        sys.exit(f"Erro: não foi possível abrir {Path(caminho).name} — {e}")

    def val(path):
        try:
            v = f[path][()]
        except KeyError:
            return None
        if isinstance(v, np.ndarray):
            if v.size != 1:
                return None
            v = v.reshape(-1)[0]
        if isinstance(v, (bytes, np.bytes_)):
            v = v.decode("utf-8", "ignore")
        return v

    def put(k, v):
        if v is not None and not (isinstance(v, float) and not np.isfinite(v)):
            meta[k] = str(v)

    def iso(v):
        d = data_pt(v) if v else None
        return f"{d:%m/%d/%Y %H:%M:%S} UTC" if d else None

    e, i, ins = "entry/", "entry/info/", "entry/instrument/"
    put("Scan name", val(e + "title"))
    put("Experiment name", val(e + "experiment_description"))
    put("Sample name", val("entry/sample/name"))
    put("Sample comment", val("entry/sample/description"))
    v = val("entry/sample/Z")
    if v is not None and np.isfinite(v):
        put("Sample Z", f"{v * 1000:.3f} mm")
    v = val("entry/sample/scan_omega_offset")
    if v is not None and np.isfinite(v):
        put("Omega offset", f"{v:g}°")
    v = val("entry/sample/temperature_SP")
    if v is not None and np.isfinite(v):
        put("Sample temperature", f"{v - 273.15:.2f} °C")
    st = val(i + "scan_type")
    put("Scan type", re.sub(r"(?<=[a-z])(?=[A-Z])", " ", st) if st else None)
    put("Scan mode", val(i + "mode"))
    for k, p in (("Start position", "start"), ("End position", "stop")):
        v = val(i + p)
        put(k, f"{v:.4f}°" if v is not None else None)
    v = val(i + "step_size")
    put("2θ step size", f"{v:.4f}°" if v is not None else None)
    v = val(i + "time_per_step")
    put("Time per step", f"{v:.3f} s" if v is not None else None)
    t0, t1 = val(e + "start_time"), val(e + "end_time")
    put("Scan start time", iso(t0))
    put("Scan end time", iso(t1))
    if t0 and t1 and data_pt(t0) and data_pt(t1):
        sec = int((data_pt(t1) - data_pt(t0)).total_seconds())
        put("Scan duration", f"{sec // 3600:02d}:{sec % 3600 // 60:02d}:{sec % 60:02d}")
    put("Scan status", val(i + "status"))
    prog = val(e + "program_name")
    try:
        ver = f[e + "program_name"].attrs.get("version")
        ver = ver[0].decode() if isinstance(ver, np.ndarray) else ver
    except Exception:
        ver = None
    put("Software", f"{prog} {ver}" if prog and ver else prog)
    put("User name", val("entry/user/name"))
    put("Source type", val(ins + "source/type"))
    alvo = val(ins + "source/target_material")
    if alvo:
        m = re.match(r"^\s*([A-Z][a-z]?)\s*K\s*\*?\s*([a-zα]?)", str(alvo), re.I)
        if m:
            ka = (m.group(2) or "").lower() in ("", "a", "α")
            put("Source target material", f"{m.group(1)} K{'α' if ka else 'β'}")
        else:
            put("Source target material", alvo)
    put("Source name", val(ins + "source/name"))
    v = val(ins + "source/voltage")
    put("Generator voltage", f"{v / 1000:g} kV" if v is not None else None)
    v = val(ins + "source/current")
    put("Generator current", f"{v * 1000:g} mA" if v is not None else None)
    put("Beam geometry", val(ins + "instrument_info/beam_geometry"))
    put("Last instrument alignment", iso(val(ins + "instrument_info/last_instrument_alignment")))
    put("Sample stage", val(ins + "instrument_info/stage"))
    put("Sample stage configuration", val(ins + "instrument_info/stage_configuration"))
    sp = val(ins + "instrument_info/sample_spinning")
    put("Sample spinning", "True" if sp else "False" if sp is not None else None)
    v = val(ins + "detector/distance")
    put("Goniometer radius", f"{v * 1000:g} mm" if v is not None else None)
    put("Detector model", val(ins + "detector/description"))
    put("Detector sensor material", val(ins + "detector/sensor_material"))
    lay = val(ins + "detector/layout")
    put("Detector mode", "Linear (1D)" if lay == "linear" else lay)
    put("Detector S/N", val(ins + "detector/detector_number"))
    try:
        roi = val(ins + "detector/ROI/height")
        full = f[ins + "detector/ROI"].attrs.get("full_detector_used")
        full = bool(full.reshape(-1)[0]) if isinstance(full, np.ndarray) else bool(full)
        put("Detector region", "Full detector region" if roi is not None and full else None)
    except Exception:
        pass
    put("Divergence slit type", val(ins + "divergence_slit/type"))
    put("Divergence slit mode", val(ins + "divergence_slit/mode"))
    v = val(ins + "divergence_slit/divergence_y")
    put("Divergence slit opening", f"{v:g}°" if v is not None else None)
    v = val(ins + "beam_mask/gap_x")
    put("Beam mask opening", f"{v * 1000:.3f} mm" if v is not None else None)
    put("Anti-scatter slit mode", val(ins + "anti_scatter_slit/mode"))
    v = val(ins + "anti_scatter_slit/gap_y")
    put("Anti-scatter slit opening", f"{v * 1000:.3f} mm" if v is not None else None)
    v = val(i + "fluorescence_suppression")
    put("Fluorescence reduction", "True" if v else "False" if v is not None else None)
    for k, p in (("Primary Soller slit", "primary_soller"), ("Secondary Soller slit", "secondary_soller")):
        v = val(ins + p + "/soller_angle")
        put(k, f"{v:g} rad" if v is not None else None)
    put("Attenuator / filter type", val(ins + "attenuator/type"))
    put("instrumento", val(ins + "name"))
    f.close()
    return meta


def ler_info(caminho):
    """Info.txt: linha 1 = equipamento - S/N; linha 2 = software versão;
    'Analista: ...'; texto livre após linha em branco = comentários."""
    info = dict(instrumento=None, serial=None, software=None, versao=None,
                fonte=None, analista=None, textos=[])
    if not caminho or not Path(caminho).exists():
        return info
    raw = Path(caminho).read_bytes()
    try:
        txt = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        txt = raw.decode("cp1252", "ignore")
    linhas = txt.splitlines()
    i = 0
    while i < len(linhas) and not linhas[i].strip():
        i += 1
    cab = []
    while i < len(linhas) and linhas[i].strip():
        cab.append(linhas[i].strip())
        i += 1
    resto = linhas[i:]
    for l in cab:
        m = re.match(r"^analista\s*[:=]\s*(.+)$", l, re.I)
        if m:
            info["analista"] = m.group(1).strip()
            continue
        m = re.match(r"^(?:fonte|radia\S*|[âa]nodo)\s*[:=]\s*(.+)$", l, re.I)
        if m:
            fonte = normalizar_fonte(m.group(1))
            if fonte:
                info["fonte"] = fonte
            else:
                info["textos"].append(l)
            continue
        m = re.match(r"^(.*?)\s*[-–—,;|]\s*(?:S\s*/?\s*N|serial(?:\s*number)?|n[úu]mero de s[ée]rie)\s*\.?\s*[:#]?\s*(\S+)\s*$", l, re.I)
        if m and not info["instrumento"]:
            info["instrumento"], info["serial"] = m.group(1).strip(), m.group(2).strip()
            continue
        m = re.match(r"^(.*?)\s+v?(\d+(?:\.\d+){1,3})\s*$", l)
        if m and not info["software"]:
            info["software"], info["versao"] = m.group(1).strip(), m.group(2)
            continue
        info["textos"].append(l)       # linha não reconhecida do cabeçalho vira texto
    par = []
    for l in resto + [""]:
        m = re.match(r"^analista\s*[:=]\s*(.+)$", l.strip(), re.I)
        if m:
            info["analista"] = m.group(1).strip()
            continue
        m = re.match(r"^(?:fonte|radia\S*|[âa]nodo)\s*[:=]\s*(.+)$", l.strip(), re.I)
        if m:
            fonte = normalizar_fonte(m.group(1))
            if fonte:
                info["fonte"] = fonte
                continue
        if l.strip():
            par.append(l.strip())
        elif par:
            info["textos"].append(" ".join(par))
            par = []
    return info


# =============================================================================
# Leitura das planilhas
# =============================================================================
def localizar_arquivos(pasta):
    import openpyxl
    achados, vistos = {}, []
    todas = {**ABAS_OBRIGATORIAS, **ABAS_OPCIONAIS}
    for f in sorted(Path(pasta).glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        try:
            wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
            vistos.append((f.name, wb.sheetnames))
            for aba in wb.sheetnames:
                if aba in todas and todas[aba] not in achados:
                    achados[todas[aba]] = (f, aba)
            wb.close()
        except Exception as e:
            print(f"Aviso: ignorando {f.name}: {e}")
    faltando = [a for a, k in ABAS_OBRIGATORIAS.items() if k not in achados]
    if faltando:
        det = "; ".join(f"{n}: {s}" for n, s in vistos)
        sys.exit("Erro: abas não encontradas: " + ", ".join(faltando) + f"\n  Arquivos vistos -> {det}")
    return achados


def achar_col(df, prefixo, obrigatorio=True):
    for c in df.columns:
        if str(c).strip().lower().startswith(prefixo.lower()):
            return c
    if obrigatorio:
        raise KeyError(f"coluna '{prefixo}' não encontrada em {list(df.columns)}")
    return None


def ler_fases(caminho, aba):
    import openpyxl
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    linhas = [list(r) for r in wb[aba].iter_rows(values_only=True)]
    wb.close()
    cab = [str(c) if c is not None else "" for c in linhas[0]]
    fases, atual, modo, cab_at, cab_or = [], None, None, [], []
    for r in linhas[1:]:
        if r[0] is not None:
            atual = {"formula_raw": str(r[0]), "atomos": [], "orient": []}
            for i, c in enumerate(cab[1:], start=1):
                atual[c] = r[i]
            fases.append(atual)
            modo = None
            continue
        rot = r[1]
        if atual is None or all(v is None for v in r):
            modo = None
        elif rot == "Atom":
            modo, cab_at = "atomos", [str(c) for c in r[1:]]
        elif rot == "Method":
            modo, cab_or = "orient", [str(c) for c in r[1:]]
        elif modo == "atomos":
            atual["atomos"].append(dict(zip(cab_at, r[1:])))
        elif modo == "orient":
            atual["orient"].append(dict(zip(cab_or, r[1:])))
    return fases


def get(d, prefixo, padrao=None):
    for k, v in d.items():
        if str(k).strip().lower().startswith(prefixo.lower()):
            return v
    return padrao


def carregar(pasta, hdf5=None, info_txt=None, fonte=None, sem_hdf5=False):
    arqs = localizar_arquivos(pasta)
    D = {}
    D["fases"] = ler_fases(*arqs["fases"])
    D["instr"] = pd.read_excel(*arqs["instr"]).iloc[0].to_dict()
    rv = pd.read_excel(*arqs["rvalues"]).dropna()
    D["rv"] = {" ".join(str(n).split()): float(v) for n, v in zip(rv.iloc[:, 0], rv.iloc[:, 1])}
    D["massa"] = pd.read_excel(*arqs["massa"]).dropna(how="all")
    pk = pd.read_excel(*arqs["picos"])
    D["picos"] = pk.dropna(subset=[achar_col(pk, "Phase")])

    pr = pd.read_excel(*arqs["padrao"])
    cols = list(pr.columns)
    c_bg = achar_col(pr, "Background")
    c_fit = achar_col(pr, "Quantitative Fit")
    c_obs = [c for c in cols[1:] if c not in (c_bg, c_fit)][0]
    c_fases = [c for c in cols[1:] if c not in (c_bg, c_fit, c_obs)]
    D["tt"] = pr.iloc[:, 0].to_numpy(float)
    D["yo"] = pr[c_obs].to_numpy(float)
    D["bg"] = pr[c_bg].to_numpy(float)
    D["yc"] = pr[c_fit].to_numpy(float)
    D["yfases"] = {str(c): pr[c].to_numpy(float) for c in c_fases}
    D["arquivo_dados"] = str(c_obs)
    D["dif"] = pd.read_excel(*arqs["dif"]).iloc[:, 1].to_numpy(float)

    D["fundo_coef"] = None
    if "fundo" in arqs:
        ch = pd.read_excel(*arqs["fundo"])
        D["fundo_coef"] = ch.iloc[:, 1].to_numpy(float)
        D["fundo_tipo"] = arqs["fundo"][1]
    verificar_fundo(D)

    # metadados: HDF5 quando disponível + Info.txt para dados do laboratório
    D["meta"] = {}
    D["fontes_meta"] = []
    h5 = None
    if not sem_hdf5:
        if hdf5:
            h5 = Path(hdf5)
            if not h5.is_file():
                sys.exit(f"Erro: arquivo HDF5 não encontrado: {h5}")
        else:
            h5 = achar_hdf5(pasta, D["arquivo_dados"])
            if h5 is None:
                achados = sorted(p.name for p in
                                 list(Path(pasta).glob("*.hdf5")) + list(Path(pasta).glob("*.h5")))
                if len(achados) > 1:
                    sys.exit("Erro: vários arquivos HDF5 na pasta (" + ", ".join(achados) +
                             "). Indique um com --hdf5 ou use --sem-hdf5.")
        if h5 is not None:
            D["meta"] = ler_hdf5(h5)
            D["fontes_meta"] = [Path(h5).name]
    if not D["meta"]:
        print("Aviso: HDF5 não fornecido — gerando relatório apenas com as planilhas e o Info.txt.")
    D["fonte_solicitada"] = normalizar_fonte(fonte)
    if info_txt is None:
        achados = [p for p in Path(pasta).glob("*") if p.name.lower() == "info.txt"]
        info_txt = achados[0] if achados else None
    D["info"] = ler_info(info_txt)
    D["info_arq"] = Path(info_txt).name if info_txt and Path(info_txt).exists() else None
    if not D["info_arq"]:
        print("Aviso: Info.txt não encontrado — S/N do instrumento e software de análise não serão exibidos.")
    return D


def verificar_fundo(D):
    """Confere se os coeficientes de Chebyshev reproduzem a curva de fundo exportada."""
    c = D.get("fundo_coef")
    if c is None:
        return
    from numpy.polynomial import chebyshev as C
    tt = D["tt"]
    x = (2 * tt - (tt.max() + tt.min())) / (tt.max() - tt.min())
    rec = C.chebval(x, c)
    err = np.abs(D["bg"] - rec)
    D["fundo_erro"] = (float(err.max()), float(np.sqrt((err ** 2).mean())),
                       float(err.max() / np.mean(D["bg"]) * 100))
    if D["fundo_erro"][2] > 1.0:
        print(f"[aviso] Os coeficientes de Chebyshev exportados não reproduzem a curva de fundo "
              f"exportada (erro máx. {D['fundo_erro'][0]:.0f} cts = {D['fundo_erro'][2]:.1f} % do fundo médio). "
              f"O relatório usa a curva exportada e não lista os coeficientes.")


# =============================================================================
# COD (Crystallography Open Database)
# =============================================================================
def cod_id(ref):
    """Referências 99 + 7 dígitos do XRDanalysis PRO correspondem ao COD (ex.: 999009666 -> 9009666)."""
    s = str(ref).strip()
    return s[2:] if re.fullmatch(r"99\d{7}", s) else None


def _cif_valor(txt, tag):
    m = re.search(rf"^{re.escape(tag)}[ \t]*(.*)$", txt, re.M)
    if not m:
        return None
    v = m.group(1).strip()
    if not v:                                  # valor na linha seguinte
        n = re.search(rf"^{re.escape(tag)}[ \t]*\n[ \t]*(.+)$", txt, re.M)
        v = n.group(1).strip() if n else ""
    if v[:1] in ("'", '"'):
        q = v[0]
        m2 = re.match(rf"{q}(.*?){q}(?:\s|$)", v)
        return m2.group(1) if m2 else v.strip(q)
    return v.split()[0] if v else None


def parse_cif(txt):
    g = lambda *tags: next((x for x in (_cif_valor(txt, t) for t in tags) if x not in (None, "?", ".")), None)
    aut = []
    m = re.search(r"loop_\s*\n\s*_publ_author_name\s*\n(.*?)(?=\n\s*(?:_|loop_|#|data_)|\Z)", txt, re.S)
    if m:
        for l in m.group(1).splitlines():
            l = l.strip().strip("'\"")
            if l:
                aut.append(l)
    z = g("_cell_formula_units_Z")
    return dict(
        formula=g("_chemical_formula_sum"),
        hm=g("_symmetry_space_group_name_H-M", "_space_group_name_H-M_alt"),
        sg_num=g("_space_group_IT_number", "_symmetry_Int_Tables_number"),
        Z=int(float(z)) if z else None,
        mineral=g("_chemical_name_mineral"),
        nome=g("_chemical_name_systematic"),
        autores=aut, revista=g("_journal_name_full"), ano=g("_journal_year"))


def cod_consultar(cid, pasta_cache, offline=False):
    """Devolve dict do CIF do COD (cache local; consulta online se permitido)."""
    arq = Path(pasta_cache) / "cod_cache.json"
    cache = {}
    try:
        if arq.exists():
            cache = json.loads(arq.read_text(encoding="utf-8"))
    except Exception:
        cache = {}
    if cid in cache:
        return cache[cid]
    if offline:
        return None
    try:
        req = urllib.request.Request(COD_URL_CIF.format(cid),
                                     headers={"User-Agent": "gerar_relatorio_drx/2.0"})
        with urllib.request.urlopen(req, timeout=8) as r:
            txt = r.read().decode("utf-8", "ignore")
        info = parse_cif(txt)
        if not info.get("formula"):
            raise ValueError("CIF sem fórmula")
    except Exception as e:
        print(f"Aviso: COD {cid} não consultado ({e.__class__.__name__}); "
              f"seguindo sem grupo espacial/Z/referência (ou use INFO_FASES).")
        return None
    cache[cid] = info
    try:
        arq.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception:
        pass
    return info


def citacao(info):
    if not info or not info.get("autores"):
        return None
    a = info["autores"][0].split(",")[0].strip()
    et = " et al." if len(info["autores"]) > 1 else ""
    partes = [f"{a}{et}"]
    if info.get("revista"):
        partes.append(info["revista"])
    if info.get("ano"):
        partes.append(f"({info['ano']})")
    return ", ".join(partes[:2]) + (f" {partes[2]}" if len(partes) > 2 else "")


# =============================================================================
# Análise
# =============================================================================
def ref_formula_do_pico(txt):
    partes = [p.strip() for p in str(txt).split("|")]
    return partes[0], partes[-1].strip("'\" ")


def analisar(D, pasta_cache, offline=False):
    A = {}
    tt, yo, yc, bg = D["tt"], D["yo"], D["yc"], D["bg"]
    pk = D["picos"]
    c2t, cd = achar_col(pk, "2θ"), achar_col(pk, "d [")
    chh, ck, cl = achar_col(pk, "h"), achar_col(pk, "k"), achar_col(pk, "l")
    cm, cf2 = achar_col(pk, "Multiplicity"), achar_col(pk, "F")
    cw, cph = achar_col(pk, "FWHM"), achar_col(pk, "Phase")

    # --- radiação e fonte ---------------------------------------------------
    t2 = pk[c2t].to_numpy(float)
    lam_all = 2 * pk[cd].to_numpy(float) * np.sin(np.radians(t2 / 2))
    lam = float(np.median(lam_all[t2 > np.percentile(t2, 70)]))
    fonte_h5 = normalizar_fonte(D["meta"].get("Source target material"))
    fonte_solic = normalizar_fonte(D.get("fonte_solicitada"))
    fonte_info = normalizar_fonte(D["info"].get("fonte"))
    fonte_inf = inferir_fonte_por_lambda(lam)
    if fonte_h5:
        fonte, origem = fonte_h5, "HDF5"
    elif fonte_solic:
        fonte, origem = fonte_solic, "interface/CLI"
    elif fonte_info:
        fonte, origem = fonte_info, "Info.txt"
    elif fonte_inf:
        fonte, origem = fonte_inf, "inferida do PeakList"
    else:
        sys.exit("Erro: não foi possível determinar a fonte de raios X. "
                 "Informe 'Fonte: Cu|Co|Cr|Mo|Ag' no Info.txt ou use --fonte.")
    A["fonte"], A["fonte_origem"] = fonte, origem
    A["avisos_fonte"] = []
    geom = D["meta"].get("Beam geometry", "").lower()
    mono = "monochromator" in geom
    if abs(lam - KALFA1[fonte]) <= TOL_FONTE:
        l_ = KALFA1[fonte]
        A["lambda"] = l_
        A["radiacao"] = (f"{fonte} Kα<sub>1</sub> · λ = {num(l_, 4)}{NBSP}Å" if mono else
                         f"{fonte} Kα · λ(Kα<sub>1</sub>) = {num(l_, 4)}{NBSP}Å")
    else:
        A["lambda"] = lam
        A["radiacao"] = f"{fonte} · λ medido ≈ {num(lam, 4)}{NBSP}Å"
        A["avisos_fonte"].append(
            f"A fonte informada ({fonte}) não confere com o comprimento de onda "
            f"calculado das reflexões ({num(lam, 4)}{NBSP}Å); confira a fonte usada "
            "na medição.")
    if origem == "inferida do PeakList":
        A["avisos_fonte"].append(
            f"Fonte de raios X inferida das reflexões do PeakList: {fonte} Kα.")

    # --- índices de qualidade ------------------------------------------------
    rv = D["rv"]
    A["Rp"], A["Rwp"] = rv["Rp [%]"], rv["Rwp [%]"]
    A["Rexp"], A["chi"] = rv["RExp [%]"], rv["χ [rel.]"]
    A["Rp_b"], A["Rwp_b"] = rv["Rp - Background [%]"], rv["Rwp - Background [%]"]
    A["Rexp_b"], A["chi_b"] = rv["RExp - Background [%]"], rv["χ - Background [rel.]"]
    w = 1 / np.clip(yo, 1e-9, None)
    rp = np.abs(yo - yc).sum() / yo.sum() * 100
    rwp = np.sqrt((w * (yo - yc) ** 2).sum() / (w * yo ** 2).sum()) * 100
    print(f"[verificação] Rp recalculado {rp:.3f} % (exportado {A['Rp']:.3f}); "
          f"Rwp recalculado {rwp:.3f} % (exportado {A['Rwp']:.3f})")

    # --- resíduos ------------------------------------------------------------
    d = D["dif"]
    i = int(np.argmax(np.abs(d)))
    A["dmax"] = (float(tt[i]), float(d[i]), float(abs(d[i]) / yo[i] * 100))
    faixa = tt.max() - tt.min()
    larg = 10 if faixa / 10 <= 8 else 15
    bordas = list(np.arange(tt.min(), tt.max(), larg)) + [tt.max() + 1e-9]
    reg = []
    for lo, hi in zip(bordas[:-1], bordas[1:]):
        s = (tt >= lo) & (tt < hi)
        if s.sum() > 5:
            r = np.sqrt((w[s] * d[s] ** 2).sum() / (w[s] * yo[s] ** 2).sum()) * 100
            reg.append((lo, min(hi, tt.max()), float(r)))
    A["regioes"] = reg

    fw = pk[cw].to_numpy(float)
    k = 5
    ds = np.convolve(d, np.ones(k) / k, mode="same")
    S = ds / np.sqrt(np.clip(yo, 1, None) / k)
    dist = np.min(np.abs(tt[:, None] - t2[None, :]) / (0.5 * fw[None, :]), axis=1)
    cand = [j for j in range(1, len(tt) - 1)
            if tt[j] > tt.min() + 1.5 and tt[j] < tt.max() - 0.5
            and S[j] >= 6 and dist[j] >= 6 and ds[j] >= ds[j - 1] and ds[j] >= ds[j + 1]]
    cand.sort(key=lambda j: -S[j])
    feic = []
    for j in cand:
        if all(abs(tt[j] - f[0]) > 0.5 for f in feic):
            feic.append((float(tt[j]), float(ds[j]), float(S[j])))
    A["feicoes"] = feic[:3]

    # --- reflexões -----------------------------------------------------------
    th = np.radians(t2 / 2)
    lp = (1 + np.cos(2 * th) ** 2) / (np.sin(th) ** 2 * np.cos(th))
    I = pk[cm].to_numpy(float) * pk[cf2].to_numpy(float) * lp
    A["refl"] = pd.DataFrame({
        "n": np.arange(1, len(pk) + 1),
        "ref": [ref_formula_do_pico(x)[0] for x in pk[cph]],
        "formula": [ref_formula_do_pico(x)[1] for x in pk[cph]],
        "h": pk[chh].astype(int).to_numpy(), "k": pk[ck].astype(int).to_numpy(),
        "l": pk[cl].astype(int).to_numpy(), "t2": t2, "d": pk[cd].to_numpy(float),
        "fwhm": fw, "I": 100 * I / I.max()})

    # --- fases ---------------------------------------------------------------
    fases, usou_cod, avisos_cod, cod_ausentes = [], False, [], []
    for idx, f in enumerate(D["fases"]):
        txt, html = formula_bonita(f["formula_raw"])
        key = canon(f["formula_raw"])
        manual = next((v for k_, v in INFO_FASES.items() if canon(k_) == key), {})
        nome, ref, massa_rel = "", "", get(f, "Rel. Mass")
        for _, r in D["massa"].iterrows():
            if canon(str(r.iloc[1])) == key:
                ref = str(int(r.iloc[0])) if pd.notna(r.iloc[0]) else ""
                nome = str(r.iloc[3] if pd.notna(r.iloc[3]) else r.iloc[2]
                           if pd.notna(r.iloc[2]) else "").strip("'\" ")
                massa_rel = float(r.iloc[4]) if pd.notna(r.iloc[4]) else massa_rel
        cid = cod_id(ref)
        cod, cod_ok = None, None
        if cid:
            cod = cod_consultar(cid, pasta_cache, offline)
            if cod:
                cod_ok = canon(cod.get("formula") or "") == key
                if not cod_ok:
                    avisos_cod.append(f"A fórmula do CIF do COD {cid} ({cod.get('formula')}) difere da fase "
                                      f"{txt}; grupo espacial e Z do COD não foram usados.")
            else:
                cod_ausentes.append(f"{txt} (COD {cid})")
        use = cod if (cod and cod_ok) else {}
        usou_cod = usou_cod or bool(use)
        a, b, c = (float(get(f, x)) for x in ("a [", "b [", "c ["))
        al, be, ga = (float(get(f, x)) for x in ("α", "β", "γ"))
        V = float(get(f, "V ["))
        M = massa_molar(f["formula_raw"])
        Z = manual.get("Z") or use.get("Z")
        sg_html = (esc(manual["grupo_espacial"]) if manual.get("grupo_espacial")
                   else grupo_espacial_html(use.get("hm")))
        if sg_html and not manual.get("grupo_espacial") and use.get("sg_num"):
            sg_html += f" (n.º {use['sg_num']})"
        dens = Z * M / (V * 1e-24 * 6.02214076e23) if (Z and M) else None
        atomos = f["atomos"]
        Bs = [float(get(at, "B (Debye")) for at in atomos if get(at, "B (Debye") is not None]
        orient = None
        if f["orient"]:
            o = f["orient"][0]
            r_ = float(get(o, "r", 1))
            metodo = str(get(o, "Method", "None"))
            hkl = f"({int(get(o, 'h'))}{NBSP}{int(get(o, 'k'))}{NBSP}{int(get(o, 'l'))})"
            orient = ("Sem correção (r = 1)" if metodo.lower() == "none" or abs(r_ - 1) < 1e-9
                      else f"{esc(metodo.replace('-', '–'))}, r = {num(r_, 3)}, plano {hkl}")
        fases.append(dict(
            txt=txt, html=html, key=key, ref=ref, cid=cid, cod=cod, cod_ok=cod_ok,
            nome=nome or (use.get("mineral") or use.get("nome") or ""),
            massa=massa_rel, vol=float(get(f, "Rel. Volume", np.nan)),
            a=a, b=b, c=c, al=al, be=be, ga=ga, V=V, celula=sistema_celula(a, b, c, al, be, ga),
            sg=sg_html, Z=Z, M=M, dens=dens, atomos=atomos, orient=orient,
            fonte_sg=("analista" if manual else "COD" if use else None),
            cit=citacao(use) if use else None, B_fixo=bool(Bs) and all(abs(x - 1) < 1e-9 for x in Bs)))
    if cod_ausentes:
        avisos_cod.append(
            "Dados do COD não disponíveis para " + "; ".join(cod_ausentes) +
            "; grupo espacial e Z foram omitidos quando não havia dados no cache "
            "ou em INFO_FASES.")
        print("Aviso: COD sem dados para " + "; ".join(cod_ausentes) + ".")
    A["fases"] = fases
    A["usou_cod"], A["avisos_cod"] = usou_cod, avisos_cod
    A["n_pts"], A["passo"] = len(tt), float(np.median(np.diff(tt)))
    return A


# =============================================================================
# Fontes
# =============================================================================
def registrar_fontes(pasta_dados=None):
    base = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
    raizes = [base, Path.cwd()] + ([Path(pasta_dados)] if pasta_dados else [])
    pastas = []
    for r in raizes:
        for sub in ("fonts", ".fonts", "Fonts", ""):
            pastas.append(r / sub if sub else r)
    home = Path.home()
    pastas += [home / ".fonts", home / ".local" / "share" / "fonts",
               Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts",
               Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts",
               home / "Library" / "Fonts", Path("/Library/Fonts")]
    nomes = {"Light": ["Roboto-Light.ttf", "Roboto Light.ttf", "RobotoLight.ttf"],
             "Regular": ["Roboto-Regular.ttf", "Roboto Regular.ttf", "RobotoRegular.ttf"],
             "Medium": ["Roboto-Medium.ttf", "Roboto Medium.ttf", "RobotoMedium.ttf"]}

    def achar(lista):
        alvo = {n.lower() for n in lista}
        for p in pastas:
            try:
                if p.is_dir():
                    for f in p.iterdir():
                        if f.name.lower() in alvo:
                            return f
            except OSError:
                pass
        for f in font_manager.findSystemFonts():
            if Path(f).name.lower() in alvo:
                return Path(f)
        return None

    caminhos = {k: achar(v) for k, v in nomes.items()}
    dj = Path(matplotlib.get_data_path()) / "fonts" / "ttf"
    if all(caminhos.values()):
        for k, p in caminhos.items():
            pdfmetrics.registerFont(TTFont(f"RB-{k}", str(p)))
            font_manager.fontManager.addfont(str(p))
        ent = next((e for e in font_manager.fontManager.ttflist
                    if Path(e.fname) == caminhos["Light"]), None)
        familia = ent.name if ent else "Roboto"
        peso = ent.weight if ent else 300
        print("Fonte Roboto carregada de:", caminhos["Light"].parent)
    else:
        print("Aviso: Roboto não encontrada — usando DejaVu Sans. Procurei em:",
              ", ".join(str(p) for p in pastas[:6]), "...")
        pdfmetrics.registerFont(TTFont("RB-Light", str(dj / "DejaVuSans.ttf")))
        pdfmetrics.registerFont(TTFont("RB-Regular", str(dj / "DejaVuSans.ttf")))
        pdfmetrics.registerFont(TTFont("RB-Medium", str(dj / "DejaVuSans-Bold.ttf")))
        familia, peso = "DejaVu Sans", 400
    pdfmetrics.registerFont(TTFont("RB-Chi", str(dj / "DejaVuSans.ttf")))
    pdfmetrics.registerFontFamily("RB-Light", normal="RB-Light", bold="RB-Medium",
                                  italic="RB-Light", boldItalic="RB-Medium")
    pdfmetrics.registerFontFamily("RB-Regular", normal="RB-Regular", bold="RB-Medium",
                                  italic="RB-Regular", boldItalic="RB-Medium")
    plt.rcParams.update({"font.family": familia, "font.weight": peso,
                         "axes.unicode_minus": True, "mathtext.fontset": "custom",
                         "mathtext.rm": familia, "mathtext.it": familia,
                         "mathtext.bf": familia, "mathtext.default": "regular"})


# =============================================================================
# Figuras
# =============================================================================
def _estilo_eixo(ax, grade=True):
    for s in ax.spines.values():
        s.set_color(COR["cinza_claro"])
        s.set_linewidth(0.6)
    ax.tick_params(colors=COR["cinza"], labelsize=6.5, width=0.5, length=2.5)
    ax.xaxis.label.set_color(COR["cinza"])
    ax.yaxis.label.set_color(COR["cinza"])
    if grade:
        ax.grid(True, color=COR["linha"], lw=0.5)
        ax.set_axisbelow(True)


def _fmt_milhar(x, _):
    return f"{int(x):d}".replace("-", MINUS)


def _desenhar_padrao(fig, spec, D, A, xlim, ymax=None, legenda=True):
    tt, yo, yc, bg, d = D["tt"], D["yo"], D["yc"], D["bg"], D["dif"]
    s = (tt >= xlim[0]) & (tt <= xlim[1])
    gs = GridSpecFromSubplotSpec(3, 1, subplot_spec=spec, height_ratios=[8, 0.9, 2.6], hspace=0.06)
    ax = fig.add_subplot(gs[0])
    axt = fig.add_subplot(gs[1], sharex=ax)
    axd = fig.add_subplot(gs[2], sharex=ax)

    ax.plot(tt[s], yo[s], ls="none", marker="o", ms=1.7, mfc="none", mec=COR["obs"], mew=0.35)
    ax.plot(tt[s], yc[s], color=COR["vermelho"], lw=0.9)
    ax.plot(tt[s], bg[s], color=COR["cinza"], lw=0.7, ls=(0, (4, 2)))
    if ymax is None:
        ymax = float(np.max(yo[s])) * 1.06
    ax.set_ylim(0, ymax)
    ax.set_xlim(*xlim)
    ax.set_ylabel("Intensidade [cts]", fontsize=7)
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(_fmt_milhar))
    plt.setp(ax.get_xticklabels(), visible=False)
    _estilo_eixo(ax)

    handles = [Line2D([], [], ls="none", marker="o", ms=3, mfc="none", mec=COR["obs"], mew=0.5),
               Line2D([], [], color=COR["vermelho"], lw=1.1),
               Line2D([], [], color=COR["cinza"], lw=0.8, ls=(0, (4, 2)))]
    nomes = ["Y$_{obs}$", "Y$_{calc}$", "Fundo"]
    for i, f in enumerate(A["fases"]):
        cor = PALETA_FASES[i % len(PALETA_FASES)]
        r = A["refl"][A["refl"]["formula"].map(canon) == f["key"]]
        pos = r["t2"].to_numpy()
        pos = pos[(pos >= xlim[0]) & (pos <= xlim[1])]
        # cada fase numa faixa vertical própria quando há mais de uma
        n = len(A["fases"])
        y0, y1 = 1 - (i + 1) / n + 0.08 / n, 1 - i / n - 0.08 / n
        axt.vlines(pos, y0, y1, color=cor, lw=0.7)
        handles.append(Line2D([], [], ls="none", marker="|", ms=7, mec=cor, mew=0.9))
        nomes.append("Bragg – " + re.sub(r"(\d+)", r"$_{\1}$", f["txt"]))
    axt.set_ylim(0, 1)
    axt.set_yticks([])
    axt.set_xlim(*xlim)
    plt.setp(axt.get_xticklabels(), visible=False)
    _estilo_eixo(axt, grade=False)
    if legenda:
        leg = ax.legend(handles, nomes, loc="upper right", fontsize=6.8, frameon=True,
                        ncol=min(len(nomes), 6), handlelength=1.4, columnspacing=1.0,
                        borderpad=0.4, framealpha=0.95, edgecolor=COR["linha"])
        for t in leg.get_texts():
            t.set_color(COR["cinza"])

    axd.plot(tt[s], d[s], color=COR["azul"], lw=0.6)
    axd.axhline(0, color=COR["cinza_claro"], lw=0.5)
    lim = float(np.max(np.abs(d[s]))) * 1.12
    axd.set_ylim(-lim, lim)
    axd.set_ylabel("Y$_{obs}$ − Y$_{calc}$", fontsize=7)
    axd.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(_fmt_milhar))
    axd.set_xlabel("2θ [°]", fontsize=7)
    axd.yaxis.set_major_locator(matplotlib.ticker.MaxNLocator(3, symmetric=True))
    _estilo_eixo(axd)
    return ax


def fig_padrao_completo(D, A):
    fig = plt.figure(figsize=(6.7, 4.15))
    gs = GridSpec(1, 1, figure=fig, left=0.085, right=0.985, top=0.985, bottom=0.095)
    _desenhar_padrao(fig, gs[0], D, A, (D["tt"].min(), D["tt"].max()))
    return fig


def _ymax_zoom(D, xlim):
    """Trunca a escala só se um pico dominar (> 3× o terceiro maior)."""
    tt, yo, yc, bg = D["tt"], D["yo"], D["yc"], D["bg"]
    s = (tt >= xlim[0]) & (tt <= xlim[1])
    y = (yc - bg)[s]
    loc = np.where((y[1:-1] > y[:-2]) & (y[1:-1] >= y[2:]))[0] + 1
    alt = np.sort(y[loc])[::-1]
    base = float(np.median(bg[s]))
    mx = float(np.max(yo[s]))
    if len(alt) >= 3 and (mx - base) > 3 * alt[2]:
        return base + float(alt[2]) * 1.2, True
    return mx * 1.06, False


def zooms_auto(D):
    lo, hi = float(D["tt"].min()), float(D["tt"].max())
    corte = round((lo + 0.6 * (hi - lo)) / 5) * 5
    return [(lo, corte), (corte, hi)]


def fig_zooms(D, A):
    zooms = zooms_auto(D)
    n = len(zooms)
    fig = plt.figure(figsize=(6.7, 2.55 * n + 0.2))
    gs = GridSpec(n, 1, figure=fig, left=0.085, right=0.985, top=0.99, bottom=0.07, hspace=0.32)
    info = []
    for i, z in enumerate(zooms):
        ymax, trunc = _ymax_zoom(D, z)
        # eixo y a partir de 0, mas aproxima do fundo para enxergar picos fracos
        _desenhar_padrao(fig, gs[i], D, A, z, ymax=ymax, legenda=(i == 0))
        info.append((z, ymax, trunc))
    return fig, info


def fig_diagnostico(D, A):
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(6.7, 2.5),
                                 gridspec_kw=dict(left=0.075, right=0.985, top=0.93,
                                                  bottom=0.2, wspace=0.28))
    reg = A["regioes"]
    x = np.arange(len(reg))
    a1.bar(x, [r[2] for r in reg], color=COR["cinza_claro"], width=0.7)
    a1.axhline(A["Rwp"], color=COR["vermelho"], lw=0.9, ls=(0, (4, 2)),
               label=f"R$_{{wp}}$ global = {num(A['Rwp'], 1)} %")
    lg1 = a1.legend(fontsize=6.3, frameon=False, loc="upper left")
    for t in lg1.get_texts():
        t.set_color(COR["vermelho"])
    for xi, r in zip(x, reg):
        a1.text(xi, r[2] / 2, num(r[2], 1), ha="center", va="center", fontsize=6.2, color=COR["preto"])
    a1.set_xticks(x)
    a1.set_xticklabels([f"{lo:.0f}–{hi:.0f}" for lo, hi, _ in reg], fontsize=5.8,
                       rotation=0 if len(reg) <= 8 else 35, ha="center" if len(reg) <= 8 else "right")
    a1.set_xlabel("Região angular 2θ [°]", fontsize=7)
    a1.set_ylabel("R$_{wp}$ local [%]", fontsize=7)
    a1.set_ylim(0, max(r[2] for r in reg) * 1.3)
    _estilo_eixo(a1)
    a1.grid(False, axis="x")
    a1.set_title("(a) Ajuste por região angular", fontsize=7.2, color=COR["cinza"], loc="left")

    pk = A["refl"]
    th = np.radians(np.linspace(D["tt"].min(), D["tt"].max(), 300) / 2)
    I = D["instr"]
    U, V, W = (float(get(I, k, 0)) for k in ("U", "V", "W"))
    inst = np.sqrt(np.clip(U * np.tan(th) ** 2 + V * np.tan(th) + W, 0, None))
    a2.plot(np.degrees(th) * 2, inst, color=COR["cinza"], lw=1.0, label="Função instrumental")
    multi = len(A["fases"]) > 1
    if multi:
        for i, f in enumerate(A["fases"]):
            r = pk[pk["formula"].map(canon) == f["key"]]
            a2.plot(r["t2"], r["fwhm"], ls="none", marker="o", ms=2.4,
                    color=PALETA_FASES[i % len(PALETA_FASES)],
                    label=re.sub(r"(\d+)", r"$_{\1}$", f["txt"]))
    else:
        a2.plot(pk["t2"], pk["fwhm"], ls="none", marker="o", ms=2.6, color=COR["vermelho"],
                label="Reflexões do modelo")
    a2.set_xlabel("2θ [°]", fontsize=7)
    a2.set_ylabel("FWHM [°]", fontsize=7)
    a2.set_ylim(0, None)
    a2.set_title("(b) Largura de pico × função instrumental", fontsize=7.2, color=COR["cinza"], loc="left")
    leg = a2.legend(fontsize=6.3, frameon=False, loc="upper left")
    for t in leg.get_texts():
        t.set_color(COR["cinza"])
    _estilo_eixo(a2)
    return fig


def fig_pizza(A):
    fases = A["fases"]
    vals = [max(f["massa"], 0) for f in fases]
    fig, ax = plt.subplots(figsize=(2.45, 2.45))
    fig.subplots_adjust(left=0.02, right=0.98, top=0.98, bottom=0.02)
    cores = [PALETA_FASES[i % len(PALETA_FASES)] for i in range(len(fases))]
    ws, _ = ax.pie(vals, colors=cores, startangle=90, counterclock=False,
                   wedgeprops=dict(width=0.42, edgecolor="white", linewidth=1.4))
    for w, f, v in zip(ws, fases, vals):
        if v < 4:
            continue
        ang = np.radians((w.theta1 + w.theta2) / 2)
        r = 0.79
        ax.text(r * np.cos(ang), r * np.sin(ang), f"{num(v, 1)}%", ha="center", va="center",
                fontsize=6.8, color="white", fontweight="medium")
    ax.text(0, 0.07, "Massa", ha="center", va="center", fontsize=7.5, color=COR["cinza"])
    ax.text(0, -0.11, "relativa", ha="center", va="center", fontsize=7.5, color=COR["cinza"])
    ax.set_aspect("equal")
    return fig


def para_png(fig, dpi=300):
    w, h = fig.get_size_inches()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf, w, h


# =============================================================================
# PDF
# =============================================================================
PAGE_W, PAGE_H = A4
MARG_E = MARG_D = 20 * mm
LARG = PAGE_W - MARG_E - MARG_D


def estilos():
    S = {}
    S["corpo"] = ParagraphStyle("corpo", fontName="RB-Light", fontSize=9.5, leading=13.6, textColor=colors.black)
    S["peq"] = ParagraphStyle("peq", parent=S["corpo"], fontSize=8, leading=11, textColor=colors.HexColor(COR["cinza"]))
    S["titulo"] = ParagraphStyle("titulo", fontName="RB-Light", fontSize=36, leading=40, textColor=colors.HexColor(COR["vermelho"]))
    S["sub"] = ParagraphStyle("sub", parent=S["corpo"], fontSize=10.5, leading=14)
    S["h0"] = ParagraphStyle("h0", fontName="RB-Light", fontSize=27, leading=31, textColor=colors.HexColor(COR["vermelho"]), spaceBefore=6, spaceAfter=8)
    S["h1"] = ParagraphStyle("h1", fontName="RB-Light", fontSize=21, leading=25, textColor=colors.HexColor(COR["vermelho"]), spaceBefore=0, spaceAfter=10)
    S["h2"] = ParagraphStyle("h2", fontName="RB-Light", fontSize=13.5, leading=17, textColor=colors.HexColor(COR["vermelho"]), spaceBefore=10, spaceAfter=4)
    S["cap"] = ParagraphStyle("cap", parent=S["peq"], spaceBefore=3, spaceAfter=6)
    S["rot"] = ParagraphStyle("rot", parent=S["corpo"], fontSize=9, leading=12.2, textColor=colors.HexColor(COR["cinza"]))
    S["val"] = ParagraphStyle("val", parent=S["corpo"], fontName="RB-Regular", fontSize=9, leading=12.2)
    S["th"] = ParagraphStyle("th", fontName="RB-Medium", fontSize=8, leading=10, textColor=colors.HexColor(COR["cinza"]))
    S["th_c"] = ParagraphStyle("th_c", parent=S["th"], fontSize=7.2, leading=8.6)
    S["td"] = ParagraphStyle("td", fontName="RB-Light", fontSize=8.6, leading=11, textColor=colors.black)
    S["td_c"] = ParagraphStyle("td_c", fontName="RB-Light", fontSize=7.8, leading=9.2, textColor=colors.black)
    S["kpi_n"] = ParagraphStyle("kpi_n", fontName="RB-Light", fontSize=22, leading=25, textColor=colors.HexColor(COR["vermelho"]))
    S["kpi_t"] = ParagraphStyle("kpi_t", fontName="RB-Light", fontSize=8, leading=10, textColor=colors.HexColor(COR["cinza"]))
    S["bul"] = ParagraphStyle("bul", parent=S["corpo"], fontSize=8.8, leading=12, leftIndent=11, bulletIndent=0, spaceAfter=2.5)
    return S


def kv(linhas, S, larg_rot=46 * mm, larg_total=LARG):
    dados = [[Paragraph(r, S["rot"]), Paragraph(v, S["val"])] for r, v in linhas if v not in (None, "")]
    if not dados:
        return Spacer(1, 1)
    t = Table(dados, colWidths=[larg_rot, larg_total - larg_rot], hAlign="LEFT")
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                           ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                           ("TOPPADDING", (0, 0), (-1, -1), 1.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.2)]))
    return t


def tabela(cab, linhas, larguras, S, alinh=None, fonte=None, pad=2.4, cab_style=None):
    alinh = alinh or ["LEFT"] * len(cab)
    est = lambda al, base: ParagraphStyle("x", parent=base, alignment={"LEFT": 0, "CENTER": 1, "RIGHT": 2}[al])
    th = [est(a, cab_style or S["th"]) for a in alinh]
    td = [est(a, S["td"] if fonte is None else fonte) for a in alinh]
    dados = [[Paragraph(c, th[i]) for i, c in enumerate(cab)]]
    for r in linhas:
        dados.append([Paragraph(str(c), td[i]) for i, c in enumerate(r)])
    t = Table(dados, colWidths=larguras, repeatRows=1, hAlign="LEFT")
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                           ("TOPPADDING", (0, 0), (-1, -1), pad), ("BOTTOMPADDING", (0, 0), (-1, -1), pad),
                           ("LINEBELOW", (0, 0), (-1, 0), 0.9, colors.HexColor(COR["vermelho"])),
                           ("LINEBELOW", (0, 1), (-1, -1), 0.4, colors.HexColor(COR["linha"]))]))
    return t


def duas_colunas(esq, dir_, gap=10 * mm, w_esq=None):
    w1 = w_esq or (LARG - gap) / 2
    w2 = LARG - gap - w1
    t = Table([[esq, "", dir_]], colWidths=[w1, gap, w2], hAlign="LEFT")
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                           ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                           ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    return t


class CanvasNumerado(rl_canvas.Canvas):
    info = {}

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self._paginas = []

    def showPage(self):
        self._paginas.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        n = len(self._paginas)
        for est in self._paginas:
            self.__dict__.update(est)
            self._decorar(n)
            super().showPage()
        super().save()

    def _decorar(self, total):
        i = self.info
        larg = 46 * mm if self._pageNumber == 1 else 36 * mm
        alt = larg * i["logo_h"] / i["logo_w"]
        self.drawImage(ImageReader(io.BytesIO(i["logo"])), PAGE_W - MARG_D - larg,
                       PAGE_H - 13 * mm - alt, larg, alt, mask="auto")
        self.setStrokeColor(colors.HexColor(COR["cinza"]))
        self.setLineWidth(0.5)
        self.line(MARG_E, 18 * mm, PAGE_W - MARG_D, 18 * mm)
        self.setFillColor(colors.HexColor(COR["cinza"]))
        self.setFont("RB-Light", 8)
        self.drawString(MARG_E, 13.2 * mm, i["rodape1"])
        if i["rodape2"]:
            self.drawString(MARG_E, 9.2 * mm, i["rodape2"])
        self.drawRightString(PAGE_W - MARG_D, 13.2 * mm, f"Página {self._pageNumber} de {total}")


def imagem(png, w_in, h_in, largura=LARG):
    return Image(png, width=largura, height=largura * h_in / w_in)


def ref_html(f):
    """Referência da fase: 'COD 9009666' (link) ou o número bruto."""
    if f["cid"]:
        return (f'<a href="{COD_URL_PAG.format(f["cid"])}" color="{COR["vermelho"]}">COD {f["cid"]}</a>')
    return esc(f["ref"]) if f["ref"] else "–"


def blocos_medicao(meta, info, A, D, S):
    m = meta
    g = lambda k: m.get(k)

    def com_unid(k, unid, nd=None):
        v = numero_em(g(k)) if g(k) else None
        if v is None:
            return None
        return f"{num(v, nd if nd is not None else (0 if v == int(v) else 3))}{NBSP}{unid}"

    instr = info.get("instrumento") or m.get("instrumento")
    if instr and info.get("serial"):
        instr = f"{esc(instr)} (S/N {esc(info['serial'])})"
    elif instr:
        instr = esc(instr)

    kv_, kvh = (LARG - 10 * mm) / 2, 40 * mm
    esq = []

    def secao(titulo, linhas):
        linhas = [(r, v) for r, v in linhas if v not in (None, "")]
        if not linhas:
            return
        if esq:
            esq.append(Spacer(1, 3 * mm))
        esq.append(Paragraph(titulo, S["h2"]))
        esq.append(kv(linhas, S, kvh, kv_))

    tubo = g("Source type")
    tubo = esc(tr(tubo)) if tubo else None
    if tubo and g("Source name"):
        tubo += f" ({esc(g('Source name'))})"
    kv_gerador = None
    if g("Generator voltage") and g("Generator current"):
        v_, c_ = numero_em(g("Generator voltage")), numero_em(g("Generator current"))
        if v_ and c_:
            kv_gerador = f"{num(v_, 0)}{NBSP}kV / {num(c_, 0)}{NBSP}mA"
    stage = g("Sample stage")
    stage_cfg = g("Sample stage configuration")
    stage_txt = None
    if stage:
        stage_txt = esc(tr_estagio(stage))
    if stage_cfg:
        cfg = esc(tr_estagio(stage_cfg))
        stage_txt = f"{stage_txt} ({cfg})" if stage_txt else cfg
    if stage_txt and g("Sample spinning") == "True" and not re.search(r"rota|spin", stage_txt, re.I):
        stage_txt += ", em rotação"
    secao("Instrumento e fonte", [
        ("Instrumento", instr),
        ("Geometria", esc(tr(g("Beam geometry"))) if g("Beam geometry") else None),
        ("Raio do goniômetro", com_unid("Goniometer radius", "mm")),
        ("Fonte de raios X", A.get("fonte")),
        ("Tubo de raios X", tubo),
        ("Radiação", A["radiacao"]),
        ("Gerador", kv_gerador),
        ("Porta-amostra", stage_txt),
        ("Último alinhamento", (f"{data_pt(g('Last instrument alignment')):%d/%m/%Y}"
                                if g("Last instrument alignment") and data_pt(g("Last instrument alignment")) else None)),
    ])

    anti = None
    if g("Anti-scatter slit mode"):
        anti = esc(tr(g("Anti-scatter slit mode")))
        if g("Anti-scatter slit opening") and numero_em(g("Anti-scatter slit opening")):
            anti += f", {num(numero_em(g('Anti-scatter slit opening')), 3)}{NBSP}mm"
    elif g("Anti-scatter slit opening") and numero_em(g("Anti-scatter slit opening")):
        anti = f"{num(numero_em(g('Anti-scatter slit opening')), 3)}{NBSP}mm"
    filtro = g("Attenuator / filter type")
    if filtro:
        mf = re.match(r"^\s*(\w+)\s+K\s*beta\s+filter\s*\((\w+)\)", filtro, re.I)
        filtro = (f"Filtro de Kβ ({mf.group(1)}) para tubo de {mf.group(2)}" if mf else
                  None if "slit" in filtro.lower() else esc(filtro))
    det = None
    if g("Detector model"):
        det = esc(g("Detector model"))
        extra = []
        if g("Detector sensor material"):
            extra.append(f"sensor de {esc(g('Detector sensor material'))}")
        if g("Detector mode"):
            extra.append(esc(tr(g("Detector mode"))))
        if g("Detector region"):
            extra.append(esc(tr(g("Detector region"))))
        if g("Detector S/N"):
            extra.append(f"S/N {esc(g('Detector S/N'))}")
        det += f" ({', '.join(extra)})" if extra else ""
    soller = None
    if g("Primary Soller slit"):
        soller = f"{g('Primary Soller slit').replace('.', ',')} (primário)"
        if g("Secondary Soller slit"):
            soller += f" / {g('Secondary Soller slit').replace('.', ',')} (secundário)"
    fluo = g("Fluorescence reduction")
    div = None
    if g("Divergence slit mode"):
        div = esc(tr(g("Divergence slit mode")))
        if g("Divergence slit opening"):
            div += f", {str(g('Divergence slit opening')).replace('.', ',')}"
    elif g("Divergence slit opening"):
        div = str(g("Divergence slit opening")).replace(".", ",")
        if g("Divergence slit type"):
            div += f" (tipo {esc(g('Divergence slit type'))})"
    secao("Óptica e detector", [
        ("Fenda de divergência", div),
        ("Máscara do feixe", com_unid("Beam mask opening", "mm", 1)),
        ("Fenda anti-espalhamento", anti),
        ("Colimadores Soller", soller),
        ("Filtro", filtro),
        ("Detector", det),
        ("Redução de fluorescência", ("Ativada" if fluo == "True" else "Desativada") if fluo else None),
    ])

    ini, fim = data_pt(g("Scan start time") or ""), data_pt(g("Scan end time") or "")
    quando = None
    if ini and fim:
        quando = (f"{ini:%d/%m/%Y}, {ini:%H:%M}–{fim:%H:%M}" if ini.date() == fim.date()
                  else f"{ini:%d/%m/%Y %H:%M} – {fim:%d/%m/%Y %H:%M}")
    passo = numero_em(g("2θ step size"))
    ini_pos, fim_pos = numero_em(g("Start position")), numero_em(g("End position"))
    if ini_pos is not None and fim_pos is not None:
        faixa_2t = f"{num(ini_pos, 2)}° – {num(fim_pos, 2)}°"
    else:
        faixa_2t = f"{num(D['tt'].min(), 2)}° – {num(D['tt'].max(), 2)}°"
    passo_dados = float(np.median(np.diff(D["tt"])))
    passo_exibir = passo if passo is not None else passo_dados
    soft = g("Software")
    if soft:
        soft = re.sub(r"\s+ReleaseVersion$", "", soft)
    analise = f"{info['software']} {info['versao']}" if info.get("software") else None
    secao("Varredura", [
        ("Tipo / modo", (f"{tr(g('Scan type'))} · {tr(g('Scan mode'))}" if g("Scan type") and g("Scan mode") else
                         tr(g("Scan type")) if g("Scan type") else None)),
        ("Faixa 2θ", faixa_2t),
        ("Passo", f"{num(passo_exibir, 3)}°"),
        ("Tempo por passo", f"{num(numero_em(g('Time per step')), 1)}{NBSP}s" if g("Time per step") else None),
        ("Duração", duracao_pt(g("Scan duration")) if g("Scan duration") else None),
        ("Data", quando),
        ("Status", esc(tr(g("Scan status"))) if g("Scan status") else None),
        ("Usuário", esc(g("User name")) if g("User name") else None),
        ("Aquisição", esc(soft) if soft else None),
        ("Análise", esc(analise) if analise else None),
    ])

    # --- direita: amostra/experimento, modelo e parâmetros globais ------------
    amostra = []
    if g("Experiment name"):
        amostra.append(("Experimento", esc(g("Experiment name"))))
    z = numero_em(g("Sample Z"))
    if z is not None:
        amostra.append(("Altura da amostra (Z)", f"{num(z, 3)}{NBSP}mm"))
    omega = numero_em(g("Omega offset"))
    if omega is not None:
        amostra.append(("Offset de ômega", f"{num(omega, 2)}{NBSP}°"))
    temp = numero_em(g("Sample temperature"))
    if temp is not None:
        amostra.append(("Temperatura (setpoint)", f"{num(temp, 0 if temp == int(temp) else 2)}{NBSP}°C"))
    dr = []
    if amostra:
        dr += [Paragraph("Amostra e experimento", S["h2"]),
               kv(amostra, S, 30 * mm, kv_),
               Spacer(1, 3 * mm)]
    dr.append(Paragraph("Modelo e dados do refinamento", S["h2"]))
    linhas = [("Arquivo de dados", esc(D["arquivo_dados"])),
              ("Perfil de pico", "pseudo-Voigt (η = fração lorentziana)"),
              ("Função de peso", "1/Y<sub>obs</sub>")]
    if D.get("fundo_coef") is not None:
        linhas.append(("Fundo", f"Polinômio de {esc(D.get('fundo_tipo', 'Chebyshev'))}, "
                                f"{len(D['fundo_coef'])} coeficientes"))
    # passo/faixa só aparecem quando diferem da aquisição (senão são redundantes)
    if passo and abs(passo - A["passo"]) > 1e-3:
        linhas.append(("Passo dos dados", f"{num(A['passo'], 3)}° (difere da aquisição)"))
    ini_, fim_ = numero_em(g("Start position")), numero_em(g("End position"))
    if ini_ is not None and fim_ is not None and (abs(ini_ - D["tt"].min()) > 0.05 or abs(fim_ - D["tt"].max()) > 0.05):
        linhas.append(("Faixa 2θ dos dados", f"{num(D['tt'].min(), 2)}° – {num(D['tt'].max(), 2)}° (difere da aquisição)"))
    dr.append(kv(linhas, S, 30 * mm, kv_))
    I = D["instr"]
    U, V, W = (float(get(I, k, 0)) for k in ("U", "V", "W"))
    dr.append(Spacer(1, 3 * mm))
    dr.append(Paragraph("Parâmetros globais", S["h2"]))
    dr.append(kv([
        ("Erro de zero", f"{num(float(get(I, 'Zero-Point', 0)), 3)}°"),
        ("Erro de altura (rel.)", sci(float(get(I, "Height Error", 0)))),
        ("Caglioti U", num(U, 5)), ("Caglioti V", num(V, 5)), ("Caglioti W", num(W, 5)),
        ("η instrumental", num(float(get(I, "η", 0)), 2)),
    ], S, 30 * mm, kv_))
    dr.append(Spacer(1, 4))
    dr.append(Paragraph("Caglioti: FWHM<super>2</super> = U·tan<super>2</super>θ + V·tanθ + W (graus<super>2</super>).", S["peq"]))
    return esq, dr


# =============================================================================
# Montagem do relatório
# =============================================================================
def construir(D, A, saida, logo_png, figs):
    S = estilos()
    meta, info = D["meta"], D["info"]
    scan = meta.get("Scan name") or D["arquivo_dados"].rsplit(".", 1)[0]
    amostra = meta.get("Sample name", "")
    agora = dt.datetime.now()
    analista = info.get("analista") or ANALISTA

    from PIL import Image as PILImage
    lw, lh = PILImage.open(io.BytesIO(logo_png)).size
    CanvasNumerado.info = dict(logo=logo_png, logo_w=lw, logo_h=lh,
                               rodape1=f"{scan} · gerado em {agora:%d/%m/%Y %H:%M}",
                               rodape2=analista)
    story = []
    fases = A["fases"]
    f0 = fases[0]
    multi = len(fases) > 1

    # ------------------------------------------------------------ página 1
    instr = info.get("instrumento") or meta.get("instrumento") or ""
    sub = esc(instr) + (f" · S/N {esc(info['serial'])}" if info.get("serial") else "")
    soft_an = f"{info['software']} {info['versao']}" if info.get("software") else ""
    if soft_an:
        sub += ("&nbsp;" * 6 if sub else "") + esc(soft_an)
    story += [Spacer(1, 12 * mm), Paragraph(TITULO_RELATORIO, S["titulo"]),
              Paragraph(sub or "Análise de difração de raios X", S["sub"]),
              Spacer(1, 7 * mm), Paragraph(esc(scan), S["h0"])]

    ini = data_pt(meta.get("Scan start time", ""))
    linhas = []
    if amostra and amostra.strip().lower() not in ("sample", "amostra"):
        linhas.append(("Amostra", formula_livre_html(amostra)))
    if meta.get("Sample comment"):
        linhas.append(("Comentário", esc(meta["Sample comment"])))
    linhas += [("Data da medição", f"{ini:%d/%m/%Y}" if ini else None),
               ("Radiação", A["radiacao"]),
               ("Faixa angular", f"{num(D['tt'].min(), 1)}° – {num(D['tt'].max(), 1)}° (2θ)"),
               ("Método de análise", "Refinamento de Rietveld")]
    story.append(kv(linhas, S, 38 * mm))
    story.append(Spacer(1, 6 * mm))

    cartao = lambda n, t: [Paragraph(n, S["kpi_n"]), Paragraph(t, S["kpi_t"])]
    kp = [cartao(f"{len(fases)}", "fases cristalinas" if multi else "fase cristalina"),
          cartao(f"{num(A['Rwp'], 1)}{NBSP}%", "R<sub>wp</sub> (ponderado)"),
          cartao(num(A["chi"], 2), "χ = R<sub>wp</sub>/R<sub>exp</sub> (GoF)"),
          cartao(f"{num(A['Rp'], 1)}{NBSP}%", "R<sub>p</sub>")]
    tk = Table([kp], colWidths=[LARG / 4] * 4)
    tk.setStyle(TableStyle([("LINEABOVE", (0, 0), (-1, 0), 0.8, colors.HexColor(COR["cinza_claro"])),
                            ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.HexColor(COR["cinza_claro"])),
                            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                            ("LEFTPADDING", (0, 0), (-1, -1), 0), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [tk, Spacer(1, 5 * mm)]

    if not multi:
        resumo = (f"Foi identificada uma única fase cristalina, <b>{f0['html']}</b>"
                  + (f" ({esc(f0['nome'])})" if f0["nome"] else "")
                  + f", com célula {f0['celula']}" + (f" e grupo espacial {f0['sg']}" if f0["sg"] else "")
                  + f". Parâmetros de rede: a = {num(f0['a'], 4)}, b = {num(f0['b'], 4)}, "
                    f"c = {num(f0['c'], 4)}{NBSP}Å (V = {num(f0['V'], 2)}{NBSP}Å<super>3</super>). ")
    else:
        lista = "; ".join(f"<b>{f['html']}</b> ({num(f['massa'], 1)}{NBSP}%)" for f in fases)
        resumo = f"Foram identificadas {len(fases)} fases cristalinas (fração em massa): {lista}. "
    resumo += (f"O refinamento do padrão ({num(D['tt'].min(), 0)}–{num(D['tt'].max(), 0)}° 2θ) convergiu para "
               f"R<sub>wp</sub> = {num(A['Rwp'], 1)}{NBSP}% e χ = {num(A['chi'], 2)}.")
    story += [Paragraph(resumo, S["corpo"]), Spacer(1, 4 * mm)]
    png, w_in, h_in = figs["fig1"]
    story += [imagem(png, w_in, h_in),
              Paragraph("Figura 1 — Padrão de difração: intensidade observada (Y<sub>obs</sub>), calculada "
                        "(Y<sub>calc</sub>), fundo, posições de Bragg e diferença.", S["cap"]),
              PageBreak()]

    # ------------------------------------------------------------ condições
    story.append(Paragraph("Condições de medição", S["h1"]))
    esq, dr = blocos_medicao(meta, info, A, D, S)
    wcol = (LARG - 10 * mm) / 2
    tbl = lambda itens: Table([[x] for x in itens], colWidths=[wcol],
                              style=TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                                ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    story += [duas_colunas(tbl(esq), tbl(dr)), PageBreak()]

    # ------------------------------------------------------------ resultados
    story.append(Paragraph("Resultados do refinamento", S["h1"]))
    story.append(Paragraph("Fases e quantificação", S["h2"]))
    if multi:
        linhas = []
        for i, f in enumerate(fases):
            cor = PALETA_FASES[i % len(PALETA_FASES)]
            sub_ = " · ".join(x for x in (esc(f["nome"]) if f["nome"] else "", ref_html(f)) if x)
            linhas.append([f'<font name="RB-Chi" color="{cor}">■</font> <b>{f["html"]}</b><br/>'
                           f'<font size="7.5" color="{COR["cinza"]}">{sub_}</font>',
                           num(f["massa"], 1), num(f["vol"], 1) if np.isfinite(f["vol"]) else "–"])
        tf = tabela(["Fase", "Massa rel. [%]", "Vol. rel. [%]"], linhas, [58 * mm, 21 * mm, 21 * mm], S,
                    ["LEFT", "RIGHT", "RIGHT"])
        pz, wp, hp = figs["pizza"]
        pie = Image(pz, width=52 * mm, height=52 * mm * hp / wp)
        story.append(duas_colunas(tf, pie, gap=12 * mm, w_esq=100 * mm))
    else:
        linhas = [[f"<b>{f['html']}</b>", esc(f["nome"]) or "–", ref_html(f), num(f["massa"], 1),
                   num(f["vol"], 1) if np.isfinite(f["vol"]) else "–"] for f in fases]
        story.append(tabela(["Fase", "Nome", "Referência", "Massa rel. [%]", "Volume rel. [%]"], linhas,
                            [30 * mm, 56 * mm, 26 * mm, 29 * mm, 29 * mm], S,
                            ["LEFT", "LEFT", "LEFT", "RIGHT", "RIGHT"]))
    story.append(Paragraph("Frações relativas à soma das fases incluídas no modelo; conteúdo amorfo e fases não "
                           "modeladas não são contabilizados.", S["cap"]))

    for f in fases:
        ref_txt = None
        if f["cid"]:
            ref_txt = ref_html(f) + (f" — {esc(f['cit'])}" if f.get("cit") else "")
        elif f["ref"]:
            ref_txt = esc(f["ref"])
        esq_kv = kv([
            ("Célula", f["celula"].capitalize()),
            ("Grupo espacial", f["sg"]),
            ("a, b, c [Å]", f"{num(f['a'], 4)} · {num(f['b'], 4)} · {num(f['c'], 4)}"),
            ("α, β, γ [°]", f"{num(f['al'], 1)} · {num(f['be'], 1)} · {num(f['ga'], 1)}"),
            ("Volume [Å<super>3</super>]", num(f["V"], 2)),
            ("Z", num(f["Z"], 0) if f["Z"] else None),
            ("Densidade calc.", (f"{num(f['dens'], 3)}{NBSP}g/cm<super>3</super>") if f["dens"] else None),
            ("Orient. preferencial", f["orient"]),
            ("Referência", ref_txt),
        ], S, 33 * mm, 88 * mm)
        linhas = [[esc(at.get("Atom", "")), num(float(get(at, "x")), 5), num(float(get(at, "y")), 5),
                   num(float(get(at, "z")), 5), num(float(get(at, "Occupation")), 2),
                   num(float(get(at, "B (Debye")), 3)] for at in f["atomos"]]
        dir_at = tabela(["Átomo", "x", "y", "z", "Ocup.", "B<sub>iso</sub>"], linhas,
                        [12 * mm, 16 * mm, 16 * mm, 16 * mm, 11 * mm, 12 * mm], S,
                        ["LEFT", "RIGHT", "RIGHT", "RIGHT", "RIGHT", "RIGHT"])
        bloco = [Paragraph(f"Estrutura cristalina — {f['html']}", S["h2"]),
                 duas_colunas(esq_kv, dir_at, gap=3 * mm, w_esq=88 * mm)]
        story.append(CondPageBreak(62 * mm))
        story.append(KeepTogether(bloco))
    story.append(Paragraph("Átomos: coordenadas fracionárias (x, y, z), ocupação e deslocamento isotrópico "
                           "B<sub>iso</sub> (Debye–Waller, Å<super>2</super>).", S["cap"]))

    story.append(CondPageBreak(70 * mm))
    qa = [Paragraph("Qualidade do ajuste", S["h2"])]
    linhas = [["R<sub>p</sub> [%]", num(A["Rp"], 2), num(A["Rp_b"], 2)],
              ["R<sub>wp</sub> [%]", num(A["Rwp"], 2), num(A["Rwp_b"], 2)],
              ["R<sub>exp</sub> [%]", num(A["Rexp"], 2), num(A["Rexp_b"], 2)],
              ["χ = R<sub>wp</sub>/R<sub>exp</sub> (GoF)", num(A["chi"], 2), num(A["chi_b"], 2)],
              ["χ<super>2</super>", num(A["chi"] ** 2, 2), num(A["chi_b"] ** 2, 2)]]
    qa.append(tabela(["Índice", "Convencional", "Sem fundo"], linhas, [70 * mm, 40 * mm, 40 * mm], S,
                     ["LEFT", "RIGHT", "RIGHT"]))
    qa.append(Paragraph("“Sem fundo”: índices calculados sobre o sinal após a subtração do fundo (mais sensíveis ao "
                        "ajuste dos picos). χ e χ<super>2</super> são adimensionais.", S["cap"]))
    story.append(KeepTogether(qa))

    # --- observações automáticas ---------------------------------------------
    obs = ["A fração de 100 % decorre do modelo monofásico adotado; fases minoritárias não incluídas no "
           "modelo e conteúdo amorfo não são contabilizados." if not multi else
           "As frações referem-se apenas às fases incluídas no modelo; conteúdo amorfo não é contabilizado."]
    t_d, d_d, p_d = A["dmax"]
    regs = [r[2] for r in A["regioes"]]
    obs.append(f"Maior desvio entre Y<sub>obs</sub> e Y<sub>calc</sub>: {num(abs(d_d), 0)} cts em 2θ = "
               f"{num(t_d, 2)}° ({num(p_d, 1)} % da intensidade observada nesse ponto). "
               f"R<sub>wp</sub> local entre {num(min(regs), 1)} % e {num(max(regs), 1)} % nas regiões angulares.")
    if A["feicoes"]:
        txt = "; ".join(f"2θ ≈ {num(t, 1)}° (+{num(a, 0)} cts, S/N ≈ {num(s, 0)})" for t, a, s in A["feicoes"])
        obs.append(f"Feição(ões) residual(is) positiva(s) fora das reflexões calculadas: {txt}. Sem atribuição; "
                   "verificar se corresponde a fase minoritária ou a ruído/fundo.")
    else:
        obs.append("Não há picos residuais isolados (S/N ≥ 6) afastados das reflexões calculadas.")
    ps = meta.get("2θ step size")
    pm = numero_em(ps) if ps else None
    if pm and abs(pm - A["passo"]) > 1e-3:
        obs.append(f"Os dados refinados têm passo de {num(A['passo'], 3)}°, enquanto a aquisição registra "
                   f"{num(pm, 3)}°; confirme que o arquivo refinado ({esc(D['arquivo_dados'])}) corresponde "
                   "à medição descrita.")
    fx = [f["html"] for f in fases if f["B_fixo"]]
    if fx:
        obs.append(f"B<sub>iso</sub> = 1,000 Å<super>2</super> em todos os átomos de {', '.join(fx)}: valor típico de "
                   "parâmetro não refinado (padrão).")
    origem = []
    if any(f["fonte_sg"] == "COD" for f in fases):
        origem.append("Grupo espacial, Z, nomes e referências estruturais vêm das entradas do COD indicadas")
    if any(f["fonte_sg"] == "analista" for f in fases):
        origem.append("Grupo espacial/Z de algumas fases foram definidos pelo analista")
    if origem:
        obs.append("; ".join(origem) + " (não constam nas exportações do XRDanalysis PRO).")
    obs += A["avisos_cod"]
    obs += A.get("avisos_fonte", [])
    if not D.get("fontes_meta"):
        obs.append("O arquivo HDF5 não foi fornecido; as condições completas do instrumento, "
                   "a data da varredura e outros metadados experimentais podem estar ausentes.")
    obs.append("As exportações não trazem incertezas (desvios-padrão) nem indicam quais parâmetros estruturais "
               "foram refinados.")
    ob = [Paragraph("Observações", S["h2"])] + [Paragraph(o, S["bul"], bulletText="•") for o in obs]
    story.append(KeepTogether(ob))

    if info["textos"]:
        cm = [Paragraph("Comentários", S["h2"])] + [Paragraph(esc(t), S["bul"], bulletText="•") for t in info["textos"]]
        story.append(KeepTogether(cm))

    # ------------------------------------------------------------ diagnóstico
    png2, w2, h2, info2 = figs["fig2"]
    cap = " ".join(f"({chr(97 + i)}) {z[0]:g}–{z[1]:g}°"
                   + (f": eixo de intensidade limitado a {num(ym, 0)} cts (picos mais intensos truncados)." if tr_ else ".")
                   for i, (z, ym, tr_) in enumerate(info2))
    png3, w3, h3 = figs["fig3"]
    diagnostico = [
        Paragraph("Diagnóstico do ajuste", S["h1"]),
        imagem(png2, w2, h2),
        Paragraph(f"Figura 2 — Ampliações do padrão. {cap}", S["cap"]),
        imagem(png3, w3, h3),
        Paragraph("Figura 3 — (a) R<sub>wp</sub> calculado por região angular; (b) FWHM das "
                  "reflexões do modelo comparada à função instrumental (Caglioti).", S["cap"]),
    ]
    story.append(KeepTogether(diagnostico))

    # ------------------------------------------------------------ apêndice
    story.append(PageBreak())
    story.append(Paragraph("Apêndice — Lista de reflexões", S["h1"]))
    refl = A["refl"]
    total = len(refl)
    refl = refl[refl["I"] >= LIMITE_APENDICE]
    omit = total - len(refl)
    n = len(refl)
    if n <= 98:
        por_pag = 49
    else:
        paginas = (n + 97) // 98
        por_pag = min(49, max(1, (n + 2 * paginas - 1) // (2 * paginas)))
    cab = ["#", "h k l", "2θ [°]", "d [Å]", "I/I<sub>1</sub> [%]", "FWHM [°]"]
    larg = [8 * mm, 15 * mm, 16 * mm, 15 * mm, 15 * mm, 12 * mm]
    alin = ["LEFT", "CENTER", "RIGHT", "RIGHT", "RIGHT", "RIGHT"]
    if multi:
        cab.insert(1, "Fase")
        larg = [7 * mm, 13 * mm, 14 * mm, 13.5 * mm, 12 * mm, 11 * mm, 11 * mm]
        alin.insert(1, "LEFT")

    def linhas_refl(df):
        out = []
        for _, r in df.iterrows():
            hkl = f"{r['h']}{NBSP}{r['k']}{NBSP}{r['l']}".replace("-", MINUS)
            cel = [str(int(r["n"])), hkl, num(r["t2"], 3), num(r["d"], 4), num(r["I"], 1), num(r["fwhm"], 3)]
            if multi:
                cel.insert(1, formula_bonita(r["formula"])[1])
            out.append(cel)
        return out

    def tab_refl(df):
        return tabela(cab, linhas_refl(df), larg, S, alin, fonte=S["td_c"], pad=1.3, cab_style=S["th_c"])

    for ini_ in range(0, n, 2 * por_pag):
        bloco = refl.iloc[ini_:ini_ + 2 * por_pag]
        e, d_ = bloco.iloc[:por_pag], bloco.iloc[por_pag:]
        story.append(duas_colunas(tab_refl(e), tab_refl(d_) if len(d_) else Spacer(1, 1), gap=8 * mm))
        if ini_ + 2 * por_pag < n:
            story.append(PageBreak())
    nota = (f"{n} reflexões listadas de {total}. " if omit else f"Total de {total} reflexões. ")
    if omit:
        nota += f"Omitidas {omit} com I/I<sub>1</sub> &lt; {num(LIMITE_APENDICE, 1)} % (lista completa na exportação PeakList). "
    nota += ("I/I<sub>1</sub>: intensidade relativa calculada como m·F<super>2</super>·(1+cos<super>2</super>2θ)/"
             "(sin<super>2</super>θ·cosθ), normalizada para a reflexão mais intensa (todas as fases); não inclui "
             "correção de orientação preferencial.")
    story.append(KeepTogether([Spacer(1, 3 * mm), Paragraph(nota, S["cap"])]))

    doc = BaseDocTemplate(str(saida), pagesize=A4, leftMargin=MARG_E, rightMargin=MARG_D,
                          topMargin=32 * mm, bottomMargin=24 * mm,
                          title=f"{TITULO_RELATORIO} — {scan}", author=analista or "Anton Paar",
                          subject="Refinamento de Rietveld")
    frame = Frame(MARG_E, 24 * mm, LARG, PAGE_H - 32 * mm - 24 * mm, id="f",
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    doc.addPageTemplates([PageTemplate(id="p", frames=[frame])])
    doc.build(story, canvasmaker=CanvasNumerado)


# =============================================================================
def executar_geracao(pasta, saida=None, hdf5=None, info=None, offline=False,
                     analista=None, comentarios=None, fonte=None, sem_hdf5=False):
    """Gera o relatório completo e devolve o caminho do PDF."""
    pasta = Path(pasta)
    if not pasta.is_dir():
        sys.exit(f"Erro: pasta '{pasta}' não encontrada.")
    registrar_fontes(pasta)
    D = carregar(pasta, hdf5, info, fonte=fonte, sem_hdf5=sem_hdf5)
    if analista is not None:
        D["info"]["analista"] = analista.strip() or None
    if comentarios is not None:
        if isinstance(comentarios, str):
            comentarios = re.split(r"\n\s*\n", comentarios)
        D["info"]["textos"] = [str(t).strip() for t in comentarios if str(t).strip()]
    print("Metadados HDF5:", ", ".join(D["fontes_meta"]) or "nenhum (somente planilhas)",
          "| Info.txt:", D["info_arq"] or "não encontrado")
    base = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
    A = analisar(D, base, offline)
    print(f"Fonte de raios X: {A['fonte']} ({A['fonte_origem']})")

    figs = {"fig1": para_png(fig_padrao_completo(D, A))}
    f2, info2 = fig_zooms(D, A)
    w2, h2 = f2.get_size_inches()
    figs["fig2"] = (para_png(f2)[0], w2, h2, info2)
    figs["fig3"] = para_png(fig_diagnostico(D, A))
    if len(A["fases"]) > 1:
        figs["pizza"] = para_png(fig_pizza(A))

    scan = D["meta"].get("Scan name") or D["arquivo_dados"].rsplit(".", 1)[0]
    saida = Path(saida) if saida else pasta / f"Relatorio_DRX_{so_ascii(scan) or 'DRX'}.pdf"
    construir(D, A, saida, base64.b64decode(LOGO_B64), figs)
    print("Relatório gerado:", saida)
    return saida


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Gera relatório PDF de DRX a partir das exportações do XRDanalysis PRO.")
    ap.add_argument("pasta", nargs="?", default=".", help="pasta com os .xlsx (+ .hdf5 e Info.txt opcionais)")
    ap.add_argument("-o", "--saida", help="PDF de saída")
    grupo_h5 = ap.add_mutually_exclusive_group()
    grupo_h5.add_argument("--hdf5", help="arquivo .hdf5 da medição (padrão: achado na pasta)")
    grupo_h5.add_argument("--sem-hdf5", action="store_true",
                          help="ignorar HDF5 e usar somente as planilhas (+ Info.txt)")
    ap.add_argument("--info", help="caminho do Info.txt (padrão: Info.txt na pasta)")
    ap.add_argument("--fonte", choices=FONTES_RAIO_X,
                    help="fonte de raios X usada (padrão: HDF5, Info.txt ou inferência)")
    ap.add_argument("--offline", action="store_true", help="não consultar o COD pela internet")
    a = ap.parse_args()
    executar_geracao(a.pasta, a.saida, a.hdf5, a.info, a.offline,
                     fonte=a.fonte, sem_hdf5=a.sem_hdf5)


LOGO_B64 = (
    "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"
    "wS1bMk45RfbXij0qyX73bs9nnxHoh8HfUghGqW/NmsCaNRnTp6OjoPfCOTCWe/XV7qVL1WFZ160jarjdTc8/X/rLXybzQUsJlHKf"
    "r+GJJ5I8ViBESqllZGRfdJG6IPxUkWO2qu/bJ8PhJAv6M2ZyXnrLLdkXXig5T67+llFfz01T1/X0n9BJCQCh7dvDbndy4Ysqnqv4"
    "2msdJ53UZR3VejIe2rhRtbMfyElTvX03DQQCQgh6tLVY6ahe2pxX+aAYi9qgnIf37Wv65z+9X37pX7HClPLoPqjWhD/BueowTazW"
    "jAkT8q67zjF5smPGDC2uvna0QmWCIbxxCire9aQqVQSDQbfb7XQ6DcMIhUI+n49z3qYRbayLV0tLSyAQcMSq2VAqOdcKCx0nnujf"
    "ubN9tL3vq69yL7+8r+sopdsTyeVTCCFUPfT0Cn7CmDRNy9Chw59/fudFF6UzZFlKousRvz9SU5Nxyin9NRhaScSm557jnPeHzrwd"
    "fo6aZvj9ns8/z5g+HR0F3W7DpfakZJ93nrWoyGxqSqLrpSQECGl6/vnC73+fFRR01fOvIiaa/v73SHNzcs8CoZRznnfddVp+Phbf"
    "R47l7gwQ2r5dcJ5crKwyYzLPOQeS6t+ozJvwzp3ddTynIoD8/uRvEecaY4W33qpiBZPTcr5VqziA3i/3zX4mpTqsNN1uzvcyKSWE"
    "VLOz1SPkX7XKv2aN85VXAuvXq1p8nfmgWuWT5FwJGwCwDx9unzw5+7zzcq+4Qh80KHbceFiDcNXE6WjCCdq5njjnPp8vGAwGAoGm"
    "pibld+Kcc84JIZRSTdNUZFr8ZxGTYZzzmpqa8vLyw0SClAXf/W7z66/Hj04KQQFcCxaU/fa3sSYMfRqHw0ETMxeUlOqOa1D1HrNn"
    "zx70858fePzxdCZNSQkAwu3ux1suoVRGIt7Fi7uxkeIxH6VpMoDm554rvfde0l/7mPeCuUR1Pez1hior7ZMnJ2e9Aee2sWMzpk+P"
    "Zj929bCZc8ZYYM+euoceGvrUU9IwEj+xUl6s4ObNtY8/TihNUg4SQjQt9+qroTXGCecFcqww6+tV6ZQuP4yESMPQ7XbryJGqEnoy"
    "z7KU7vff79Y6RmaKW7Ou66WlSbZe0DQRDEYqKwd4zYk+IKWCwWCCU5AQ0ivK97XWKCeapuamceBA04svepcu9X31lREKqVp87Eg+"
    "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"
    "dP/982fOWKZHy1qesXke+x9+OP7BD2op0focAAAA4EZszrWUu2CPaP473xn77GdzL7xARObiz61EmSFRtGxPFGlNa28scV2RSKSe"
    "4T+wicGd80KhIKVc7GLHORFF7rpr4a2vnD/c48kPDeX/8R99PT1aKbYja1PlcrlYLNbZkU8pZVmWaVCxniiltQk/o5/4RPKxxywh"
    "dH3VsMUwx7kj5f4vfrHpt39b2zbbqJ7sAAAAADftxe2GP+LCtijO888+e/nd7774K7+Se+EFzrngXDuO6SLAhHCkVFpHTp8+/N3v"
    "bkYlquYykoii0Wg0GsUav+2itbYsK5VKpdNpqlrjR0Tezs7g4KCqfuuVYkTF8+epsqVqp70WIkqn06lUqp72fWarWDQaXf+3lJJZ"
    "1sRDD419/vNCiLVV6hgjzh0puz7zmdZ//a+14yBHAQAAAOysKOWWm8pvvDHysY+d/5mfST7+OGeMc05KLdQWONdaO1JG3/KW3q9/"
    "vf/pp2P33beBe6JWv5aNRCKoSm0vx3FyudySm5Qixhre+U5VtfLNtESf/7u/I6V27HapbDZbf4lJCNHS0rLuHEWWNfvooyMPPMA5"
    "11JS/acxY4xzKWXX5z7X/uCDWNcHAAAAsFE26KJKKTMwlBxn8stfnvyv/7U4PCyILCG06TDGmBllo4lCR460P/hg46/+KvN6aem4"
    "3s1jrnc7OjquXr2KNLWd2Z3zZDLZ2dm5mJqUYpwHBgfZ0jOKE+VfesmemvK0te20lujmyU9NTdXZE9IMKV7nRinHIcua++Y3X//w"
    "h7kQbE2nL2NEZOpR7f/xPy78rAEAAADAhlzZ3ugDaE2VRuep733v/M/8zNDHP14aHq5p0Ke1dhwnePx4zxe+cOzFF+Pvex/zehfb"
    "923VVXI0Gg0EAohS28XUBmdmZqq73pmL++CpU4GeHlXTm1up3DPPLJxmO0ypVEqn0/W3QY/FYp51LKur1KMuv+99rFw21d36Y6sm"
    "Isvq/epX2x98kEyOwupWAAAAgJ0QpbSUxBgJUTh79vX3vOfiL/9y9rnnBOcLa5DcECWlp6mp6w/+oP/MmcTHP878frNCaYsHg5rr"
    "+Pb2dsdxsF1quzDGtNbZbJbc7VKca6W8nZ2+Q4dU1c4oZlmO1qm/+zvaYdOlzNOemZmpp3efeclSyubm5sWXXCfHISFMPYqEYJyv"
    "IVJyrpViWu//sz9rvv9+LSVt4WcWAAAAAIhSq4YorZkQcnZ27NOffvX222e+9S3GuTCtqysxyZFSxGLtv//7A6+80v6pT4l43LQd"
    "Y9t3VdfU1IQ+ftsbpWzbnpqaWnKrUsRY8PhxRksKUIyoeOkSmV12O+wtMxOl6szkjLFQKLSmuFZTj2JrqUcxIZRSIh4/8MgjTb/9"
    "29pxsK4PAAAAYAdEKdOUWQhiLPmNb7x6110jn/ucyuctIUhrbZoEMOZISZbV8uEPHzlzpuu//BcrkSCTvrZvy7u56o3FYg0NDShM"
    "bSOtdSaToapdQ+ZCP3rffdzt4kikpeRExddeK168SHV0yduyJ88YK5VK2WzWsqzrPitTkjLdI6n+jVJKrbsexYSQUnoTiYPf+lb8"
    "gx80rf9QjwIAAADY5ihlVvQxIfLPP3/xHe94/UMfyp8967Es08ecMUacK6W0Uk3vfOfRH//4wF/+ZXBwcKEH+g5YX2TW+DU3NyNH"
    "bSPLsgqFQrFYrA4cRBR805u4368rvyWtuWUVJycLL720kC52jLm5OdMGvb5YpILBoM/nq/McJSlJiOQjj6ynHmVZjpSBvr7DP/hB"
    "5Od/XjsObe0yWgAAAABEqZWu8JRiQqhsduxTn3r1rrvmvv99zpg7LYpZllRKKhU+derg44/3fec7wVOnSCm9kz4UNwmqs7PT5/Nh"
    "jd82ptlMJmMKU9XTpazGxtDtty9GKfOWEZXeeIPW0fhuM0+h8fHx+nv3cc4bGxupno1SplGhEBOf//wbH/4w43wN9SjGmGU5jhN7"
    "29uO/MM/BAYGCH3PAQAAALY5SpmPyRkjzuf+9m/Pnz599Q//UOVyFuekta5si7Idx3/48P4vfan/xz9ueNe7yPRA55ztvA/FPR5P"
    "R0cH1vhtb6CqnS4lJXEe/aVfqu48oZViROnvf98seNshT75UKiWTSV737iPOeWtr63XToFnZqMvl0QceGPnEJ5jWCytm6/seRCQd"
    "p/Wf/bND3/uep7NzRx0xAAAAgJsySillPiYvXbx4+Vd/9dJ735s/f14IwRjTlWYAjpTM7+968MH+M2daP/pRZlm6Er125kU8EbW3"
    "t6P5xDa+BUKImZkZWjqTl4j8R48u6TyhNSPK/b//58zNEW1/S3TzJEdHRx3HqSdKMcaUUg0NDWai1Oo/aEwIUmro/vtHH3qIm9Ww"
    "9b1exrlWSmvd+eCDPV/7Gg8GdU1PeQAAAADY0ihVKSvpUmni858/f/p08tvf5pxzxshsi2LMUYp5PK0f+tDAK690fOYznra2rR8V"
    "tVbm2j0ajcbjcRSmtvFdSKVSK0yXetOblkyX0poR6VKpeOHCTohSJhrNzMzUH8KllKYktdqXSEmcOzMzF37hF6a+8Q3L4zE9Wup6"
    "SkJIpTwtLQe//e32z3zGfLrBUI8CAAAA2K4otfCpNueZf/iH1+66a/gTn3Dm5y0hFhqdW5ZUSmkdu/vuwz/4wf6vf93X20tSLnyy"
    "vkt0d3dzfHK/rZkknU5TzXSp7u7a6VKcSynT3//+Gla7bQ7zPOfn5+fm5urp3We+xOPxmN591+Q4pup74W1vm/8//8cSQtt2nQeR"
    "hHCkDN166+Gnnmr4lV9Z6MyBTwcAAAAAtidKVaZC2RMTw//qX114+9szzzxjWdbCij4hNJHtOIEjR3q//vUjTz4ZvvNO8yUkxO5a"
    "U9Tc3NzU1FRdGIGtjFKO40xPT9eee0TBEyeWrPHjXBPln3+eGNvekGAKmGNjY/VP5nUcxzTfp2ttlFKKLCv1P//nqz/3c7mzZy0h"
    "3F7w13lwIUhrKWXi/vv7K00mzBwCnF0AAAAA2xGlKlOh5h577LWf//mJP/kT0lpYljYL4RhzpBShUNenPtV/5kzThz5keqDvxgs4"
    "U1I4cOAAFvht41tgOk/UTJeK/JN/smS6lFKMyL56VaZSZjfRNj7nQqEwOTlZf0mKMWY67698f62J8+Rf/dWl97ynPDEh6s9RluVI"
    "ycPhA3/yJz1f/SqPREz/dJxUAAAAANsUpbQmzksXLlx+3/su/cZvFF57TRBprZXjEJFUSivV/N73Hv3JTzr+4A+sRGJxW9QuZC7f"
    "4/F4c3OzNFu/YGsJIYrFYrlcrn5XiCh4yy1Lyptac85z58/nX3yR6uknvpnZe3R01K5z9V3lNba3t9OKJSkpibHC2bPDv//7uljk"
    "nNeTo8yPm+M4kTe/+fBTT7X87u8u7KpCjgIAAADYtiilNRHNPvrohfvum//mN61QyBOJ8GBQBIMiHGahUPjUqUPf/W7vY4/5Dx/e"
    "dduiVglUPT09yFHbkkyEEKlUasl2KcaIyJNIBE+eVESLnSeEkLZdunTJPVG3/tkyxmzbnpycpPomXDHGpJTxeDwQCKxyt/T/+l/l"
    "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"
)

if __name__ == "__main__":
    main()
