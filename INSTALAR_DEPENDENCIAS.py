#!/usr/bin/env python3
"""
INSTALAR_DEPENDENCIAS.py
========================
Launcher com interface grafica (tkinter) para instalar todas as
dependencias do Separador de ASOs sem abrir o CMD.

Instala:
  - Tesseract OCR (4 estrategias em cascata)
  - Poppler (download direto)
  - Bibliotecas Python (pip)

Execute este arquivo antes de usar o SEPARACAODEASO.py.
"""

import os
import sys
import threading
import queue
import shutil
import platform
import subprocess
import tempfile
import ctypes
import time
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk

# ─────────────────────────────────────────────────────────────
#  CONFIGURACOES
# ─────────────────────────────────────────────────────────────

PASTA_APP = Path(__file__).parent

TESSERACT_URL = (
    "https://github.com/UB-Mannheim/tesseract/releases/download/"
    "v5.3.3.20231005/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
)
POPPLER_URL = (
    "https://github.com/oschwartz10612/poppler-windows/releases/"
    "download/v24.08.0-0/Release-24.08.0-0.zip"
)

TESSERACT_PATHS = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    Path(os.environ.get("LOCALAPPDATA", "")) / "Tesseract-OCR" / "tesseract.exe",
    Path(os.environ.get("APPDATA", ""))      / "Tesseract-OCR" / "tesseract.exe",
    PASTA_APP / "tesseract_bin" / "tesseract.exe",
]

LIBS_PYTHON = [
    "pdfplumber", "pypdf", "pdf2image",
    "pytesseract", "openpyxl", "pillow",
]

# ─────────────────────────────────────────────────────────────
#  CORES
# ─────────────────────────────────────────────────────────────

COR_BG       = "#0F1117"
COR_PAINEL   = "#1A1D27"
COR_BORDA    = "#2A2D3E"
COR_ENTRADA  = "#252836"
COR_TEXTO    = "#E8EAF0"
COR_DIM      = "#6B7280"
COR_PRIMARIA = "#4F8EF7"
COR_SUCESSO  = "#2ECC71"
COR_ERRO     = "#E74C3C"
COR_AVISO    = "#F39C12"
COR_HOVER    = "#3D4166"

FONTE_TITULO = ("Segoe UI", 13, "bold")
FONTE_LABEL  = ("Segoe UI", 9)
FONTE_BOLD   = ("Segoe UI", 9, "bold")
FONTE_MONO   = ("Consolas", 9)
FONTE_PEQ    = ("Segoe UI", 8)


# ─────────────────────────────────────────────────────────────
#  UTILITARIOS
# ─────────────────────────────────────────────────────────────

def eh_windows():
    return platform.system().lower() == "windows"

def tem_admin():
    if eh_windows():
        try: return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except: return False
    return os.geteuid() == 0

def run_cmd(cmd, timeout=300):
    """Executa comando sem janela, retorna (returncode, stdout, stderr)."""
    kwargs = {"capture_output": True, "text": True, "timeout": timeout}
    if eh_windows():
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = subprocess.SW_HIDE
        kwargs["startupinfo"] = si
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    try:
        r = subprocess.run(cmd, **kwargs)
        return r.returncode, r.stdout, r.stderr
    except subprocess.TimeoutExpired:
        return 1, "", "timeout"
    except FileNotFoundError:
        return 1, "", "comando nao encontrado"
    except Exception as e:
        return 1, "", str(e)

def adicionar_path_permanente(pasta: str):
    """Adiciona pasta ao PATH permanente do sistema (Windows)."""
    if not eh_windows(): return
    try:
        code, out, _ = run_cmd(
            ["powershell", "-NoProfile", "-Command",
             '[System.Environment]::GetEnvironmentVariable("Path","Machine")']
        )
        path_atual = out.strip()
        if pasta.lower() not in path_atual.lower():
            novo = f"{path_atual};{pasta}"
            run_cmd(["powershell", "-NoProfile", "-Command",
                     f'[System.Environment]::SetEnvironmentVariable("Path","{novo}","Machine")'])
        # Atualiza sessao atual
        os.environ["PATH"] = pasta + os.pathsep + os.environ.get("PATH", "")
    except Exception:
        pass

def baixar_arquivo(url, destino, cb_progresso=None):
    """Download com callback de progresso (0.0 a 1.0)."""
    def reporthook(blocos, tam, total):
        if cb_progresso and total > 0:
            cb_progresso(min(blocos * tam / total, 1.0))
    try:
        urllib.request.urlretrieve(url, destino, reporthook)
        return True
    except Exception as e:
        return False

def tesseract_encontrar():
    """Retorna caminho do tesseract.exe ou None."""
    exe = shutil.which("tesseract")
    if exe: return exe
    for p in TESSERACT_PATHS:
        if p.exists(): return str(p)
    # Busca no registro
    if eh_windows():
        try:
            import winreg
            chave = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Tesseract-OCR")
            pasta, _ = winreg.QueryValueEx(chave, "InstallDir")
            t = Path(pasta) / "tesseract.exe"
            if t.exists(): return str(t)
        except Exception:
            pass
    return None

def poppler_encontrar():
    """Retorna pasta bin do Poppler ou None."""
    exe = shutil.which("pdftoppm") or shutil.which("pdftoppm.exe")
    if exe: return os.path.dirname(exe)
    locais = [
        r"C:\ProgramData\chocolatey\lib\poppler\tools\Library\bin",
        r"C:\ProgramData\chocolatey\bin",
        r"C:\poppler\bin",
        os.path.join(os.environ.get("LOCALAPPDATA",""), "poppler", "bin"),
    ]
    for loc in locais:
        if os.path.isfile(os.path.join(loc, "pdftoppm.exe")): return loc
    # Busca generica em C:\poppler
    base = Path(r"C:\poppler")
    if base.exists():
        for p in base.rglob("pdftoppm.exe"):
            return str(p.parent)
    return None


# ─────────────────────────────────────────────────────────────
#  LOGICA DE INSTALACAO
# ─────────────────────────────────────────────────────────────

class InstaladorLogica:
    """Toda a logica de instalacao, sem UI. Comunica via callbacks."""

    def __init__(self, log_cb, prog_cb, status_cb):
        self.log    = log_cb      # log(msg, tag)
        self.prog   = prog_cb     # prog(valor 0.0-1.0, texto)
        self.status = status_cb   # status(etapa, ok/erro/aviso)

    # ── Tesseract ────────────────────────────────────────────

    def instalar_tesseract(self):
        self.log("━━ Verificando Tesseract OCR...", "titulo")
        exe = tesseract_encontrar()
        if exe:
            self.log(f"✔ Tesseract ja instalado: {exe}", "ok")
            self._configurar_pytesseract(exe)
            self.status("tesseract", "ok")
            return True

        self.log("Tesseract nao encontrado. Iniciando instalacao...", "aviso")

        # Estrategia 1 — winget
        self.log("[1/4] Tentando winget...", "info")
        if shutil.which("winget"):
            code, _, _ = run_cmd([
                "winget", "install", "--id", "UB-Mannheim.TesseractOCR",
                "--silent", "--accept-package-agreements",
                "--accept-source-agreements"
            ], timeout=180)
            time.sleep(2)
            exe = tesseract_encontrar()
            if exe:
                self.log(f"✔ Tesseract instalado via winget: {exe}", "ok")
                self._configurar_pytesseract(exe)
                self.status("tesseract", "ok")
                adicionar_path_permanente(str(Path(exe).parent))
                return True
            self.log("winget nao conseguiu instalar. Tentando proxima...", "aviso")
        else:
            self.log("winget nao disponivel.", "aviso")

        # Estrategia 2 — Chocolatey
        self.log("[2/4] Tentando Chocolatey...", "info")
        if shutil.which("choco"):
            code, _, _ = run_cmd(["choco", "install", "tesseract", "-y", "--no-progress"], timeout=180)
            time.sleep(1)
            exe = tesseract_encontrar()
            if exe:
                self.log(f"✔ Tesseract instalado via Chocolatey: {exe}", "ok")
                self._configurar_pytesseract(exe)
                self.status("tesseract", "ok")
                adicionar_path_permanente(str(Path(exe).parent))
                return True
            self.log("Chocolatey nao conseguiu instalar. Tentando proxima...", "aviso")
        else:
            self.log("Chocolatey nao instalado.", "aviso")

        # Estrategia 3 — Download direto + instalador silencioso (requer admin)
        self.log("[3/4] Download direto do instalador oficial...", "info")
        with tempfile.TemporaryDirectory() as tmp:
            dest = Path(tmp) / "tess_setup.exe"
            self.log(f"Baixando instalador ({TESSERACT_URL[-40:]})...", "info")

            def prog_tess(p):
                self.prog(p * 0.8, f"Baixando Tesseract... {int(p*100)}%")

            ok_dl = baixar_arquivo(TESSERACT_URL, dest, prog_tess)
            if ok_dl and dest.exists():
                self.prog(0.85, "Instalando Tesseract silenciosamente...")
                self.log("Instalando... (pode demorar 1-2 minutos)", "info")
                install_dir = Path(r"C:\Program Files\Tesseract-OCR")
                code, _, err = run_cmd(
                    [str(dest), "/S", f"/D={install_dir}"], timeout=300
                )
                time.sleep(2)
                exe = tesseract_encontrar()
                if exe:
                    self.log(f"✔ Tesseract instalado com sucesso: {exe}", "ok")
                    self._configurar_pytesseract(exe)
                    self.status("tesseract", "ok")
                    adicionar_path_permanente(str(Path(exe).parent))
                    return True
                self.log(f"Instalador direto falhou (cod={code}). Tentando portatil...", "aviso")
            else:
                self.log("Download falhou. Tentando instalacao portatil...", "aviso")

        # Estrategia 4 — Portatil (sem admin, extrai na pasta do app)
        self.log("[4/4] Instalacao portatil (sem necessidade de admin)...", "info")
        pasta_portatil = PASTA_APP / "tesseract_bin"
        pasta_portatil.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory() as tmp:
            dest = Path(tmp) / "tess_setup.exe"

            def prog_port(p):
                self.prog(p * 0.8, f"Baixando Tesseract portatil... {int(p*100)}%")

            ok_dl = baixar_arquivo(TESSERACT_URL, dest, prog_port)
            if ok_dl and dest.exists():
                self.prog(0.85, "Extraindo Tesseract portatil...")
                self.log(f"Extraindo em {pasta_portatil}...", "info")
                code, _, _ = run_cmd(
                    [str(dest), "/S", f"/D={pasta_portatil}"], timeout=300
                )
                time.sleep(2)
                # Busca recursiva
                encontrados = list(pasta_portatil.rglob("tesseract.exe"))
                if encontrados:
                    exe = str(encontrados[0])
                    self.log(f"✔ Tesseract portatil disponivel: {exe}", "ok")
                    self._configurar_pytesseract(exe)
                    self.status("tesseract", "ok")
                    return True

        self.log("✖ Todas as estrategias falharam para o Tesseract.", "erro")
        self.log("  Instale manualmente: github.com/UB-Mannheim/tesseract/wiki", "erro")
        self.status("tesseract", "erro")
        return False

    def _configurar_pytesseract(self, caminho_exe: str):
        """Cria config_tesseract.py ao lado do app para persistir a configuracao."""
        config = PASTA_APP / "config_tesseract.py"
        config.write_text(
            f"# Gerado automaticamente pelo instalador\n"
            f"import pytesseract\n"
            f"pytesseract.pytesseract.tesseract_cmd = r\"{caminho_exe}\"\n",
            encoding="utf-8"
        )
        pasta = str(Path(caminho_exe).parent)
        if pasta not in os.environ.get("PATH", ""):
            os.environ["PATH"] = pasta + os.pathsep + os.environ.get("PATH", "")

    # ── Poppler ──────────────────────────────────────────────

    def instalar_poppler(self):
        self.log("━━ Verificando Poppler...", "titulo")
        bin_dir = poppler_encontrar()
        if bin_dir:
            self.log(f"✔ Poppler ja instalado: {bin_dir}", "ok")
            self.status("poppler", "ok")
            return True

        self.log("Poppler nao encontrado. Baixando...", "aviso")
        destino_zip = Path(r"C:\poppler\poppler.zip")
        destino_zip.parent.mkdir(parents=True, exist_ok=True)

        def prog_popp(p):
            self.prog(p * 0.85, f"Baixando Poppler... {int(p*100)}%")

        ok_dl = baixar_arquivo(POPPLER_URL, destino_zip, prog_popp)
        if not ok_dl or not destino_zip.exists():
            self.log("✖ Download do Poppler falhou.", "erro")
            self.status("poppler", "erro")
            return False

        self.prog(0.9, "Extraindo Poppler...")
        self.log("Extraindo Poppler...", "info")
        code, _, err = run_cmd([
            "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-Command",
            f"Expand-Archive -Path '{destino_zip}' -DestinationPath 'C:\\poppler\\' -Force"
        ], timeout=120)
        try: destino_zip.unlink()
        except: pass

        bin_dir = poppler_encontrar()
        if bin_dir:
            self.log(f"✔ Poppler instalado: {bin_dir}", "ok")
            self.status("poppler", "ok")
            adicionar_path_permanente(bin_dir)
            return True

        self.log("✖ Poppler nao encontrado apos extracao.", "erro")
        self.status("poppler", "erro")
        return False

    # ── Bibliotecas Python ───────────────────────────────────

    def instalar_libs(self):
        self.log("━━ Instalando bibliotecas Python...", "titulo")
        faltando = []
        for lib in LIBS_PYTHON:
            nome_import = lib.replace("-", "_")
            if nome_import == "pillow": nome_import = "PIL"
            if nome_import == "pypdf":  nome_import = "pypdf"
            try:
                __import__(nome_import)
            except ImportError:
                faltando.append(lib)

        if not faltando:
            self.log("✔ Todas as bibliotecas ja instaladas.", "ok")
            self.status("libs", "ok")
            return True

        self.log(f"Instalando: {', '.join(faltando)}", "info")
        self.prog(0.3, "Instalando bibliotecas Python...")

        code, out, err = run_cmd(
            [sys.executable, "-m", "pip", "install", "--upgrade"] + faltando,
            timeout=300
        )
        if code == 0:
            self.log("✔ Bibliotecas instaladas com sucesso.", "ok")
            self.status("libs", "ok")
            return True
        else:
            self.log(f"✖ Erro ao instalar bibliotecas: {err[:300]}", "erro")
            self.status("libs", "erro")
            return False

    # ── Orquestrador ─────────────────────────────────────────

    def instalar_tudo(self):
        resultados = {}

        self.prog(0.0, "Iniciando instalacao...")
        self.log(f"Sistema: {platform.system()} {platform.release()}", "info")
        self.log(f"Python: {platform.python_version()}", "info")
        self.log(f"Admin: {'Sim' if tem_admin() else 'Nao'}", "info")
        self.log("", "")

        # Libs Python (rapido, faz primeiro)
        self.prog(0.05, "Verificando bibliotecas Python...")
        resultados["libs"] = self.instalar_libs()
        self.prog(0.3, "Bibliotecas concluidas.")
        self.log("", "")

        # Tesseract
        self.prog(0.35, "Verificando Tesseract...")
        resultados["tesseract"] = self.instalar_tesseract()
        self.prog(0.75, "Tesseract concluido.")
        self.log("", "")

        # Poppler
        self.prog(0.80, "Verificando Poppler...")
        resultados["poppler"] = self.instalar_poppler()
        self.prog(1.0, "Concluido!")
        self.log("", "")

        # Resumo final
        self.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━", "titulo")
        ok_total = all(resultados.values())
        for item, ok in resultados.items():
            simbolo = "✔" if ok else "✖"
            tag     = "ok" if ok else "erro"
            self.log(f"  {simbolo} {item.capitalize()}", tag)

        if ok_total:
            self.log("", "")
            self.log("Tudo instalado! Pode abrir o Separador de ASOs.", "ok")
        else:
            self.log("", "")
            self.log("Alguns itens falharam. Veja os detalhes acima.", "aviso")

        return ok_total


# ─────────────────────────────────────────────────────────────
#  INTERFACE GRAFICA
# ─────────────────────────────────────────────────────────────

class Launcher(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Instalador — Separador de ASOs")
        self.geometry("680x560")
        self.minsize(600, 500)
        self.configure(bg=COR_BG)
        self.resizable(True, True)

        self.fila = queue.Queue()
        self._rodando = False

        self._construir_ui()
        self.after(100, self._processar_fila)
        self.after(300, self._verificar_status_inicial)

    # ── UI ───────────────────────────────────────────────────

    def _construir_ui(self):
        # Cabecalho
        hdr = tk.Frame(self, bg=COR_PAINEL, pady=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⚙  Instalador de Dependências",
                 font=FONTE_TITULO, bg=COR_PAINEL, fg=COR_TEXTO).pack(side="left", padx=20)
        tk.Label(hdr, text="Separador de ASOs",
                 font=FONTE_PEQ, bg=COR_PAINEL, fg=COR_DIM).pack(side="right", padx=20)

        # Status das dependencias
        status_frame = tk.Frame(self, bg=COR_BG, pady=12)
        status_frame.pack(fill="x", padx=20)

        tk.Label(status_frame, text="STATUS DAS DEPENDÊNCIAS",
                 font=("Segoe UI", 8, "bold"), bg=COR_BG, fg=COR_DIM).pack(anchor="w")

        cards = tk.Frame(status_frame, bg=COR_BG)
        cards.pack(fill="x", pady=(6, 0))

        self.status_labels = {}
        for nome, rotulo in [("libs", "Bibliotecas Python"),
                               ("tesseract", "Tesseract OCR"),
                               ("poppler", "Poppler")]:
            card = tk.Frame(cards, bg=COR_PAINEL, padx=14, pady=10,
                            highlightthickness=1, highlightbackground=COR_BORDA)
            card.pack(side="left", fill="x", expand=True, padx=(0, 8))
            tk.Label(card, text=rotulo, font=FONTE_PEQ,
                     bg=COR_PAINEL, fg=COR_DIM).pack(anchor="w")
            lbl = tk.Label(card, text="● Verificando...",
                           font=FONTE_BOLD, bg=COR_PAINEL, fg=COR_AVISO)
            lbl.pack(anchor="w", pady=(3, 0))
            self.status_labels[nome] = lbl

        # Barra de progresso
        prog_frame = tk.Frame(self, bg=COR_BG, padx=20)
        prog_frame.pack(fill="x", pady=(4, 0))

        self.lbl_prog_texto = tk.Label(prog_frame, text="Aguardando...",
                                        font=FONTE_PEQ, bg=COR_BG, fg=COR_DIM)
        self.lbl_prog_texto.pack(anchor="w")

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Inst.Horizontal.TProgressbar",
                         troughcolor=COR_ENTRADA, background=COR_PRIMARIA,
                         bordercolor=COR_ENTRADA, lightcolor=COR_PRIMARIA,
                         darkcolor=COR_PRIMARIA, thickness=12)

        self.progressbar = ttk.Progressbar(prog_frame, style="Inst.Horizontal.TProgressbar",
                                            mode="determinate", maximum=100)
        self.progressbar.pack(fill="x", pady=(4, 0))

        # Log
        tk.Frame(self, bg=COR_BORDA, height=1).pack(fill="x", padx=20, pady=10)
        tk.Label(self, text="LOG DE INSTALAÇÃO",
                 font=("Segoe UI", 8, "bold"), bg=COR_BG, fg=COR_DIM).pack(anchor="w", padx=20)

        log_outer = tk.Frame(self, bg=COR_BORDA, padx=1, pady=1)
        log_outer.pack(fill="both", expand=True, padx=20, pady=(4, 0))

        style.configure("Log.TScrollbar", troughcolor=COR_ENTRADA,
                         background=COR_BORDA, bordercolor=COR_ENTRADA,
                         arrowcolor=COR_DIM)

        self.txt_log = tk.Text(log_outer, bg=COR_ENTRADA, fg=COR_TEXTO,
                                font=FONTE_MONO, wrap="word", state="disabled",
                                relief="flat", padx=10, pady=8, spacing1=2)
        scroll = ttk.Scrollbar(log_outer, style="Log.TScrollbar",
                                command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.txt_log.pack(fill="both", expand=True)

        self.txt_log.tag_config("ok",     foreground=COR_SUCESSO)
        self.txt_log.tag_config("erro",   foreground=COR_ERRO)
        self.txt_log.tag_config("aviso",  foreground=COR_AVISO)
        self.txt_log.tag_config("info",   foreground=COR_PRIMARIA)
        self.txt_log.tag_config("titulo", foreground=COR_DIM)

        # Botoes
        rodape = tk.Frame(self, bg=COR_PAINEL, pady=12)
        rodape.pack(fill="x", side="bottom")

        self.btn_instalar = tk.Button(
            rodape, text="⬇  INSTALAR TUDO",
            font=FONTE_BOLD, bg=COR_PRIMARIA, fg="white",
            relief="flat", padx=22, pady=9, cursor="hand2",
            activebackground="#3A7BE0", activeforeground="white",
            command=self._iniciar_instalacao
        )
        self.btn_instalar.pack(side="right", padx=(0, 20))

        self.btn_abrir_app = tk.Button(
            rodape, text="▶  ABRIR SEPARADOR DE ASOs",
            font=FONTE_BOLD, bg=COR_SUCESSO, fg="white",
            relief="flat", padx=22, pady=9, cursor="hand2",
            activebackground="#27AE60", activeforeground="white",
            state="disabled",
            command=self._abrir_app
        )
        self.btn_abrir_app.pack(side="right", padx=(0, 8))

        self.lbl_admin = tk.Label(
            rodape,
            text="⚠ Sem admin — instalacao pode ser limitada" if not tem_admin() else "✔ Admin OK",
            font=FONTE_PEQ, bg=COR_PAINEL,
            fg=COR_AVISO if not tem_admin() else COR_SUCESSO
        )
        self.lbl_admin.pack(side="left", padx=20)

    # ── Fila ─────────────────────────────────────────────────

    def _processar_fila(self):
        try:
            while True:
                msg = self.fila.get_nowait()
                t = msg.get("tipo")

                if t == "log":
                    self._log(msg["texto"], msg.get("tag", ""))

                elif t == "prog":
                    v = int(msg["valor"] * 100)
                    self.progressbar["value"] = v
                    self.lbl_prog_texto.config(text=msg.get("texto", ""))

                elif t == "status":
                    nome = msg["nome"]
                    ok   = msg["ok"]
                    if nome in self.status_labels:
                        if ok == "ok":
                            self.status_labels[nome].config(
                                text="● Instalado", fg=COR_SUCESSO)
                        elif ok == "erro":
                            self.status_labels[nome].config(
                                text="● Falhou", fg=COR_ERRO)
                        elif ok == "aviso":
                            self.status_labels[nome].config(
                                text="● Parcial", fg=COR_AVISO)
                        elif ok == "verificando":
                            self.status_labels[nome].config(
                                text="● Verificando...", fg=COR_AVISO)

                elif t == "fim":
                    self._rodando = False
                    sucesso = msg.get("sucesso", False)
                    self.btn_instalar.config(
                        text="⬇  INSTALAR TUDO", state="normal",
                        bg=COR_PRIMARIA
                    )
                    if sucesso:
                        self.btn_abrir_app.config(state="normal")
                        self.lbl_prog_texto.config(text="✔ Instalação concluída com sucesso!")
                    else:
                        self.lbl_prog_texto.config(text="⚠ Alguns itens falharam — veja o log.")

        except queue.Empty:
            pass
        self.after(80, self._processar_fila)

    def _emit(self, **kwargs):
        self.fila.put(kwargs)

    # ── Log ──────────────────────────────────────────────────

    def _log(self, msg, tag=""):
        self.txt_log.configure(state="normal")
        if msg == "":
            self.txt_log.insert("end", "\n")
        else:
            ts = datetime.now().strftime("%H:%M:%S")
            self.txt_log.insert("end", f"[{ts}] ", "titulo")
            self.txt_log.insert("end", msg + "\n", tag)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    # ── Verificacao inicial ───────────────────────────────────

    def _verificar_status_inicial(self):
        def run():
            for nome in self.status_labels:
                self._emit(tipo="status", nome=nome, ok="verificando")

            # Libs
            faltando = []
            for lib in LIBS_PYTHON:
                ni = lib.replace("-","_")
                if ni == "pillow": ni = "PIL"
                try: __import__(ni)
                except ImportError: faltando.append(lib)
            self._emit(tipo="status", nome="libs",
                       ok="ok" if not faltando else "erro")

            # Tesseract
            exe = tesseract_encontrar()
            self._emit(tipo="status", nome="tesseract",
                       ok="ok" if exe else "erro")

            # Poppler
            bin_dir = poppler_encontrar()
            self._emit(tipo="status", nome="poppler",
                       ok="ok" if bin_dir else "erro")

            # Habilita botao abrir se tudo ok
            if not faltando and exe and bin_dir:
                self.fila.put({"tipo": "fim", "sucesso": True})
                self._emit(tipo="log",
                           texto="Todas as dependencias ja estao instaladas!",
                           tag="ok")
                self._emit(tipo="prog", valor=1.0, texto="Pronto!")
            else:
                if faltando:
                    self._emit(tipo="log",
                               texto=f"Faltam libs: {', '.join(faltando)}", tag="aviso")
                if not exe:
                    self._emit(tipo="log", texto="Tesseract nao encontrado.", tag="aviso")
                if not bin_dir:
                    self._emit(tipo="log", texto="Poppler nao encontrado.", tag="aviso")
                self._emit(tipo="log",
                           texto="Clique em 'INSTALAR TUDO' para instalar automaticamente.",
                           tag="info")

        threading.Thread(target=run, daemon=True).start()

    # ── Instalacao ───────────────────────────────────────────

    def _iniciar_instalacao(self):
        if self._rodando: return
        self._rodando = True

        self.btn_instalar.config(text="⏳ Instalando...", state="disabled", bg="#555")
        self.btn_abrir_app.config(state="disabled")
        self.progressbar["value"] = 0

        # Limpa log
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

        def log_cb(msg, tag=""):
            self._emit(tipo="log", texto=msg, tag=tag)

        def prog_cb(valor, texto=""):
            self._emit(tipo="prog", valor=valor, texto=texto)

        def status_cb(nome, estado):
            self._emit(tipo="status", nome=nome, ok=estado)

        def run():
            logica = InstaladorLogica(log_cb, prog_cb, status_cb)
            sucesso = logica.instalar_tudo()
            self._emit(tipo="fim", sucesso=sucesso)

        threading.Thread(target=run, daemon=True).start()

    # ── Abrir app ────────────────────────────────────────────

    def _abrir_app(self):
        app_py = PASTA_APP / "SEPARACAODEASO.py"
        app_exe = PASTA_APP / "Separador_ASOs.exe"

        if app_exe.exists():
            subprocess.Popen([str(app_exe)],
                             creationflags=subprocess.CREATE_NO_WINDOW
                             if eh_windows() else 0)
        elif app_py.exists():
            subprocess.Popen([sys.executable, str(app_py)],
                             creationflags=subprocess.CREATE_NO_WINDOW
                             if eh_windows() else 0)
        else:
            import tkinter.messagebox as mb
            mb.showerror("Erro",
                         "Nao foi possivel encontrar SEPARACAODEASO.py ou Separador_ASOs.exe\n"
                         f"na pasta: {PASTA_APP}")
            return
        self.destroy()


# ─────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = Launcher()
    app.mainloop()
