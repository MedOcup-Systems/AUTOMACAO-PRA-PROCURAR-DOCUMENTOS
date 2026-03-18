#!/usr/bin/env python3
"""
=============================================================
  AUTOMACAO DE SEPARACAO DE ASOs
=============================================================
  - Agrupa funcionarios do mesmo dia: cada PDF e aberto uma
    unica vez, buscando todos os nomes do dia de uma vez.
  - Score de similaridade: aceita nomes com pequenas diferencas
    de OCR ou digitacao (ex: "JOSE SILVA" vs "JOSE DA SILVA").
  - Para imediatamente apos encontrar todos os nomes do grupo.

  Estrutura esperada de pastas:
    PASTA_MES/
      06.02/  <- dia
        arquivo.pdf
      10.02/
        ...

  Como usar:
    1. Preencha a planilha com nome (col A) e data DD/MM (col B)
    2. Execute: python separar_asos.py
    3. Selecione a planilha, pasta do mes e pasta de saida

  Dependencias (instalar uma vez):
    pip install pdfplumber pypdf pdf2image pytesseract openpyxl

  Programas externos necessarios no Windows:
    1. Tesseract OCR:
       https://github.com/UB-Mannheim/tesseract/wiki
       Instalar em: C:\Program Files\Tesseract-OCR\

    2. Poppler (para converter PDF em imagem):
       https://github.com/oschwartz10612/poppler-windows/releases
       Extrair e adicionar a pasta "bin" ao PATH do Windows,
       ou informar o caminho em POPPLER_PATH abaixo.
=============================================================
"""

import os
import re
import sys
import pdfplumber
import pytesseract
from pypdf import PdfWriter, PdfReader
from openpyxl import load_workbook
from datetime import datetime
import unicodedata
from tkinter import Tk, messagebox
from tkinter.filedialog import askdirectory, askopenfilename

# ─────────────────────────────────────────────────
#  CONFIGURACOES — EDITE AQUI SE NECESSARIO
# ─────────────────────────────────────────────────

# DPI para OCR — 200 e suficiente. NAO aumente para 300,
# PDFs de 150+ paginas podem travar o computador.
DPI_OCR = 200

# Caminho do Tesseract (Windows)
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Caminho da pasta "bin" do Poppler (deixe None para busca automatica)
POPPLER_PATH = None

# Score minimo de similaridade para aceitar um nome (0.0 a 1.0)
#   1.00 = identico
#   0.80 = aceita pequenas diferencas (recomendado)
#   0.65 = mais tolerante (util se nomes estiverem muito abreviados)
SCORE_MINIMO = 0.80

# ─────────────────────────────────────────────────
#  PALAVRAS-CHAVE DOS DOCUMENTOS
# ─────────────────────────────────────────────────

PALAVRAS_ASO = [
    "ASO - ATESTADO DE SAUDE OCUPACIONAL",
    "ATESTADO DE SAUDE OCUPACIONAL",
]

PALAVRAS_FICHA = [
    "PRONTUARIO MEDICO",
    "ANAMNESE",
    "FICHA CLINICA",
]

# ─────────────────────────────────────────────────
#  VARIAVEL GLOBAL DO POPPLER
# ─────────────────────────────────────────────────

_POPPLER_BIN = None  # preenchido em verificar_dependencias()


# =============================================================
#  UTILITARIOS
# =============================================================

def normalizar(texto):
    """Remove acentos e coloca em maiusculo para comparacao robusta."""
    if not texto:
        return ""
    texto = str(texto).upper().strip()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in texto if not unicodedata.combining(c))


def nome_para_arquivo(texto):
    """Converte nome para formato seguro de nome de arquivo."""
    texto = normalizar(texto)
    texto = re.sub(r"[^\w\s\-]", "", texto)
    return re.sub(r"\s+", "_", texto.strip())


def parsear_data(valor):
    """
    Aceita objeto datetime (Excel), string DD/MM ou YYYY-MM-DD.
    Retorna (dia_str, mes_str) com zero a esquerda.
    """
    if isinstance(valor, datetime):
        return valor.strftime("%d"), valor.strftime("%m")
    texto = str(valor).strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", texto)
    if m:
        return m.group(3), m.group(2)
    partes = re.split(r"[/\-\.]", texto)
    if len(partes) >= 2 and len(partes[0]) <= 2:
        return partes[0].zfill(2), partes[1].zfill(2)
    return None, None


def barra_progresso(atual, total, label="", largura=35):
    """Barra de progresso que se atualiza na mesma linha do terminal."""
    pct = atual / total if total > 0 else 0
    preenchido = int(largura * pct)
    barra = "█" * preenchido + "░" * (largura - preenchido)
    linha = f"\r  [{barra}] {atual}/{total}  {label}"
    print(linha.ljust(80), end="", flush=True)
    if atual >= total:
        print()


# =============================================================
#  SIMILARIDADE DE NOMES
# =============================================================

def levenshtein(a, b):
    """Calcula distancia de edicao entre duas strings."""
    if len(a) < len(b):
        return levenshtein(b, a)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a):
        curr = [i + 1]
        for j, cb in enumerate(b):
            curr.append(min(prev[j + 1] + 1, curr[j] + 1,
                            prev[j] + (0 if ca == cb else 1)))
        prev = curr
    return prev[-1]


def score_nome(nome_planilha, texto_pagina):
    """
    Calcula o score de similaridade entre o nome da planilha
    e o texto extraido da pagina do PDF.

    Estrategia em camadas (retorna o maior score):
      1. Nome exato no texto                          → 1.00
      2. Todas as palavras do nome estao no texto     → 0.95
      3. Proporcao de palavras encontradas            → 0.50–0.90
      4. Proporcao de palavras com erro de 1 char
         (tipico de OCR: O→0, l→1, rn→m, etc.)       → 0.50–0.88
      5. Levenshtein entre nome e cada linha          → 0.0–1.0

    Retorna float entre 0.0 e 1.0.
    """
    np_norm = normalizar(nome_planilha)
    tx_norm = normalizar(texto_pagina)

    # 1. Nome completo exato no texto
    if np_norm in tx_norm:
        return 1.00

    palavras_nome = [p for p in np_norm.split() if len(p) > 2]
    if not palavras_nome:
        return 0.0

    palavras_tx = set(tx_norm.split())

    # 2. Todas as palavras batem exatamente
    encontradas_exatas = [p for p in palavras_nome if p in palavras_tx]
    proporcao_exata = len(encontradas_exatas) / len(palavras_nome)
    if proporcao_exata >= 1.0:
        return 0.95

    # 3. Palavras com tolerancia de 1 erro de caractere (OCR)
    def palavra_proxima(palavra, conjunto):
        if palavra in conjunto:
            return True
        for p in conjunto:
            if abs(len(p) - len(palavra)) <= 1 and levenshtein(palavra, p) <= 1:
                return True
        return False

    encontradas_fuzzy = [p for p in palavras_nome if palavra_proxima(p, palavras_tx)]
    proporcao_fuzzy = len(encontradas_fuzzy) / len(palavras_nome)

    if proporcao_fuzzy >= 1.0:
        return 0.90
    if proporcao_fuzzy >= 0.5:
        score_fuzzy = round(0.50 + proporcao_fuzzy * 0.38, 2)
    else:
        score_fuzzy = 0.0

    if proporcao_exata >= 0.5:
        score_exato = round(0.50 + proporcao_exata * 0.40, 2)
    else:
        score_exato = 0.0

    # 4. Levenshtein linha a linha
    melhor_lev = 0.0
    for linha in tx_norm.split("\n"):
        linha = linha.strip()
        if len(linha) < 5 or len(linha) > len(np_norm) * 2.5:
            continue
        dist = levenshtein(np_norm, linha)
        s = 1.0 - dist / max(len(np_norm), len(linha), 1)
        if s > melhor_lev:
            melhor_lev = s
            if melhor_lev > 0.95:
                break

    return round(max(score_exato, score_fuzzy, melhor_lev), 2)


# =============================================================
#  POPPLER / DEPENDENCIAS
# =============================================================

def descobrir_poppler():
    """
    Localiza o Poppler automaticamente:
      1. POPPLER_PATH definido manualmente
      2. Variavel de ambiente PATH do Windows
      3. Locais de instalacao comuns no Windows
    """
    import shutil

    if POPPLER_PATH and os.path.isdir(POPPLER_PATH):
        if os.path.isfile(os.path.join(POPPLER_PATH, "pdftoppm.exe")):
            return POPPLER_PATH

    exe = shutil.which("pdftoppm") or shutil.which("pdftoppm.exe")
    if exe:
        return os.path.dirname(exe)

    for base in [
        r"C:\poppler",
        r"C:\Program Files\poppler",
        r"C:\Program Files (x86)\poppler",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "poppler"),
        os.path.join(os.environ.get("USERPROFILE",   ""), "poppler"),
    ]:
        if os.path.isdir(base):
            for raiz, _, arqs in os.walk(base):
                if "pdftoppm.exe" in arqs:
                    return raiz
    return None


def verificar_dependencias():
    global _POPPLER_BIN
    erros = []

    # Tesseract
    if os.path.exists(TESSERACT_PATH):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
    else:
        try:
            pytesseract.get_tesseract_version()
        except Exception:
            erros.append(
                "TESSERACT NAO ENCONTRADO\n"
                "Instale em: https://github.com/UB-Mannheim/tesseract/wiki\n"
                f"Caminho esperado: {TESSERACT_PATH}"
            )

    # Poppler
    _POPPLER_BIN = descobrir_poppler()
    if _POPPLER_BIN:
        print(f"Poppler encontrado em: {_POPPLER_BIN}")
    else:
        try:
            from pdf2image import convert_from_bytes
            convert_from_bytes(
                b"%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n"
                b"2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n"
                b"3 0 obj<</Type/Page/MediaBox[0 0 3 3]/Parent 2 0 R>>endobj\n"
                b"xref\n0 4\n0000000000 65535 f\n0000000009 00000 n\n"
                b"0000000058 00000 n\n0000000115 00000 n\n"
                b"trailer<</Size 4/Root 1 0 R>>\nstartxref\n190\n%%EOF",
                dpi=72
            )
            print("Poppler encontrado no PATH do sistema.")
        except Exception:
            erros.append(
                "POPPLER NAO ENCONTRADO\n"
                "Baixe em: https://github.com/oschwartz10612/poppler-windows/releases\n"
                "Extraia e adicione a pasta 'bin' ao PATH do Windows.\n"
                "Ou defina POPPLER_PATH no topo do script.\n"
                "Exemplo: POPPLER_PATH = r'C:\\poppler\\Library\\bin'"
            )

    if erros:
        msg = "Erros antes de iniciar:\n\n" + "\n\n".join(erros)
        print(msg)
        try:
            root = Tk(); root.withdraw()
            messagebox.showerror("Erro de configuracao", msg)
        except Exception:
            pass
        sys.exit(1)

    print("Dependencias OK (Tesseract + Poppler encontrados)")


# =============================================================
#  PDF / OCR
# =============================================================

def encontrar_pdfs_do_dia(pasta_mes, dia, mes):
    """Tenta formatos DD.MM, D.MM, DD, D para o nome da subpasta."""
    for tentativa in [f"{dia}.{mes}", f"{int(dia)}.{mes}", dia, str(int(dia))]:
        pasta_dia = os.path.join(pasta_mes, tentativa)
        if os.path.isdir(pasta_dia):
            pdfs = sorted([
                os.path.join(pasta_dia, f)
                for f in os.listdir(pasta_dia)
                if f.lower().endswith(".pdf")
            ])
            if pdfs:
                print(f"  Pasta '{tentativa}' — {len(pdfs)} PDF(s)")
            return pdfs
    return []


def pagina_para_imagem(caminho_pdf, numero_pagina):
    """Converte UMA pagina do PDF para imagem para OCR."""
    from pdf2image import convert_from_path
    kwargs = {"poppler_path": _POPPLER_BIN} if _POPPLER_BIN else {}
    imgs = convert_from_path(
        caminho_pdf, dpi=DPI_OCR,
        first_page=numero_pagina + 1,
        last_page=numero_pagina + 1,
        **kwargs
    )
    return imgs[0] if imgs else None


def e_pagina_aso(texto):
    """Retorna True se o texto pertence a um ASO (nao a prontuario/ficha)."""
    t = normalizar(texto)
    if any(normalizar(p) in t for p in PALAVRAS_FICHA):
        return False
    return any(normalizar(p) in t for p in PALAVRAS_ASO)


def buscar_grupo_no_pdf(caminho_pdf, grupo):
    """
    Varre o PDF buscando todos os nomes do grupo de uma vez.

    Parametros:
      caminho_pdf : caminho para o arquivo PDF
      grupo       : dict {nome_original: nome_normalizado}

    Retorna dict {nome_original: pagina_encontrada (0-indexed)}.
    Para de varrer assim que todos os nomes forem encontrados.
    """
    pendentes = dict(grupo)   # copia — remove conforme acha
    resultado = {}

    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            total = len(pdf.pages)
            nome_pdf = os.path.basename(caminho_pdf)
            print(f"\n    {nome_pdf} ({total} pags) | {len(pendentes)} nome(s) pendente(s)")

            # Detecta se o PDF tem texto digital (evita OCR desnecessario)
            tem_texto = any(
                (p.extract_text() or "").strip()
                for p in pdf.pages[:3]
            )
            if not tem_texto:
                print("    PDF escaneado — OCR ativo pagina a pagina...")

            for i, pagina in enumerate(pdf.pages):
                if not pendentes:
                    break  # todos encontrados — para imediatamente

                barra_progresso(
                    i + 1, total,
                    f"pag {i+1}/{total} | faltam {len(pendentes)}"
                )

                # Extrai texto da pagina
                texto = (pagina.extract_text() or "").strip()
                if not texto and not tem_texto:
                    img = pagina_para_imagem(caminho_pdf, i)
                    if not img:
                        continue
                    texto = pytesseract.image_to_string(
                        img, lang="eng", config="--psm 3 --oem 1"
                    )
                    fonte = "OCR"
                elif not texto:
                    continue
                else:
                    fonte = "digital"

                # Filtra: so processa paginas de ASO
                if not e_pagina_aso(texto):
                    continue

                # Verifica cada nome pendente contra esta pagina
                achados_agora = []
                for nome_orig, nome_norm in list(pendentes.items()):
                    score = score_nome(nome_norm, texto)
                    if score >= SCORE_MINIMO:
                        achados_agora.append((nome_orig, score))

                # Registra os achados e remove dos pendentes
                for nome_orig, score in achados_agora:
                    resultado[nome_orig] = i
                    del pendentes[nome_orig]
                    pct = int(score * 100)
                    print(
                        f"\r    [{fonte}] pag {i+1}/{total}: "
                        f"'{nome_orig}' — score {pct}%".ljust(80)
                    )

            # Garante quebra de linha apos a barra
            print()

    except Exception as e:
        print(f"\r    ERRO em {os.path.basename(caminho_pdf)}: {e}".ljust(80))

    return resultado


def extrair_pagina(caminho_pdf, pagina, caminho_saida):
    """Extrai uma pagina especifica e salva como novo PDF."""
    try:
        reader = PdfReader(caminho_pdf)
        writer = PdfWriter()
        writer.add_page(reader.pages[pagina])
        with open(caminho_saida, "wb") as f:
            writer.write(f)
        return True
    except Exception as e:
        print(f"    Erro ao salvar: {e}")
        return False


# =============================================================
#  PLANILHA E AGRUPAMENTO
# =============================================================

def ler_planilha(planilha_path):
    """
    Le nome (col A) e data DD/MM (col B) da planilha.
    Remove automaticamente duplicatas de mesmo nome + mesma data,
    mantendo apenas a primeira ocorrencia e reportando as ignoradas.
    """
    wb = load_workbook(planilha_path, data_only=True)
    ws = wb.active
    funcionarios = []
    vistos = {}      # chave: (nome_normalizado, dia, mes) → numero da linha
    duplicatas = []  # lista de (linha, nome, data) ignorados

    for num_linha, linha in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        nome = linha[0] if len(linha) > 0 else None
        data = linha[1] if len(linha) > 1 else None

        if not nome or str(nome).strip() == "":
            continue

        nome_str = str(nome).strip()
        dia, mes = parsear_data(data)

        # Chave de deduplicacao: nome normalizado + data (dia+mes ou None)
        chave = (normalizar(nome_str), dia, mes)

        if chave in vistos:
            duplicatas.append((num_linha, nome_str, data))
            continue

        vistos[chave] = num_linha
        funcionarios.append({
            "nome": nome_str,
            "data_raw": data,
        })

    # Reporta duplicatas encontradas
    if duplicatas:
        print(f"\n  [DUPLICATAS] {len(duplicatas)} linha(s) ignorada(s) por ja existirem:")
        for num_linha, nome, data in duplicatas:
            print(f"    Linha {num_linha}: '{nome}' [{data}] — ignorada")

    return funcionarios


def agrupar_por_dia(funcionarios):
    """
    Agrupa funcionarios pelo mesmo dia/mes para processar
    cada PDF uma unica vez por dia.

    Retorna:
      grupos_ordenados : lista de ((dia, mes), [func, ...])
      invalidos        : lista de funcs com data invalida
    """
    grupos   = {}
    invalidos = []

    for func in funcionarios:
        dia, mes = parsear_data(func["data_raw"])
        if not dia:
            invalidos.append(func)
            continue
        grupos.setdefault((dia, mes), []).append(func)

    # Ordena cronologicamente: mes primeiro, depois dia
    grupos_ordenados = sorted(grupos.items(), key=lambda x: (x[0][1], x[0][0]))
    return grupos_ordenados, invalidos


# =============================================================
#  RELATORIO
# =============================================================

def gerar_relatorio(pasta_saida, encontrados, nao_encontrados, agora):
    """
    Gera um arquivo Excel de relatorio na pasta de saida com:
      - Aba 'Encontrados'     : lista de ASOs extraidos com sucesso
      - Aba 'Nao Encontrados' : lista de nomes nao localizados nos PDFs
      - Aba 'Resumo'          : totais e data/hora de execucao
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()

    COR_VERDE      = "1E7E34"
    COR_VERMELHO   = "C0392B"
    COR_AZUL       = "1F4E79"
    COR_VERDE_BG   = "D4EDDA"
    COR_VERM_BG    = "FADBD8"
    COR_LINHA_PAR  = "F2F2F2"
    FONTE_BRANCA   = "FFFFFF"

    def estilo_header(cell, cor_fundo):
        cell.font      = Font(bold=True, color=FONTE_BRANCA, name="Arial", size=11)
        cell.fill      = PatternFill("solid", start_color=cor_fundo)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        borda           = Side(style="thin", color="FFFFFF")
        cell.border    = Border(left=borda, right=borda, top=borda, bottom=borda)

    def estilo_dado(cell, cor_bg=None):
        cell.font      = Font(name="Arial", size=10)
        cell.alignment = Alignment(vertical="center")
        if cor_bg:
            cell.fill  = PatternFill("solid", start_color=cor_bg)
        borda           = Side(style="thin", color="DDDDDD")
        cell.border    = Border(left=borda, right=borda, top=borda, bottom=borda)

    # ── Aba 1: Encontrados ────────────────────────
    ws_ok = wb.active
    ws_ok.title = "Encontrados"
    ws_ok.append(["#", "NOME DO FUNCIONARIO", "STATUS"])
    ws_ok.row_dimensions[1].height = 28
    for col, header in enumerate(["#", "NOME DO FUNCIONARIO", "STATUS"], 1):
        estilo_header(ws_ok.cell(row=1, column=col), COR_VERDE)

    for i, nome in enumerate(sorted(encontrados), 1):
        ws_ok.append([i, nome, "ASO extraido com sucesso"])
        bg = COR_VERDE_BG if i % 2 == 0 else "FFFFFF"
        for col in range(1, 4):
            estilo_dado(ws_ok.cell(row=i + 1, column=col), bg)
        ws_ok.row_dimensions[i + 1].height = 18

    ws_ok.column_dimensions["A"].width = 6
    ws_ok.column_dimensions["B"].width = 40
    ws_ok.column_dimensions["C"].width = 28

    # ── Aba 2: Nao Encontrados ────────────────────
    ws_nao = wb.create_sheet("Nao Encontrados")
    ws_nao.append(["#", "NOME DO FUNCIONARIO", "STATUS"])
    ws_nao.row_dimensions[1].height = 28
    for col, header in enumerate(["#", "NOME DO FUNCIONARIO", "STATUS"], 1):
        estilo_header(ws_nao.cell(row=1, column=col), COR_VERMELHO)

    for i, nome in enumerate(sorted(nao_encontrados), 1):
        ws_nao.append([i, nome, "Nao localizado nos PDFs"])
        bg = COR_VERM_BG if i % 2 == 0 else "FFFFFF"
        for col in range(1, 4):
            estilo_dado(ws_nao.cell(row=i + 1, column=col), bg)
        ws_nao.row_dimensions[i + 1].height = 18

    ws_nao.column_dimensions["A"].width = 6
    ws_nao.column_dimensions["B"].width = 40
    ws_nao.column_dimensions["C"].width = 28

    # ── Aba 3: Resumo ─────────────────────────────
    ws_res = wb.create_sheet("Resumo")
    linhas = [
        ("Data/Hora",         agora.strftime("%d/%m/%Y %H:%M:%S")),
        ("Total processados", len(encontrados) + len(nao_encontrados)),
        ("Encontrados",       len(encontrados)),
        ("Nao encontrados",   len(nao_encontrados)),
        ("Taxa de sucesso",   f"{int(len(encontrados)/(len(encontrados)+len(nao_encontrados))*100) if (encontrados or nao_encontrados) else 0}%"),
    ]
    for i, (chave, valor) in enumerate(linhas, 1):
        ws_res.cell(row=i, column=1, value=chave).font  = Font(bold=True, name="Arial", size=11)
        ws_res.cell(row=i, column=2, value=valor).font  = Font(name="Arial", size=11)
        ws_res.cell(row=i, column=1).fill = PatternFill("solid", start_color="EBF5FB")
        ws_res.row_dimensions[i].height = 22

    ws_res.column_dimensions["A"].width = 22
    ws_res.column_dimensions["B"].width = 30

    # Salva o arquivo
    nome_arquivo = f"RELATORIO_ASOs_{agora.strftime('%d-%m-%Y_%H-%M')}.xlsx"
    caminho = os.path.join(pasta_saida, nome_arquivo)
    wb.save(caminho)
    print(f"\n  Relatorio salvo: {nome_arquivo}")


def processar():
    verificar_dependencias()

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    planilha = askopenfilename(
        title="Selecione a planilha de funcionarios (.xlsx)",
        filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")]
    )
    if not planilha:
        print("Selecao cancelada.")
        return

    pasta_mes = askdirectory(title="Selecione a pasta do MES (ex: .../2026/02)")
    if not pasta_mes:
        print("Selecao cancelada.")
        return

    pasta_saida = askdirectory(title="Selecione a pasta de SAIDA dos ASOs")
    if not pasta_saida:
        print("Selecao cancelada.")
        return

    print("\n" + "="*58)
    print("  SEPARACAO AUTOMATICA DE ASOs")
    print("="*58)
    print(f"\n Planilha  : {planilha}")
    print(f" Pasta mes : {pasta_mes}")
    print(f" Saida     : {pasta_saida}")
    print(f" OCR DPI   : {DPI_OCR}")
    print(f" Score min : {int(SCORE_MINIMO * 100)}%")

    os.makedirs(pasta_saida, exist_ok=True)
    funcionarios = ler_planilha(planilha)
    grupos, invalidos = agrupar_por_dia(funcionarios)

    total_funcs  = len(funcionarios)
    total_grupos = len(grupos)
    print(f"\n-> {total_funcs} funcionario(s) agrupados em {total_grupos} dia(s)\n")

    encontrados     = []
    nao_encontrados = []

    # Reporta invalidos imediatamente
    for func in invalidos:
        print(f"  [AVISO] '{func['nome']}' — data invalida ou vazia")
        nao_encontrados.append(func["nome"])

    funcs_processados = len(invalidos)

    for (dia, mes), grupo in grupos:
        nomes_do_grupo = [f["nome"] for f in grupo]
        data_fmt = f"{dia}-{mes}"

        print(f"\n{'─'*58}")
        print(f"  DIA {dia}/{mes} — {len(nomes_do_grupo)} funcionario(s):")
        for n in nomes_do_grupo:
            print(f"    • {n}")

        # Barra geral de progresso
        funcs_processados += len(nomes_do_grupo)
        barra_progresso(funcs_processados, total_funcs, f"{funcs_processados}/{total_funcs} total")

        pdfs = encontrar_pdfs_do_dia(pasta_mes, dia, mes)
        if not pdfs:
            print(f"  Nenhum PDF na pasta do dia '{dia}.{mes}'")
            nao_encontrados.extend(nomes_do_grupo)
            continue

        # {nome_original: nome_normalizado} — remove conforme acha
        pendentes = {n: normalizar(n) for n in nomes_do_grupo}
        achados   = {}  # {nome: (pdf_path, pagina)}

        for pdf_path in pdfs:
            if not pendentes:
                break  # todos encontrados, nao precisa abrir mais PDFs

            resultado_pdf = buscar_grupo_no_pdf(pdf_path, pendentes)

            for nome, pagina in resultado_pdf.items():
                achados[nome] = (pdf_path, pagina)
                del pendentes[nome]

        # Salva ASOs encontrados
        for nome, (pdf_path, pagina) in achados.items():
            nome_base = nome_para_arquivo(nome)
            saida = os.path.join(pasta_saida, f"ASO_{nome_base}_{data_fmt}.pdf")
            if extrair_pagina(pdf_path, pagina, saida):
                print(f"  Salvo : {os.path.basename(saida)}")
            encontrados.append(nome)

        # Registra nao encontrados
        for nome in pendentes:
            print(f"  AVISO : NAO ENCONTRADO — {nome}")
            nao_encontrados.append(nome)

    # ── Resumo final ──────────────────────────────
    agora = datetime.now()
    print("\n" + "="*58)
    print(f"  CONCLUIDO — {agora.strftime('%d/%m/%Y %H:%M:%S')}")
    print("="*58)
    print(f"  Encontrados    : {len(encontrados)}")
    print(f"  Nao encontrados: {len(nao_encontrados)}")
    if nao_encontrados:
        print("\n  Nao encontrados:")
        for n in nao_encontrados:
            print(f"    - {n}")
    print(f"\n  ASOs salvos em: {pasta_saida}")
    print("="*58 + "\n")

    # ── Gera arquivo de relatorio ──────────────────
    gerar_relatorio(pasta_saida, encontrados, nao_encontrados, agora)

    resumo = (
        f"Concluido!\n\n"
        f"Encontrados    : {len(encontrados)}\n"
        f"Nao encontrados: {len(nao_encontrados)}\n"
        f"\nRelatorio salvo na pasta de saida."
    )
    if nao_encontrados:
        resumo += "\nNao encontrados:\n" + "\n".join(f"  - {n}" for n in nao_encontrados)
    messagebox.showinfo("ASOs — Resultado", resumo)


if __name__ == "__main__":
    processar()