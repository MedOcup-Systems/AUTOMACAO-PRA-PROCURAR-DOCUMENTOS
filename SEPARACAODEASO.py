#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================
  SEPARADOR DE ASOs  -  v3.0
=============================================================
  - Interface escura (mesma da versao anterior) com progresso
    geral + progresso do PDF atual, log colorido, PAUSAR/PARAR.
  - Aceita a pasta do ANO ou do MES: o programa acha sozinho a
    subpasta do dia (06.02, 6.02, 06-02, 06_02, "DIA 06"...).
  - Cada PDF e aberto uma unica vez por dia e todos os nomes do
    dia sao procurados juntos; para assim que acha todos.
  - Score de similaridade tolerante a OCR (mesma logica antiga):
    nome exato > todas as palavras > palavras com 1 erro >
    Levenshtein linha a linha.
  - Checkpoint: se fechar/travar, continua de onde parou.
  - Nao precisa de Poppler nem pdfplumber/pypdf.

  Dependencias:
      pip install PyMuPDF pytesseract pillow openpyxl
  Programa externo:
      Tesseract OCR (com o pacote de idioma "por" de preferencia)
=============================================================
"""

import os
import re
import sys
import gc
import json
import time
import queue
import threading
import traceback
import unicodedata
from datetime import datetime, date

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.filedialog import askdirectory, askopenfilename

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.datetime import from_excel

try:
    import pytesseract
except ImportError:
    pytesseract = None
try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None
try:
    from PIL import Image
except ImportError:
    Image = None

# ─────────────────────────────────────────────────
#  CONFIGURACOES
# ─────────────────────────────────────────────────
DPI_OCR = 200                      # 200 = mesmo valor da versao que funcionava
TESSERACT_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
SCORE_MINIMO = 0.80                # aceita o nome em paginas de ASO
SCORE_SEGUNDA_CHANCE = 0.95        # nome quase identico em pagina sem palavra-chave de ASO
MIN_CHARS_TEXTO = 30               # abaixo disso a pagina e tratada como imagem -> OCR
PROFUNDIDADE_PASTAS = 3            # quantos niveis abaixo da pasta escolhida procurar o dia
NOME_CHECKPOINT = "checkpoint_progresso.json"
ARQUIVO_CONFIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config_separador.json")

IDIOMA_OCR = "eng"                 # ajustado em verificar_dependencias()

# ─────────────────────────────────────────────────
#  TEXTO / NORMALIZACAO
# ─────────────────────────────────────────────────
PREPOSICOES = {"DE", "DA", "DO", "DOS", "DAS", "E"}
MESES = ["JANEIRO", "FEVEREIRO", "MARCO", "ABRIL", "MAIO", "JUNHO",
         "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]


def limpar(texto):
    """MAIUSCULO, sem acentos, sem pontuacao, espacos unificados."""
    if texto is None:
        return ""
    t = unicodedata.normalize("NFKD", str(texto).upper())
    t = "".join(c for c in t if not unicodedata.combining(c))
    return " ".join(re.findall(r"[A-Z0-9]+", t))


def nome_para_arquivo(texto):
    return limpar(texto).replace(" ", "_")


# Titulos que identificam um ASO com certeza
_ASO_FORTE = [limpar(p) for p in [
    "ASO - ATESTADO DE SAUDE OCUPACIONAL",
    "ATESTADO DE SAUDE OCUPACIONAL",
    "ATESTADO MEDICO OCUPACIONAL",
    "ATESTADO OCUPACIONAL",
]]
# Termos mais fracos (podem aparecer em cabecalho de clinica)
_ASO_FRACO = [limpar(p) for p in [
    "SAUDE OCUPACIONAL", "EXAME MEDICO OCUPACIONAL", "EXAME OCUPACIONAL",
    "AVALIACAO CLINICA OCUPACIONAL", "ATESTADO DE SAUDE",
    "APTO PARA A FUNCAO", "APTO PARA A ATIVIDADE", "APTO PARA O TRABALHO",
    "INAPTO PARA A FUNCAO", "CONCLUSAO DO EXAME", "PARECER APTO",
    "MEDICO COORDENADOR", "MEDICO EXAMINADOR",
]]
_FICHA = [limpar(p) for p in [
    "PRONTUARIO MEDICO", "ANAMNESE CLINICA", "ANAMNESE OCUPACIONAL",
    "FICHA CLINICA OCUPACIONAL", "FICHA CLINICA", "QUESTIONARIO DE SAUDE",
]]


def e_pagina_aso(tx):
    """tx = texto ja passado por limpar()."""
    if any(k in tx for k in _ASO_FORTE):
        return True
    if any(k in tx for k in _FICHA):
        return False
    if any(k in tx for k in _ASO_FRACO):
        return True
    if re.search(r"\bASO\b", tx) and ("EXAME" in tx or "MEDICO" in tx or "FUNCAO" in tx):
        return True
    return False


# ─────────────────────────────────────────────────
#  SIMILARIDADE DE NOMES
# ─────────────────────────────────────────────────
def levenshtein(a, b):
    if len(a) < len(b):
        return levenshtein(b, a)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a):
        curr = [i + 1]
        for j, cb in enumerate(b):
            curr.append(min(prev[j + 1] + 1, curr[j] + 1, prev[j] + (0 if ca == cb else 1)))
        prev = curr
    return prev[-1]


def preparar_pagina(texto):
    """Pre-processa o texto da pagina UMA vez (usado para todos os nomes)."""
    linhas = [l for l in (limpar(x) for x in str(texto).split("\n")) if l]
    tx = " ".join(linhas)
    tokens = set(tx.split())
    por_tam = {}
    for t in tokens:
        por_tam.setdefault(len(t), set()).add(t)
    return {"linhas": linhas, "tx": tx, "txp": f" {tx} ", "tokens": tokens, "por_tam": por_tam}


def _palavra_proxima(p, por_tam):
    if p in por_tam.get(len(p), ()):
        return True
    if len(p) < 4:
        return False
    for L in (len(p) - 1, len(p), len(p) + 1):
        for t in por_tam.get(L, ()):
            if levenshtein(p, t) <= 1:
                return True
    return False


def score_nome(alvo, pag, rapido=False):
    """
    alvo: nome ja passado por limpar(). pag: retorno de preparar_pagina().
      1.00  nome completo no texto
      0.95  todas as palavras significativas presentes
      0.90  todas presentes com ate 1 erro de OCR
      0.50-0.88  proporcao de palavras
      Levenshtein linha a linha
    rapido=True: so as etapas exatas (usado em paginas fora do padrao ASO).
    """
    if not alvo or not pag["tx"]:
        return 0.0
    if f" {alvo} " in pag["txp"]:
        return 1.00

    palavras = [p for p in alvo.split() if len(p) >= 2]
    signif = [p for p in palavras if p not in PREPOSICOES] or palavras
    if not signif:
        return 0.0

    exatas = [p for p in signif if p in pag["tokens"]]
    prop_exata = len(exatas) / len(signif)
    if prop_exata >= 1.0:
        return 0.95
    if rapido:
        return 0.0

    faltando = [p for p in signif if p not in pag["tokens"]]
    fuzzy_ok = len(exatas) + sum(1 for p in faltando if _palavra_proxima(p, pag["por_tam"]))
    prop_fuzzy = fuzzy_ok / len(signif)
    if prop_fuzzy >= 1.0:
        return 0.90

    score_fuzzy = round(0.50 + prop_fuzzy * 0.38, 2) if prop_fuzzy >= 0.5 else 0.0
    score_exato = round(0.50 + prop_exata * 0.40, 2) if prop_exata >= 0.5 else 0.0

    melhor_lev = 0.0
    if prop_fuzzy >= 0.5:  # so vale a pena o calculo caro se ja ha palavras parecidas
        n_pal = max(1, len(alvo.split()))
        for linha in pag["linhas"]:
            if len(linha) < 5:
                continue
            if len(linha) > len(alvo) * 2.5:
                toks = linha.split()
                candidatos = (" ".join(toks[k:k + n_pal]) for k in range(max(1, len(toks) - n_pal + 1)))
            else:
                candidatos = (linha,)
            for c in candidatos:
                s = 1.0 - levenshtein(alvo, c) / max(len(alvo), len(c), 1)
                if s > melhor_lev:
                    melhor_lev = s
            if melhor_lev > 0.95:
                break

    return round(max(score_exato, score_fuzzy, melhor_lev), 2)


# ─────────────────────────────────────────────────
#  DATAS / PLANILHA
# ─────────────────────────────────────────────────
def parsear_data(valor):
    """Aceita datetime, numero (serial do Excel ou 6.02), 'DD/MM', 'DD/MM/AAAA', 'AAAA-MM-DD'."""
    if valor is None:
        return None, None
    if isinstance(valor, (datetime, date)):
        return f"{valor.day:02d}", f"{valor.month:02d}"
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        try:
            if valor >= 1000:  # serial de data do Excel
                dt = from_excel(valor)
                return f"{dt.day:02d}", f"{dt.month:02d}"
            valor = f"{float(valor):.2f}"  # 6.02 digitado como numero = 06/02
        except Exception:
            return None, None
    texto = str(valor).strip()
    if not texto:
        return None, None
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", texto)
    if m:
        d, mes = int(m.group(3)), int(m.group(2))
    else:
        m = re.search(r"(\d{1,2})\s*[/\-\.]\s*(\d{1,2})", texto)
        if not m:
            return None, None
        d, mes = int(m.group(1)), int(m.group(2))
    if 1 <= d <= 31 and 1 <= mes <= 12:
        return f"{d:02d}", f"{mes:02d}"
    return None, None


def ler_planilha(caminho):
    """Retorna (funcionarios, duplicatas). Nome = coluna A, data = coluna B."""
    wb = load_workbook(caminho, data_only=True)
    ws = wb.active
    funcs, vistos, dups = [], set(), []
    for n, linha in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        nome = linha[0] if len(linha) > 0 else None
        raw = linha[1] if len(linha) > 1 else None
        if nome is None or not str(nome).strip():
            continue
        nome = " ".join(str(nome).split())
        dia, mes = parsear_data(raw)
        chave = (limpar(nome), dia, mes)
        if chave in vistos:
            dups.append((n, nome))
            continue
        vistos.add(chave)
        funcs.append({"nome": nome, "raw": raw, "dia": dia, "mes": mes})
    wb.close()
    return funcs, dups


def agrupar_por_dia(funcs):
    grupos, invalidos = {}, []
    for f in funcs:
        if not f["dia"]:
            invalidos.append(f)
        else:
            grupos.setdefault((f["dia"], f["mes"]), []).append(f)
    ordenados = sorted(grupos.items(), key=lambda x: (x[0][1], x[0][0]))
    return ordenados, invalidos


# ─────────────────────────────────────────────────
#  LOCALIZACAO DAS PASTAS (ANO OU MES)
# ─────────────────────────────────────────────────
def indexar_pastas(raiz):
    """Lista todas as subpastas ate PROFUNDIDADE_PASTAS niveis: [(caminho, nome, pai)]."""
    indice = []
    for atual, dirs, _ in os.walk(raiz):
        rel = os.path.relpath(atual, raiz)
        prof = 0 if rel == "." else rel.count(os.sep) + 1
        dirs.sort()
        if prof >= PROFUNDIDADE_PASTAS:
            dirs[:] = []
        for d in dirs:
            indice.append((os.path.join(atual, d), d, atual))
    return indice


def _parece_mes(nome_norm, m):
    return bool(re.match(rf"^0*{m}(?!\d)", nome_norm)) or MESES[m - 1] in nome_norm


def localizar_pdfs_do_dia(indice, raiz, dia, mes):
    """Retorna (lista_de_pdfs, [descricao das pastas usadas])."""
    d, m = int(dia), int(mes)
    pat_dm = re.compile(rf"^(?:DIA\s*)?0*{d}\s*[./\-_ ]+\s*0*{m}(?!\d)")
    pat_d = re.compile(rf"^(?:DIA\s*)?0*{d}$")
    exatas, so_dia = [], []
    for caminho, nome, pai in indice:
        n = nome.upper().strip()
        n = "".join(c for c in unicodedata.normalize("NFKD", n) if not unicodedata.combining(c))
        if pat_dm.match(n):
            exatas.append(caminho)
        elif pat_d.match(n):
            nome_pai = limpar(os.path.basename(pai.rstrip("/\\")))
            pai_e_raiz = os.path.normpath(pai) == os.path.normpath(raiz)
            if _parece_mes(nome_pai, m) or (pai_e_raiz and not re.match(r"^\d{4}$", nome_pai)):
                so_dia.append(caminho)
    pastas = exatas or so_dia
    pdfs = []
    for p in pastas:
        for atual, dirs, arqs in os.walk(p):
            dirs.sort()
            pdfs.extend(os.path.join(atual, a) for a in sorted(arqs) if a.lower().endswith(".pdf"))
    desc = [os.path.relpath(p, raiz) for p in pastas]
    return pdfs, desc


# ─────────────────────────────────────────────────
#  CHECKPOINT
# ─────────────────────────────────────────────────
class Checkpoint:
    def __init__(self, pasta_saida):
        self.pasta = pasta_saida
        self.caminho = os.path.join(pasta_saida, NOME_CHECKPOINT)
        self.dados = {"versao": "3.0", "encontrados": {}}
        try:
            if os.path.isfile(self.caminho):
                with open(self.caminho, "r", encoding="utf-8") as f:
                    d = json.load(f)
                if isinstance(d, dict) and "encontrados" in d:
                    self.dados = d
        except Exception:
            pass

    @staticmethod
    def chave(nome, dia, mes):
        return f"{limpar(nome)}_{dia}_{mes}"

    def total(self):
        return len(self.dados["encontrados"])

    def limpar_tudo(self):
        self.dados["encontrados"] = {}
        self.salvar()

    def salvar(self):
        try:
            os.makedirs(self.pasta, exist_ok=True)
            tmp = self.caminho + f".tmp{os.getpid()}"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.dados, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self.caminho)
        except Exception:
            pass

    def registrar(self, nome, dia, mes, saida, pdf, pagina):
        self.dados["encontrados"][self.chave(nome, dia, mes)] = {
            "nome": nome, "dia": dia, "mes": mes, "arquivo": saida,
            "pdf_origem": os.path.basename(pdf), "pagina": pagina,
            "quando": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        self.salvar()

    def esta_pronto(self, nome, dia, mes):
        info = self.dados["encontrados"].get(self.chave(nome, dia, mes))
        if info:
            p = info.get("arquivo", "")
            if os.path.isfile(p) and os.path.getsize(p) > 500:
                return True
        esperado = os.path.join(self.pasta, f"ASO_{nome_para_arquivo(nome)}_{dia}-{mes}.pdf")
        return os.path.isfile(esperado) and os.path.getsize(esperado) > 500


# ─────────────────────────────────────────────────
#  CONTEXTO (ponte thread -> GUI, pausa e cancelamento)
# ─────────────────────────────────────────────────
class Ctx:
    def __init__(self, q):
        self.q = q
        self.pausa = threading.Event()
        self.pausa.set()
        self.cancelar = threading.Event()

    def log(self, msg, tag="dim"):
        self.q.put(("log", msg, tag))

    def geral(self, a, b):
        self.q.put(("geral", a, b))

    def pdf(self, a, b):
        self.q.put(("pdf", a, b))

    def checar(self):
        """Bloqueia enquanto pausado. Retorna False se foi cancelado."""
        while not self.pausa.is_set():
            if self.cancelar.is_set():
                return False
            time.sleep(0.2)
        return not self.cancelar.is_set()


# ─────────────────────────────────────────────────
#  LEITURA DOS PDFs
# ─────────────────────────────────────────────────
def texto_da_pagina(doc, idx):
    """Texto digital; se for pagina de imagem faz OCR (apenas nessa pagina)."""
    texto, fonte, erro = "", "digital", None
    pagina = doc[idx]
    try:
        texto = pagina.get_text("text") or ""
    except Exception:
        texto = ""
    if len(texto.strip()) < MIN_CHARS_TEXTO:
        try:
            pix = pagina.get_pixmap(dpi=DPI_OCR, alpha=False, colorspace=fitz.csGRAY)
            img = Image.frombytes("L", [pix.width, pix.height], pix.samples)
            ocr = pytesseract.image_to_string(img, lang=IDIOMA_OCR, config="--psm 3 --oem 1")
            if len(ocr.strip()) > len(texto.strip()):
                texto, fonte = ocr, "OCR"
            del pix, img
        except Exception as e:
            erro = str(e)
    return texto.strip(), fonte, erro


def buscar_no_pdf(ctx, caminho_pdf, pendentes):
    """
    pendentes: {nome_original: nome_limpo}
    Retorna (achados {nome: pagina}, candidatos {nome: (score, pagina)}).
    Para assim que todos os nomes forem achados.
    """
    achados, candidatos = {}, {}
    pend = dict(pendentes)
    try:
        doc = fitz.open(caminho_pdf)
    except Exception as e:
        ctx.log(f"✗ Nao foi possivel abrir o PDF: {e}", "err")
        return achados, candidatos

    total = len(doc)
    ctx.log(f"{os.path.basename(caminho_pdf)}  ({total} pags) | buscando {len(pend)} nome(s)", "dim")
    n_ocr = n_aso = n_vazias = 0
    aviso_ocr = False

    for i in range(total):
        if not pend:
            break
        if not ctx.checar():
            break
        ctx.pdf(i + 1, total)
        try:
            texto, fonte, erro = texto_da_pagina(doc, i)
        except Exception as e:
            ctx.log(f"⚠ Erro na pag {i + 1}: {e}", "warn")
            continue
        if erro and not aviso_ocr:
            ctx.log(f"⚠ Erro no OCR (pag {i + 1}): {erro}", "warn")
            aviso_ocr = True
        if fonte == "OCR":
            n_ocr += 1
        if not texto:
            n_vazias += 1
            continue

        pag = preparar_pagina(texto)
        eh_aso = e_pagina_aso(pag["tx"])
        if eh_aso:
            n_aso += 1

        for nome, alvo in list(pend.items()):
            if eh_aso:
                sc = score_nome(alvo, pag)
                if sc >= SCORE_MINIMO:
                    achados[nome] = i
                    del pend[nome]
                    ctx.log(f"✓ Encontrado: {nome}  (pag {i + 1}, {fonte}, score {int(sc * 100)}%)", "ok")
            else:
                sc = score_nome(alvo, pag, rapido=True)
                if sc >= SCORE_SEGUNDA_CHANCE:
                    ant = candidatos.get(nome)
                    if not ant or sc > ant[0]:
                        candidatos[nome] = (sc, i)
        if i and i % 25 == 0:
            gc.collect()

    doc.close()
    ctx.pdf(total, total)
    ctx.log(f"Lidas: {min(i + 1, total) if total else 0}/{total} pags | OCR: {n_ocr} | "
            f"sem texto: {n_vazias} | reconhecidas como ASO: {n_aso}", "dim")
    if total and n_aso == 0 and pend:
        ctx.log("⚠ Nenhuma pagina foi reconhecida como ASO neste PDF (confira se e o PDF certo "
                "ou se o OCR esta lendo bem).", "warn")
    return achados, candidatos


def extrair_pagina(caminho_pdf, pagina, saida):
    src = novo = None
    try:
        src = fitz.open(caminho_pdf)
        novo = fitz.open()
        novo.insert_pdf(src, from_page=pagina, to_page=pagina)
        novo.save(saida, garbage=3, deflate=True)
        return True, ""
    except Exception as e:
        return False, str(e)
    finally:
        for d in (novo, src):
            try:
                if d:
                    d.close()
            except Exception:
                pass


# ─────────────────────────────────────────────────
#  RELATORIO EXCEL
# ─────────────────────────────────────────────────
def gerar_relatorio(pasta_saida, detalhes, nao_enc, agora):
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatorio"
    cab = ["#", "FUNCIONARIO", "DATA", "STATUS", "ORIGEM", "OBSERVACAO"]
    ws.append(cab)
    for c in range(1, len(cab) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF", name="Arial")
        cell.fill = PatternFill("solid", start_color="1F4E79")
        cell.alignment = Alignment(horizontal="center")

    linhas = [(d["nome"], d["data"], "ASO salvo", d["origem"], d.get("obs", ""), True) for d in detalhes]
    linhas += [(n, dt, "Nao encontrado", "-", motivo, False) for n, dt, motivo in nao_enc]
    linhas.sort(key=lambda x: limpar(x[0]))
    for i, (nome, dt, status, origem, obs, ok) in enumerate(linhas, 1):
        ws.append([i, nome, dt, status, origem, obs])
        cor = "D4EDDA" if ok else "FADBD8"
        for c in range(1, len(cab) + 1):
            ws.cell(row=i + 1, column=c).fill = PatternFill("solid", start_color=cor)
    for col, w in zip("ABCDEF", (6, 44, 10, 18, 26, 46)):
        ws.column_dimensions[col].width = w

    rs = wb.create_sheet("Resumo")
    total = len(linhas)
    ok = len(detalhes)
    dados = [("Data/Hora", agora.strftime("%d/%m/%Y %H:%M:%S")), ("Total", total),
             ("Encontrados", ok), ("Nao encontrados", total - ok),
             ("Taxa de sucesso", f"{int(ok / total * 100) if total else 0}%")]
    for i, (k, v) in enumerate(dados, 1):
        rs.cell(row=i, column=1, value=k).font = Font(bold=True, name="Arial")
        rs.cell(row=i, column=2, value=v)
    rs.column_dimensions["A"].width = 22
    rs.column_dimensions["B"].width = 24

    nome = f"RELATORIO_ASOs_{agora.strftime('%d-%m-%Y_%H-%M')}.xlsx"
    try:
        wb.save(os.path.join(pasta_saida, nome))
        return nome
    except Exception:
        return None


# ─────────────────────────────────────────────────
#  PROCESSAMENTO PRINCIPAL (roda em thread)
# ─────────────────────────────────────────────────
def executar(ctx, planilha, raiz, saida, usar_cp):
    try:
        _executar(ctx, planilha, raiz, saida, usar_cp)
    except Exception:
        ctx.log("ERRO CRITICO:\n" + traceback.format_exc(), "err")
    finally:
        ctx.q.put(("fim",))


def _executar(ctx, planilha, raiz, saida, usar_cp):
    ctx.log("Lendo planilha...", "info")
    funcs, dups = ler_planilha(planilha)
    for linha, nome in dups:
        ctx.log(f"Duplicata ignorada (linha {linha}): {nome}", "warn")
    grupos, invalidos = agrupar_por_dia(funcs)
    total = len(funcs)
    ctx.log(f"{total} funcionário(s) em {len(grupos)} dia(s)", "info")
    if total == 0:
        ctx.log("Nenhum funcionário encontrado na planilha.", "warn")
        return

    detalhes, nao_enc = [], []
    for f in invalidos:
        ctx.log(f"Data vazia/inválida: {f['nome']}   (valor lido: {f['raw']!r})", "warn")
        nao_enc.append((f["nome"], "-", "Data vazia/invalida na planilha"))

    os.makedirs(saida, exist_ok=True)
    cp = Checkpoint(saida)
    if not usar_cp:
        cp.limpar_tudo()

    ja = set()
    if usar_cp:
        for f in funcs:
            if f["dia"] and cp.esta_pronto(f["nome"], f["dia"], f["mes"]):
                ja.add(Checkpoint.chave(f["nome"], f["dia"], f["mes"]))
                detalhes.append({"nome": f["nome"], "data": f"{f['dia']}/{f['mes']}",
                                 "origem": "Recuperado do checkpoint"})
        if ja:
            ctx.log(f"[CHECKPOINT] {len(ja)} ASO(s) já prontos — retomando de onde parou.", "warn")

    ctx.log("Indexando subpastas...", "dim")
    indice = indexar_pastas(raiz)
    ctx.log(f"{len(indice)} subpasta(s) indexada(s) em '{os.path.basename(raiz.rstrip('/\\'))}'", "dim")

    feitos = len(invalidos) + len(ja)
    ctx.geral(feitos, total)
    cancelado = False

    for (dia, mes), grupo in grupos:
        if not ctx.checar():
            cancelado = True
            break

        nomes = [f["nome"] for f in grupo
                 if Checkpoint.chave(f["nome"], dia, mes) not in ja]
        if not nomes:
            continue

        ctx.log(f"—— DIA {dia}/{mes} — {len(nomes)} funcionário(s)", "dia")
        for n in nomes:
            ctx.log(f"  • {n}", "dim")
        ctx.pdf(0, 0)

        pdfs, pastas = localizar_pdfs_do_dia(indice, raiz, dia, mes)
        if not pdfs:
            ctx.log(f"⚠ Nenhuma pasta/PDF encontrado para o dia {dia}/{mes}.", "warn")
            ctx.log("  Nomes aceitos: 06.02, 6.02, 06-02, 06_02, DIA 06 (dentro da pasta do ano ou do mês).", "dim")
            amostra = [os.path.relpath(c, raiz) for c, _, _ in indice[:12]]
            if amostra:
                ctx.log("  Subpastas vistas: " + " | ".join(amostra), "dim")
            else:
                ctx.log("  A pasta escolhida não tem subpastas. Selecione a pasta do ANO ou do MÊS.", "warn")
            for n in nomes:
                nao_enc.append((n, f"{dia}/{mes}", "Pasta/PDF do dia nao encontrado"))
            feitos += len(nomes)
            ctx.geral(feitos, total)
            continue

        ctx.log(f"Pasta encontrada: {' + '.join(pastas)} — {len(pdfs)} PDF(s)", "dim")

        pendentes = {n: limpar(n) for n in nomes}
        achados, candidatos = {}, {}
        for pdf in pdfs:
            if not pendentes:
                break
            if not ctx.checar():
                cancelado = True
                break
            ctx.log(f"📄 {os.path.basename(pdf)}", "pasta")
            res, cand = buscar_no_pdf(ctx, pdf, pendentes)
            for nome, pag in res.items():
                achados[nome] = (pdf, pag, "")
                pendentes.pop(nome, None)
            for nome, (sc, pag) in cand.items():
                if nome not in candidatos or sc > candidatos[nome][0]:
                    candidatos[nome] = (sc, pdf, pag)
        if cancelado:
            break

        # Segunda chance: nome quase identico em pagina sem palavra-chave de ASO
        for nome in list(pendentes):
            if nome in candidatos:
                sc, pdf, pag = candidatos[nome]
                achados[nome] = (pdf, pag, "Pagina sem palavra-chave de ASO - CONFIRA")
                pendentes.pop(nome)
                ctx.log(f"⚠ {nome}: nome achado na pag {pag + 1} de {os.path.basename(pdf)} "
                        f"sem palavra-chave de ASO — confira o arquivo salvo.", "warn")

        for nome, (pdf, pag, obs) in achados.items():
            arq = os.path.join(saida, f"ASO_{nome_para_arquivo(nome)}_{dia}-{mes}.pdf")
            ok, erro = extrair_pagina(pdf, pag, arq)
            if ok:
                ctx.log(f"💾 Salvo: {os.path.basename(arq)}", "ok")
                cp.registrar(nome, dia, mes, arq, pdf, pag)
                detalhes.append({"nome": nome, "data": f"{dia}/{mes}",
                                 "origem": f"{os.path.basename(pdf)} (pag {pag + 1})", "obs": obs})
            else:
                ctx.log(f"✗ Erro ao salvar {nome}: {erro}", "err")
                nao_enc.append((nome, f"{dia}/{mes}", f"Erro ao salvar: {erro}"))

        for nome in pendentes:
            ctx.log(f"✗ Não encontrado: {nome}", "err")
            nao_enc.append((nome, f"{dia}/{mes}", "Nao localizado nos PDFs do dia"))

        feitos += len(nomes)
        ctx.geral(feitos, total)
        gc.collect()

    if cancelado or ctx.cancelar.is_set():
        ctx.log("⏹ Processamento CANCELADO. O progresso salvo no checkpoint foi preservado.", "warn")
        return

    agora = datetime.now()
    ctx.log("=" * 50, "info")
    ctx.log(f"CONCLUÍDO — {agora.strftime('%d/%m/%Y %H:%M:%S')}", "info")
    ctx.log(f"Total: {total} | Encontrados: {len(detalhes)} | Não encontrados: {len(nao_enc)}", "info")
    for n, dt, motivo in nao_enc:
        ctx.log(f"   - {n}  [{dt}]  {motivo}", "dim")
    ctx.log(f"Pasta de saída: {saida}", "dim")
    nome_rel = gerar_relatorio(saida, detalhes, nao_enc, agora)
    if nome_rel:
        ctx.log(f"Relatório salvo: {nome_rel}", "ok")
    else:
        ctx.log("⚠ Não foi possível salvar o relatório Excel (arquivo aberto?).", "warn")
    ctx.geral(total, total)


# ─────────────────────────────────────────────────
#  DEPENDENCIAS
# ─────────────────────────────────────────────────
def verificar_dependencias():
    global IDIOMA_OCR
    faltando = []
    if fitz is None:
        faltando.append("PyMuPDF  (pip install PyMuPDF)")
    if Image is None:
        faltando.append("Pillow  (pip install pillow)")
    if pytesseract is None:
        faltando.append("pytesseract  (pip install pytesseract)")
    else:
        if os.path.exists(TESSERACT_PATH):
            pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
        try:
            langs = pytesseract.get_languages(config="")
            IDIOMA_OCR = "por+eng" if "por" in langs else "eng"
        except Exception:
            faltando.append("Tesseract OCR  (https://github.com/UB-Mannheim/tesseract/wiki)")
    return faltando


# ─────────────────────────────────────────────────
#  INTERFACE (mesmo visual da versao anterior)
# ─────────────────────────────────────────────────
BG = "#12161c"
HEADER = "#0e1116"
FIELD = "#1b2028"
LOGBG = "#1b2129"
BORDA = "#242b35"
TEXT = "#c9d1d9"
SUB = "#7d8590"
BLUE = "#3b82f6"
GREEN = "#2ea043"
AMBER = "#e3a53a"
RED = "#e5534b"
BTN = "#2a3039"
MONO = ("Consolas", 9)
UI = ("Segoe UI", 9)

TAGS = {
    "info": "#58a6ff", "dia": "#6f9bff", "dim": "#8b949e", "pasta": "#a5b4c8",
    "warn": "#d29922", "err": "#f0616b", "ok": "#3fb950", "ts": "#6e7681",
}


class App:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.ctx = None
        self.rodando = False
        self.pausado = False

        root.title("Separador de ASOs")
        root.configure(bg=BG)
        W, H = 1000, 760
        root.minsize(820, 640)
        root.geometry(f"{W}x{H}+{(root.winfo_screenwidth() - W) // 2}+{max(0, (root.winfo_screenheight() - H) // 2 - 20)}")

        self.v_planilha = tk.StringVar()
        self.v_raiz = tk.StringVar()
        self.v_saida = tk.StringVar()
        self._carregar_config()

        self._montar()
        falt = verificar_dependencias()
        self.deps_ok = not falt
        if falt:
            self.lbl_dep.config(text="● Dependências ausentes", fg=RED)
            self.log("Dependências ausentes:\n  - " + "\n  - ".join(falt), "err")
            self.btn_iniciar.config(state="disabled", bg="#3a3f47")
        else:
            self.lbl_dep.config(text="● Dependências OK", fg="#3fb950")
            self.log(f"Sistema pronto (OCR: {IDIOMA_OCR}). Selecione os arquivos e clique em INICIAR.", "dim")

        root.protocol("WM_DELETE_WINDOW", self._fechar)
        root.after(60, self._poll)

    # ── layout ──────────────────────────────────
    def _montar(self):
        r = self.root

        head = tk.Frame(r, bg=HEADER, height=62)
        head.pack(fill="x")
        head.pack_propagate(False)
        tk.Label(head, text="🗂", font=("Segoe UI Emoji", 16), bg=HEADER, fg="white").pack(side="left", padx=(24, 10))
        tk.Label(head, text="SEPARADOR DE ASOs", font=("Segoe UI", 13, "bold"), bg=HEADER, fg="white").pack(side="left")
        self.lbl_dep = tk.Label(head, text="● Verificando...", font=UI, bg=HEADER, fg=SUB)
        self.lbl_dep.pack(side="right", padx=24)

        body = tk.Frame(r, bg=BG)  # empacotado por ultimo (depois do rodape)

        self._secao(body, "ARQUIVOS E PASTAS")
        self._campo(body, "📊", "Planilha de Funcionários (.xlsx)", self.v_planilha, "arquivo")
        self._campo(body, "📁", "Pasta do Ano ou Mês (ex: .../2026 ou .../2026/03)", self.v_raiz, "pasta")
        self._campo(body, "💾", "Pasta de Saída dos ASOs", self.v_saida, "pasta")

        tk.Frame(body, height=1, bg=BORDA).pack(fill="x", pady=(14, 10))

        self._secao(body, "PROGRESSO")
        st = ttk.Style()
        st.theme_use("clam")
        st.configure("aso.Horizontal.TProgressbar", troughcolor=FIELD, background=BLUE,
                     bordercolor=FIELD, lightcolor=BLUE, darkcolor=BLUE, thickness=8)
        self.pb_geral, self.lb_geral = self._barra(body, "Geral")
        self.pb_pdf, self.lb_pdf = self._barra(body, "PDF atual")

        tk.Frame(body, height=1, bg=BORDA).pack(fill="x", pady=(10, 8))
        self._secao(body, "LOG")

        moldura = tk.Frame(body, bg=LOGBG)
        moldura.pack(fill="both", expand=True, pady=(0, 10))
        self.txt = tk.Text(moldura, height=8, font=MONO, bg=LOGBG, fg=TEXT, relief="flat", wrap="word",
                           padx=12, pady=10, state="disabled", highlightthickness=0, bd=0,
                           insertbackground=TEXT, selectbackground=BORDA)
        sb = tk.Scrollbar(moldura, command=self.txt.yview)
        self.txt.config(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt.pack(side="left", fill="both", expand=True)
        for nome, cor in TAGS.items():
            self.txt.tag_config(nome, foreground=cor)

        rod = tk.Frame(r, bg=HEADER, height=68)
        rod.pack(fill="x", side="bottom")
        rod.pack_propagate(False)
        tk.Button(rod, text="🗑", font=("Segoe UI Emoji", 11), bg=BTN, fg=SUB, relief="flat",
                  activebackground="#39414d", cursor="hand2", width=4,
                  command=self._limpar_log).pack(side="left", padx=(22, 0), pady=14, ipady=4)

        self.btn_iniciar = tk.Button(rod, text="▶   INICIAR", font=("Segoe UI", 10, "bold"), bg=GREEN, fg="white",
                                     relief="flat", activebackground="#3fb950", activeforeground="white",
                                     cursor="hand2", width=16, command=self.iniciar)
        self.btn_iniciar.pack(side="right", padx=(8, 22), pady=14, ipady=6)
        self.btn_parar = tk.Button(rod, text="⏹   PARAR", font=("Segoe UI", 10, "bold"), bg=RED, fg="white",
                                   relief="flat", activebackground="#c9443c", activeforeground="white",
                                   cursor="hand2", width=14, state="disabled", command=self.parar)
        self.btn_parar.pack(side="right", padx=8, pady=14, ipady=6)
        self.btn_pausar = tk.Button(rod, text="⏸   PAUSAR", font=("Segoe UI", 10, "bold"), bg=AMBER, fg="#1a1a1a",
                                    relief="flat", activebackground="#f0b445", activeforeground="#1a1a1a",
                                    cursor="hand2", width=14, state="disabled", command=self.pausar)
        self.btn_pausar.pack(side="right", padx=8, pady=14, ipady=6)

        # corpo empacotado por ultimo: ocupa so o espaco que sobra e nunca empurra o rodape para fora
        body.pack(fill="both", expand=True, padx=22, pady=(14, 0))

    def _secao(self, parent, titulo):
        tk.Label(parent, text=titulo, font=("Segoe UI", 8, "bold"), bg=BG, fg=SUB, anchor="w").pack(fill="x", pady=(0, 6))

    def _campo(self, parent, icone, rotulo, var, modo):
        tk.Label(parent, text=f"{icone}  {rotulo}", font=UI, bg=BG, fg=TEXT, anchor="w").pack(fill="x", pady=(2, 3))
        linha = tk.Frame(parent, bg=BG)
        linha.pack(fill="x", pady=(0, 6))
        tk.Entry(linha, textvariable=var, font=MONO, bg=FIELD, fg=TEXT, readonlybackground=FIELD,
                 relief="flat", bd=0, highlightthickness=0, state="readonly").pack(
            side="left", fill="x", expand=True, ipady=7, padx=(0, 8))

        def procurar():
            atual = var.get()
            ini = os.path.dirname(atual) if modo == "arquivo" and atual else (atual or None)
            if modo == "arquivo":
                p = askopenfilename(title=rotulo, initialdir=ini, filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")])
            else:
                p = askdirectory(title=rotulo, initialdir=ini)
            if p:
                var.set(p)

        tk.Button(linha, text="Procurar", font=UI, bg=BTN, fg=TEXT, relief="flat", activebackground="#39414d",
                  activeforeground=TEXT, cursor="hand2", width=9, command=procurar).pack(side="right", ipady=4)

    def _barra(self, parent, rotulo):
        linha = tk.Frame(parent, bg=BG)
        linha.pack(fill="x", pady=3)
        tk.Label(linha, text=rotulo, font=UI, bg=BG, fg=SUB, width=9, anchor="w").pack(side="left")
        pb = ttk.Progressbar(linha, style="aso.Horizontal.TProgressbar", maximum=100, mode="determinate")
        pb.pack(side="left", fill="x", expand=True, padx=(4, 10))
        lb = tk.Label(linha, text="0/0", font=UI, bg=BG, fg=SUB, width=9, anchor="e")
        lb.pack(side="right")
        return pb, lb

    # ── config ──────────────────────────────────
    def _carregar_config(self):
        try:
            with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as f:
                c = json.load(f)
            self.v_planilha.set(c.get("planilha", ""))
            self.v_raiz.set(c.get("raiz", ""))
            self.v_saida.set(c.get("saida", ""))
        except Exception:
            pass

    def _salvar_config(self):
        try:
            with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
                json.dump({"planilha": self.v_planilha.get(), "raiz": self.v_raiz.get(),
                           "saida": self.v_saida.get()}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ── log / progresso ─────────────────────────
    def log(self, msg, tag="dim"):
        self.txt.config(state="normal")
        self.txt.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] ", "ts")
        self.txt.insert("end", msg + "\n", tag)
        if int(self.txt.index("end-1c").split(".")[0]) > 6000:
            self.txt.delete("1.0", "1000.0")
        self.txt.see("end")
        self.txt.config(state="disabled")

    def _limpar_log(self):
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")

    @staticmethod
    def _set_barra(pb, lb, a, b):
        pb["value"] = (a / b * 100) if b else 0
        lb.config(text=f"{a}/{b}")

    def _poll(self):
        try:
            for _ in range(300):
                ev = self.q.get_nowait()
                if ev[0] == "log":
                    self.log(ev[1], ev[2])
                elif ev[0] == "geral":
                    self._set_barra(self.pb_geral, self.lb_geral, ev[1], ev[2])
                elif ev[0] == "pdf":
                    self._set_barra(self.pb_pdf, self.lb_pdf, ev[1], ev[2])
                elif ev[0] == "fim":
                    self._finalizar()
        except queue.Empty:
            pass
        self.root.after(60, self._poll)

    # ── botoes ──────────────────────────────────
    def iniciar(self):
        p, m, s = self.v_planilha.get().strip(), self.v_raiz.get().strip(), self.v_saida.get().strip()
        if not p or not os.path.isfile(p):
            return messagebox.showwarning("Atenção", "Selecione a planilha de funcionários (.xlsx).", parent=self.root)
        if not m or not os.path.isdir(m):
            return messagebox.showwarning("Atenção", "Selecione a pasta do ano ou do mês.", parent=self.root)
        if not s:
            return messagebox.showwarning("Atenção", "Selecione a pasta de saída.", parent=self.root)
        self._salvar_config()

        usar_cp = True
        try:
            n = Checkpoint(s).total()
            if n > 0:
                usar_cp = messagebox.askyesno(
                    "Checkpoint detectado",
                    f"{n} ASO(s) já foram processados nesta pasta de saída.\n\n"
                    "• Sim  → continuar de onde parou (recomendado)\n"
                    "• Não  → recomeçar tudo do zero", parent=self.root)
        except Exception:
            usar_cp = True

        self._limpar_log()
        self._set_barra(self.pb_geral, self.lb_geral, 0, 0)
        self._set_barra(self.pb_pdf, self.lb_pdf, 0, 0)
        self.ctx = Ctx(self.q)
        self.rodando, self.pausado = True, False
        self.btn_iniciar.config(state="disabled", text="⏳  Processando...", bg="#3a3f47", fg=SUB)
        self.btn_pausar.config(state="normal", text="⏸   PAUSAR", bg=AMBER, fg="#1a1a1a")
        self.btn_parar.config(state="normal")
        threading.Thread(target=executar, args=(self.ctx, p, m, s, usar_cp), daemon=True).start()

    def pausar(self):
        if not self.ctx:
            return
        if not self.pausado:
            self.ctx.pausa.clear()
            self.pausado = True
            self.btn_pausar.config(text="▶   RETOMAR", bg=GREEN, fg="white")
            self.log("⏸ Processamento PAUSADO.", "warn")
        else:
            self.ctx.pausa.set()
            self.pausado = False
            self.btn_pausar.config(text="⏸   PAUSAR", bg=AMBER, fg="#1a1a1a")
            self.log("▶ Processamento RETOMADO.", "info")

    def parar(self):
        if self.ctx and messagebox.askyesno(
                "Parar?", "Deseja parar o processamento?\nO progresso já salvo no checkpoint será preservado.",
                parent=self.root):
            self.ctx.cancelar.set()
            self.ctx.pausa.set()
            self.btn_parar.config(state="disabled")
            self.btn_pausar.config(state="disabled")
            self.log("⏹ Parando... (termina a página atual)", "warn")

    def _finalizar(self):
        self.rodando = False
        self.btn_iniciar.config(state="normal", text="▶   INICIAR", bg=GREEN, fg="white")
        self.btn_pausar.config(state="disabled", text="⏸   PAUSAR", bg=AMBER, fg="#1a1a1a")
        self.btn_parar.config(state="disabled")

    def _fechar(self):
        if self.ctx:
            self.ctx.cancelar.set()
            self.ctx.pausa.set()
        self.root.destroy()


def main():
    try:  # texto nitido em telas com escala de DPI (Windows)
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()