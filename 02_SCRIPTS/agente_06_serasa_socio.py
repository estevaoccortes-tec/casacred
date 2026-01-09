#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
agente_06_serasa_socio.py

Organizador SERASA (SÓCIO) -> Excel (STRICT/DUAL-SOURCE)

SAÍDA:
- 1 .xlsx com 1 aba "Planilha1"
- Colunas fixas: Campo | Informação | CPF
- Fonte PDF (exclusivo): Identificação, Participações Societárias, Anotações Negativas, CPF
- Fonte IMAGEM (exclusivo): Consultas por mês (gráfico)
"""

from __future__ import annotations

import argparse
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd
import pdfplumber

# OCR (só para IMAGEM e para páginas do PDF que vierem como imagem)
import pytesseract
from pdf2image import convert_from_path

# OpenCV é opcional, mas recomendado (melhor para detectar barras azuis).
# Se não existir no ambiente, o script ainda gera o XLSX sem Consultas.
try:
    import cv2
    import numpy as np

    HAS_CV2 = True
except Exception:
    HAS_CV2 = False


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = BASE_DIR / "01_INPUT"
DEFAULT_OUTPUT = BASE_DIR / "03_OUTPUT" / "6. SERASA SÓCIO"
DEFAULT_POPPLER = Path(r"C:\Users\Usuario\Desktop\poppler-25.12.0\Library\bin")

try:
    from agente_02_serasacedente import (
        extract_one_pdf as cedente_extract_one_pdf,
        POPPLER_PATH as CEDENTE_POPPLER,
        START_LABEL as CEDENTE_START_LABEL,
        EXPECTED_BARS as CEDENTE_EXPECTED_BARS,
        DPI_GRAFICO as CEDENTE_DPI_GRAFICO,
    )
except Exception:
    cedente_extract_one_pdf = None
    CEDENTE_POPPLER = str(DEFAULT_POPPLER)
    CEDENTE_START_LABEL = "Nov/2024"
    CEDENTE_EXPECTED_BARS = 13
    CEDENTE_DPI_GRAFICO = 200


# =========================
# Regex / Constantes
# =========================

CPF_RE = re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")
CNPJ_RE = re.compile(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b")
DATE_RE = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")
PERCENT_RE = re.compile(r"\b\d{1,3}(?:[.,]\d{1,2})?\s*%")
MONEY_RE = re.compile(r"(R\$\s*)?\d{1,3}(?:[\.,]\d{3})*(?:[\.,]\d{2})")

UF_SET = {
    "AC",
    "AL",
    "AP",
    "AM",
    "BA",
    "CE",
    "DF",
    "ES",
    "GO",
    "MA",
    "MT",
    "MS",
    "MG",
    "PA",
    "PB",
    "PR",
    "PE",
    "PI",
    "RJ",
    "RN",
    "RS",
    "RO",
    "RR",
    "SC",
    "SP",
    "SE",
    "TO",
}

PT_MONTHS = {
    "jan": "01",
    "fev": "02",
    "mar": "03",
    "abr": "04",
    "mai": "05",
    "jun": "06",
    "jul": "07",
    "ago": "08",
    "set": "09",
    "out": "10",
    "nov": "11",
    "dez": "12",
}

DEFAULT_START_LABEL = "Dez/2024"

IDENT_FIELDS = [
    "Situação na Receita Federal",
    "Nome completo",
    "Município/UF",
    "Serasa Score",
    "Liminar",
    "Probabilidade de pagamento em 12 meses",
    "Total em anotações negativas",
]

# Regras de Liminar (somente olhando texto do bloco Anotações Negativas do PDF)
LIMINAR_KEYS = [
    "NADA CONSTA",
    "LIMINAR",
    "ART. 43",
    "DECISÃO JUDICIAL",
    "DECISAO JUDICIAL",
    "RJ",
    "SUSPENSÃO",
    "SUSPENSAO",
    "BLOQUEIO JUDICIAL",
]


# Validação de Campo: somente whitelist/padrões do prompt
def _campo_valido(campo: str) -> bool:
    if "–" in campo or "—" in campo:
        return False

    exatos = {
        "Situação na Receita Federal",
        "Nome completo",
        "Município/UF",
        "Serasa Score",
        "Liminar",
        "Probabilidade de pagamento em 12 meses",
        "Total em anotações negativas",
        "Cheques - Motivo",
    }
    if campo in exatos:
        return True

    # Consultas - MM/AAAA
    if re.fullmatch(r"Consultas - \d{2}/\d{4}", campo):
        return True
    # Consulta N (tabela de consultas por data)
    if re.fullmatch(r"Consulta \d+", campo):
        return True

    # Participação N - ...
    if re.fullmatch(r"Participação \d+ - (CNPJ|Capital|Situação Cadastral|UF|Razão Social)", campo):
        return True

    # PEFIN / REFIN
    if re.fullmatch(r"(PEFIN|REFIN) - Registro \d+ \((Modalidade|Valor|Origem|Data)\)", campo):
        return True

    # Protestos
    if re.fullmatch(r"Protestos - Registro \d+ \((Data|Valor)\)", campo):
        return True

    # Ações Judiciais
    if re.fullmatch(r"Ações Judiciais - Registro \d+ \((Data|Valor|Natureza|Cidade)\)", campo):
        return True

    return False


def _strip(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\u00a0", " ")).strip()


def _no_accents(s: str) -> str:
    s = _strip(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s


def _no_accents_upper(s: str) -> str:
    return _no_accents(s).upper()


def _safe_filename(s: str) -> str:
    s2 = _no_accents(s)
    s2 = re.sub(r"[^\w\s\-\.]", "", s2, flags=re.UNICODE)
    s2 = _strip(s2).replace(" ", "_")
    return s2 or "SOCIO"


def _ptbr_money_from_any(raw: str) -> Optional[str]:
    """
    Normaliza dinheiro para PT-BR "x.xxx.xxx,xx" como TEXTO.
    Aceita "R$ 1.234,56", "1,234.56", "1234,56" etc.
    """
    s = _strip(raw)
    if not s:
        return None

    s = s.replace("R$", "").replace(" ", "")
    if s in {"-", "--", "–", "—"}:
        return None
    s2 = re.sub(r"[^0-9\.,]", "", s)
    if not s2:
        return None

    last_comma = s2.rfind(",")
    last_dot = s2.rfind(".")

    if last_comma == -1 and last_dot == -1:
        try:
            val = int(s2)
            inteiro_pt = f"{val:,}".replace(",", "X").replace(".", ",").replace("X", ".")
            return f"{inteiro_pt},00"
        except Exception:
            return None

    dec_sep = "," if last_comma > last_dot else "."
    thou_sep = "." if dec_sep == "," else ","

    parts = s2.split(dec_sep)
    inteiro = parts[0].replace(thou_sep, "")
    if not inteiro.isdigit():
        return None

    frac = parts[1] if len(parts) > 1 else ""
    frac = re.sub(r"\D", "", frac)
    if len(frac) == 0:
        frac = "00"
    elif len(frac) == 1:
        frac = frac + "0"
    else:
        frac = frac[:2]

    val_int = int(inteiro)
    inteiro_pt = f"{val_int:,}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{inteiro_pt},{frac}"


# =========================
# PDF -> Texto (com OCR pontual por página)
# =========================


def _pdf_page_text_native(page) -> str:
    # layout=True costuma melhorar "cards" no Serasa
    return page.extract_text(layout=True, x_tolerance=1, y_tolerance=2) or ""


def _pdf_page_text_ocr(pdf_path: Path, page_index0: int, poppler_path: Path, dpi: int = 260) -> str:
    # OCR só de 1 página específica
    imgs = convert_from_path(
        str(pdf_path),
        dpi=dpi,
        poppler_path=str(poppler_path),
        first_page=page_index0 + 1,
        last_page=page_index0 + 1,
        fmt="png",
    )
    if not imgs:
        return ""
    img = imgs[0]
    txt = pytesseract.image_to_string(img, config="--oem 3 --psm 6") or ""
    return txt


def _text_signal(text: str) -> int:
    t = text or ""
    t_up = _no_accents_upper(t)
    keys = [
        "SITUACAO NA RECEITA FEDERAL",
        "NOME COMPLETO",
        "MUNICIPIO/UF",
        "SERASA SCORE",
        "PROBABILIDADE DE PAGAMENTO",
    ]
    score = sum(1 for k in keys if k in t_up)
    if CPF_RE.search(t):
        score += 1
    return score


def extract_pdf_page_texts(pdf_path: Path, poppler_path: Optional[Path], ocr_dpi: int = 260) -> List[str]:
    """
    Estratégia:
    - tenta texto nativo página a página
    - se uma página vier "vazia" (muito pouco texto) e tiver poppler configurado, faz OCR só nela
    """
    parts: List[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for i, page in enumerate(pdf.pages):
            t = _pdf_page_text_native(page)
            if poppler_path is None:
                parts.append(t)
                continue
            if len(_strip(t)) < 80:
                t_ocr = _pdf_page_text_ocr(pdf_path, i, poppler_path=poppler_path, dpi=ocr_dpi)
                use = t_ocr if len(_strip(t_ocr)) > len(_strip(t)) else t
                parts.append(use)
                continue

            t_ocr = _pdf_page_text_ocr(pdf_path, i, poppler_path=poppler_path, dpi=ocr_dpi)
            score_native = _text_signal(t)
            score_ocr = _text_signal(t_ocr)
            if score_ocr > score_native:
                parts.append(t_ocr)
            else:
                parts.append(t)
    return parts


def extract_pdf_text(pdf_path: Path, poppler_path: Optional[Path], ocr_dpi: int = 260) -> str:
    return "\n".join(extract_pdf_page_texts(pdf_path, poppler_path=poppler_path, ocr_dpi=ocr_dpi))


# =========================
# CPF (somente PDF)
# =========================


def extract_cpf_from_pdf_text(text: str) -> str:
    m = CPF_RE.search(text or "")
    return m.group(0) if m else "A confirmar"


def extract_cpf_from_any(text: str) -> Optional[str]:
    m = CPF_RE.search(text or "")
    return m.group(0) if m else None


# =========================
# Identificação (somente PDF)
# =========================


def _value_after_label(text: str, label: str) -> Optional[str]:
    """
    Procura padrão: "<label> <valor>" no texto linear.
    Funciona bem para o Serasa porque o PDF geralmente imprime card + valor na mesma região.
    """
    t = text or ""
    # tolera ":" e quebra
    pat = rf"{re.escape(label)}\s*[:\-]?\s*(.+)"
    m = re.search(pat, t, flags=re.IGNORECASE)
    if not m:
        return None
    val = _strip(m.group(1))
    # corta quando "gruda" outro rótulo ao lado
    val = re.split(r"\s{2,}|\bOcultar\b|\bParticipa", val, flags=re.IGNORECASE)[0]
    val = _strip(val)
    return val or None


def _is_bad_nome_line(line: str) -> bool:
    up = _no_accents_upper(line)
    if not line:
        return True
    bad_keys = [
        "CPF",
        "DATA DE NASCIMENTO",
        "NOME DA MAE",
        "SEXO",
        "SITUACAO",
        "SITUAÇÃO",
    ]
    return any(k in up for k in bad_keys)


def _clean_nome(line: str) -> str:
    if not line:
        return ""
    cut = re.split(
        r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b|\b\d{2}/\d{2}/\d{4}\b|\b\d{8}\b",
        line,
        flags=re.IGNORECASE,
    )[0]
    cut = re.split(
        r"\bCPF\b|\bData\s+de\s+Nascimento\b|\bData\s+de\s+nasc\b|\bNome\s+da\s+M[ãa]e\b|\bSexo\b",
        cut,
        flags=re.IGNORECASE,
    )[0]
    return _strip(cut)


def _extract_nome_completo(text_pdf: str) -> Optional[str]:
    lines = [l.strip() for l in (text_pdf or "").splitlines() if l.strip()]
    for i, ln in enumerate(lines):
        up = _no_accents_upper(ln)
        if "NOME COMPLETO" in up:
            tail = re.split(r"nome completo", ln, flags=re.IGNORECASE)[-1]
            tail = _strip(tail.strip(" :-"))
            tail = _clean_nome(tail)
            if tail and not _is_bad_nome_line(tail):
                return tail
            for j in range(i + 1, min(i + 4, len(lines))):
                cand = _clean_nome(_strip(lines[j]))
                if cand and not _is_bad_nome_line(cand):
                    return cand
    return None


def extract_identificacao(text_pdf: str) -> Dict[str, str]:
    out: Dict[str, str] = {}

    # Situação na Receita Federal
    sit = _value_after_label(text_pdf, "Situação na Receita Federal")
    if sit:
        # normalmente vem "REGULAR"
        sit = _strip(sit.split(" ")[0] if len(sit.split()) > 3 else sit)
    out["Situação na Receita Federal"] = sit or "A confirmar"

    # Nome completo
    nome = _extract_nome_completo(text_pdf) or _clean_nome(_value_after_label(text_pdf, "Nome completo") or "")
    if not nome:
        # fallback: muitas vezes o nome aparece logo no topo como "Detalhes do CPF ... <NOME>"
        m = re.search(
            r"Detalhes\s+do\s+CPF\s+.+?\s{1,}([A-ZÁÉÍÓÚÂÊÔÃÕÇ][^\n]+)",
            text_pdf,
            flags=re.IGNORECASE,
        )
        if m:
            nome = _strip(m.group(1))
            nome = re.split(r"\bCPF\b|\bRegular\b|\bIrregular\b", nome, flags=re.IGNORECASE)[0].strip()
            if _is_bad_nome_line(nome):
                nome = None
    out["Nome completo"] = nome or "A confirmar"

    # Município/UF
    mun = _value_after_label(text_pdf, "Município/UF")
    if not mun:
        mun = _value_after_label(text_pdf, "Municipio/UF")
    out["Município/UF"] = mun or "A confirmar"

    # Serasa Score
    score = _value_after_label(text_pdf, "Serasa Score")
    if score:
        m = re.search(r"\b(\d{1,4})\b", score)
        score = m.group(1) if m else score
    out["Serasa Score"] = score or "A confirmar"

    # Probabilidade de pagamento em 12 meses
    prob = _value_after_label(text_pdf, "Probabilidade de pagamento em 12 meses")
    if prob:
        m = PERCENT_RE.search(prob)
        prob = m.group(0).replace(" ", "") if m else prob
    out["Probabilidade de pagamento em 12 meses"] = prob or "A confirmar"

    # Total em anotações negativas (precisa ser MONETÁRIO)
    tot_neg = _value_after_label(text_pdf, "Total em anotações negativas")
    val_money = None
    if tot_neg:
        mm = MONEY_RE.search(tot_neg)
        if mm:
            val_money = _ptbr_money_from_any(mm.group(0))
    # fallback: alguns PDFs mostram "Total de dívidas: R$ ..." dentro do bloco de anotações
    if val_money is None:
        m2 = re.search(
            r"Total\s+de\s+d[ií]vidas\s*[:\-]?\s*(R\$\s*[\d\.,]+)",
            text_pdf,
            flags=re.IGNORECASE,
        )
        if m2:
            val_money = _ptbr_money_from_any(m2.group(1))
    bloco_anot = extract_block(
        text_pdf,
        "Anotações negativas",
        stop_any=[
            "Participações societárias",
            "Consultas por mês",
            "Consultas por mes",
            "Dados do CPF",
            "Detalhes das participações",
        ],
    )
    sem_regs = _has_sem_registros(bloco_anot) if bloco_anot else False
    if val_money is not None:
        if val_money == "0,00":
            out["Total em anotações negativas"] = "Sem registro"
            sem_regs = True
        else:
            out["Total em anotações negativas"] = val_money
    else:
        out["Total em anotações negativas"] = "Sem registro" if sem_regs else "A confirmar"

    # Liminar (somente do bloco Anotações Negativas)
    if not bloco_anot:
        out["Liminar"] = "Sem registro"
    elif sem_regs:
        out["Liminar"] = "Sem registro"
    else:
        up = _no_accents_upper(bloco_anot)
        found = any(_no_accents_upper(k) in up for k in LIMINAR_KEYS)
        out["Liminar"] = "Sim" if found else "Não"

    return out


def _words_to_lines(words: List[dict], y_tol: float = 2.8) -> List[Tuple[float, str]]:
    if not words:
        return []
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines: List[List[dict]] = []
    cur: List[dict] = []
    cur_top: Optional[float] = None
    for w in words:
        if cur_top is None or abs(w["top"] - cur_top) <= y_tol:
            cur.append(w)
            cur_top = w["top"] if cur_top is None else (cur_top + w["top"]) / 2.0
        else:
            lines.append(sorted(cur, key=lambda x: x["x0"]))
            cur = [w]
            cur_top = w["top"]
    if cur:
        lines.append(sorted(cur, key=lambda x: x["x0"]))

    out: List[Tuple[float, str]] = []
    for ln in lines:
        text = _strip(" ".join(w.get("text", "") for w in ln))
        if text:
            out.append((ln[0]["top"], text))
    return out


def extract_identificacao_from_pdf(pdf_path: Path, page_indices: List[int]) -> Dict[str, str]:
    out: Dict[str, str] = {k: "A confirmar" for k in IDENT_FIELDS}
    if not page_indices:
        return out

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_idx in page_indices:
            if page_idx < 0 or page_idx >= len(pdf.pages):
                continue
            page = pdf.pages[page_idx]
            words = page.extract_words(x_tolerance=1, y_tolerance=2) or []
            lines = _words_to_lines(words)
            if not lines:
                continue

            # Nome completo e CPF (linha logo abaixo do cabeçalho "Nome completo")
            for i, (_, line) in enumerate(lines):
                if "NOME COMPLETO" in _no_accents_upper(line):
                    for j in range(i + 1, min(i + 5, len(lines))):
                        _, val_line = lines[j]
                        mcpf = CPF_RE.search(val_line)
                        if not mcpf:
                            continue
                        nome_raw = val_line.split(mcpf.group(0))[0]
                        nome = _clean_nome(nome_raw)
                        if nome:
                            out["Nome completo"] = nome
                        break
                    break

            # Situação na Receita Federal + Município/UF (linha logo abaixo do cabeçalho)
            for i, (_, line) in enumerate(lines):
                if "MUNICIPIO/UF" in _no_accents_upper(line):
                    if i + 1 < len(lines):
                        _, val_line = lines[i + 1]
                        m_mun = re.search(r"[A-ZÁÉÍÓÚÂÊÔÃÕÇ]+(?:\s+[A-ZÁÉÍÓÚÂÊÔÃÕÇ]+)*/[A-Z]{2}", val_line)
                        if m_mun:
                            out["Município/UF"] = _strip(m_mun.group(0))
                        m_sit = re.search(r"\b(REGULAR|IRREGULAR|INAPTA|SUSPENSA|ATIVA|ATIVO)\b", _no_accents_upper(val_line))
                        if m_sit:
                            out["Situação na Receita Federal"] = m_sit.group(1).title()
                    break

            # Serasa Score + Probabilidade
            for _, line in lines:
                up_line = _no_accents_upper(line)
                has_hint = "CHANCE DE PAGAMENTO" in up_line or "SERASA SCORE" in up_line
                m_prob = re.search(r"\b(\d{1,3},\d{2})\b", line)
                if has_hint or m_prob:
                    nums = re.findall(r"\b(\d{3})\b", line)
                    score = None
                    if nums:
                        if "CHANCE DE PAGAMENTO" in up_line:
                            score = nums[0]
                        elif "SERASA SCORE" in up_line:
                            score = nums[0]
                        else:
                            for n in nums:
                                if n not in {"701", "800"}:
                                    score = n
                                    break
                    if score:
                        out["Serasa Score"] = score
                    if m_prob:
                        out["Probabilidade de pagamento em 12 meses"] = f"{m_prob.group(1)}%"
                    if score or m_prob:
                        break

            # Sai assim que preencher os principais
            if out.get("Nome completo") != "A confirmar" or out.get("Serasa Score") != "A confirmar":
                break

    return out


# Consultas (tabela: data + consultante)
def _extract_consultas_tabela_ocr(
    pdf_path: Path,
    page_indices: List[int],
    poppler_path: Optional[Path],
) -> List[Tuple[str, str]]:
    if poppler_path is None:
        return []

    pages = sorted(set(page_indices))
    if not pages:
        return []

    first_page = min(pages) + 1
    last_page = max(pages) + 1

    try:
        images = convert_from_path(
            str(pdf_path),
            dpi=300,
            poppler_path=str(poppler_path),
            first_page=first_page,
            last_page=last_page,
            fmt="png",
        )
    except Exception:
        return []

    results: List[Tuple[str, str]] = []
    for img in images:
        txt = pytesseract.image_to_string(img, lang="por+eng", config="--oem 3 --psm 6") or ""
        for line in (txt.splitlines() or []):
            line = _strip(line)
            if not line:
                continue
            m = DATE_RE.search(line)
            if not m:
                continue
            date = m.group(0)
            rest = line[m.end() :].strip()
            rest = CNPJ_RE.sub(" ", rest)
            rest = re.sub(r"\b\d{1,3}\b$", " ", rest)
            rest = re.sub(r"\b\d+\b", " ", rest)
            rest = _strip(rest)
            if rest:
                results.append((date, rest))

    # dedup preserve order
    seen = set()
    out: List[Tuple[str, str]] = []
    for d, n in results:
        key = f"{d}|{n}"
        if key in seen:
            continue
        seen.add(key)
        out.append((d, n))
    return out


def _extract_consultas_tabela_text(pdf_text: str) -> List[Tuple[str, str]]:
    block = extract_block(
        pdf_text or "",
        "Consultas",
        stop_any=[
            "Participações societárias",
            "Anotações negativas",
            "Consultas por mês",
            "Consultas por mes",
            "Dados do CPF",
            "Detalhes das participações",
        ],
    )
    if not block:
        return []

    lines = [_strip(l) for l in (block.splitlines() or [])]
    results: List[Tuple[str, str]] = []

    for i, line in enumerate(lines):
        if not line:
            continue
        m = DATE_RE.search(line)
        if not m:
            continue
        date = m.group(0)
        rest = _strip(line[m.end() :])
        if not rest:
            j = i + 1
            while j < len(lines) and not _strip(lines[j]):
                j += 1
            if j < len(lines):
                rest = _strip(lines[j])
        rest = CPF_RE.sub(" ", rest)
        rest = CNPJ_RE.sub(" ", rest)
        rest = re.sub(r"\b\d{1,4}\b", " ", rest)
        rest = _strip(rest)
        if rest:
            results.append((date, rest))

    seen = set()
    out: List[Tuple[str, str]] = []
    for d, n in results:
        key = f"{d}|{n}"
        if key in seen:
            continue
        seen.add(key)
        out.append((d, n))
    return out


def extract_consultas_tabela(
    pdf_path: Path,
    page_indices: List[int],
    poppler_path: Optional[Path],
    pdf_text: str = "",
) -> List[Tuple[str, str]]:
    resultados: List[Tuple[str, str]] = []

    def _is_header_row(row: List[str]) -> bool:
        header = " ".join([c or "" for c in row]).lower()
        return "data da consulta" in header or ("data" in header and "consulta" in header)

    def _rows_from_table(table: List[List[str]]) -> List[Tuple[str, str]]:
        rows_out: List[Tuple[str, str]] = []
        if not table:
            return rows_out
        start_idx = 1 if table and _is_header_row(table[0]) else 0
        for row in table[start_idx:]:
            if not row:
                continue
            cells = [(c or "").strip() for c in row]
            date = None
            for c in cells:
                m = DATE_RE.search(c)
                if m:
                    date = m.group(0)
                    break
            if not date:
                continue
            consultante = ""
            for c in reversed(cells):
                if not c:
                    continue
                if DATE_RE.search(c):
                    continue
                if re.search(r"[A-Za-zÁÉÍÓÚÂÊÔÃÕÇ]", c):
                    consultante = c
                    break
            consultante = _strip(consultante)
            if consultante:
                rows_out.append((date, consultante))
        return rows_out

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_idx in page_indices:
            if page_idx < 0 or page_idx >= len(pdf.pages):
                continue
            page = pdf.pages[page_idx]
            tables = page.extract_tables()
            for table in tables or []:
                resultados.extend(_rows_from_table(table))

    if not resultados:
        resultados = _extract_consultas_tabela_ocr(pdf_path, page_indices, poppler_path)
    if len(resultados) < 5 and pdf_text:
        extra = _extract_consultas_tabela_text(pdf_text)
        if extra:
            seen = set()
            merged: List[Tuple[str, str]] = []
            for d, n in resultados + extra:
                key = f"{d}|{n}"
                if key in seen:
                    continue
                seen.add(key)
                merged.append((d, n))
            resultados = merged
    return resultados


# =========================
# Utilitário: extrair blocos por âncoras
# =========================


def extract_block(text: str, start: str, stop_any: List[str], max_chars: int = 25000) -> str:
    t = text or ""
    s_up = _no_accents_upper(start)
    idx = _no_accents_upper(t).find(s_up)
    if idx < 0:
        return ""
    cut = t[idx : idx + max_chars]
    # corta no primeiro stop encontrado
    cut_up = _no_accents_upper(cut)
    stops = []
    for st in stop_any:
        j = cut_up.find(_no_accents_upper(st))
        if j > 0:
            stops.append(j)
    if stops:
        cut = cut[: min(stops)]
    return cut

# =========================
# Participações Societárias (somente PDF)
# =========================


@dataclass
class Participacao:
    cnpj: str
    razao: str
    capital: str
    uf: str
    situacao: str


def _join_cnpj_fragments(tokens: List[str]) -> Optional[str]:
    """
    Une casos do tipo: "32.434.675/0001-" + "35" => "32.434.675/0001-35"
    """
    for i in range(len(tokens) - 1):
        a = tokens[i]
        b = tokens[i + 1]
        if re.fullmatch(r"\d{2}\.\d{3}\.\d{3}/\d{4}-", a) and re.fullmatch(r"\d{2}", b):
            c = a + b
            if CNPJ_RE.fullmatch(c):
                return c
    return None


def _parse_participacao_record(record_text: str) -> Participacao:
    """
    Record_text é um "bloco" de linhas concatenadas para 1 participação.
    A regra é: só registrar o que estiver explícito.
    """
    s = _strip(record_text)

    tokens = [t for t in re.split(r"\s+", s) if t]
    cnpj = None
    m = CNPJ_RE.search(s)
    if m:
        cnpj = m.group(0)
    else:
        cnpj = _join_cnpj_fragments(tokens)

    # capital
    capital = "A confirmar"
    perc = PERCENT_RE.findall(s)
    if perc:
        # escolhe o percentual mais "coerente" (0-100)
        best = None
        for p in perc:
            p2 = p.replace(" ", "").replace("%", "")
            try:
                v = float(p2.replace(",", "."))
                if 0 <= v <= 100:
                    best = p.replace(" ", "")
                    break
            except Exception:
                continue
        capital = best or "A confirmar"

    # UF
    uf = "A confirmar"
    for tk in tokens:
        up = _no_accents_upper(tk)
        if up in UF_SET:
            uf = up
            break

    # situação cadastral (Ativa / Inapta etc.)
    situacao = "A confirmar"
    m_sit = re.search(r"\b(ATIVA|ATIVO|INAPTA|BAIXADA|SUSPENSA|NULA)\b", _no_accents_upper(s))
    if m_sit:
        # preserva forma "Ativa" como no PDF? Aqui guardo com inicial maiúscula
        raw = m_sit.group(1)
        situacao = "Ativa" if raw.startswith("ATIV") else raw.capitalize()

    # razão social: remove cnpj, capital, datas, uf, situação e restos óbvios
    razao = s
    if cnpj:
        razao = re.sub(re.escape(cnpj), " ", razao)
    razao = PERCENT_RE.sub(" ", razao)
    razao = DATE_RE.sub(" ", razao)
    if uf != "A confirmar":
        razao = re.sub(rf"\b{re.escape(uf)}\b", " ", razao)
    if situacao != "A confirmar":
        razao = re.sub(rf"\b{re.escape(situacao)}\b", " ", razao, flags=re.IGNORECASE)
    # remove números soltos longos
    razao = re.sub(r"\b\d{4,}\b", " ", razao)
    razao = _strip(re.sub(r"\s+", " ", razao))
    # se ficar vazio, "A confirmar"
    razao = razao if razao else "A confirmar"

    return Participacao(
        cnpj=cnpj or "A confirmar",
        razao=razao,
        capital=capital,
        uf=uf,
        situacao=situacao,
    )


def extract_participacoes(pdf_path: Path, page_indices: Optional[Iterable[int]] = None) -> List[Participacao]:
    """
    Extrai participações via posições de coluna (CNPJ/Razão/Capital/UF/Situação).
    """
    participacoes: List[Participacao] = []
    page_set = set(page_indices) if page_indices is not None else None

    stop_keys = {
        "DOCUMENTOS ROUBADOS",
        "ANOTACOES NEGATIVAS",
        "ANOTAÇÕES NEGATIVAS",
        "PROTESTOS",
        "PEFIN",
        "REFIN",
        "CONSULTAS",
        "CHEQUES",
    }

    with pdfplumber.open(str(pdf_path)) as pdf:
        for idx, page in enumerate(pdf.pages):
            if page_set is not None and idx not in page_set:
                continue
            text = _pdf_page_text_native(page)
            if "Participações societárias" not in (text or "") and "Participacoes societarias" not in _no_accents(
                text
            ):
                continue

            words = page.extract_words(x_tolerance=1, y_tolerance=2) or []
            if not words:
                continue

            # localizar posições do header
            header_pos: Dict[str, float] = {}
            for w in words:
                t = _no_accents_upper(w.get("text", ""))
                if t == "CNPJ":
                    header_pos.setdefault("cnpj", w["x0"])
                elif t == "RAZAO":
                    header_pos.setdefault("razao", w["x0"])
                elif t == "CAPITAL":
                    header_pos.setdefault("capital", w["x0"])
                elif t == "UF":
                    header_pos.setdefault("uf", w["x0"])
                elif t == "SITUACAO":
                    header_pos.setdefault("situacao", w["x0"])

            if "cnpj" not in header_pos or "razao" not in header_pos:
                continue

            header_top = min(w["top"] for w in words if _no_accents_upper(w.get("text", "")) in {"CNPJ", "RAZAO"})

            # definir ranges de coluna
            cols = sorted([(k, v) for k, v in header_pos.items()], key=lambda kv: kv[1])
            bounds: Dict[str, Tuple[float, float]] = {}
            for i, (k, x0) in enumerate(cols):
                left = -1e9 if i == 0 else (cols[i - 1][1] + x0) / 2.0
                right = 1e9 if i == len(cols) - 1 else (x0 + cols[i + 1][1]) / 2.0
                bounds[k] = (left, right)

            data_words = [w for w in words if w["top"] > header_top + 6]
            data_words = sorted(data_words, key=lambda w: (w["top"], w["x0"]))

            # agrupa por linha
            lines: List[Dict[str, object]] = []
            cur: List[dict] = []
            cur_top: Optional[float] = None
            y_tol = 2.8
            for w in data_words:
                if cur_top is None or abs(w["top"] - cur_top) <= y_tol:
                    cur.append(w)
                    cur_top = w["top"] if cur_top is None else (cur_top + w["top"]) / 2.0
                else:
                    lines.append({"top": cur_top, "words": cur})
                    cur = [w]
                    cur_top = w["top"]
            if cur:
                lines.append({"top": cur_top, "words": cur})

            # linhas -> textos por coluna
            line_items: List[Dict[str, object]] = []
            for item in lines:
                ln_words = sorted(item["words"], key=lambda w: w["x0"])
                cols_text = {k: [] for k in bounds.keys()}
                full_line = []
                for w in ln_words:
                    txt = w.get("text", "")
                    full_line.append(txt)
                    x = w["x0"]
                    for col, (lo, hi) in bounds.items():
                        if lo <= x < hi:
                            cols_text[col].append(txt)
                            break
                line_text = _strip(" ".join(full_line))
                line_up = _no_accents_upper(line_text)
                if any(k in line_up for k in stop_keys):
                    break
                if "EXIBINDO" in line_up or "DETALHES DAS PARTICIPACOES" in line_up:
                    continue
                if "RECEITA FEDERAL" in line_up:
                    continue
                line_items.append(
                    {
                        "top": item["top"],
                        "text": line_text,
                        "cnpj": _strip(" ".join(cols_text.get("cnpj", []))),
                        "razao": _strip(" ".join(cols_text.get("razao", []))),
                        "capital": _strip(" ".join(cols_text.get("capital", []))),
                        "uf": _strip(" ".join(cols_text.get("uf", []))),
                        "situacao": _strip(" ".join(cols_text.get("situacao", []))),
                    }
                )

            # criar registros a partir de linhas com CNPJ
            records: List[Dict[str, object]] = []
            for ln in line_items:
                if not CNPJ_RE.search(ln["cnpj"]):
                    continue
                rec = {
                    "top": ln["top"],
                    "cnpj": CNPJ_RE.search(ln["cnpj"]).group(0),
                    "capital": "A confirmar",
                    "uf": "A confirmar",
                    "situacao": "A confirmar",
                    "razao_parts": [],
                }
                if ln["razao"]:
                    rec["razao_parts"].append(ln["razao"])
                perc = PERCENT_RE.findall(ln["text"])
                if perc:
                    rec["capital"] = perc[0].replace(" ", "")
                for tk in re.split(r"\s+", ln["text"]):
                    if _no_accents_upper(tk) in UF_SET:
                        rec["uf"] = _no_accents_upper(tk)
                        break
                m_sit = re.search(r"\b(ATIVA|ATIVO|INAPTA|BAIXADA|SUSPENSA|NULA)\b", _no_accents_upper(ln["text"]))
                if m_sit:
                    raw = m_sit.group(1)
                    rec["situacao"] = "Ativa" if raw.startswith("ATIV") else raw.capitalize()
                records.append(rec)

            if not records:
                continue

            # atribui linhas de razão social ao registro mais próximo
            for ln in line_items:
                if CNPJ_RE.search(ln["cnpj"]):
                    continue
                if not ln["razao"]:
                    continue
                # ignora linhas com números pesados
                if re.search(r"\d{2}/\d{2}/\d{4}", ln["text"]):
                    continue
                nearest = min(records, key=lambda r: abs(float(r["top"]) - float(ln["top"])))
                nearest["razao_parts"].append(ln["razao"])

            for rec in records:
                razao = _strip(" ".join(rec["razao_parts"])) or "A confirmar"
                participacoes.append(
                    Participacao(
                        cnpj=rec["cnpj"],
                        razao=razao,
                        capital=rec["capital"],
                        uf=rec["uf"],
                        situacao=rec["situacao"],
                    )
                )

    return participacoes


# =========================
# Anotações Negativas (somente PDF)
# =========================


def _has_sem_registros(block: str) -> bool:
    return bool(re.search(r"\bSem\s+registros\b", block or "", flags=re.IGNORECASE))


def extract_anotacoes(text_pdf: str) -> List[Tuple[str, str]]:
    """
    Retorna lista de pares (Campo, Informação) já no formato final.
    Regras BI:
    - Se não houver registros na categoria -> ZERO linhas (não criar "Sem registros")
    """
    out: List[Tuple[str, str]] = []

    anot = extract_block(
        text_pdf,
        "Anotações negativas",
        stop_any=[
            "Participações societárias",
            "Consultas por mês",
            "Consultas por mes",
            "Dados do CPF",
        ],
    )
    if not anot:
        return out

    # seções por categoria
    # (cada uma termina na próxima)
    def subsec(title: str, stops: List[str]) -> str:
        return extract_block(anot, title, stop_any=stops, max_chars=20000)

    # PEFIN
    pefin = subsec("PEFIN", ["REFIN", "Protestos", "Ações judiciais", "Acoes judiciais", "Cheques"])
    if pefin and not _has_sem_registros(pefin):
        # tentativa genérica: 1 registro = precisa ter pelo menos 1 data ou 1 valor
        recs = _parse_registros_4campos(pefin)
        for i, r in enumerate(recs, start=1):
            out.append((f"PEFIN - Registro {i} (Modalidade)", r.get("Modalidade", "A confirmar")))
            out.append((f"PEFIN - Registro {i} (Valor)", r.get("Valor", "A confirmar")))
            out.append((f"PEFIN - Registro {i} (Origem)", r.get("Origem", "A confirmar")))
            out.append((f"PEFIN - Registro {i} (Data)", r.get("Data", "A confirmar")))

    # REFIN
    refin = subsec("REFIN", ["Protestos", "Ações judiciais", "Acoes judiciais", "Cheques"])
    if refin and not _has_sem_registros(refin):
        recs = _parse_registros_4campos(refin)
        for i, r in enumerate(recs, start=1):
            out.append((f"REFIN - Registro {i} (Modalidade)", r.get("Modalidade", "A confirmar")))
            out.append((f"REFIN - Registro {i} (Valor)", r.get("Valor", "A confirmar")))
            out.append((f"REFIN - Registro {i} (Origem)", r.get("Origem", "A confirmar")))
            out.append((f"REFIN - Registro {i} (Data)", r.get("Data", "A confirmar")))

    # Protestos
    prot = subsec("Protestos", ["Ações judiciais", "Acoes judiciais", "Cheques"])
    if prot and not _has_sem_registros(prot):
        recs = _parse_protestos(prot)
        for i, r in enumerate(recs, start=1):
            out.append((f"Protestos - Registro {i} (Data)", r.get("Data", "A confirmar")))
            out.append((f"Protestos - Registro {i} (Valor)", r.get("Valor", "A confirmar")))

    # Ações judiciais
    acoes = subsec("Ações judiciais", ["Cheques"])
    if not acoes:
        acoes = subsec("Acoes judiciais", ["Cheques"])
    if acoes and not _has_sem_registros(acoes):
        recs = _parse_acoes(acoes)
        for i, r in enumerate(recs, start=1):
            out.append((f"Ações Judiciais - Registro {i} (Data)", r.get("Data", "A confirmar")))
            out.append((f"Ações Judiciais - Registro {i} (Valor)", r.get("Valor", "A confirmar")))
            out.append((f"Ações Judiciais - Registro {i} (Natureza)", r.get("Natureza", "A confirmar")))
            out.append((f"Ações Judiciais - Registro {i} (Cidade)", r.get("Cidade", "A confirmar")))

    # Cheques (motivos)
    cheq = subsec("Cheques", [])
    if cheq and not _has_sem_registros(cheq):
        motivos = _parse_cheques_motivos(cheq)
        for mot in motivos:
            out.append(("Cheques - Motivo", mot))

    return out

def _parse_registros_4campos(block: str) -> List[Dict[str, str]]:
    """
    Parser genérico "conservador":
    Só cria registro se encontrar um par (Data OU Valor) explícito em um agrupamento.
    """
    lines = [l for l in (block or "").splitlines() if _strip(l)]
    # remove cabeçalhos comuns
    cleaned = []
    for l in lines:
        up = _no_accents_upper(l)
        if "EXIBINDO" in up and "REGISTRO" in up:
            continue
        if "MODALIDADE" in up and "VALOR" in up:
            continue
        if "ORIGEM" in up and "DATA" in up:
            continue
        cleaned.append(_strip(l))

    records: List[Dict[str, str]] = []
    buf: List[str] = []

    def flush_buf() -> None:
        nonlocal buf
        if not buf:
            return
        s = _strip(" ".join(buf))
        # critério mínimo para existir registro real:
        if not (DATE_RE.search(s) or MONEY_RE.search(s)):
            buf = []
            return
        rec = {}
        # Data
        mdt = DATE_RE.search(s)
        rec["Data"] = mdt.group(0) if mdt else "A confirmar"
        # Valor
        mm = MONEY_RE.search(s)
        rec["Valor"] = _ptbr_money_from_any(mm.group(0)) if mm else "A confirmar"
        if rec["Valor"] is None:
            rec["Valor"] = "A confirmar"
        # Modalidade / Origem: heurística simples (não inventa)
        # pega texto antes do valor como modalidade; depois tenta achar origem perto de palavras-chave
        tmp = s
        if mm:
            tmp = tmp.replace(mm.group(0), " ")
        if mdt:
            tmp = tmp.replace(mdt.group(0), " ")
        tmp = _strip(tmp)
        rec["Modalidade"] = tmp[:80] if tmp else "A confirmar"
        rec["Origem"] = "A confirmar"
        records.append(rec)
        buf = []

    for l in cleaned:
        # se começa um novo registro por detectar um valor/data em linha "forte"
        if DATE_RE.search(l) or MONEY_RE.search(l):
            buf.append(l)
            flush_buf()
        else:
            # acumula texto até achar valor/data
            buf.append(l)

    flush_buf()
    return records


def _parse_protestos(block: str) -> List[Dict[str, str]]:
    lines = [l for l in (block or "").splitlines() if _strip(l)]
    recs: List[Dict[str, str]] = []
    for l in lines:
        s = _strip(l)
        if "EXIBINDO" in _no_accents_upper(s):
            continue
        mdt = DATE_RE.search(s)
        mm = MONEY_RE.search(s)
        if not (mdt and mm):
            continue
        val = _ptbr_money_from_any(mm.group(0)) or "A confirmar"
        recs.append({"Data": mdt.group(0), "Valor": val})
    return recs


def _parse_acoes(block: str) -> List[Dict[str, str]]:
    lines = [l for l in (block or "").splitlines() if _strip(l)]
    recs: List[Dict[str, str]] = []
    buf: List[str] = []
    for l in lines:
        s = _strip(l)
        if "EXIBINDO" in _no_accents_upper(s):
            continue
        if DATE_RE.search(s) and MONEY_RE.search(s):
            # tenta fechar buffer anterior
            if buf:
                buf = []
            buf = [s]
            joined = _strip(" ".join(buf))
            mdt = DATE_RE.search(joined)
            mm = MONEY_RE.search(joined)
            rec = {
                "Data": mdt.group(0) if mdt else "A confirmar",
                "Valor": _ptbr_money_from_any(mm.group(0)) if mm else "A confirmar",
                "Natureza": "A confirmar",
                "Cidade": "A confirmar",
            }
            if rec["Valor"] is None:
                rec["Valor"] = "A confirmar"
            recs.append(rec)
            buf = []
        else:
            buf.append(s)
    return recs


def _parse_cheques_motivos(block: str) -> List[str]:
    lines = [l for l in (block or "").splitlines() if _strip(l)]
    motivos: List[str] = []
    for l in lines:
        s = _strip(l)
        up = _no_accents_upper(s)
        if "SEM REGISTROS" in up:
            return []
        # procura motivos típicos em texto
        if re.search(r"\b(SUSTADO|SEM FUNDOS|DEVOLVIDO|IRREGULAR)\b", up):
            # pega a palavra mais relevante
            m = re.search(r"\b(SUSTADO|SEM FUNDOS|DEVOLVIDO|IRREGULAR)\b", up)
            if m:
                motivos.append(m.group(1).title())
    # remove duplicados preservando ordem
    seen = set()
    out = []
    for m in motivos:
        if m not in seen:
            seen.add(m)
            out.append(m)
    return out


# =========================
# Consultas por mês (somente IMAGEM)
# =========================


@dataclass
class Bar:
    x: int
    y: int
    w: int
    h: int


def _blue_mask_hsv(
    img_bgr: "np.ndarray",
    lower: Optional["np.ndarray"] = None,
    upper: Optional["np.ndarray"] = None,
) -> "np.ndarray":
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
    if lower is None:
        lower = np.array([90, 40, 40], dtype=np.uint8)
    if upper is None:
        upper = np.array([150, 255, 255], dtype=np.uint8)
    mask = cv2.inRange(hsv, lower, upper)
    k1 = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    k2 = cv2.getStructuringElement(cv2.MORPH_RECT, (9, 9))
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, k1, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, k2, iterations=2)
    return mask


def _detect_bars(
    img_bgr: "np.ndarray",
    lower: Optional["np.ndarray"] = None,
    upper: Optional["np.ndarray"] = None,
) -> List[Bar]:
    mask = _blue_mask_hsv(img_bgr, lower=lower, upper=upper)
    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    h, w = img_bgr.shape[:2]
    bars: List[Bar] = []
    for c in cnts:
        x, y, bw, bh = cv2.boundingRect(c)
        if bw < max(6, w // 300):
            continue
        if bh < max(12, h // 300):
            continue
        if bh > int(h * 0.9):
            continue
        bars.append(Bar(x, y, bw, bh))
    bars.sort(key=lambda b: b.x)

    merged: List[Bar] = []
    for b in bars:
        if not merged:
            merged.append(b)
            continue
        prev = merged[-1]
        if abs(b.x - prev.x) < 10 and abs((b.y + b.h) - (prev.y + prev.h)) < 40:
            x1 = min(prev.x, b.x)
            y1 = min(prev.y, b.y)
            x2 = max(prev.x + prev.w, b.x + b.w)
            y2 = max(prev.y + prev.h, b.y + b.h)
            merged[-1] = Bar(x1, y1, x2 - x1, y2 - y1)
        else:
            merged.append(b)

    return merged


def _remove_blue(
    img_bgr: "np.ndarray",
    lower: Optional["np.ndarray"] = None,
    upper: Optional["np.ndarray"] = None,
) -> "np.ndarray":
    mask = _blue_mask_hsv(img_bgr, lower=lower, upper=upper)
    out = img_bgr.copy()
    out[mask > 0] = (255, 255, 255)
    return out


def _preprocess_for_digits(roi_bgr: "np.ndarray", mode: str) -> "np.ndarray":
    gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.resize(gray, None, fx=4, fy=4, interpolation=cv2.INTER_CUBIC)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    if mode == "otsu":
        _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    elif mode == "adapt":
        th = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 41, 15
        )
    elif mode == "adapt_inv":
        th = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 41, 15
        )
        th = 255 - th
    else:
        _, th = cv2.threshold(gray, 170, 255, cv2.THRESH_BINARY)
    return th


def _tighten_binary(th_img: "np.ndarray") -> "np.ndarray":
    ys, xs = np.where(th_img < 250)
    if xs.size == 0 or ys.size == 0:
        return th_img

    x1, x2 = xs.min(), xs.max() + 1
    y1, y2 = ys.min(), ys.max() + 1
    crop = th_img[y1:y2, x1:x2]

    pad = max(4, int(min(crop.shape[:2]) * 0.10))
    crop = cv2.copyMakeBorder(crop, pad, pad, pad, pad, cv2.BORDER_CONSTANT, value=255)

    h, w = crop.shape[:2]
    scale = 1
    if max(h, w) < 80:
        scale = 2
    if max(h, w) < 40:
        scale = 3
    if scale > 1:
        crop = cv2.resize(crop, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    return crop


def _largest_component(th_img: "np.ndarray") -> "np.ndarray":
    inv = (th_img < 128).astype(np.uint8)
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(inv, connectivity=8)
    if num_labels <= 1:
        return th_img
    idx = 1 + int(np.argmax(stats[1:, cv2.CC_STAT_AREA]))
    mask = (labels == idx).astype(np.uint8)
    out = np.full(th_img.shape, 255, dtype=np.uint8)
    out[mask > 0] = 0
    return out


def _ocr_digits(th_img: "np.ndarray", psm: int) -> str:
    cfg = f"--oem 3 --psm {psm} -c tessedit_char_whitelist=0123456789"
    txt = pytesseract.image_to_string(th_img, config=cfg) or ""
    txt = re.sub(r"\D", "", txt)
    return txt


def _best_digit_ocr(roi_bgr: "np.ndarray") -> Optional[int]:
    roi_nb = _remove_blue(roi_bgr)
    tries = [
        ("otsu", 7),
        ("adapt", 7),
        ("adapt_inv", 7),
        ("otsu", 10),
        ("adapt", 10),
        ("otsu", 8),
        ("adapt", 8),
        ("otsu", 13),
    ]

    best_txt = ""
    for mode, psm in tries:
        th = _preprocess_for_digits(roi_nb, mode=mode)
        k = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        th2 = cv2.morphologyEx(th, cv2.MORPH_CLOSE, k, iterations=1)
        th3 = _tighten_binary(th2)
        txt = _ocr_digits(th3, psm=psm)
        if not txt:
            th4 = _largest_component(th2)
            th3 = _tighten_binary(th4)
            txt = _ocr_digits(th3, psm=psm)

        if len(txt) > len(best_txt):
            best_txt = txt
        elif len(txt) == len(best_txt) and txt > best_txt:
            best_txt = txt

        if len(best_txt) >= 2:
            break

    if not best_txt:
        gray = cv2.cvtColor(roi_nb, cv2.COLOR_BGR2GRAY)
        gray = cv2.resize(gray, None, fx=6, fy=6, interpolation=cv2.INTER_CUBIC)
        _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        k = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        th = cv2.morphologyEx(th, cv2.MORPH_CLOSE, k, iterations=1)
        th = 255 - th
        th = _largest_component(th)
        th = _tighten_binary(th)
        for psm in (10, 8, 7):
            txt = _ocr_digits(th, psm=psm)
            if txt:
                best_txt = txt
                break

    if not best_txt:
        return None

    try:
        val = int(best_txt)
        if val < 0 or val > 500:
            return None
        return val
    except Exception:
        return None


def _normalize_month_text(s: str) -> str:
    s = (s or "").strip()
    s = s.replace("\\", "/").replace("|", "/")
    s = re.sub(r"\s+", "", s)
    return s


def _parse_start_label(label: str) -> datetime:
    s = (label or "").strip().lower()
    m = re.search(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\s*[/\\-]\s*(20\d{2})", s)
    if not m:
        raise ValueError(f'--start-label invalido: "{label}". Use tipo "Nov/2024".')
    mon = int(PT_MONTHS[m.group(1)])
    year = int(m.group(2))
    return datetime(year, mon, 1)


def _add_months(dt: datetime, n: int) -> datetime:
    y = dt.year
    m = dt.month + n
    y += (m - 1) // 12
    m = ((m - 1) % 12) + 1
    return datetime(y, m, 1)


def _infer_start_month(ocr_months: List[Optional[datetime]], fallback_start: datetime) -> datetime:
    candidates = []
    for i, dt in enumerate(ocr_months):
        if dt is None:
            continue
        candidates.append((i, dt))

    if not candidates:
        return fallback_start

    best_start = fallback_start
    best_score = -1
    best_exact = -1

    for i, dt in candidates:
        start = _add_months(dt, -i)
        score = 0
        exact = 0
        for j, dtj in enumerate(ocr_months):
            if dtj is None:
                continue
            exp = _add_months(start, j)
            if dtj.year == exp.year and dtj.month == exp.month:
                score += 2
                exact += 1
            elif dtj.month == exp.month:
                score += 1
        if score > best_score or (score == best_score and exact > best_exact):
            best_score = score
            best_exact = exact
            best_start = start

    return best_start


def _parse_month_label(s: str) -> Optional[datetime]:
    s2 = _normalize_month_text(s).lower()
    m = re.search(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez).*(20\d{2})", s2)
    if not m:
        return None
    mon = int(PT_MONTHS[m.group(1)])
    year = int(m.group(2))
    return datetime(year, mon, 1)


def _ocr_month_under_bar(chart_nb_bgr: "np.ndarray", bar: Bar) -> Optional[str]:
    h, w = chart_nb_bgr.shape[:2]
    x1 = max(0, bar.x - 40)
    x2 = min(w, bar.x + bar.w + 40)
    y1 = max(0, bar.y + bar.h + 25)
    y2 = min(h, bar.y + bar.h + 140)

    roi = chart_nb_bgr[y1:y2, x1:x2]
    if roi.size == 0:
        return None

    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    gray = cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 41, 15)
    cfg = "--oem 3 --psm 7"
    txt = pytesseract.image_to_string(th, config=cfg) or ""
    txt = _normalize_month_text(txt)
    return txt if len(txt) >= 5 else None


def _crop_chart_roi(img_bgr: "np.ndarray", bars: List[Bar]) -> Tuple["np.ndarray", Tuple[int, int, int, int]]:
    h, w = img_bgr.shape[:2]
    xs = [b.x for b in bars]
    xe = [b.x + b.w for b in bars]
    ys = [b.y for b in bars]
    ye = [b.y + b.h for b in bars]

    def clamp(v: int, lo: int, hi: int) -> int:
        return max(lo, min(hi, v))

    x1 = clamp(min(xs) - 160, 0, w - 1)
    x2 = clamp(max(xe) + 160, 0, w)
    y1 = clamp(min(ys) - 260, 0, h - 1)
    y2 = clamp(max(ye) + 260, 0, h)

    roi = img_bgr[y1:y2, x1:x2].copy()
    return roi, (x1, y1, x2 - x1, y2 - y1)


def _extract_consultas_from_image_array(
    img_bgr: "np.ndarray",
    lower: Optional["np.ndarray"] = None,
    upper: Optional["np.ndarray"] = None,
) -> List[Tuple[str, str]]:
    bars_page = _detect_bars(img_bgr, lower=lower, upper=upper)
    if len(bars_page) < 3:
        return []

    roi_bgr, _ = _crop_chart_roi(img_bgr, bars_page)
    bars_roi = _detect_bars(roi_bgr, lower=lower, upper=upper)
    if not bars_roi:
        return []

    base_ys = np.array([b.y + b.h for b in bars_roi], dtype=np.int32)
    if len(base_ys) > 0:
        base_ref = int(np.median(base_ys))
        bars_roi = [
            b for b in bars_roi if abs((b.y + b.h) - base_ref) < int(roi_bgr.shape[0] * 0.25)
        ]

    bars_roi.sort(key=lambda b: b.x)

    chart_nb = _remove_blue(roi_bgr, lower=lower, upper=upper)
    h, w = chart_nb.shape[:2]

    def clamp(v: int, lo: int, hi: int) -> int:
        return max(lo, min(hi, v))

    results: List[Tuple[str, str]] = []
    seen = set()
    ocr_months: List[Optional[datetime]] = []
    bar_vals: List[Optional[int]] = []

    for b in bars_roi:
        x1 = clamp(b.x - 55, 0, w - 1)
        x2 = clamp(b.x + b.w + 55, 0, w)
        up = max(140, int(h * 0.22))
        y1 = clamp(b.y - up, 0, h - 1)
        y2 = clamp(b.y - 5, 0, h)
        roi_num = chart_nb[y1:y2, x1:x2]
        val = _best_digit_ocr(roi_num)

        mtxt = _ocr_month_under_bar(chart_nb, b)
        mdt = _parse_month_label(mtxt) if mtxt else None
        ocr_months.append(mdt)
        bar_vals.append(val)

    if not bar_vals:
        return []

    valid_years = [dt.year for dt in ocr_months if dt and dt.year >= 2024]
    target_year = max(valid_years) if valid_years else None
    ocr_fixed: List[Optional[datetime]] = []
    for dt in ocr_months:
        if dt and dt.year < 2024 and target_year:
            ocr_fixed.append(datetime(target_year, dt.month, 1))
        else:
            ocr_fixed.append(dt)

    dt_fallback = _parse_start_label(DEFAULT_START_LABEL)
    if sum(1 for dt in ocr_fixed if dt is not None) >= 2:
        dt0 = _infer_start_month(ocr_fixed, fallback_start=dt_fallback)
    else:
        dt0 = dt_fallback
    months = [_add_months(dt0, i) for i in range(len(bar_vals))]

    final_results: List[Tuple[str, str]] = []
    for i, val in enumerate(bar_vals):
        if val is None or val <= 0:
            continue
        base_dt = ocr_fixed[i] if ocr_fixed[i] is not None else months[i]
        mm_yyyy = f"{base_dt.month:02d}/{base_dt.year}"
        campo = f"Consultas - {mm_yyyy}"
        if campo in seen:
            continue
        seen.add(campo)
        final_results.append((campo, str(val)))

    def key_month(campo: str) -> Tuple[int, int]:
        m = re.search(r"(\d{2})/(\d{4})", campo)
        if not m:
            return (9999, 99)
        return (int(m.group(2)), int(m.group(1)))

    final_results.sort(key=lambda x: key_month(x[0]))
    return final_results


def _extract_consultas_from_image(img_path: Path) -> List[Tuple[str, str]]:
    if not HAS_CV2:
        return []
    img_bgr = cv2.imread(str(img_path))
    if img_bgr is None:
        return []
    hsv_ranges = [
        (np.array([90, 40, 40], dtype=np.uint8), np.array([150, 255, 255], dtype=np.uint8)),
        (np.array([85, 30, 30], dtype=np.uint8), np.array([160, 255, 255], dtype=np.uint8)),
        (np.array([80, 20, 20], dtype=np.uint8), np.array([170, 255, 255], dtype=np.uint8)),
    ]

    for lower, upper in hsv_ranges:
        res = _extract_consultas_from_image_array(img_bgr, lower=lower, upper=upper)
        if res:
            return res
    return []


def _extract_consultas_from_pdf(
    pdf_path: Path,
    page_indices: Iterable[int],
    poppler_path: Optional[Path],
    dpi: int = 200,
) -> List[Tuple[str, str]]:
    pages = list(sorted(set(page_indices)))
    if not pages:
        return []

    if not HAS_CV2 or poppler_path is None:
        return []

    first_page = min(pages) + 1
    last_page = max(pages) + 1
    hsv_ranges = [
        (np.array([90, 40, 40], dtype=np.uint8), np.array([150, 255, 255], dtype=np.uint8)),
        (np.array([85, 30, 30], dtype=np.uint8), np.array([160, 255, 255], dtype=np.uint8)),
        (np.array([80, 20, 20], dtype=np.uint8), np.array([170, 255, 255], dtype=np.uint8)),
    ]

    for dpi_try in (dpi, 250, 300):
        images = convert_from_path(
            str(pdf_path),
            dpi=dpi_try,
            poppler_path=str(poppler_path),
            first_page=first_page,
            last_page=last_page,
            fmt="png",
        )

        for lower, upper in hsv_ranges:
            best_score = -1
            best_img = None
            for idx, img in enumerate(images):
                pno = first_page + idx
                if (pno - 1) not in pages:
                    continue
                img_rgb = np.array(img)
                img_bgr = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2BGR)
                mask = _blue_mask_hsv(img_bgr, lower=lower, upper=upper)
                score = int(mask.sum())
                if score > best_score:
                    best_score = score
                    best_img = img_bgr

            if best_img is None:
                continue
            res = _extract_consultas_from_image_array(best_img, lower=lower, upper=upper)
            if res:
                return res

    return []

# =========================
# Montagem final + validações STRICT
# =========================


def build_rows(
    pdf_text: str,
    pdf_path: Optional[Path],
    consultas_img: Optional[Path],
    poppler_path: Optional[Path],
    page_indices: Optional[List[int]] = None,
) -> Tuple[List[Tuple[str, str]], str, str]:
    """
    Retorna:
    - rows: [(Campo, Informação)]
    - cpf
    - nome_base (para nome do arquivo)
    """
    if pdf_text:
        cpf = extract_cpf_from_pdf_text(pdf_text)
        ident = extract_identificacao(pdf_text)
        if pdf_path and page_indices:
            ident_pdf = extract_identificacao_from_pdf(pdf_path, page_indices=page_indices)
            for k in ("Situação na Receita Federal", "Nome completo", "Município/UF", "Serasa Score", "Probabilidade de pagamento em 12 meses"):
                v = ident_pdf.get(k)
                if v and v != "A confirmar":
                    ident[k] = v
        nome_base = ident.get("Nome completo", "") if ident else ""
    else:
        cpf = "A confirmar"
        ident = {k: "A confirmar" for k in IDENT_FIELDS}
        nome_base = ""

    rows: List[Tuple[str, str]] = []

    # A) IDENTIFICAÇÃO (sempre gera as 7 linhas nessa ordem; ausente => A confirmar)
    for k in IDENT_FIELDS:
        rows.append((k, ident.get(k, "A confirmar")))

    # A.1) CONSULTAS (tabela de consultas - 5 linhas fixas)
    if pdf_path and page_indices:
        consultantes = extract_consultas_tabela(
            pdf_path,
            page_indices=page_indices,
            poppler_path=poppler_path,
            pdf_text=pdf_text or "",
        )
        if len(consultantes) < 5:
            raise RuntimeError("ERRO: Consultas insuficientes no SERASA SOCIO (esperado 5).")
        for i in range(1, 6):
            data, nome = consultantes[i - 1]
            rows.append((f"Consulta {i}", f"{data} - {nome}"))
    else:
        raise RuntimeError("ERRO: Consultas insuficientes no SERASA SOCIO (esperado 5).")

    # B) CONSULTAS (somente imagem; se não existir, gera ZERO linhas)
    if consultas_img and consultas_img.exists():
        cons = _extract_consultas_from_image(consultas_img)
        for campo, info in cons:
            rows.append((campo, info))
    elif pdf_path and page_indices:
        # fallback: gráfico dentro do PDF
        cons = _extract_consultas_from_pdf(pdf_path, page_indices, poppler_path=poppler_path)
        for campo, info in cons:
            rows.append((campo, info))

    # C) PARTICIPAÇÕES (somente PDF; se não houver, zero linhas)
    if pdf_path and pdf_path.exists() and pdf_text:
        parts = extract_participacoes(pdf_path, page_indices=page_indices)
        for idx, p in enumerate(parts, start=1):
            rows.append((f"Participação {idx} - CNPJ", p.cnpj))
            rows.append((f"Participação {idx} - Capital", p.capital))
            rows.append((f"Participação {idx} - Situação Cadastral", p.situacao))
            rows.append((f"Participação {idx} - UF", p.uf))
            rows.append((f"Participação {idx} - Razão Social", p.razao))

    # D) ANOTAÇÕES NEGATIVAS (somente PDF; se não houver registros, zero linhas)
    if pdf_text:
        anot = extract_anotacoes(pdf_text)
        rows.extend(anot)

    # Normalizações finais (moeda PT-BR quando possível)
    def norm_info(campo: str, info: str) -> str:
        s = _strip(info or "")
        if not s:
            return "A confirmar"
        # não usar travessão especial
        s = s.replace("—", "-").replace("–", "-")
        # tenta normalizar dinheiro onde for esperado por conteúdo monetário
        if "Valor" in campo or "Total em anotações negativas" == campo:
            m = MONEY_RE.search(s)
            if m:
                v = _ptbr_money_from_any(m.group(0))
                if v:
                    return v
        return s

    rows = [(c, norm_info(c, i)) for (c, i) in rows]

    # validações STRICT
    for campo, _ in rows:
        if not _campo_valido(campo):
            raise RuntimeError(
                'ERRO: "Fontes cruzadas / CPF ausente / Campo inválido (caractere divergente). '
                'Consultas: somente IMAGEM. Identificação/Participações/Anotações: somente PDF."'
            )

    return rows, cpf, (nome_base or "SOCIO")


def write_xlsx(out_xlsx: Path, rows: List[Tuple[str, str]], cpf: str) -> None:
    rows = [(c, i) for (c, i) in rows]
    df = pd.DataFrame(rows, columns=["Campo", "Informação"])
    df["CPF"] = cpf
    df = df[["Campo", "Informação", "CPF"]]

    # valida cabeçalho
    if list(df.columns) != ["Campo", "Informação", "CPF"]:
        raise RuntimeError(
            'ERRO: "Fontes cruzadas / CPF ausente / Campo inválido (caractere divergente). '
            'Consultas: somente IMAGEM. Identificação/Participações/Anotações: somente PDF."'
        )

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Planilha1", index=False)
        wb = writer.book
        ws = wb["Planilha1"]
        for row in ws.iter_rows():
            for cell in row:
                cell.number_format = "@"  # tudo texto


# =========================
# Auto-scan (01_INPUT)
# =========================


def _pick_pdf_socio_files(folder: Path) -> List[Path]:
    pdfs = [p for p in folder.glob("*.pdf")]
    out = []
    for p in pdfs:
        name = _no_accents(p.name).lower()
        if "serasa" in name and "socio" in name:
            out.append(p)
    return sorted(out)


def _pick_consultas_img(folder: Path, pdf_path: Optional[Path]) -> Optional[Path]:
    imgs = list(folder.glob("*.png")) + list(folder.glob("*.jpg")) + list(folder.glob("*.jpeg"))
    if not imgs:
        return None

    target = ""
    if pdf_path:
        tokens = _no_accents(pdf_path.stem).lower().split("_")
        if len(tokens) >= 3:
            target = tokens[-1]

    best = None
    best_score = -999
    for p in imgs:
        name = _no_accents(p.name).lower()
        score = 0
        if "consulta" in name:
            score += 3
        if "graf" in name:
            score += 2
        if "serasa" in name:
            score += 1
        if "socio" in name:
            score += 2
        if target and target in name:
            score += 3
        if score > best_score:
            best_score = score
            best = p
    return best if best_score >= 3 else None


def _group_pages_by_cpf(page_texts: List[str]) -> Dict[str, List[int]]:
    groups: Dict[str, List[int]] = {}
    current_cpf = None

    for idx, text in enumerate(page_texts):
        cpf = extract_cpf_from_any(text)
        if cpf:
            current_cpf = cpf
        if not current_cpf:
            current_cpf = "A confirmar"
        groups.setdefault(current_cpf, []).append(idx)

    return groups


def _join_pages(page_texts: List[str], page_indices: List[int]) -> str:
    return "\n".join([page_texts[i] for i in page_indices if 0 <= i < len(page_texts)])


def _first_name(raw: str) -> str:
    s = _strip(raw or "")
    s = re.sub(r"[^A-Za-zÀ-ÿ\s\-]", " ", s)
    s = _strip(re.sub(r"\s+", " ", s))
    return s.split(" ")[0] if s else ""


def _output_name(nome_base: str, cpf: str, pdf_path: Optional[Path]) -> str:
    if nome_base and nome_base != "A confirmar":
        up = _no_accents_upper(nome_base)
        if "DATA DE NASCIMENTO" not in up and "NOME DA MAE" not in up and "CPF" not in up:
            return _safe_filename(nome_base)
    if pdf_path:
        tokens = _no_accents(pdf_path.stem).lower().split("_")
        if len(tokens) >= 3:
            first = _first_name(tokens[-1])
            return _safe_filename(first) if first else _safe_filename(tokens[-1])
        return _safe_filename(pdf_path.stem)
    if cpf and cpf != "A confirmar":
        return _safe_filename(cpf)
    return "SOCIO"


def run_pdf(
    pdf_path: Path,
    outdir: Path,
    poppler_path: Optional[Path],
    img_path: Optional[Path],
) -> List[Path]:
    page_texts = extract_pdf_page_texts(pdf_path, poppler_path=poppler_path, ocr_dpi=260)
    groups = _group_pages_by_cpf(page_texts)

    outputs: List[Path] = []
    for cpf, pages in groups.items():
        pdf_text = _join_pages(page_texts, pages)
        rows, cpf_out, nome_base = build_rows(
            pdf_text=pdf_text,
            pdf_path=pdf_path,
            consultas_img=img_path,
            poppler_path=poppler_path,
            page_indices=pages,
        )
        safe_nome = _output_name(nome_base, cpf_out, pdf_path)
        out_xlsx = outdir / f"SERASA_SOCIO_{safe_nome}.xlsx"
        write_xlsx(out_xlsx, rows=rows, cpf=cpf_out)
        outputs.append(out_xlsx)

    return outputs


def run_folder(folder: Path, outdir: Path, poppler_path: Optional[Path]) -> List[Path]:
    outputs: List[Path] = []
    for pdf_path in _pick_pdf_socio_files(folder):
        img_path = _pick_consultas_img(folder, pdf_path)
        outputs.extend(run_pdf(pdf_path, outdir=outdir, poppler_path=poppler_path, img_path=img_path))
    return outputs


# =========================
# CLI / Orquestração
# =========================


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", default="", help="Caminho do PDF SERASA (SÓCIO).")
    ap.add_argument("--img", default="", help="Caminho da IMAGEM (png/jpg) do gráfico de Consultas.")
    ap.add_argument("--input", default=str(DEFAULT_INPUT), help="Pasta base 01_INPUT (varre subpastas).")
    ap.add_argument("--outdir", default=str(DEFAULT_OUTPUT), help=r"Pasta de saída (ex.: ...\03_OUTPUT\6. SERASA SOCIO)")
    ap.add_argument("--poppler", default=str(DEFAULT_POPPLER), help="Pasta do Poppler bin (para OCR pontual no PDF).")
    ap.add_argument("--tesseract", default="", help="Caminho do tesseract.exe (opcional).")
    args = ap.parse_args()

    pdf_path = Path(args.pdf) if args.pdf else None
    img_path = Path(args.img) if args.img else None
    input_dir = Path(args.input)
    outdir = Path(args.outdir)

    poppler_path = Path(args.poppler) if args.poppler else None
    if poppler_path and not poppler_path.exists():
        poppler_path = None
    if args.tesseract:
        pytesseract.pytesseract.tesseract_cmd = args.tesseract

    if pdf_path:
        outputs = run_pdf(pdf_path, outdir=outdir, poppler_path=poppler_path, img_path=img_path)
        for out in outputs:
            print("Arquivo gerado.")
            print(str(out))
        return

    if not input_dir.exists():
        print("ERRO: Pasta de input não encontrada.")
        return

    folders = [p for p in input_dir.iterdir() if p.is_dir()]
    if not folders:
        if _pick_pdf_socio_files(input_dir):
            outputs = run_folder(input_dir, outdir=outdir, poppler_path=poppler_path)
            for out in outputs:
                print("Arquivo gerado.")
                print(out.name)
            return
        print("ERRO: Nenhuma pasta encontrada em 01_INPUT.")
        return

    for folder in folders:
        outputs = run_folder(folder, outdir=outdir, poppler_path=poppler_path)
        for out in outputs:
            print("Arquivo gerado.")
            print(out.name)


if __name__ == "__main__":
    main()
