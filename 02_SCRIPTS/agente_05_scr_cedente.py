#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
AGENTE_05 SCR CEDENTE — Extrator SCR → Excel (STRICT/BI)

- Entrada: PDFs SCR do cedente, nomeados: scr_cedente_[empresa]_[ano].pdf
- Saída: SCR_[RAZAO_SOCIAL].xlsx com 4 abas:
  1) Cards: Campo | Valor | CNPJ | Data base (consulta)
  2) Modalidades: Modalidade | Valor Total (R$) | Ano | CNPJ | Razão Social | Data base (consulta)
  3) LinhasModalidade: Modalidade | Produto/Linha | Tipo | Faixa de prazo | Valor (R$) | Ano | CNPJ | Razão Social | Data base (consulta)
  4) Prazos (faixas): Faixa de prazo | Valor Total (R$) | Referência (dias) | Ano | CNPJ | Razão Social | Data base (consulta)

Prioridade: texto nativo -> OCR pontual (somente páginas-alvo).
Regras: não inventar faixas/linhas; só registrar pares explícitos (faixa + R$).

Dependências:
  pip install pdfplumber pdf2image pillow pytesseract opencv-python numpy pandas openpyxl

Requisitos SO:
  - Poppler (pdftoppm/pdfinfo)
  - Tesseract (opcional; se precisar, informar --tesseract)
"""

from __future__ import annotations

import argparse
import logging
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import unicodedata

# OCR (pontual)
import numpy as np
import cv2
import pytesseract
from pdf2image import convert_from_path

# Texto nativo
try:
    import pdfplumber
except Exception:
    pdfplumber = None
else:
    logging.getLogger("pdfminer").setLevel(logging.ERROR)
    logging.getLogger("pdfminer.pdfinterp").setLevel(logging.ERROR)


# -----------------------------
# Constantes / Canon
# -----------------------------
CARDS_ORDER = [
    "Razão Social",
    "Valor Total",
    "Limite de Crédito",
    "Classificação de risco",
    "Total Instituições",
    "Total Operações",
    "Volume processado (%)",
    "Doc processado (%)",
]

MOD_CANON = [
    "Empréstimos",
    "Financiamentos",
    "Adiantamentos a depositantes",
    "Outros créditos",
    "Títulos descontado Direitos creditórios descontados",
    "Títulos de crédito (fora da carteira classificada)",
    "Limite",
]

ASCII_COL_MAP = {
    "Razão Social": "Razão Social",
    "Limite de Crédito": "Limite de Crédito",
    "Classificação de risco": "Classificação de risco",
    "Total Instituições": "Total Instituições",
    "Total Operações": "Total Operações",
    "Referência (dias)": "Referência (dias)",
}


def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns=ASCII_COL_MAP)


def normalize_card_fields(df: pd.DataFrame) -> pd.DataFrame:
    if "Campo" in df.columns:
        df["Campo"] = df["Campo"].map(lambda v: ASCII_COL_MAP.get(v, v))
    return df

# match tolerante (sem catálogo extra)
MOD_MATCH = [
    (re.compile(r"emprest", re.I), "Empréstimos"),
    (re.compile(r"financi", re.I), "Financiamentos"),
    (re.compile(r"adiant.*deposit", re.I), "Adiantamentos a depositantes"),
    (re.compile(r"outros", re.I), "Outros créditos"),
    (re.compile(r"titulos?.*descontad.*direit", re.I), "Títulos descontado Direitos creditórios descontados"),
    (re.compile(r"direit.*creditor.*descont", re.I), "Títulos descontado Direitos creditórios descontados"),
    (re.compile(r"fora.*carteira", re.I), "Títulos de crédito (fora da carteira classificada)"),
    (re.compile(r"limite", re.I), "Limite"),
]

PT_MON_ABBR = {
    "jan": 1,
    "fev": 2,
    "mar": 3,
    "abr": 4,
    "mai": 5,
    "jun": 6,
    "jul": 7,
    "ago": 8,
    "set": 9,
    "out": 10,
    "nov": 11,
    "dez": 12,
}


# -----------------------------
# Utils
# -----------------------------
def norm_space(s: str) -> str:
    s = (s or "").replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    return s.strip()


def _no_accents_upper(s: str) -> str:
    s = norm_space(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.upper()


def money_ptbr(raw: str) -> str:
    """
    Recebe "396.672.543,59" ou "0,00" ou "1.000" e devolve texto pt-br "x.xxx.xxx,xx".
    Mantém sempre como TEXTO.
    """
    if raw is None:
        return "--"
    s = raw.strip()
    s = s.replace("R$", "").strip()
    s = re.sub(r"[^\d\.,]", "", s)

    if re.fullmatch(r"\d+", s):
        return f"{s},00"

    if "," not in s and "." in s:
        return f"{s},00"

    if re.fullmatch(r".*,\d$", s):
        return s + "0"

    if re.fullmatch(r".*,\d{2}$", s):
        return s

    m = re.search(r"(\d[\d\.]*)(?:,(\d{1,2}))?$", s)
    if not m:
        return "--"
    inte = m.group(1)
    dec = (m.group(2) or "00").ljust(2, "0")[:2]
    return f"{inte},{dec}"


def find_first_money(text: str) -> Optional[str]:
    m = re.search(r"R\$\s*([\d\.\,]+)", text)
    return m.group(1) if m else None


def only_digits(s: str) -> str:
    return re.sub(r"\D", "", s or "")


def parse_cnpj(text: str) -> Optional[str]:
    m = re.search(r"\b(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\b", text)
    if m:
        return m.group(1)
    m = re.search(r"\b(\d{3}\.\d{3}\.\d{3}-\d{2})\b", text)
    return m.group(1) if m else None


def parse_razao_social(text: str) -> Optional[str]:
    """
    Pega linhas após "Razão social" (ou "Razao social") até quebrar.
    """
    m = re.search(
        r"Raz[aã]o\s+social\s*\n(.+?)(?:\n[A-Z][^\n]{0,40}:|\nData base|\nCNPJ|\Z)",
        text,
        flags=re.I | re.S,
    )
    if m:
        raw = m.group(1).strip()
        lines = [norm_space(x) for x in raw.splitlines() if norm_space(x)]
        if lines:
            return norm_space(" ".join(lines))

    m2 = re.search(
        r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b\s+(.+?)\s+([A-Za-z]{3})\s*/\s*(\d{2,4})\b",
        text,
    )
    if m2:
        return norm_space(m2.group(1))

    lines = [norm_space(x) for x in (text or "").splitlines() if norm_space(x)]
    for i, ln in enumerate(lines):
        if re.search(r"\bNome:\s*", ln, flags=re.I):
            name_parts: List[str] = []
            if i + 1 < len(lines):
                name_parts.append(lines[i + 1])
            if i + 2 < len(lines):
                nxt = lines[i + 2]
                if not re.search(r"\d", nxt) and not re.search(
                    r"\bData base\b|\bDoc\.|\bVol\.|\bInicio do relacionamento\b|:", nxt, flags=re.I
                ):
                    name_parts.append(nxt)
            if name_parts:
                raw = " ".join(name_parts)
                raw = re.sub(r"\b[A-Za-z]{3}\s*/\s*\d{2,4}\b", "", raw)
                raw = re.sub(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b", "", raw)
                raw = re.sub(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b", "", raw)
                raw = norm_space(raw)
                if raw:
                    return raw

    return None


def parse_database_consulta(text: str) -> Optional[datetime]:
    """
    "Data base consultada:\nNov/25" => 01/11/2025
    """
    m = re.search(r"Data base consultada:\s*\n?\s*([A-Za-z]{3})\s*/\s*(\d{2,4})", text, flags=re.I)
    if not m:
        m = re.search(
            r"Data base consultada:.*?([A-Za-z]{3})\s*/\s*(\d{2,4})",
            text,
            flags=re.I | re.S,
        )
    if not m:
        m = re.search(
            r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b\s+.+?\s+([A-Za-z]{3})\s*/\s*(\d{2,4})\b",
            text,
        )
    if not m:
        m = re.search(r"\b([A-Za-z]{3})\s*/\s*(\d{2,4})\b", text)
    if not m:
        return None
    mon3 = m.group(1).lower()
    yy = m.group(2)
    if mon3 not in PT_MON_ABBR:
        return None
    month = PT_MON_ABBR[mon3]
    if len(yy) == 2:
        year = 2000 + int(yy)
    else:
        year = int(yy)
    return datetime(year, month, 1)


def fmt_datebase(dt: Optional[datetime]) -> str:
    if not dt:
        return "A confirmar"
    return dt.strftime("01/%m/%Y")


def year_from_dt(dt: Optional[datetime]) -> str:
    if not dt:
        return "A confirmar"
    return str(dt.year)


def referencia_dias(faixa: str) -> str:
    nums = [int(x) for x in re.findall(r"\d+", faixa or "")]
    if not nums:
        return "A confirmar"
    return str(max(nums))


def faixa_texto_exato(texto_original: str) -> str:
    s = norm_space(texto_original or "")
    s = re.sub(r"\s{2,}", " ", s).strip()

    s = re.sub(r"^(A vencer|a vencer)\s+", "", s, flags=re.I)
    s = re.sub(r"^(Vencidos?|vencid\w*)\s+", "", s, flags=re.I)

    if not re.search(r"\d", s) and not re.search(r"(até|de|a|dias)", s, flags=re.I):
        return "não-faixa"

    return s


def prefixar_faixa(tipo: str, faixa_limpa: str) -> str:
    if faixa_limpa == "não-faixa":
        if tipo in {"Prejuízo", "A liberar"}:
            return tipo
        return "não-faixa"
    if tipo == "A Vencer":
        return f"A vencer {faixa_limpa}"
    if tipo == "Vencido":
        return f"Vencidos {faixa_limpa}"
    if tipo == "Prejuízo":
        return faixa_limpa if faixa_limpa.lower().startswith("preju") else f"Prejuízo {faixa_limpa}"
    if tipo == "A liberar":
        return faixa_limpa if faixa_limpa.lower().startswith("a liberar") else "A liberar"
    return faixa_limpa


def canonical_modalidade(texto: str) -> str:
    s = texto or ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    for rx, canon in MOD_MATCH:
        if rx.search(s):
            if canon == "Outros créditos":
                s_clean = re.sub(r"\s+", " ", s).strip().lower()
                if not re.fullmatch(r"outros( creditos)?", s_clean):
                    continue
            return canon
    return "Outros créditos"


def is_modalidade_line(texto: str) -> bool:
    s = texto or ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    for rx, canon in MOD_MATCH:
        if rx.search(s):
            if canon == "Outros créditos":
                s_clean = re.sub(r"\s+", " ", s).strip().lower()
                if not re.fullmatch(r"outros( creditos)?", s_clean):
                    return False
            return True
    return False


# -----------------------------
# OCR pontual (páginas alvo)
# -----------------------------
def ocr_pdf_pages(
    pdf_path: Path,
    poppler: Path,
    dpi: int,
    page_numbers_1based: List[int],
    debug_dir: Optional[Path] = None,
) -> Dict[int, str]:
    """
    OCR só das páginas necessárias. Retorna dict {page_1based: text}
    """
    texts: Dict[int, str] = {}
    if not page_numbers_1based:
        return texts

    pages = convert_from_path(
        str(pdf_path),
        dpi=dpi,
        poppler_path=str(poppler),
        fmt="png",
        first_page=min(page_numbers_1based),
        last_page=max(page_numbers_1based),
    )

    start = min(page_numbers_1based)
    for i, pil_img in enumerate(pages):
        pno = start + i
        if pno not in page_numbers_1based:
            continue
        img_rgb = np.array(pil_img)
        img_bgr = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2BGR)

        gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        th = cv2.adaptiveThreshold(
            gray,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            41,
            15,
        )

        try:
            txt = pytesseract.image_to_string(th, lang="por+eng") or ""
        except Exception:
            txt = pytesseract.image_to_string(th, lang="eng") or ""

        texts[pno] = txt

        if debug_dir:
            debug_dir.mkdir(parents=True, exist_ok=True)
            cv2.imwrite(str(debug_dir / f"ocr_p{pno:02d}.png"), th)

    return texts


# -----------------------------
# Leitura PDF (texto nativo)
# -----------------------------
def read_pdf_text_native(pdf_path: Path) -> Tuple[List[str], str]:
    if pdfplumber is None:
        raise RuntimeError("pdfplumber não está instalado. Rode: pip install pdfplumber")

    page_texts: List[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for p in pdf.pages:
            t = p.extract_text() or ""
            page_texts.append(t)
    full = "\n".join(page_texts)
    return page_texts, full


# -----------------------------
# Páginas-alvo
# -----------------------------
@dataclass
class TargetPages:
    carteira_pages: List[int]
    detalhamento_pages: List[int]
    limite_pages: List[int]
    has_detalhamento: bool


def locate_target_pages(page_texts: List[str]) -> TargetPages:
    carteira = []
    detal = []
    limite = []
    has_detal = False

    for idx0, t in enumerate(page_texts):
        pno = idx0 + 1
        u = t.upper()
        if "CARTEIRA ATIVA" in u:
            carteira.append(pno)
        if "DETALHAMENTO" in u:
            detal.append(pno)
            has_detal = True
        if "LIMITE TOTAL" in u:
            limite.append(pno)

    return TargetPages(
        carteira_pages=sorted(set(carteira)),
        detalhamento_pages=sorted(set(detal)),
        limite_pages=sorted(set(limite)),
        has_detalhamento=has_detal,
    )


# -----------------------------
# Cards
# -----------------------------
def extract_cards(full_text: str) -> Dict[str, str]:
    out: Dict[str, str] = {}

    cnpj = parse_cnpj(full_text) or "A confirmar"
    razao = parse_razao_social(full_text) or "A confirmar"
    dt = parse_database_consulta(full_text)
    data_base = fmt_datebase(dt)

    m = re.search(r"CARTEIRA ATIVA\s*\(A\)\s*R\$\s*([\d\.\,]+)", full_text, flags=re.I)
    valor_total = money_ptbr(m.group(1)) if m else "--"

    m2 = re.search(r"Limite Total\s*R\$\s*([\d\.\,]+)", full_text, flags=re.I)
    limite_total = money_ptbr(m2.group(1)) if m2 else "--"

    risco = "--"

    lines = [norm_space(x) for x in (full_text or "").splitlines()]

    def _norm_percent(raw: str) -> str:
        s = re.sub(r"[^\d\.,]", "", raw or "")
        if not s:
            return "--"
        s = s.replace(".", ",")
        if not s.endswith("%"):
            s = f"{s}%"
        return s

    def _extract_ints_no_percent(raw: str) -> List[str]:
        cleaned = re.sub(r"\b\d{1,3}[\.,]\d{1,2}\s*%", " ", raw or "")
        return re.findall(r"\b(\d{1,5})\b", cleaned)

    def _find_percent_after(label_re: str) -> str:
        for i, ln in enumerate(lines):
            if re.search(label_re, ln, flags=re.I):
                m = re.search(r"(\d{1,3}[\.,]\d{1,2})\s*%", ln)
                if m:
                    return _norm_percent(m.group(0))
                if i + 1 < len(lines):
                    m2 = re.search(r"(\d{1,3}[\.,]\d{1,2})\s*%", lines[i + 1])
                    if m2:
                        return _norm_percent(m2.group(0))
        return "--"

    def _find_int_after(label_re: str) -> str:
        for i, ln in enumerate(lines):
            if re.search(label_re, ln, flags=re.I):
                nums = _extract_ints_no_percent(ln)
                if nums:
                    return nums[0]
                if i + 1 < len(lines):
                    nums2 = _extract_ints_no_percent(lines[i + 1])
                    if nums2:
                        return nums2[0]
        return "--"

    vol_proc = _find_percent_after(r"Vol\.\s*processado")
    doc_proc = _find_percent_after(r"Doc\.\s*processados")

    tot_inst = _find_int_after(r"Total de institui[cç][oõ]es")
    tot_ops = _find_int_after(r"Total de opera[cç][oõ]es")

    def _extract_risco() -> str:
        for i, ln in enumerate(lines):
            if re.search(r"Classifica[cç][aã]o de risco", ln, flags=re.I):
                pool = ln
                if i + 1 < len(lines):
                    pool = pool + " " + lines[i + 1]
                pool = pool.replace("R$", " ")
                tokens = re.findall(r"\b([A-Z]{1,2})\b", pool)
                tokens = [t for t in tokens if t != "R"]
                if tokens:
                    return tokens[-1].upper()
        return "--"

    risco = _extract_risco()

    for i, ln in enumerate(lines):
        if re.search(r"Total de institui[cç][oõ]es", ln, flags=re.I) and re.search(
            r"Total de opera[cç][oõ]es", ln, flags=re.I
        ):
            if i + 1 < len(lines):
                nums = _extract_ints_no_percent(lines[i + 1])
                if len(nums) >= 2:
                    tot_inst = nums[0]
                    tot_ops = nums[1]
            break

    if vol_proc == "--" or tot_inst == "--" or tot_ops == "--":
        for i, ln in enumerate(lines):
            if re.search(r"Vol\.\s*processado", ln, flags=re.I):
                if i + 1 < len(lines):
                    cand = lines[i + 1]
                    m = re.search(r"(\d{1,3}[\.,]\d{1,2})\s*%\s+(\d+)\s+(\d+)", cand)
                    if m:
                        vol_proc = _norm_percent(m.group(1))
                        tot_inst = m.group(2)
                        tot_ops = m.group(3)
                        break

    out["Razão Social"] = razao
    out["Valor Total"] = valor_total
    out["Limite de Crédito"] = limite_total
    out["Classificação de risco"] = risco
    out["Total Instituições"] = tot_inst
    out["Total Operações"] = tot_ops
    out["Volume processado (%)"] = vol_proc
    out["Doc processado (%)"] = doc_proc

    out["_CNPJ"] = cnpj
    out["_DATA_BASE"] = data_base
    out["_ANO"] = year_from_dt(dt)
    out["_RAZAO"] = razao
    return out


# -----------------------------
# Aba 4: Prazos (CARTEIRA ATIVA)
# -----------------------------
def extract_prazos_from_text(full_text: str, meta: Dict[str, str]) -> List[Dict[str, str]]:
    lines = [norm_space(x) for x in (full_text or "").splitlines()]
    rows: List[Dict[str, str]] = []

    in_carteira = False
    mode: Optional[str] = None

    saw_vencido_header = False
    vencido_pairs = 0

    for ln in lines:
        u = ln.upper()

        if "CARTEIRA ATIVA" in u:
            in_carteira = True
            continue

        if in_carteira and ("DETALHAMENTO DOS REGISTROS" in u or "DETALHAMENTO" == u):
            break

        if not in_carteira:
            continue

        if re.search(r"\bA VENCER\b", u):
            mode = "A Vencer"
            continue

        if re.search(r"CR[ÉE]DITOS VENCIDOS|\bVENCIDO", u):
            mode = "Vencido"
            saw_vencido_header = True
            continue

        if re.search(r"\bTOTAL\b", u):
            continue
        ln_clean = re.sub(r"\s+\d{1,3}[\.,]\d{1,2}%.*$", "", ln)
        ln_clean = norm_space(ln_clean)
        if "R$" not in ln:
            continue

        m = re.search(r"^(.+?)\s+R\$\s*([\d\.\,]+)", ln_clean)
        if not m:
            continue

        faixa_raw = m.group(1)
        valor_raw = m.group(2)

        if re.search(r"Cr[ée]ditos\s+a\s+vencer", faixa_raw, flags=re.I):
            continue

        faixa_limpa = faixa_texto_exato(faixa_raw)
        if faixa_limpa == "não-faixa":
            continue

        if mode == "Vencido":
            vencido_pairs += 1

        faixa_final = prefixar_faixa(mode or "A Vencer", faixa_limpa)
        rows.append(
            {
                "Faixa de prazo": faixa_final,
                "Valor Total (R$)": money_ptbr(valor_raw),
                "Referência (dias)": referencia_dias(faixa_limpa),
                "Ano": meta["_ANO"],
                "CNPJ": meta["_CNPJ"],
                "Razão Social": meta["_RAZAO"],
                "Data base (consulta)": meta["_DATA_BASE"],
            }
        )

    if saw_vencido_header and vencido_pairs == 0:
        rows = [r for r in rows if not str(r["Faixa de prazo"]).upper().startswith("VENCIDOS ")]

    return rows


# -----------------------------
# Aba 2: Modalidades (resumo)
# -----------------------------
def extract_modalidades_from_text(full_text: str, meta: Dict[str, str]) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []

    def _to_num(raw: str) -> float:
        s = re.sub(r"[^\d\.,]", "", str(raw or ""))
        if not s:
            return 0.0
        s = s.replace(".", "").replace(",", ".")
        try:
            return float(s)
        except Exception:
            return 0.0

    def _extract_from_detalhamento() -> List[Dict[str, str]]:
        det = re.search(r"\nDETALHAMENTO\b(.+?)\nLimite Total", full_text, flags=re.I | re.S)
        det_block = det.group(1) if det else ""
        det_lines = [norm_space(x) for x in det_block.splitlines() if norm_space(x)]
        acc: Dict[str, float] = {}
        for ln in det_lines:
            if not re.search(r"\bModalidade\b", ln, flags=re.I) and "R$" in ln:
                prefix = ln.split("R$")[0]
                prefix_norm = unicodedata.normalize("NFD", prefix)
                prefix_norm = "".join(ch for ch in prefix_norm if unicodedata.category(ch) != "Mn")
                has_mod = any(rx.search(prefix_norm) for rx, _ in MOD_MATCH)
                if not has_mod:
                    continue
                canon = canonical_modalidade(prefix)
                if canon not in MOD_CANON and canon != "Outros créditos":
                    continue
                val = find_first_money(ln)
                if val:
                    acc[canon] = acc.get(canon, 0.0) + _to_num(val)
        out: List[Dict[str, str]] = []
        for canon, total in acc.items():
            out.append(
                {
                    "Modalidade": canon if canon in MOD_CANON else "Outros créditos",
                    "Valor Total (R$)": money_ptbr(f"{total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")),
                    "Ano": meta["_ANO"],
                    "CNPJ": meta["_CNPJ"],
                    "Razão Social": meta["_RAZAO"],
                    "Data base (consulta)": meta["_DATA_BASE"],
                }
            )
        return out

    m = re.search(r"Detalhamento dos registros(.+?)(?:\nDETALHAMENTO\b|\Z)", full_text, flags=re.I | re.S)
    if not m:
        return _extract_from_detalhamento()

    block = m.group(1)
    lines = [norm_space(x) for x in block.splitlines() if norm_space(x)]

    i = 0
    last_money = None
    while i < len(lines):
        ln = lines[i]
        if re.search(r"\bTotal\b", ln, flags=re.I):
            i += 1
            continue
        if not re.search(r"[A-Za-zÀ-ÿ]", ln):
            i += 1
            continue

        monies = [m.group(1) for m in re.finditer(r"R\$\s*([\d\.\,]+)", ln)]
        if monies:
            last_money = monies[-1]

        if not is_modalidade_line(ln):
            i += 1
            continue
        canon = canonical_modalidade(ln)

        if canon in MOD_CANON or canon == "Outros créditos":
            val = find_first_money(ln)
            if not val and last_money:
                val = last_money
            if not val and i + 1 < len(lines):
                val = find_first_money(lines[i + 1])

            if val:
                rows.append(
                    {
                        "Modalidade": canon if canon in MOD_CANON else "Outros créditos",
                        "Valor Total (R$)": money_ptbr(val),
                        "Ano": meta["_ANO"],
                        "CNPJ": meta["_CNPJ"],
                        "Razão Social": meta["_RAZAO"],
                        "Data base (consulta)": meta["_DATA_BASE"],
                    }
                )
        i += 1

    dedup = {}
    for r in rows:
        dedup[r["Modalidade"]] = r
    rows = list(dedup.values())

    card_total = _to_num(meta.get("Valor Total", ""))
    mod_total = sum(_to_num(r.get("Valor Total (R$)", "")) for r in rows)
    if card_total and mod_total and abs(card_total - mod_total) > 0.01:
        det_rows = _extract_from_detalhamento()
        if det_rows:
            return det_rows

    return rows


# -----------------------------
# Aba 3: LinhasModalidade (DETALHAMENTO + Limite)
# -----------------------------
def extract_linhasmodalidade_from_text(full_text: str, meta: Dict[str, str]) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []

    m = re.search(r"\nDETALHAMENTO\b(.+?)\nLimite Total", full_text, flags=re.I | re.S)
    det_block = m.group(1) if m else ""

    m2 = re.search(r"\nLimite Total(.+)\Z", full_text, flags=re.I | re.S)
    lim_block = m2.group(1) if m2 else ""

    lines = [norm_space(x) for x in det_block.splitlines() if norm_space(x)]
    current_modalidade = "Outros créditos"
    pending_produto = ""
    current_produto = ""
    pending_rows: List[Dict[str, str]] = []
    pending_prod_parts: List[str] = []

    def _has_letters(s: str) -> bool:
        return bool(re.search(r"[A-Za-zÀ-ÿ]", s or ""))

    def flush_pending_rows(prod: str) -> None:
        nonlocal pending_rows
        if not pending_rows:
            return
        for r in pending_rows:
            r["Produto/Linha"] = prod
            rows.append(r)
        pending_rows = []

    def update_recent_produto(old: str, new: str) -> None:
        for i in range(len(rows) - 1, -1, -1):
            if rows[i]["Produto/Linha"] == old and rows[i]["Modalidade"] == current_modalidade:
                rows[i]["Produto/Linha"] = new
            else:
                break

    started = False
    for ln in lines:
        ln_up = _no_accents_upper(ln)
        if not started:
            if ln_up == "DETALHAMENTO" or ln_up.startswith("MODALIDADE"):
                started = True
            continue
        if "%" in ln:
            continue
        if re.search(r"\bTotal\b", ln, flags=re.I):
            continue
        if re.search(r"\bModalidade\b", ln, flags=re.I):
            continue
        if ln.upper() == "DETALHAMENTO":
            continue

        is_mod_line = is_modalidade_line(ln)
        canon_line = canonical_modalidade(ln) if is_mod_line else ""
        if is_mod_line and canon_line in MOD_CANON and (
            canon_line != "Outros créditos" or re.search(r"\boutros\b", ln, flags=re.I)
        ):
            if canon_line != current_modalidade:
                if pending_rows:
                    flush_pending_rows("A confirmar")
                current_modalidade = canon_line
                current_produto = ""
                pending_produto = ""
                pending_prod_parts = []

        if "R$" not in ln:
            if is_mod_line and canon_line in MOD_CANON:
                canon_up = _no_accents_upper(canon_line)
                if ln_up in {canon_up, "OUTROS", "OUTROS CREDITOS"}:
                    if pending_rows and not current_produto:
                        current_produto = canon_line
                        flush_pending_rows(current_produto)
                    continue
            if not _has_letters(ln):
                continue
            if re.fullmatch(r"Não", ln, flags=re.I) or ln_up == "NAO":
                continue
            if re.fullmatch(r"\d{1,3}(?:\.\d{3})*,\d{2}", ln):
                continue
            if (
                current_modalidade == "Títulos descontado Direitos creditórios descontados"
                and "CREDITORIOS DESCONTADOS" in ln_up
            ):
                continue
            if current_produto and not pending_rows and ln[:1].islower():
                old_prod = current_produto
                current_produto = norm_space(current_produto + " " + ln)
                update_recent_produto(old_prod, current_produto)
                continue
            pending_prod_parts = [ln]
            pending_produto = norm_space(" ".join(pending_prod_parts))
            current_produto = pending_produto
            flush_pending_rows(current_produto)
            continue

        if "R$" in ln and re.search(r"\bA vencer\b|\bvencid|\bPreju[ií]zo\b|\bA liberar\b", ln, flags=re.I):
            prefix = re.split(r"\b(?:A vencer|Vencid\w*|Preju[ií]zo|A liberar)\b", ln, flags=re.I)[0]
            prefix = norm_space(prefix)
            if current_modalidade and prefix.lower().startswith(current_modalidade.lower()):
                prefix = norm_space(prefix[len(current_modalidade):])
            prefix_clean = prefix
            if "R$" in prefix_clean or re.search(r"\d", prefix_clean):
                prefix_clean = re.sub(r"R\$\s*[\d\.,]+", " ", prefix_clean)
                prefix_clean = re.sub(r"\bN[ãa]o\b", " ", prefix_clean, flags=re.I)
                prefix_clean = re.sub(r"\d", " ", prefix_clean)
                prefix_clean = norm_space(prefix_clean)
            if prefix_clean and _has_letters(prefix_clean) and "R$" not in prefix_clean:
                if pending_prod_parts:
                    pending_prod_parts.append(prefix_clean)
                elif not current_produto:
                    current_produto = prefix_clean
                else:
                    old_prod = current_produto
                    current_produto = norm_space(current_produto + " " + prefix_clean)
                    update_recent_produto(old_prod, current_produto)
                if pending_prod_parts and not current_produto:
                    current_produto = norm_space(" ".join(pending_prod_parts))
                    pending_prod_parts = []
                if current_produto:
                    flush_pending_rows(current_produto)

            pairs = re.findall(
                r"((?:A vencer|Vencid\w*|Preju[ií]zo|A liberar).{0,80}?)\s*R\$\s*([\d\.\,]+)",
                ln,
                flags=re.I,
            )
            for faixa_raw, val_raw in pairs:
                if re.search(r"preju", faixa_raw, flags=re.I):
                    tipo = "Prejuízo"
                elif re.search(r"\bliber", faixa_raw, flags=re.I):
                    tipo = "A liberar"
                elif re.search(r"vencid", faixa_raw, flags=re.I):
                    tipo = "Vencido"
                else:
                    tipo = "A Vencer"
                faixa_limpa = faixa_texto_exato(faixa_raw)
                if faixa_limpa == "não-faixa" and tipo in {"Prejuízo", "A liberar"}:
                    faixa_limpa = tipo
                if faixa_limpa == "não-faixa":
                    continue

                faixa_final = prefixar_faixa(tipo, faixa_limpa)
                row = {
                    "Modalidade": current_modalidade or "Outros créditos",
                    "Produto/Linha": current_produto or "A confirmar",
                    "Tipo": tipo if tipo in ("A Vencer", "Vencido", "Prejuízo", "A liberar") else "A Vencer",
                    "Faixa de prazo": faixa_final,
                    "Valor (R$)": money_ptbr(val_raw),
                    "Ano": meta["_ANO"],
                    "CNPJ": meta["_CNPJ"],
                    "Razão Social": meta["_RAZAO"],
                    "Data base (consulta)": meta["_DATA_BASE"],
                }
                if current_produto:
                    rows.append(row)
                else:
                    pending_rows.append(row)
            continue

    lim_lines = [norm_space(x) for x in lim_block.splitlines() if norm_space(x)]
    current_lim_prod = ""
    pending_lim_rows: List[Dict[str, str]] = []
    for idx, ln in enumerate(lim_lines):
        ln_up = _no_accents_upper(ln)
        if ln_up.startswith("LIMITE TOTAL"):
            continue

        if "R$" not in ln and not re.search(r"Limite com vencimento", ln, flags=re.I):
            if re.fullmatch(r"\d{1,3}(?:\.\d{3})*,\d{2}", ln):
                continue
            if re.search(r"\bdias\b", ln, flags=re.I) and re.search(r"\d{1,3}(?:\.\d{3})*,\d{2}", ln):
                prod_candidate = re.split(r"\bdias\b", ln, flags=re.I)[0]
                prod_candidate = re.sub(r"\d", "", prod_candidate).strip()
                if prod_candidate:
                    current_lim_prod = prod_candidate
                    if pending_lim_rows:
                        for r in pending_lim_rows:
                            r["Produto/Linha"] = current_lim_prod
                            rows.append(r)
                        pending_lim_rows = []
                continue
            if ln_up.startswith("LIMITE") and ln_up != "LIMITE":
                prod = norm_space(re.sub(r"^\s*Limite\s*", "", ln, flags=re.I))
                if prod:
                    current_lim_prod = prod
                    if pending_lim_rows:
                        for r in pending_lim_rows:
                            r["Produto/Linha"] = current_lim_prod
                            rows.append(r)
                        pending_lim_rows = []
            elif ln and not ln_up.startswith("LIMITE"):
                current_lim_prod = ln
                if pending_lim_rows:
                    for r in pending_lim_rows:
                        r["Produto/Linha"] = current_lim_prod
                        rows.append(r)
                    pending_lim_rows = []
            continue

        if re.search(r"Limite com vencimento", ln, flags=re.I):
            limite_match = re.search(r"Limite com vencimento", ln, flags=re.I)
            limite_tail = ln[limite_match.start():] if limite_match else ln
            money_hits = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}", limite_tail)
            val = money_hits[-1] if money_hits else None
            mfaixa = re.search(r"(Limite com vencimento.*?)(?:R\$|$)", limite_tail, flags=re.I)
            faixa_raw = norm_space(mfaixa.group(1)) if mfaixa else ln
            force_pending_prod = bool(
                re.search(r"Limite\s+R\$", ln, flags=re.I)
                and re.search(r"Não\s+Limite com vencimento|Nao\s+Limite com vencimento", ln, flags=re.I)
            )
            if idx + 1 < len(lim_lines):
                nxt = lim_lines[idx + 1]
                if val is None:
                    mval = re.findall(r"\d{1,3}(?:\.\d{3})*,\d{2}", nxt)
                    if mval:
                        val = mval[-1]
                nxt_prefix = re.split(r"\d{1,3}(?:\.\d{3})*,\d{2}", nxt)[0].strip()
                if nxt_prefix and re.search(r"\bdias\b", nxt_prefix, flags=re.I):
                    if re.search(r"[A-Za-zÀ-ÿ]", nxt_prefix) and not re.match(r"^\d", nxt_prefix):
                        if not current_lim_prod:
                            prod_candidate = re.split(r"\bdias\b", nxt_prefix, flags=re.I)[0].strip()
                            prod_candidate = re.sub(r"\d", "", prod_candidate).strip()
                            if prod_candidate:
                                current_lim_prod = prod_candidate
                                if pending_lim_rows:
                                    for r in pending_lim_rows:
                                        r["Produto/Linha"] = current_lim_prod
                                        rows.append(r)
                                    pending_lim_rows = []
                        if not re.search(r"\bdias\b", faixa_raw, flags=re.I):
                            faixa_raw = norm_space(faixa_raw + " dias")
                    else:
                        if not re.search(r"\bdias\b", faixa_raw, flags=re.I):
                            faixa_raw = norm_space(faixa_raw + " " + nxt_prefix)
            faixa_limpa = faixa_texto_exato(faixa_raw)
            if val and faixa_limpa != "não-faixa":
                row = {
                    "Modalidade": "Limite",
                    "Produto/Linha": current_lim_prod or "A confirmar",
                    "Tipo": "Limite",
                    "Faixa de prazo": prefixar_faixa("Limite", faixa_limpa),
                    "Valor (R$)": money_ptbr(val),
                    "Ano": meta["_ANO"],
                    "CNPJ": meta["_CNPJ"],
                    "Razão Social": meta["_RAZAO"],
                    "Data base (consulta)": meta["_DATA_BASE"],
                }
                if current_lim_prod and not force_pending_prod:
                    rows.append(row)
                else:
                    pending_lim_rows.append(row)
            continue

        if re.search(r"\bLimite\b", ln, flags=re.I) and "R$" in ln:
            before = ln.split("R$")[0]
            before = re.sub(r"^\s*Limite\s*", "", before, flags=re.I)
            prod = norm_space(before)
            if prod:
                current_lim_prod = prod
                if pending_lim_rows:
                    for r in pending_lim_rows:
                        r["Produto/Linha"] = current_lim_prod
                        rows.append(r)
                    pending_lim_rows = []

    for r in rows:
        if r["Tipo"] not in ("A Vencer", "Vencido", "Prejuízo", "A liberar", "Limite"):
            raise RuntimeError(
                "ERRO: Tipo inválido em LinhasModalidade (deve ser A Vencer/Vencido/Prejuízo/A liberar/Limite)."
            )

    return rows


# -----------------------------
# Montagem Excel (4 abas)
# -----------------------------
def build_xlsx(
    out_xlsx: Path,
    cards: Dict[str, str],
    modalidades: List[Dict[str, str]],
    linhas: List[Dict[str, str]],
    prazos: List[Dict[str, str]],
) -> None:
    cnpj = cards.get("_CNPJ", "A confirmar")
    data_base = cards.get("_DATA_BASE", "A confirmar")

    rows_cards = []
    for k in CARDS_ORDER:
        rows_cards.append(
            {
                "Campo": k,
                "Valor": cards.get(k, "--"),
                "CNPJ": cnpj,
                "Data base (consulta)": data_base,
            }
        )
    df_cards = pd.DataFrame(rows_cards, columns=["Campo", "Valor", "CNPJ", "Data base (consulta)"])

    df_mod = (
        pd.DataFrame(
            modalidades,
            columns=["Modalidade", "Valor Total (R$)", "Ano", "CNPJ", "Razão Social", "Data base (consulta)"],
        )
        if modalidades
        else pd.DataFrame(
            columns=["Modalidade", "Valor Total (R$)", "Ano", "CNPJ", "Razão Social", "Data base (consulta)"]
        )
    )

    df_linhas = (
        pd.DataFrame(
            linhas,
            columns=[
                "Modalidade",
                "Produto/Linha",
                "Tipo",
                "Faixa de prazo",
                "Valor (R$)",
                "Ano",
                "CNPJ",
                "Razão Social",
                "Data base (consulta)",
            ],
        )
        if linhas
        else pd.DataFrame(
            columns=[
                "Modalidade",
                "Produto/Linha",
                "Tipo",
                "Faixa de prazo",
                "Valor (R$)",
                "Ano",
                "CNPJ",
                "Razão Social",
                "Data base (consulta)",
            ]
        )
    )

    df_prazos = (
        pd.DataFrame(
            prazos,
            columns=[
                "Faixa de prazo",
                "Valor Total (R$)",
                "Referência (dias)",
                "Ano",
                "CNPJ",
                "Razão Social",
                "Data base (consulta)",
            ],
        )
        if prazos
        else pd.DataFrame(
            columns=[
                "Faixa de prazo",
                "Valor Total (R$)",
                "Referência (dias)",
                "Ano",
                "CNPJ",
                "Razão Social",
                "Data base (consulta)",
            ]
        )
    )

    df_cards = normalize_card_fields(normalize_headers(df_cards))
    df_mod = normalize_headers(df_mod)
    df_linhas = normalize_headers(df_linhas)
    df_prazos = normalize_headers(df_prazos)

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df_cards.to_excel(writer, index=False, sheet_name="Cards")
        df_mod.to_excel(writer, index=False, sheet_name="Modalidades")
        df_linhas.to_excel(writer, index=False, sheet_name="LinhasModalidade")
        df_prazos.to_excel(writer, index=False, sheet_name="Prazos (faixas)")


# -----------------------------
# Pipeline por PDF
# -----------------------------
def process_one_scr_pdf(
    pdf_path: Path,
    outdir: Path,
    poppler: Path,
    dpi: int,
    tesseract: str = "",
    debug: bool = False,
) -> Tuple[bool, str]:
    """
    Retorna (ok, msg).
    Se falha dura: ok=False e msg = "ERRO: ...".
    """
    if tesseract:
        pytesseract.pytesseract.tesseract_cmd = tesseract

    page_texts, full_text = read_pdf_text_native(pdf_path)

    targets = locate_target_pages(page_texts)

    def is_poor(txt: str) -> bool:
        return len(norm_space(txt)) < 50

    pages_need_ocr = []
    for p in (targets.carteira_pages + targets.detalhamento_pages + targets.limite_pages):
        if p <= 0 or p > len(page_texts):
            continue
        if is_poor(page_texts[p - 1]):
            pages_need_ocr.append(p)

    if pages_need_ocr:
        debug_dir = (outdir / (pdf_path.stem + "_debug")) if debug else None
        ocr_texts = ocr_pdf_pages(
            pdf_path,
            poppler=poppler,
            dpi=dpi,
            page_numbers_1based=sorted(set(pages_need_ocr)),
            debug_dir=debug_dir,
        )
        for p, txt in ocr_texts.items():
            page_texts[p - 1] = txt
        full_text = "\n".join(page_texts)

    cards = extract_cards(full_text)

    prazos = extract_prazos_from_text(full_text, meta=cards)

    modalidades = extract_modalidades_from_text(full_text, meta=cards)

    linhas = extract_linhasmodalidade_from_text(full_text, meta=cards)

    if targets.has_detalhamento and len(linhas) == 0:
        det_pages = targets.detalhamento_pages[:]
        if det_pages:
            debug_dir = (outdir / (pdf_path.stem + "_debug_det")) if debug else None
            ocr_texts = ocr_pdf_pages(
                pdf_path,
                poppler=poppler,
                dpi=max(dpi, 300),
                page_numbers_1based=det_pages,
                debug_dir=debug_dir,
            )
            for p, txt in ocr_texts.items():
                page_texts[p - 1] = txt
            full_text2 = "\n".join(page_texts)
            cards2 = extract_cards(full_text2)
            linhas = extract_linhasmodalidade_from_text(full_text2, meta=cards2)

        if len(linhas) == 0:
            return (False, "ERRO: Falha na extração SCR (prazos/linhas). Não gerar arquivo.")

    razao = cards.get("_RAZAO", "A_confirmar")
    safe = re.sub(r"[\\/:*?\"<>|]+", " ", razao).strip()
    safe = re.sub(r"\s+", " ", safe)
    if not safe or safe.lower() in {"a confirmar", "a_confirmar"}:
        safe = pdf_path.stem

    ano = cards.get("_ANO", "A confirmar")
    suffix = f"_{ano}" if ano and ano != "A confirmar" else ""
    out_xlsx = outdir / f"SCR_{safe}{suffix}.xlsx"

    build_xlsx(out_xlsx, cards, modalidades, linhas, prazos)
    return (True, f"OK: {out_xlsx.name}")


# -----------------------------
# Varredura de INPUT (scr_cedente_*.pdf)
# -----------------------------
def find_scr_cedente_pdfs(input_dir: Path) -> List[Path]:
    pdfs: List[Path] = []
    for root, _, files in os.walk(str(input_dir)):
        for fn in files:
            low = fn.lower()
            if low.endswith(".pdf") and "scr_cedente" in low:
                pdfs.append(Path(root) / fn)
    pdfs.sort()
    return pdfs


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="Pasta base 01_INPUT (varre subpastas)")
    ap.add_argument("--outdir", required=True, help=r"Pasta de saída (ex: ...\03_OUTPUT\4. SCR CEDENTE)")
    ap.add_argument("--poppler", required=True, help="Pasta bin do Poppler (pdfinfo/pdftoppm)")
    ap.add_argument("--dpi", type=int, default=250, help="DPI para OCR pontual (use 250-350)")
    ap.add_argument("--tesseract", default="", help="Caminho do tesseract.exe (opcional)")
    ap.add_argument("--debug", action="store_true", help="Salva imagens OCR de debug (somente quando OCR rodar)")
    args = ap.parse_args()

    input_dir = Path(args.input)
    outdir = Path(args.outdir)
    poppler = Path(args.poppler)

    pdfs = find_scr_cedente_pdfs(input_dir)
    if not pdfs:
        print("ERRO: Nenhum PDF scr_cedente_*.pdf encontrado no input.")
        return

    ok_count = 0
    for pdf_path in pdfs:
        ok, msg = process_one_scr_pdf(
            pdf_path=pdf_path,
            outdir=outdir,
            poppler=poppler,
            dpi=args.dpi,
            tesseract=args.tesseract,
            debug=args.debug,
        )
        print(f"[{pdf_path.name}] {msg}")
        if ok:
            ok_count += 1

    print(f"\nConcluído: {ok_count}/{len(pdfs)} arquivos gerados em: {outdir}")


if __name__ == "__main__":
    main()
