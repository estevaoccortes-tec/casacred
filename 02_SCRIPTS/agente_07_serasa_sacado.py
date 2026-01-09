#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
agente_07_serasa_sacado.py

Organizador SERASA SACADO -> Excel (STRICT/DUAL-SOURCE)

SAIDA:
- 1 .xlsx com 1 aba "Planilha1"
- Colunas fixas: Campo | Informação | CNPJ  (titulo fixo)
- Fonte PDF (exclusivo): Identificacao + Liminar + Total em anotacoes negativas + DETALHES (PEFIN/REFIN/Protestos/Acoes/Cheques) + Identificador (CNPJ/CPF)
- Fonte IMAGEM (exclusivo): Consultas - MM/AAAA (grafico) e Consulta 1..5 (tabela)

VALIDACAO (falha dura):
- Campo fora da whitelist / padrao exato
- Campo contendo "–" ou "—"
- Colunas fora da ordem exata
"""

from __future__ import annotations

import argparse
import os
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import pdfplumber

import pytesseract
from pdf2image import convert_from_path

# OpenCV e recomendado para o grafico (barras).
# Se nao houver, o script ainda gera XLSX sem Consultas/Consulta i.
try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False


# =========================
# Regex / Constantes
# =========================

BASE_DIR = Path(__file__).resolve().parents[1]
PASTA_INPUT = BASE_DIR / "01_INPUT"
PASTA_DESTINO = BASE_DIR / "03_OUTPUT" / "10. SERASA SACADO"

CNPJ_RE = re.compile(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b")
CPF_RE = re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")
DATE_RE = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")
MONEY_RE = re.compile(r"(R\$\s*)?\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2})")
PERCENT_RE = re.compile(r"\b\d{1,3}(?:,\d{1,2})?\s*%\b")

UF_SET = {
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS",
    "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC",
    "SP", "SE", "TO",
}

PT_MONTHS = {
    "jan": "01", "fev": "02", "mar": "03", "abr": "04", "mai": "05", "jun": "06",
    "jul": "07", "ago": "08", "set": "09", "out": "10", "nov": "11", "dez": "12",
}
PT_MONTHS_NUM = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}
INV_PT = {v: k for k, v in PT_MONTHS_NUM.items()}

START_LABEL = "Nov/2024"
EXPECTED_BARS = 13
DPI_GRAFICO = 200

IDENT_FIELDS_ORDER = [
    "Nome",
    "CNPJ",
    "CPF",
    "Serasa Score",
    "CNAE",
    "Localização",
    "Liminar",
    "Total em anotações negativas",
]

LIMINAR_KEYS = [
    "NADA CONSTA",
    "LIMINAR",
    "ART. 43",
    "DECISAO JUDICIAL",
    "RJ",
    "SUSPENSAO",
    "BLOQUEIO JUDICIAL",
]


# =========================
# Campo permitido (STRICT)
# =========================

def _campo_valido(campo: str) -> bool:
    if "–" in campo or "—" in campo:
        return False

    exatos = {
        "Nome",
        "CNPJ",
        "CPF",
        "Serasa Score",
        "CNAE",
        "Localização",
        "Liminar",
        "Total em anotações negativas",
        "Cheques - Motivo",
        "Consulta 1",
        "Consulta 2",
        "Consulta 3",
        "Consulta 4",
        "Consulta 5",
    }
    if campo in exatos:
        return True

    if re.fullmatch(r"Consultas - \d{2}/\d{4}", campo):
        return True

    if re.fullmatch(r"(PEFIN|REFIN) - Registro \d+ \((Modalidade|Valor|Origem|Data)\)", campo):
        return True

    if re.fullmatch(r"Protestos - Registro \d+ \((Data|Valor)\)", campo):
        return True

    if re.fullmatch(r"Ações Judiciais - Registro \d+ \((Data|Valor|Natureza|Cidade)\)", campo):
        return True

    return False


def _fail() -> None:
    raise RuntimeError(
        'ERRO: "Campo invalido (caractere divergente) ou fontes cruzadas. '
        'Consultas: somente IMAGEM. Identificação/Anotações: somente PDF."'
    )


# =========================
# Utils
# =========================

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
    return s2 or "SACADO"

def _clean_nome(val: str) -> str:
    v = _strip(val)
    v = re.sub(r"^(nome\s+fantasia\s*[:\-]\s*)", "", v, flags=re.IGNORECASE)
    v = re.sub(r"^(fantasia\s*[:\-]\s*)", "", v, flags=re.IGNORECASE)
    return _strip(v)

def _ptbr_money_from_any(raw: str) -> Optional[str]:
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


def _ptbr_money_with_rs(raw: str) -> Optional[str]:
    v = _ptbr_money_from_any(raw)
    return f"R$ {v}" if v else None


def _parse_date_br(s: str) -> Optional[datetime]:
    m = DATE_RE.search(s or "")
    if not m:
        return None
    try:
        return datetime.strptime(m.group(0), "%d/%m/%Y")
    except Exception:
        return None


def _clamp(v: int, lo: int, hi: int) -> int:
    return max(lo, min(hi, v))


def _parse_start_label(label: str) -> datetime:
    s = (label or "").strip().lower()
    m = re.search(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\s*[/\-]\s*(20\d{2})", s)
    if not m:
        raise ValueError(f'--start-label invalido: "{label}". Use tipo "Nov/2024".')
    mon = PT_MONTHS_NUM[m.group(1)]
    year = int(m.group(2))
    return datetime(year, mon, 1)


def _add_months(dt: datetime, n: int) -> datetime:
    y = dt.year
    m = dt.month + n
    y += (m - 1) // 12
    m = ((m - 1) % 12) + 1
    return datetime(y, m, 1)


def _fmt_label(dt: datetime) -> str:
    return f"{INV_PT[dt.month].capitalize()}/{dt.year}"


def _mes_ref(dt: datetime) -> str:
    return dt.strftime("%Y-%m-01")


def _list_pdf_sacado(input_dir: Path) -> List[Path]:
    pdfs: List[Path] = []
    for root, _, files in os.walk(str(input_dir)):
        for f in files:
            if not f.lower().endswith(".pdf"):
                continue
            name = f.lower()
            if "serasa" in name and "sacado" in name:
                pdfs.append(Path(root) / f)
    return pdfs


def _choose_image_for_pdf(pdf_path: Path) -> Optional[Path]:
    folder = pdf_path.parent
    pdf_stem = pdf_path.stem.lower()
    candidates: List[Path] = []
    for ext in (".png", ".jpg", ".jpeg"):
        candidates.extend(folder.glob(f"*{ext}"))

    if not candidates:
        return None

    def score(p: Path) -> int:
        name = p.stem.lower()
        s = 0
        if name == pdf_stem:
            s += 6
        if "consulta" in name:
            s += 3
        if "grafico" in name or "gráfico" in name:
            s += 3
        if "sacado" in name:
            s += 2
        if "serasa" in name:
            s += 1
        return s

    scored = sorted(((score(p), p) for p in candidates), key=lambda t: t[0], reverse=True)
    best_score, best_path = scored[0]
    if best_score <= 0 and len(candidates) > 1:
        return None
    return best_path


# =========================
# PDF -> texto (OCR so em paginas "vazias")
# =========================

def _pdf_page_text_native(page) -> str:
    return page.extract_text(layout=True, x_tolerance=1, y_tolerance=2) or ""


def _pdf_page_text_ocr(pdf_path: Path, page_index0: int, poppler_path: Path, dpi: int = 260) -> str:
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
    txt = pytesseract.image_to_string(imgs[0], config="--oem 3 --psm 6") or ""
    return txt


def extract_pdf_text(pdf_path: Path, poppler_path: Optional[Path], ocr_dpi: int = 260) -> str:
    parts: List[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for i, page in enumerate(pdf.pages):
            t = _pdf_page_text_native(page)
            if len(_strip(t)) >= 80:
                parts.append(t)
                continue
            if poppler_path is None:
                parts.append(t)
                continue
            t_ocr = _pdf_page_text_ocr(pdf_path, i, poppler_path=poppler_path, dpi=ocr_dpi)
            use = t_ocr if len(_strip(t_ocr)) > len(_strip(t)) else t
            parts.append(use)
    return "\n".join(parts)


# =========================
# Identificador do sacado (3a coluna "CNPJ")
# =========================

def extract_ids(pdf_text: str) -> Tuple[str, str, str]:
    """
    Retorna (cnpj_linha, cpf_linha, id_escolhido_3a_coluna)
    Regras:
      - 3a coluna (CNPJ) usa CNPJ se existir, senao CPF, senao "A confirmar"
      - Linhas do bloco A: se PJ, CPF vira "-", se PF, CNPJ vira "-"
    """
    cnpj = CNPJ_RE.search(pdf_text or "")
    cpf = CPF_RE.search(pdf_text or "")

    if cnpj:
        cnpj_v = cnpj.group(0)
        cpf_v = "-"
        chosen = cnpj.group(0)
    elif cpf:
        cnpj_v = "-"
        cpf_v = cpf.group(0)
        chosen = cpf.group(0)
    else:
        cnpj_v = "-"
        cpf_v = "-"
        chosen = "A confirmar"

    return cnpj_v, cpf_v, chosen


# =========================
# Extracao identificacao (somente PDF)
# =========================

def _value_after_label(text: str, label: str) -> Optional[str]:
    t = text or ""
    pat = rf"{re.escape(label)}\s*[:\-]?\s*(.+)"
    m = re.search(pat, t, flags=re.IGNORECASE)
    if not m:
        return None
    val = _strip(m.group(1))
    val = re.split(r"\s{2,}|\bOcultar\b|\bAnota", val, flags=re.IGNORECASE)[0]
    return _strip(val) or None


def _extract_nome(pdf_text: str) -> Optional[str]:
    m_header = re.search(r"CNPJ:\s*\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\s*\|\s*([^\n]+)", pdf_text or "", flags=re.IGNORECASE)
    if m_header:
        cand = _clean_nome(m_header.group(1))
        if cand and cand != "-":
            return cand

    nome = _value_after_label(pdf_text, "Nome")
    if nome and len(nome) >= 3:
        nome = _clean_nome(nome)
        if nome and nome != "-":
            return nome
        return None

    m = re.search(r"(Dados\s+do\s+sacado|Dados\s+do\s+Sacado)\s*[:\-]?\s*([^\n]+)", pdf_text, flags=re.IGNORECASE)
    if m:
        cand = _strip(m.group(2))
        if len(cand) >= 3:
            cand = _clean_nome(cand)
            if cand and cand != "-":
                return cand

    lines = [l for l in (pdf_text or "").splitlines() if _strip(l)]
    for l in lines[:20]:
        if len(_strip(l)) >= 8 and not CNPJ_RE.search(l) and not CPF_RE.search(l):
            up = _no_accents_upper(l)
            if "SERASA" in up or "RELATORIO" in up or "RELATÓRIO" in up:
                continue
            cand = _clean_nome(l)
            if cand and cand != "-":
                return cand
    return None


def _extract_serasa_score(pdf_text: str) -> Optional[str]:
    score = _value_after_label(pdf_text, "Serasa Score")
    if score:
        m2 = re.search(r"\b(\d{1,4})\b", score)
        if m2:
            v = int(m2.group(1))
            return str(v) if 0 <= v <= 1000 else None
        score = None

    txt = pdf_text or ""
    m2 = re.search(r"faixa\s+de\s+(\d{3,4})", txt, flags=re.IGNORECASE)
    if m2:
        v = int(m2.group(1))
        if 0 <= v <= 1000:
            return str(v)
    m = re.search(r"Serasa\s+Score\s+Empresas(.{0,600})", txt, flags=re.IGNORECASE | re.DOTALL)
    if m:
        block = m.group(1)
        nums = [int(n) for n in re.findall(r"\b(\d{3,4})\b", block)]
        nums = [n for n in nums if n not in {0, 500, 1000}]
        if nums:
            return str(nums[-1])
    return None


def _extract_cnae(pdf_text: str) -> Optional[str]:
    cnae = _value_after_label(pdf_text, "CNAE")
    if cnae:
        return cnae
    m = re.search(r"\bCNAE\b.*?(\d{4}-\d/\d{2})", pdf_text, flags=re.IGNORECASE)
    return m.group(1) if m else None


def _extract_localizacao(pdf_text: str) -> Optional[str]:
    lines = [l.strip() for l in (pdf_text or "").splitlines() if l.strip()]
    for i, ln in enumerate(lines):
        if "munic" in ln.lower() and "uf" in ln.lower():
            val = re.sub(r"(?i).*munic[ií]pio/uf\s*[:\-]?\s*", "", ln).strip()
            candidates = [val] if val else []
            for j in range(i + 1, min(i + 4, len(lines))):
                candidates.append(lines[j].strip())
            for cand in candidates:
                if not cand:
                    continue
                bad = {"anos", "ano", "funcionarios", "sociedade", "ramo", "opcao", "opção"}
                if any(b in cand.lower() for b in bad):
                    continue
                mm = re.search(r"([A-Za-zÀ-ÿ .'\-]{2,})\s*[-/]\s*([A-Z]{2})\b", cand)
                if mm:
                    return f"{_strip(mm.group(1))}/{mm.group(2)}"
    for ln in lines:
        if "endere" in ln.lower():
            mm = re.search(r"([A-Za-zÀ-ÿ .'\-]{2,})\s*[-/]\s*([A-Z]{2})\b", ln)
            if mm:
                return f"{_strip(mm.group(1))}/{mm.group(2)}"
    loc = _value_after_label(pdf_text, "Localização") or _value_after_label(pdf_text, "Localizacao")
    if loc:
        mm = re.search(r"([A-Za-zÀ-ÿ .'\-]{2,})\s*[-/]\s*([A-Z]{2})\b", loc)
        if mm:
            return f"{_strip(mm.group(1))}/{mm.group(2)}"
        return loc
    return None


def _extract_total_anot(pdf_text: str) -> Optional[str]:
    v = _value_after_label(pdf_text, "Total em anotações negativas")
    if v:
        m = MONEY_RE.search(v)
        if m:
            return _ptbr_money_with_rs(m.group(0))
    v = _value_after_label(pdf_text, "Total em anotacoes negativas")
    if v:
        m = MONEY_RE.search(v)
        if m:
            return _ptbr_money_with_rs(m.group(0))
    m2 = re.search(r"Total\s+de\s+d[ií]vidas\s*[:\-]?\s*(R\$\s*[\d\.,]+)", pdf_text, flags=re.IGNORECASE)
    if m2:
        return _ptbr_money_with_rs(m2.group(1))
    return None


def _extract_liminar(pdf_text: str) -> str:
    if not pdf_text or len(_strip(pdf_text)) < 40:
        return "a preencher"
    up = _no_accents_upper(pdf_text)
    found = any(_no_accents_upper(k) in up for k in LIMINAR_KEYS)
    return "Sim" if found else "Sem registro"


def extract_identificacao(
    pdf_text: str,
    cnpj_row: str,
    cpf_row: str,
    pdf_path: Optional[Path] = None,
    poppler_path: Optional[Path] = None,
) -> Dict[str, str]:
    out: Dict[str, str] = {}

    nome = _extract_nome(pdf_text)
    out["Nome"] = nome if nome else "a preencher"

    out["CNPJ"] = cnpj_row if cnpj_row else "a preencher"
    out["CPF"] = cpf_row if cpf_row else "a preencher"

    score = _extract_serasa_score(pdf_text)
    out["Serasa Score"] = score if score else "a preencher"

    cnae = _extract_cnae(pdf_text)
    out["CNAE"] = cnae if cnae else "a preencher"

    loc = _extract_localizacao(pdf_text)
    out["Localização"] = loc if loc else "a preencher"

    out["Liminar"] = _extract_liminar(pdf_text)

    total = _extract_total_anot(pdf_text)
    out["Total em anotações negativas"] = total if total else "a preencher"

    needs_ocr = [
        out["Nome"] in {"a preencher", "-", ""} or out["Nome"].strip().endswith(":-"),
        out["Serasa Score"] == "a preencher",
        out["Localização"] == "a preencher",
    ]
    if any(needs_ocr) and pdf_path and poppler_path:
        ocr_text = _pdf_page_text_ocr(pdf_path, page_index0=0, poppler_path=poppler_path, dpi=260)
        if ocr_text:
            if out["Nome"] in {"a preencher", "-", ""} or out["Nome"].strip().endswith(":-"):
                nome_ocr = _extract_nome(ocr_text)
                if nome_ocr:
                    out["Nome"] = nome_ocr
            if out["Serasa Score"] == "a preencher":
                score_ocr = _extract_serasa_score(ocr_text)
                if score_ocr:
                    out["Serasa Score"] = score_ocr
            if out["Localização"] == "a preencher":
                loc_ocr = _extract_localizacao(ocr_text)
                if loc_ocr:
                    out["Localização"] = loc_ocr
        if HAS_CV2 and any([
            out["Nome"] in {"a preencher", "-", ""} or out["Nome"].strip().endswith(":-"),
            out["Serasa Score"] == "a preencher",
        ]):
            pages_bgr = _render_pages_pdf(pdf_path, poppler_path=poppler_path, dpi=260)
            if pages_bgr:
                page_bgr = pages_bgr[0]
                if out["Nome"] in {"a preencher", "-", ""} or out["Nome"].strip().endswith(":-"):
                    nome_img = _ocr_nome_from_image(page_bgr)
                    if nome_img:
                        out["Nome"] = nome_img
                if out["Serasa Score"] == "a preencher":
                    score_img = _ocr_score_from_image(page_bgr)
                    if score_img:
                        out["Serasa Score"] = score_img

    return out


# =========================
# Anotacoes Negativas — DETALHES (somente PDF)
# =========================

def extract_block(text: str, start: str, stop_any: List[str], max_chars: int = 25000) -> str:
    t = text or ""
    s_up = _no_accents_upper(start)
    idx = _no_accents_upper(t).find(s_up)
    if idx < 0:
        return ""
    cut = t[idx: idx + max_chars]
    cut_up = _no_accents_upper(cut)
    stops = []
    for st in stop_any:
        j = cut_up.find(_no_accents_upper(st))
        if j > 0:
            stops.append(j)
    if stops:
        cut = cut[: min(stops)]
    return cut


def _has_sem_registros(block: str) -> bool:
    return bool(re.search(r"\bSem\s+registros\b", block or "", flags=re.IGNORECASE))


def _parse_registros_4campos(block: str) -> List[Dict[str, str]]:
    """
    Parser conservador: so cria registro se tiver Data ou Valor explicitos.
    """
    lines = [l for l in (block or "").splitlines() if _strip(l)]
    cleaned = []
    for l in lines:
        up = _no_accents_upper(l)
        if "EXIBINDO" in up and "REGISTRO" in up:
            continue
        cleaned.append(_strip(l))

    records: List[Dict[str, str]] = []
    buf: List[str] = []

    def flush_buf():
        nonlocal buf
        if not buf:
            return
        s = _strip(" ".join(buf))
        if not (DATE_RE.search(s) or MONEY_RE.search(s)):
            buf = []
            return

        rec: Dict[str, str] = {}
        mdt = DATE_RE.search(s)
        rec["Data"] = mdt.group(0) if mdt else "a preencher"

        mm = MONEY_RE.search(s)
        rec["Valor"] = _ptbr_money_with_rs(mm.group(0)) if mm else "a preencher"
        if rec["Valor"] is None:
            rec["Valor"] = "a preencher"

        tmp = s
        if mm:
            tmp = tmp.replace(mm.group(0), " ")
        if mdt:
            tmp = tmp.replace(mdt.group(0), " ")
        tmp = _strip(tmp)
        rec["Modalidade"] = tmp[:80] if tmp else "a preencher"
        rec["Origem"] = "a preencher"

        records.append(rec)
        buf = []

    for l in cleaned:
        if DATE_RE.search(l) or MONEY_RE.search(l):
            buf.append(l)
            flush_buf()
        else:
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
        recs.append({
            "Data": mdt.group(0),
            "Valor": _ptbr_money_with_rs(mm.group(0)) or "a preencher",
        })
    return recs


def _parse_acoes(block: str) -> List[Dict[str, str]]:
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
        recs.append({
            "Data": mdt.group(0),
            "Valor": _ptbr_money_with_rs(mm.group(0)) or "a preencher",
            "Natureza": "a preencher",
            "Cidade": "a preencher",
        })
    return recs


def _parse_cheques_motivos(block: str) -> List[str]:
    lines = [l for l in (block or "").splitlines() if _strip(l)]
    motivos: List[str] = []
    for l in lines:
        up = _no_accents_upper(l)
        if "SEM REGISTROS" in up:
            return []
        m = re.search(r"\b(SUSTADO|SEM FUNDOS|DEVOLVIDO|IRREGULAR)\b", up)
        if m:
            motivos.append(m.group(1).title())
    seen = set()
    out = []
    for m in motivos:
        if m not in seen:
            seen.add(m)
            out.append(m)
    return out


def extract_anotacoes_detalhes(pdf_text: str) -> List[Tuple[str, str]]:
    out: List[Tuple[str, str]] = []

    anot = extract_block(pdf_text, "Anotações negativas", stop_any=[
        "Consultas", "Consulta", "Participações", "Dados",
    ])
    if not anot:
        return out

    def subsec(title: str, stops: List[str]) -> str:
        return extract_block(anot, title, stop_any=stops, max_chars=20000)

    pefin = subsec("PEFIN", ["REFIN", "Protestos", "Ações judiciais", "Cheques"])
    if pefin and not _has_sem_registros(pefin):
        recs = _parse_registros_4campos(pefin)
        for i, r in enumerate(recs, start=1):
            out.append((f"PEFIN - Registro {i} (Modalidade)", r.get("Modalidade", "a preencher")))
            out.append((f"PEFIN - Registro {i} (Valor)", r.get("Valor", "a preencher")))
            out.append((f"PEFIN - Registro {i} (Origem)", r.get("Origem", "a preencher")))
            out.append((f"PEFIN - Registro {i} (Data)", r.get("Data", "a preencher")))

    refin = subsec("REFIN", ["Protestos", "Ações judiciais", "Cheques"])
    if refin and not _has_sem_registros(refin):
        recs = _parse_registros_4campos(refin)
        for i, r in enumerate(recs, start=1):
            out.append((f"REFIN - Registro {i} (Modalidade)", r.get("Modalidade", "a preencher")))
            out.append((f"REFIN - Registro {i} (Valor)", r.get("Valor", "a preencher")))
            out.append((f"REFIN - Registro {i} (Origem)", r.get("Origem", "a preencher")))
            out.append((f"REFIN - Registro {i} (Data)", r.get("Data", "a preencher")))

    prot = subsec("Protestos", ["Ações judiciais", "Cheques"])
    if prot and not _has_sem_registros(prot):
        recs = _parse_protestos(prot)
        for i, r in enumerate(recs, start=1):
            out.append((f"Protestos - Registro {i} (Data)", r.get("Data", "a preencher")))
            out.append((f"Protestos - Registro {i} (Valor)", r.get("Valor", "a preencher")))

    acoes = subsec("Ações judiciais", ["Cheques"])
    if acoes and not _has_sem_registros(acoes):
        recs = _parse_acoes(acoes)
        for i, r in enumerate(recs, start=1):
            out.append((f"Ações Judiciais - Registro {i} (Data)", r.get("Data", "a preencher")))
            out.append((f"Ações Judiciais - Registro {i} (Valor)", r.get("Valor", "a preencher")))
            out.append((f"Ações Judiciais - Registro {i} (Natureza)", r.get("Natureza", "a preencher")))
            out.append((f"Ações Judiciais - Registro {i} (Cidade)", r.get("Cidade", "a preencher")))

    cheq = subsec("Cheques", [])
    if cheq and not _has_sem_registros(cheq):
        motivos = _parse_cheques_motivos(cheq)
        if motivos:
            out.append(("Cheques - Motivo", " | ".join(motivos)))

    return out


# =========================
# IMAGEM — Consultas (grafico) + 5 ultimos (tabela)
# =========================

@dataclass
class Bar:
    x: int
    y: int
    w: int
    h: int


def _render_pages_pdf(pdf_path: Path, poppler_path: Optional[Path], dpi: int = 220) -> List[np.ndarray]:
    if poppler_path is None:
        return []
    pages = convert_from_path(
        str(pdf_path),
        dpi=dpi,
        poppler_path=str(poppler_path),
        fmt="png",
    )
    return [cv2.cvtColor(np.array(p), cv2.COLOR_RGB2BGR) for p in pages]


def _pick_page_by_blue_score(pages_bgr: List[np.ndarray]) -> Optional[np.ndarray]:
    if not pages_bgr:
        return None
    scores = []
    for i, img_bgr in enumerate(pages_bgr):
        mask = _blue_mask_hsv(img_bgr)
        score = int(mask.sum())
        scores.append((score, i))
    scores.sort(key=lambda t: t[0], reverse=True)
    return pages_bgr[scores[0][1]] if scores else None


def _ocr_nome_from_image(img_bgr: np.ndarray) -> Optional[str]:
    data = pytesseract.image_to_data(img_bgr, output_type=pytesseract.Output.DICT, config="--oem 3 --psm 6")
    n = len(data.get("text", []))
    words = []
    for i in range(n):
        txt = _strip(data["text"][i])
        if not txt:
            continue
        x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
        words.append((txt.lower(), x, y, w, h))

    if not words:
        return None

    h_img, w_img = img_bgr.shape[:2]
    for i, (txt, x, y, w, h) in enumerate(words):
        if txt != "nome":
            continue
        for j in range(i + 1, min(i + 6, len(words))):
            txt2, x2, y2, w2, h2 = words[j]
            if abs(y2 - y) > 12:
                continue
            if "fantasia" not in txt2:
                continue
            x1 = max(x + w, x2 + w2) + 5
            y1 = max(y - 3, 0)
            y2 = min(max(y + h, y2 + h2) + 8, h_img)
            roi = img_bgr[y1:y2, x1:w_img]
            if roi.size == 0:
                continue
            txt_nome = pytesseract.image_to_string(roi, config="--oem 3 --psm 7") or ""
            txt_nome = _clean_nome(txt_nome)
            txt_nome = _strip(txt_nome.replace("\n", " "))
            if txt_nome and txt_nome != "-" and re.search(r"[A-Za-zÀ-ÿ]", txt_nome) and len(txt_nome) > 2:
                return txt_nome
    return None


def _ocr_score_from_image(img_bgr: np.ndarray) -> Optional[str]:
    data = pytesseract.image_to_data(img_bgr, output_type=pytesseract.Output.DICT, config="--oem 3 --psm 6")
    n = len(data.get("text", []))
    words = []
    nums = []
    for i in range(n):
        txt = _strip(data["text"][i])
        if not txt:
            continue
        x, y, w, h = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
        t = txt.lower()
        if t == "score":
            words.append((x, y, w, h))
        if re.fullmatch(r"\d{3,4}", txt):
            v = int(txt)
            if v in {0, 500, 1000}:
                continue
            if 0 <= v <= 1000:
                nums.append((v, x, y, w, h))

    if not nums:
        return None

    if words:
        sx, sy, _, _ = words[0]
        nums.sort(key=lambda t: (abs(t[2] - sy), -t[3]))
        return str(nums[0][0])

    nums.sort(key=lambda t: -t[3])
    return str(nums[0][0])


def _blue_mask_hsv(img_bgr: np.ndarray) -> np.ndarray:
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
    lower = np.array([90, 40, 40], dtype=np.uint8)
    upper = np.array([150, 255, 255], dtype=np.uint8)
    mask = cv2.inRange(hsv, lower, upper)
    k1 = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    k2 = cv2.getStructuringElement(cv2.MORPH_RECT, (9, 9))
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, k1, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, k2, iterations=2)
    return mask


def _remove_blue(img_bgr: np.ndarray) -> np.ndarray:
    mask = _blue_mask_hsv(img_bgr)
    out = img_bgr.copy()
    out[mask > 0] = (255, 255, 255)
    return out


def _detect_bars_in_image(img_bgr: np.ndarray) -> List[Bar]:
    mask = _blue_mask_hsv(img_bgr)
    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    bars: List[Bar] = []
    H, W = img_bgr.shape[:2]

    for c in cnts:
        x, y, w, h = cv2.boundingRect(c)
        if w < max(6, W // 300):
            continue
        if h < max(12, H // 300):
            continue
        if h > int(H * 0.9):
            continue
        bars.append(Bar(x, y, w, h))

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


def _crop_chart_roi(img_bgr: np.ndarray, bars: List[Bar]) -> Tuple[np.ndarray, Tuple[int, int, int, int]]:
    H, W = img_bgr.shape[:2]
    xs = [b.x for b in bars]
    xe = [b.x + b.w for b in bars]
    ys = [b.y for b in bars]
    ye = [b.y + b.h for b in bars]

    x1 = _clamp(min(xs) - 160, 0, W - 1)
    x2 = _clamp(max(xe) + 160, 0, W)
    y1 = _clamp(min(ys) - 260, 0, H - 1)
    y2 = _clamp(max(ye) + 260, 0, H)

    roi = img_bgr[y1:y2, x1:x2].copy()
    return roi, (x1, y1, x2 - x1, y2 - y1)


def _parse_month_label(s: str) -> Optional[datetime]:
    s2 = _normalize_month_text(s).lower()
    m = re.search(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez).*(20\d{2})", s2)
    if not m:
        return None
    mon = PT_MONTHS_NUM[m.group(1)]
    year = int(m.group(2))
    return datetime(year, mon, 1)


def _ocr_month_under_bar(chart_nb_bgr: np.ndarray, bar: Bar) -> Optional[str]:
    H, W = chart_nb_bgr.shape[:2]
    x1 = _clamp(bar.x - 40, 0, W - 1)
    x2 = _clamp(bar.x + bar.w + 40, 0, W)
    y1 = _clamp(bar.y + bar.h + 25, 0, H - 1)
    y2 = _clamp(bar.y + bar.h + 140, 0, H)

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
    if len(txt) < 5:
        return None
    return txt


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


def _extract_consultas_tabela_pdf(pdf_path: Path) -> List[Tuple[str, str]]:
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
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for table in tables:
                resultados.extend(_rows_from_table(table))

    out: List[Tuple[str, str]] = []
    for i, (data, nome) in enumerate(resultados[:5], start=1):
        out.append((f"Consulta {i}", f"{data} - {nome}"))
    return out


def _extract_consultas_tabela_text(pdf_text: str) -> List[Tuple[str, str]]:
    bloco = extract_block(
        pdf_text or "",
        "Consultas",
        stop_any=[
            "Participações",
            "Anotações negativas",
            "Consultas por mês",
            "Consultas por mes",
            "Dados",
            "Detalhes",
        ],
    )
    if not bloco:
        return []

    lines = [_strip(l) for l in (bloco.splitlines() or [])]
    candidates: List[Tuple[datetime, str]] = []
    for i, line in enumerate(lines):
        if not line:
            continue
        dt = _parse_date_br(line)
        if not dt:
            continue
        rest = DATE_RE.sub(" ", line, count=1)
        rest = _strip(rest)
        if not rest:
            j = i + 1
            while j < len(lines) and not _strip(lines[j]):
                j += 1
            if j < len(lines):
                rest = _strip(lines[j])
        if rest:
            candidates.append((dt, rest))

    if not candidates:
        return []

    candidates.sort(key=lambda t: t[0], reverse=True)
    top5 = candidates[:5]
    out: List[Tuple[str, str]] = []
    for i, (dt, consultante) in enumerate(top5, start=1):
        campo = f"Consulta {i}"
        info = f"{dt.strftime('%d/%m/%Y')} - {consultante}"
        out.append((campo, info))
    return out


def _extract_consultas_grafico_pdf(
    pdf_path: Path,
    poppler_path: Optional[Path],
    dpi: int = DPI_GRAFICO,
    start_label: str = START_LABEL,
    expected_bars: int = EXPECTED_BARS,
) -> List[Tuple[str, str]]:
    if not HAS_CV2 or poppler_path is None:
        return []

    pages_bgr = _render_pages_pdf(pdf_path, poppler_path=poppler_path, dpi=dpi)
    page_bgr = _pick_page_by_blue_score(pages_bgr)
    if page_bgr is None:
        return []

    bars_page = _detect_bars_in_image(page_bgr)
    if len(bars_page) < 5:
        return []

    roi_bgr, _ = _crop_chart_roi(page_bgr, bars_page)
    bars_roi = _detect_bars_in_image(roi_bgr)

    base_ys = np.array([b.y + b.h for b in bars_roi], dtype=np.int32)
    if len(base_ys) > 0:
        base_ref = int(np.median(base_ys))
        bars_roi = [b for b in bars_roi if abs((b.y + b.h) - base_ref) < int(roi_bgr.shape[0] * 0.25)]

    bars_roi.sort(key=lambda b: b.x)
    if len(bars_roi) > expected_bars:
        bars_roi = sorted(bars_roi, key=lambda b: b.h, reverse=True)[:expected_bars]
        bars_roi.sort(key=lambda b: b.x)

    chart_nb = _remove_blue(roi_bgr)
    results = []
    ocr_months: List[Optional[datetime]] = []

    for b in bars_roi:
        H, W = chart_nb.shape[:2]
        x1 = _clamp(b.x - 55, 0, W - 1)
        x2 = _clamp(b.x + b.w + 55, 0, W)

        up = max(140, int(H * 0.22))
        y1 = _clamp(b.y - up, 0, H - 1)
        y2 = _clamp(b.y - 5, 0, H)

        num_roi = chart_nb[y1:y2, x1:x2].copy()
        val = _best_digit_ocr(num_roi)

        mtxt = _ocr_month_under_bar(chart_nb, b)
        mdt = _parse_month_label(mtxt) if mtxt else None
        ocr_months.append(mdt)

        results.append(val)

    dt_fallback = _parse_start_label(start_label)
    dt0 = _infer_start_month(ocr_months, fallback_start=dt_fallback)

    out: List[Tuple[str, str]] = []
    months = [_add_months(dt0, i) for i in range(len(results))]
    for i, dt in enumerate(months):
        campo = f"Consultas - {dt.strftime('%m/%Y')}"
        val = results[i]
        out.append((campo, "" if val is None else str(val)))

    return out


def _preprocess_for_digits(roi_bgr: np.ndarray, mode: str) -> np.ndarray:
    gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.resize(gray, None, fx=4, fy=4, interpolation=cv2.INTER_CUBIC)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)

    if mode == "otsu":
        _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    elif mode == "adapt":
        th = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY, 41, 15
        )
    elif mode == "adapt_inv":
        th = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY_INV, 41, 15
        )
        th = 255 - th
    else:
        _, th = cv2.threshold(gray, 170, 255, cv2.THRESH_BINARY)

    return th


def _tighten_binary(th_img: np.ndarray) -> np.ndarray:
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


def _largest_component(th_img: np.ndarray) -> np.ndarray:
    inv = (th_img < 128).astype(np.uint8)
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(inv, connectivity=8)
    if num_labels <= 1:
        return th_img
    idx = 1 + int(np.argmax(stats[1:, cv2.CC_STAT_AREA]))
    mask = (labels == idx).astype(np.uint8)
    out = np.full(th_img.shape, 255, dtype=np.uint8)
    out[mask > 0] = 0
    return out


def _ocr_digits(th_img: np.ndarray, psm: int) -> str:
    cfg = f"--oem 3 --psm {psm} -c tessedit_char_whitelist=0123456789"
    txt = pytesseract.image_to_string(th_img, config=cfg) or ""
    txt = re.sub(r"\D", "", txt)
    return txt


def _best_digit_ocr(roi_bgr: np.ndarray) -> Optional[int]:
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
        th = _preprocess_for_digits(roi_bgr, mode=mode)
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
        gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)
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


def _ocr_month(roi_bgr: np.ndarray) -> Optional[str]:
    gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                               cv2.THRESH_BINARY, 41, 15)
    cfg = "--oem 3 --psm 7"
    txt = pytesseract.image_to_string(th, config=cfg) or ""
    txt = _normalize_month_text(txt).lower()
    m = re.search(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez).*(20\d{2})", txt)
    if not m:
        return None
    mm = PT_MONTHS.get(m.group(1))
    yyyy = m.group(2)
    if not mm:
        return None
    return f"{mm}/{yyyy}"


def _extract_consultas_grafico_bgr(img_bgr: np.ndarray) -> List[Tuple[str, str]]:
    """
    Retorna [(Campo, Informacao)] para "Consultas - MM/AAAA" (somente se valor > 0)
    Fonte exclusiva IMAGEM.
    """
    if not HAS_CV2:
        return []

    H, W = img_bgr.shape[:2]

    mask = _blue_mask_hsv(img_bgr)

    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    bars: List[Bar] = []
    for c in cnts:
        x, y, w, h = cv2.boundingRect(c)
        if w < max(6, W // 300):
            continue
        if h < max(12, H // 300):
            continue
        if h > int(H * 0.9):
            continue
        bars.append(Bar(x, y, w, h))

    if len(bars) < 3:
        return []

    bars.sort(key=lambda b: b.x)

    def clamp(v: int, lo: int, hi: int) -> int:
        return max(lo, min(hi, v))

    results: List[Tuple[str, str]] = []
    seen = set()

    for b in bars:
        x1 = clamp(b.x - 50, 0, W - 1)
        x2 = clamp(b.x + b.w + 50, 0, W)
        y1 = clamp(b.y - int(H * 0.18), 0, H - 1)
        y2 = clamp(b.y - 2, 0, H)
        roi_num = img_bgr[y1:y2, x1:x2]
        val = _best_digit_ocr(roi_num)

        y3 = clamp(b.y + b.h + 8, 0, H - 1)
        y4 = clamp(b.y + b.h + int(H * 0.12), 0, H)
        roi_mon = img_bgr[y3:y4, x1:x2]
        mm_yyyy = _ocr_month(roi_mon)

        if mm_yyyy is None or val is None:
            continue
        if val <= 0:
            continue

        campo = f"Consultas - {mm_yyyy}"
        if campo in seen:
            continue
        seen.add(campo)
        results.append((campo, str(val)))

    def key_month(c: str):
        m = re.search(r"(\d{2})/(\d{4})", c)
        if not m:
            return (9999, 99)
        return (int(m.group(2)), int(m.group(1)))

    results.sort(key=lambda x: key_month(x[0]))
    return results


def _extract_consultas_grafico(img_path: Path) -> List[Tuple[str, str]]:
    if not HAS_CV2:
        return []
    img_bgr = cv2.imread(str(img_path))
    if img_bgr is None:
        return []
    return _extract_consultas_grafico_bgr(img_bgr)


def _extract_consultas_tabela_ultimas_bgr(img_bgr: np.ndarray) -> List[Tuple[str, str]]:
    """
    Retorna ate 5 linhas: ("Consulta i", "dd/mm/aaaa - CONSULTANTE")
    Fonte exclusiva IMAGEM.
    Estrategia: OCR por palavras (pytesseract image_to_data) e reconstrucao de linhas.
    """
    if not HAS_CV2:
        return []

    def _run_ocr_table(img: np.ndarray) -> List[Tuple[str, str]]:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY, 41, 15)

        data = pytesseract.image_to_data(th, output_type=pytesseract.Output.DICT, config="--oem 3 --psm 6")
        n = len(data.get("text", []))

        words = []
        for i in range(n):
            txt = _strip(data["text"][i])
            if not txt:
                continue
            conf = int(float(data.get("conf", ["-1"])[i]))
            if conf < 35:
                continue
            x, y = data["left"][i], data["top"][i]
            words.append((y, x, txt))

        if not words:
            return []

        words.sort(key=lambda t: (t[0], t[1]))
        lines: List[List[Tuple[int, int, str]]] = []
        cur: List[Tuple[int, int, str]] = []
        cur_y: Optional[float] = None
        y_tol = 9

        for y, x, txt in words:
            if cur_y is None or abs(y - cur_y) <= y_tol:
                cur.append((y, x, txt))
                cur_y = y if cur_y is None else (cur_y + y) / 2.0
            else:
                cur.sort(key=lambda t: t[1])
                lines.append(cur)
                cur = [(y, x, txt)]
                cur_y = y
        if cur:
            cur.sort(key=lambda t: t[1])
            lines.append(cur)

        header_idx = None
        for i, ln in enumerate(lines):
            text_line = " ".join(t[2] for t in ln)
            up = _no_accents_upper(text_line)
            if ("DATA" in up and "CONSULT" in up) or ("CONSULTANTE" in up and "DATA" in up):
                header_idx = i
                break

        start = header_idx + 1 if header_idx is not None else 0

        candidates: List[Tuple[datetime, str]] = []
        for ln in lines[start:]:
            text_line = " ".join(t[2] for t in ln)
            dt = _parse_date_br(text_line)
            if not dt:
                continue
            rest = DATE_RE.sub(" ", text_line, count=1)
            rest = _strip(rest)
            if len(rest) < 2:
                continue
            candidates.append((dt, rest))

        if not candidates:
            return []

        candidates.sort(key=lambda t: t[0], reverse=True)
        top5 = candidates[:5]

        out: List[Tuple[str, str]] = []
        for i, (dt, consultante) in enumerate(top5, start=1):
            campo = f"Consulta {i}"
            info = f"{dt.strftime('%d/%m/%Y')} - {consultante}"
            out.append((campo, info))
        return out

    out = _run_ocr_table(img_bgr)
    if out:
        return out

    h = img_bgr.shape[0]
    crop = img_bgr[int(h * 0.45):, :]
    return _run_ocr_table(crop)


def _extract_consultas_tabela_ultimas(img_path: Path) -> List[Tuple[str, str]]:
    if not HAS_CV2:
        return []
    img_bgr = cv2.imread(str(img_path))
    if img_bgr is None:
        return []
    return _extract_consultas_tabela_ultimas_bgr(img_bgr)


# =========================
# Build rows + validacoes
# =========================

def build_rows(
    pdf_text: str,
    pdf_path: Optional[Path],
    img_path: Optional[Path],
    poppler_path: Optional[Path],
) -> Tuple[List[Tuple[str, str]], str, str]:
    """
    Retorna:
      - rows: [(Campo, Informacao)] em ordem final
      - id_3a_coluna (valor replicado na coluna CNPJ)
      - nome_base (para nome do arquivo)
    """
    if pdf_text and pdf_path and pdf_path.exists():
        cnpj_row, cpf_row, chosen_id = extract_ids(pdf_text)
        ident = extract_identificacao(
            pdf_text,
            cnpj_row=cnpj_row,
            cpf_row=cpf_row,
            pdf_path=pdf_path,
            poppler_path=poppler_path,
        )
        nome_base = ident.get("Nome", "SACADO")
        anot = extract_anotacoes_detalhes(pdf_text)
    else:
        cnpj_row, cpf_row, chosen_id = ("-", "-", "A confirmar")
        ident = {k: "a preencher" for k in IDENT_FIELDS_ORDER}
        ident["CNPJ"] = cnpj_row
        ident["CPF"] = cpf_row
        nome_base = "SACADO"
        anot = []

    rows: List[Tuple[str, str]] = []

    for k in IDENT_FIELDS_ORDER:
        rows.append((k, ident.get(k, "a preencher")))

    rows.extend(anot)

    consultas_mensais: List[Tuple[str, str]] = []
    consultas_ultimas: List[Tuple[str, str]] = []

    if img_path and img_path.exists():
        consultas_mensais = _extract_consultas_grafico(img_path)
        consultas_ultimas = _extract_consultas_tabela_ultimas(img_path)
    elif pdf_path and pdf_path.exists():
        consultas_mensais = _extract_consultas_grafico_pdf(pdf_path, poppler_path=poppler_path)
        consultas_ultimas = _extract_consultas_tabela_pdf(pdf_path)

    if len(consultas_ultimas) < 5 and pdf_text:
        consultas_text = _extract_consultas_tabela_text(pdf_text)
        if len(consultas_text) > len(consultas_ultimas):
            consultas_ultimas = consultas_text

    if len(consultas_ultimas) < 5:
        raise RuntimeError("ERRO: Consultas insuficientes no SERASA SACADO (esperado 5).")

    rows.extend(consultas_mensais)
    rows.extend(consultas_ultimas)

    for campo, _ in rows:
        if not _campo_valido(campo):
            _fail()

    return rows, chosen_id, str(nome_base or "SACADO")


def write_xlsx(out_xlsx: Path, rows: List[Tuple[str, str]], id_value: str) -> None:
    rows = [(c, i) for (c, i) in rows]
    df = pd.DataFrame(rows, columns=["Campo", "Informação"])
    df["CNPJ"] = id_value
    df = df[["Campo", "Informação", "CNPJ"]]

    if list(df.columns) != ["Campo", "Informação", "CNPJ"]:
        _fail()

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Planilha1", index=False)
        wb = writer.book
        ws = wb["Planilha1"]
        for row in ws.iter_rows():
            for cell in row:
                cell.number_format = "@"


# =========================
# CLI
# =========================

def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", default="", help="PDF SERASA (sacado). Se vazio, processa lote em 01_INPUT.")
    ap.add_argument("--img", default="", help="IMAGEM (png/jpg) com grafico + tabela de consultas (modo arquivo unico).")
    ap.add_argument("--input", default=str(PASTA_INPUT), help="Pasta base para varrer subpastas (modo lote).")
    ap.add_argument("--outdir", default=str(PASTA_DESTINO), help=r"Pasta de saida (ex.: ...\03_OUTPUT\10. SERASA SACADO)")
    ap.add_argument("--poppler", default="", help="Pasta do Poppler bin (para OCR pontual em paginas vazias do PDF).")
    ap.add_argument("--tesseract", default="", help="Caminho do tesseract.exe (opcional).")
    args = ap.parse_args()

    pdf_path = Path(args.pdf) if args.pdf else None
    img_path = Path(args.img) if args.img else None
    outdir = Path(args.outdir)

    poppler_path = Path(args.poppler) if args.poppler else None
    if args.tesseract:
        pytesseract.pytesseract.tesseract_cmd = args.tesseract
    if not img_path and poppler_path is None:
        print("Aviso: sem --img e sem --poppler. Grafico nao sera extraido; tabela tenta via PDF.")

    if pdf_path and pdf_path.exists():
        pdf_text = extract_pdf_text(pdf_path, poppler_path=poppler_path, ocr_dpi=260)
        rows, id_value, nome_base = build_rows(
            pdf_text=pdf_text,
            pdf_path=pdf_path,
            img_path=img_path,
            poppler_path=poppler_path,
        )
        safe_nome = _safe_filename(nome_base)
        if safe_nome in {"SACADO", "A_CONFIRMAR"}:
            safe_nome = _safe_filename(pdf_path.stem.replace("serasa_sacado_", ""))
        out_xlsx = outdir / f"SERASA_SACADO_{safe_nome}.xlsx"
        write_xlsx(out_xlsx, rows=rows, id_value=id_value)
        print("Arquivo gerado.")
        print(str(out_xlsx))
        return

    input_dir = Path(args.input)
    pdfs = _list_pdf_sacado(input_dir)
    if not pdfs:
        print("Nenhum PDF SACADO encontrado.")
        return

    outdir.mkdir(parents=True, exist_ok=True)
    for pdf in pdfs:
        img = _choose_image_for_pdf(pdf)
        pdf_text = extract_pdf_text(pdf, poppler_path=poppler_path, ocr_dpi=260)
        rows, id_value, nome_base = build_rows(
            pdf_text=pdf_text,
            pdf_path=pdf,
            img_path=img,
            poppler_path=poppler_path,
        )
        safe_nome = _safe_filename(nome_base)
        if safe_nome in {"SACADO", "A_CONFIRMAR"}:
            safe_nome = _safe_filename(pdf.stem.replace("serasa_sacado_", ""))
        out_xlsx = outdir / f"SERASA_SACADO_{safe_nome}.xlsx"
        write_xlsx(out_xlsx, rows=rows, id_value=id_value)
        print(f"OK: {pdf.name}")


if __name__ == "__main__":
    main()
