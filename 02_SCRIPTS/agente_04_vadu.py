#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
agente_04_vadu.py

STRICT: Read VADU images (protestos/processos_areas/ultimos/cnpj) and
produce one XLSX with fixed tabs.
"""

from __future__ import annotations

import argparse
import re
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import pytesseract
from pytesseract import Output
from PIL import Image, ImageFilter, ImageOps


CNPJ_RE = re.compile(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b")
MONEY_RE = re.compile(r"(R\$\s*)?\d{1,3}(?:[\.,]\d{3})*(?:[\.,]\d{2})")
DATE_RE = re.compile(r"\b\d{2}[\/\.-]\d{2}[\/\.-]\d{2,4}\b")

PROCESSOS_FIELDS = [
    "Valor total em protesto",
    "Valor em Processos (Total R$)",
    "Total Processos (Qtde)",
    "Total Estadual (Qtde)",
    "Total Estadual (R$)",
    "Total Federal (Qtde)",
    "Total Federal (R$)",
    "Total Trabalhista (Qtde)",
    "Total Trabalhista (R$)",
]

PARTES_STATUS_CANON = [
    "ARQUIVADO ADMINISTRATIVAMENTE",
    "ARQUIVAMENTO DEFINITIVO",
    "EM TRAMITACAO",
    "EM GRAU DE RECURSO",
    "SUSPENSO",
    "TOTAL",
]


def _strip(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\u00a0", " ")).strip()


def _no_accents_upper(s: str) -> str:
    s = _strip(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.upper()


def _valid_date(dd: str, mm: str, yyyy: str) -> bool:
    try:
        d = int(dd)
        m = int(mm)
        y = int(yyyy)
    except ValueError:
        return False
    if y < 1990 or y > 2035:
        return False
    if m < 1 or m > 12:
        return False
    if d < 1 or d > 31:
        return False
    return True


def _normalize_date_str(date_str: str) -> Optional[str]:
    m = re.search(r"(\d{2})[\/\.-](\d{2})[\/\.-](\d{2,4})", date_str)
    if m:
        dd, mm, yy = m.groups()
        yyyy = yy if len(yy) == 4 else f"20{yy}"
        if _valid_date(dd, mm, yyyy):
            return f"{dd}/{mm}/{int(yyyy):04d}"
        return None
    digits = re.sub(r"\D", "", date_str or "")
    if len(digits) >= 8:
        dd, mm, yyyy = digits[:2], digits[2:4], digits[4:8]
        if _valid_date(dd, mm, yyyy):
            return f"{dd}/{mm}/{int(yyyy):04d}"
    return None


def _find_date_any(s: str) -> Optional[str]:
    if not s:
        return None
    m = DATE_RE.search(s)
    if m:
        return _normalize_date_str(m.group(0))
    return _normalize_date_str(s)


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
        val = int(s2)
        return f"{val:,}".replace(",", "X").replace(".", ",").replace("X", ".") + ",00"

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


def _max_money_in_text(raw: str) -> Optional[str]:
    vals: List[Tuple[float, str]] = []
    for m in MONEY_RE.finditer((raw or "").replace("RS", "R$")):
        val = _ptbr_money_from_any(m.group(0))
        if not val:
            continue
        try:
            num = float(val.replace(".", "").replace(",", "."))
        except ValueError:
            continue
        vals.append((num, val))
    if not vals:
        return None
    vals.sort(key=lambda t: t[0], reverse=True)
    return vals[0][1]


def _ptbr_int_text(raw: str) -> Optional[str]:
    s = _strip(raw)
    m = re.search(r"\b(\d{1,6})\b", s)
    return m.group(1) if m else None


def _find_first_cnpj(text: str) -> Optional[str]:
    m = CNPJ_RE.search(text or "")
    return m.group(0) if m else None


def _safe_company_key(s: str) -> str:
    s = _no_accents_upper(s)
    s = re.sub(r"[^A-Z0-9]+", "_", s).strip("_")
    return s


def _is_blank(s: Optional[str]) -> bool:
    return not _strip(s or "")


def _anchor_hit(line: str) -> bool:
    keys = ["PROCESSO", "CLASSE", "ASSUNTO", "DATA", "RECEBIDO", "RECEBIDO EM", "VALOR", "TIPO", "STATUS"]
    up = _no_accents_upper(line)
    hits = sum(1 for k in keys if k in up)
    return hits >= 3


# =========================
# Imagem -> texto (OCR)
# =========================


def ocr_image_text(image_path: Path, psms: Optional[List[int]] = None) -> str:
    img = Image.open(image_path)
    psms = psms or [6]
    texts = []
    for psm in psms:
        txt = pytesseract.image_to_string(img, lang="por+eng", config=f"--oem 3 --psm {psm}") or ""
        if txt.strip():
            texts.append(txt)
    return "\n".join(texts)


def ocr_image_texts(image_path: Path, psms: Optional[List[int]] = None) -> List[str]:
    img = Image.open(image_path)
    psms = psms or [6]
    texts = []
    for psm in psms:
        txt = pytesseract.image_to_string(img, lang="por+eng", config=f"--oem 3 --psm {psm}") or ""
        if txt.strip():
            texts.append(txt)
    return texts


def _prep_ocr_variants(img: Image.Image) -> List[Image.Image]:
    variants: List[Image.Image] = []
    base = img.convert("L")
    variants.append(base)
    variants.append(ImageOps.autocontrast(base))
    variants.append(ImageOps.autocontrast(base).filter(ImageFilter.SHARPEN))
    big = base.resize((base.width * 2, base.height * 2), Image.BICUBIC)
    variants.append(big)
    variants.append(ImageOps.autocontrast(big))
    variants.append(ImageOps.autocontrast(big).filter(ImageFilter.SHARPEN))
    big3 = base.resize((base.width * 3, base.height * 3), Image.BICUBIC)
    variants.append(ImageOps.autocontrast(big3))
    variants.append(ImageOps.autocontrast(big3).point(lambda p: 255 if p > 180 else 0))
    big4 = base.resize((base.width * 4, base.height * 4), Image.BICUBIC)
    variants.append(ImageOps.autocontrast(big4))
    variants.append(ImageOps.autocontrast(big4).point(lambda p: 255 if p > 180 else 0))
    return variants


def _ocr_column_lines(
    img: Image.Image,
    x0: int,
    x1: int,
    whitelist: str,
    psms: Optional[List[int]] = None,
) -> List[str]:
    psms = psms or [6]
    x0 = max(0, int(x0))
    x1 = min(img.width, int(x1))
    if x1 <= x0:
        return []
    crop = img.crop((x0, 0, x1, img.height))
    out: List[Tuple[int, str]] = []
    seen = set()
    for psm in psms:
        cfg = f"--oem 3 --psm {psm} -c tessedit_char_whitelist={whitelist}"
        data = pytesseract.image_to_data(crop, lang="por+eng", config=cfg, output_type=Output.DICT)
        lines: Dict[Tuple[int, int, int], List[Tuple[int, int, str]]] = {}
        n = len(data.get("text", []))
        for i in range(n):
            text = (data["text"][i] or "").strip()
            if not text:
                continue
            try:
                conf = float(data.get("conf", ["0"])[i])
            except ValueError:
                conf = 0.0
            if conf < 5 and not re.search(r"\d", text):
                continue
            line_id = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
            lines.setdefault(line_id, []).append((int(data["top"][i]), int(data["left"][i]), text))
        for _, items in lines.items():
            items.sort(key=lambda x: x[1])
            top = items[0][0]
            line_text = _strip(" ".join(t for _, _, t in items))
            if not line_text or line_text in seen:
                continue
            seen.add(line_text)
            out.append((top, line_text))
    out.sort(key=lambda x: x[0])
    return [t for _, t in out]


def _extract_dates_from_column(img: Image.Image, x0: int, x1: int) -> List[str]:
    lines = _ocr_column_lines(img, x0, x1, "0123456789/.-", psms=[6, 11, 4])
    dates: List[str] = []
    for line in lines:
        d = _find_date_any(line)
        if d:
            dates.append(d)
    return dates


def _extract_values_from_column(img: Image.Image, x0: int, x1: int) -> List[str]:
    lines = _ocr_column_lines(img, x0, x1, "0123456789R$.,-", psms=[6, 11, 4])
    values: List[str] = []
    for line in lines:
        val = _max_money_in_text(line) or _ptbr_money_from_any(line)
        if val:
            values.append(val)
    return values


def ocr_image_texts_variants(image_path: Path, psms: Optional[List[int]] = None) -> List[str]:
    img = Image.open(image_path)
    psms = psms or [6]
    texts: List[str] = []
    for variant in _prep_ocr_variants(img):
        for psm in psms:
            txt = pytesseract.image_to_string(variant, lang="por+eng", config=f"--oem 3 --psm {psm}") or ""
            if txt.strip():
                texts.append(txt)
    return texts


def _check_tesseract() -> None:
    try:
        _ = pytesseract.get_tesseract_version()
    except Exception as exc:
        raise RuntimeError(
            "ERRO: Tesseract nao encontrado. Instale o Tesseract OCR ou passe --tesseract com o caminho do tesseract.exe."
        ) from exc


# =========================
# Inputs: relatorio + protestos + processos
# =========================


def _classify_image(stem: str) -> Optional[Tuple[str, str, str]]:
    norm = _no_accents_upper(stem)
    tokens = re.split(r"[_\s\-]+", norm)
    role = None
    if "PROTESTOS" in tokens:
        role = "protestos"
    elif "ULTIMOSPROCESSOS" in tokens or ("ULTIMOS" in tokens and "PROCESSOS" in tokens):
        role = "ultimos_processos"
    elif "PROCESSOS" in tokens and "AREAS" in tokens:
        role = "processos_areas"
    elif "CNPJ" in tokens:
        role = "cnpj"
    else:
        return None

    drop = {"PROTESTOS", "PARTES", "PROCESSOS", "AREAS", "ULTIMOSPROCESSOS", "ULTIMOS", "VADU", "CNPJ"}
    company_tokens = [
        t
        for t in tokens
        if t
        and t not in drop
        and not re.fullmatch(r"P\d+", t)
        and not re.fullmatch(r"PG\d+", t)
    ]
    company_raw = "_".join(company_tokens) if company_tokens else stem
    key = _safe_company_key(company_raw)
    return role, key, company_raw


def collect_inputs(input_dir: Path) -> Dict[str, Dict[str, Optional[Path]]]:
    images = [p for p in input_dir.rglob("*.png") if p.is_file()]

    out: Dict[str, Dict[str, Optional[Path]]] = {}

    for p in images:
        classified = _classify_image(p.stem)
        if not classified:
            continue
        role, key, _ = classified
        if key not in out:
            out[key] = {
                "protestos": None,
                "processos_areas": None,
                "ultimos_processos": [],
                "cnpj": None,
            }
        if role == "ultimos_processos":
            out[key][role].append(p)
        else:
            out[key][role] = p

    if not out:
        raise RuntimeError("ERRO: Nenhuma imagem .png encontrada no input (nem em subpastas).")

    return out


# =========================
# Empresa / CNPJ
# =========================


def extract_company_name(text: str, fallback_stem: str) -> str:
    t = text or ""
    patterns = [
        r"Raz[aã]o\s+Social[:\s]+(.+)",
        r"Nome\s+Fantasia[:\s]+(.+)",
        r"Empresa[:\s]+(.+)",
    ]
    for pat in patterns:
        m = re.search(pat, t, flags=re.IGNORECASE)
        if m:
            val = _strip(m.group(1))
            val = re.split(r"\s{2,}|\||\u2022", val)[0]
            if len(val) >= 3:
                return val
    return _strip(fallback_stem).replace(">", "").strip("_- ")


def extract_cnpj(text: str, cnpj_user: Optional[str]) -> str:
    if cnpj_user:
        cnpj_user = _strip(cnpj_user)
        if CNPJ_RE.fullmatch(cnpj_user):
            return cnpj_user
    cnpj = _find_first_cnpj(text)
    return cnpj if cnpj else "A confirmar"


# =========================
# ABA 1: Processos (cards/totais)
# =========================


def _extract_valor_total_protesto(text_protestos: Optional[str]) -> Optional[str]:
    if not text_protestos:
        return None

    lines = [_strip(l) for l in text_protestos.splitlines() if _strip(l)]
    if not lines:
        return None

    for i, line in enumerate(lines):
        if "VALOR TOTAL" in _no_accents_upper(line):
            for j in range(i, min(i + 8, len(lines))):
                if "R$" in lines[j]:
                    m = MONEY_RE.search(lines[j])
                    if m:
                        val = _ptbr_money_from_any(m.group(0))
                        if val:
                            return val

    blob = " ".join(lines)
    m2 = re.search(r"VALOR\s+TOTAL.*?(R\$\s*[\d\.,]+)", blob, flags=re.IGNORECASE)
    if m2:
        return _ptbr_money_from_any(m2.group(1))
    vals = []
    for m3 in MONEY_RE.finditer(blob):
        val = _ptbr_money_from_any(m3.group(0))
        if val:
            num = float(val.replace(".", "").replace(",", "."))
            vals.append((num, val))
    if vals:
        vals.sort(key=lambda x: x[0], reverse=True)
        return vals[0][1]
    return None


def _find_money_after_label(lines: List[str], label_up: str) -> Optional[str]:
    for i, line in enumerate(lines):
        up = _no_accents_upper(line)
        if label_up not in up:
            continue
        for j in range(i, min(i + 8, len(lines))):
            m = MONEY_RE.search(lines[j].replace("RS", "R$").replace("R$", "R$"))
            if m:
                val = _ptbr_money_from_any(m.group(0))
                if val:
                    return val
    return None


def _extract_totais_qtde(lines: List[str]) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str]]:
    for i, line in enumerate(lines):
        up = _no_accents_upper(line)
        if all(k in up for k in ["TOTAL", "TRABALHISTA", "ESTADUAL", "FEDERAL"]):
            if i + 1 < len(lines):
                nums = re.findall(r"\d{1,6}", lines[i + 1])
                if len(nums) >= 4:
                    total, trab, est, fed = nums[0], nums[1], nums[2], nums[3]
                    return total, est, fed, trab
    return None, None, None, None


def _extract_status_juridico_block(
    lines: List[str],
) -> Tuple[Optional[Dict[str, Optional[str]]], Optional[Dict[str, Optional[str]]]]:
    block = lines[:]
    order: List[str] = []
    counts: Dict[str, Optional[str]] = {}
    values: Dict[str, Optional[str]] = {}
    header_idx = None

    for idx, line in enumerate(block):
        up = _no_accents_upper(line)
        if all(k in up for k in ["TOTAL", "TRABALHISTA", "ESTADUAL", "FEDERAL"]):
            positions = []
            for key in ["TOTAL", "TRABALHISTA", "ESTADUAL", "FEDERAL"]:
                pos = up.find(key)
                if pos >= 0:
                    positions.append((pos, key))
            order = [k for _, k in sorted(positions)]
            header_idx = idx
            break

    if not order:
        return None, None

    if header_idx is not None:
        for line in block[header_idx + 1:header_idx + 5]:
            tokens = re.findall(r"\d{1,6}|[-–—]", line)
            if len(tokens) >= len(order):
                for idx, key in enumerate(order):
                    tok = tokens[idx]
                    counts[key] = tok if tok.isdigit() else None
                break

    best_vals: List[str] = []
    for idx, line in enumerate(block):
        up = _no_accents_upper(line)
        line_norm = line.replace("RS", "R$")
        vals = [_ptbr_money_from_any(m.group(0)) for m in MONEY_RE.finditer(line_norm)]
        vals = [v for v in vals if v]
        if not vals and ("VALOR" in up or "R$" in up or "RS" in up) and idx + 1 < len(block):
            next_line = block[idx + 1].replace("RS", "R$")
            vals = [_ptbr_money_from_any(m.group(0)) for m in MONEY_RE.finditer(next_line)]
            vals = [v for v in vals if v]
        if len(vals) > len(best_vals):
            best_vals = vals
        if len(best_vals) >= len(order):
            break

    if len(best_vals) >= len(order):
        for idx, key in enumerate(order):
            values[key] = best_vals[idx]

    if not counts:
        counts = None
    if not values:
        values = None
    return counts, values


def extract_cards_processos(
    text_processos_areas: str,
    text_protestos: Optional[str],
) -> Dict[str, str]:
    lines = [l for l in (text_processos_areas or "").splitlines() if _strip(l)]

    out: Dict[str, str] = {}

    val_prot = _extract_valor_total_protesto(text_protestos)
    out["Valor total em protesto"] = val_prot if val_prot is not None else "-"

    counts, values = _extract_status_juridico_block(lines)
    total_q = counts.get("TOTAL") if counts else None
    trab_q = counts.get("TRABALHISTA") if counts else None
    est_q = counts.get("ESTADUAL") if counts else None
    fed_q = counts.get("FEDERAL") if counts else None

    total_v = values.get("TOTAL") if values else None
    trab_v = values.get("TRABALHISTA") if values else None
    est_v = values.get("ESTADUAL") if values else None
    fed_v = values.get("FEDERAL") if values else None

    if counts is None:
        total_q, est_q, fed_q, trab_q = _extract_totais_qtde(lines)

    out["Total Processos (Qtde)"] = total_q if total_q is not None else "-"
    out["Total Estadual (Qtde)"] = est_q if est_q is not None else "-"
    out["Total Federal (Qtde)"] = fed_q if fed_q is not None else "-"
    out["Total Trabalhista (Qtde)"] = trab_q if trab_q is not None else "-"

    out["Valor em Processos (Total R$)"] = (
        total_v
        or _find_money_after_label(lines, "VALOR EM PROCESSOS")
        or _find_money_after_label(lines, "VALOR EM PROCESSO")
        or "-"
    )
    out["Total Estadual (R$)"] = est_v or _find_money_after_label(lines, "TOTAL ESTADUAL") or "-"
    out["Total Federal (R$)"] = fed_v or _find_money_after_label(lines, "TOTAL FEDERAL") or "-"
    out["Total Trabalhista (R$)"] = trab_v or _find_money_after_label(lines, "TOTAL TRABALHISTA") or "-"

    return out


# =========================
# ABA 2: Partes (Passiva / Ativa)
# =========================


def extract_partes_from_text(text_partes: str) -> pd.DataFrame:
    return pd.DataFrame(columns=["Parte", "Status", "Valor (R$)"])


# =========================
# ABA 3: Areas (Processos por area)
# =========================


def extract_areas_from_text(text_relatorio: str) -> pd.DataFrame:
    lines = [l for l in (text_relatorio or "").splitlines()]
    idx0 = None
    for i, l in enumerate(lines):
        if "PROCESSOS POR AREA" in _no_accents_upper(l):
            idx0 = i
            break
    if idx0 is None:
        return pd.DataFrame(columns=["\u00c1rea", "Qtde"])

    block = lines[idx0:idx0 + 160]
    rows: List[Dict[str, str]] = []
    for l in block:
        s = _strip(l)
        if not s:
            continue
        su = _no_accents_upper(s)
        if "PROCESSOS POR AREA" in su or "ULTIMO PROCESSO" in su:
            continue
        if su in {"AREA / QTD", "AREA/QTD"} or ("AREA" in su and "QTD" in su):
            continue
        if "STATUS JURIDICO" in su:
            break
        if re.search(r"\b1\W{0,2}PROCESSO\b", su):
            continue

        tokens = s.split()
        if not tokens:
            continue
        num_idx = None
        for i, tok in enumerate(tokens):
            if re.fullmatch(r"\d{1,6}", tok):
                num_idx = i
                break
        if num_idx is None or num_idx == 0:
            continue
        area_tokens = tokens[:num_idx]
        while area_tokens and len(area_tokens[0]) == 1 and not area_tokens[0].isdigit():
            area_tokens = area_tokens[1:]
        if not area_tokens:
            continue
        area = _strip(" ".join(area_tokens))
        qtd = tokens[num_idx]
        if len(area) >= 3:
            rows.append({"Área": area, "Qtde": qtd})
    if not rows:
        return pd.DataFrame(columns=["\u00c1rea", "Qtde"])

    seen = set()
    out = []
    for r in rows:
        key = (_no_accents_upper(r["\u00c1rea"]), r["Qtde"])
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return pd.DataFrame(out, columns=["\u00c1rea", "Qtde"])


# =========================
# ABA 4: UltimosProcessos (processos_ PDF)
# =========================


def _clean_assunto(s: str) -> str:
    s = _strip(s)
    s = DATE_RE.sub(" ", s)
    s = re.sub(r"R\$\s*\d{1,3}(?:[\.,]\d{3})*(?:[\.,]\d{2})", " ", s)
    s = re.sub(r"\b(ATIVO|PASSIVO)\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\b(EM TRAMITACAO|ARQUIVADO|SUSPENSO|FINALIZADO)\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\b(DOCUMENTO|PESQUISA POR|PESQUISA)\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\b\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}\b", " ", s)
    s = re.sub(r"\b\d{4}\.\d{3}\.\d{3}-\d\b", " ", s)
    s = re.sub(r"\b\d{6,}\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s or "A confirmar"


CNJ_PATTERNS = [
    r"\b\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}\b",
    r"\b\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}\b",
    r"\b\d{1,4}(?:\.\d{1,4}){3,}\b",
    r"\b\d{4,}-\d{1,3}\b",
]


def _assunto_only(texto: str) -> str:
    s = _strip(texto)
    s = s.replace("|", " ").replace("@", " ").replace("$", " ")
    s = re.sub(r"[()]", " ", s)
    s = DATE_RE.sub(" ", s)
    s = re.split(r"\bFiltrados\b", s, flags=re.IGNORECASE)[0].strip()
    s = MONEY_RE.sub(" ", s)
    for pat in CNJ_PATTERNS:
        s = re.sub(pat, " ", s)
    s = re.sub(r"\d+", " ", s)
    s = re.sub(r"\bR\$\b", " ", s)
    s = re.sub(r"\bRECEBIDO\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\b(ATIVO|PASSIVO)\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\bTRAMITA[ÇC][AÃ]O\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\b(EM\s+TRAMITACAO|TRAMITACAO)\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\b(DOCUMENTO|PESQUISA|PESQUISA\s+POR)\b", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\s*-\s*\.\s*\.", " ", s)
    s = re.sub(r"[\.]{2,}", " ", s)
    s = _no_accents_upper(s)
    s = s.replace("PROCESUAL", "PROCESSUAL")
    s = s.replace("PROCESUALCMIL", "PROCESSUAL CIVIL")
    s = s.replace("PROCESUALCIVIL", "PROCESSUAL CIVIL")
    s = s.replace("CIVL", "CIVIL")
    drop_words = {
        "CLASSE",
        "ASSUNTO",
        "PROCESSO",
        "PESQUISA",
        "DOCUMENTO",
        "STATUS",
        "UF",
        "RECEBIDO",
        "TIPO",
        "VALOR",
        "PAPEL",
        "ATIVO",
        "ATIVA",
        "PASSIVO",
        "ARQUIVAMENTO",
        "DEFINITIVO",
        "ARQUIVADO",
        "SUSPENSO",
        "FINALIZADO",
        "BAIXADO",
        "TRAMITACAO",
        "EM",
    }
    drop_prefixes = (
        "PASSIV",
        "PASSIO",
        "ARQUIV",
        "DEFINIT",
        "DEHINIT",
        "SEFIN",
        "FINIT",
        "SUSPEN",
        "FINALIZ",
        "BAIXAD",
        "TRAMIT",
    )
    tokens = re.findall(r"[A-Z]+", s)
    keep_short = {"DE", "DO", "DA", "DAS", "DOS", "E", "EM", "POR", "A", "AO", "NA", "NO", "COM", "PARA"}
    tokens = [
        t
        for t in tokens
        if t not in drop_words and not any(t.startswith(p) for p in drop_prefixes) and (len(t) >= 3 or t in keep_short)
    ]
    # Limpeza leve de "E POR" (OCR comum em VADU)
    tokens = [t for i, t in enumerate(tokens) if not (t == "E" and i + 1 < len(tokens) and tokens[i + 1] == "POR")]

    class_prefix = {
        "PROCEDIMENTO",
        "COMUM",
        "CIVEL",
        "JUIZADO",
        "ESPECIAL",
        "CARTA",
        "PRECATORIA",
        "CUMPRIMENTO",
        "SENTENCA",
        "PROCESSUAL",
        "TRABALHO",
        "FAZENDA",
    }
    idx = 0
    while idx < len(tokens) and tokens[idx] in class_prefix and (len(tokens) - idx) >= 2:
        idx += 1
    while idx < len(tokens) and tokens[idx] in {"DE", "DO", "DA", "DAS", "DOS"} and (len(tokens) - idx) >= 2:
        idx += 1
    trimmed = tokens[idx:] if idx < len(tokens) else tokens
    cleaned = [t for t in trimmed if t not in class_prefix]
    if len(cleaned) >= 2:
        trimmed = cleaned

    if len(trimmed) >= 2:
        return " ".join(trimmed[:6]).strip()
    if len(trimmed) == 1:
        return trimmed[0]
    if len(tokens) == 1:
        raw = _no_accents_upper(_strip(texto))
        raw_tokens = re.findall(r"[A-Z]+", raw)
        raw_tokens = [t for t in raw_tokens if t not in drop_words and (len(t) >= 3 or t in keep_short)]
        if len(raw_tokens) >= 2:
            return " ".join(raw_tokens[:6]).strip()
        return tokens[0]
    return "A confirmar"


def _parse_ultimos_rows(text_ultimos: str) -> List[Dict[str, str]]:
    lines = [l for l in (text_ultimos or "").splitlines() if _strip(l)]
    chunks: List[str] = []
    current: List[str] = []

    def _find_date(s: str) -> Optional[str]:
        return _find_date_any(s)

    has_date = any(_find_date_any(l) for l in lines)
    if has_date:
        for line in lines:
            if _find_date(line):
                date_splits = list(DATE_RE.finditer(line))
                if date_splits:
                    for i, m in enumerate(date_splits):
                        start = m.start()
                        end = date_splits[i + 1].start() if i + 1 < len(date_splits) else len(line)
                        part = line[start:end].strip()
                        if current:
                            chunks.append(" ".join(current))
                        current = [part]
                    continue
                if current:
                    chunks.append(" ".join(current))
                current = [line]
            else:
                if current:
                    current.append(line)
        if current:
            chunks.append(" ".join(current))
    else:
        # fallback: separar por número de processo (CNJ) e incluir linhas anteriores
        window: List[str] = []
        for line in lines:
            window.append(line)
            if len(window) > 6:
                window.pop(0)
            if any(re.search(pat, line) for pat in CNJ_PATTERNS):
                if current:
                    chunks.append(" ".join(current))
                current = list(window)
            else:
                if current:
                    current.append(line)
        if current:
            chunks.append(" ".join(current))

    def _extract_valor(s: str) -> Tuple[str, int]:
        s_norm = s.replace("RS", "R$")
        val = _max_money_in_text(s_norm)
        if val:
            try:
                num = float(val.replace(".", "").replace(",", "."))
            except ValueError:
                num = -1.0
            if num < 1000 and "0,00" in s_norm:
                return "0,00", 1
            return val, 2
        if re.search(r"\b[-–—]\b", s_norm) or "0,00" in s_norm:
            return "0,00", 1
        m2 = re.search(r"\d{1,3}(?:\.\d{3})*,\d{2}", s_norm)
        if m2:
            return _ptbr_money_from_any(m2.group(0)) or "A confirmar", 0
        return "A confirmar", 0

    def _detect_papel(s: str) -> str:
        s_up = _no_accents_upper(s)
        if "PASSIV" in s_up or "PASSI" in s_up:
            return "PASSIVO"
        if "ATIV" in s_up:
            return "ATIVO"
        return "A confirmar"

    def _detect_status(s: str) -> str:
        s_up = _no_accents_upper(s)
        if "TRAMIT" in s_up:
            return "EM TRAMITACAO"
        if "ARQUIV" in s_up:
            return "ARQUIVAMENTO DEFINITIVO"
        if "SUSPENS" in s_up:
            return "SUSPENSO"
        if "FINALIZ" in s_up:
            return "FINALIZADO"
        if "BAIXAD" in s_up:
            return "BAIXADO"
        return "A confirmar"

    rows: List[Dict[str, str]] = []
    for s in chunks:
        s = _strip(s)
        recebido = _find_date(s)
        if not recebido and not has_date:
            recebido = "A confirmar"
        if not recebido:
            continue

        valor, val_score = _extract_valor(s)

        papel = _detect_papel(s)
        status = _detect_status(s)
        if status == "A confirmar" and papel in {"ATIVO", "PASSIVO"} and valor != "A confirmar":
            status = "EM TRAMITACAO"

        assunto = _assunto_only(s)

        if not (MONEY_RE.search(s) or papel in {"ATIVO", "PASSIVO"} or status != "A confirmar"):
            continue

        if assunto == "A confirmar" and valor == "A confirmar" and papel == "A confirmar":
            continue

        rows.append({
            "Recebido": recebido,
            "Assunto": assunto,
            "Valor (R$)": valor,
            "Papel": papel,
            "status": status,
            "_val_score": val_score,
        })

    return rows


def _select_best_ultimos(rows: List[Dict[str, str]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=["Recebido", "Assunto", "Valor (R$)", "Papel", "status"])

    date_counts: Dict[str, int] = {}
    for r in rows:
        date_counts[r["Recebido"]] = date_counts.get(r["Recebido"], 0) + 1
    if len(date_counts) > 1 and any(v > 1 for v in date_counts.values()):
        filtered = [r for r in rows if date_counts.get(r["Recebido"], 0) > 1]
        if filtered:
            rows = filtered

    def _val_num(v: str) -> float:
        try:
            return float(v.replace(".", "").replace(",", "."))
        except Exception:
            return -1.0

    grouped: Dict[str, List[Dict[str, str]]] = {}
    for r in rows:
        grouped.setdefault(r["Recebido"], []).append(r)

    final_rows = []
    for date, items in grouped.items():
        best = None
        best_score = -1.0
        best_val = -1.0
        for r in items:
            val = r["Valor (R$)"]
            val_num = _val_num(val) if val != "A confirmar" else -1.0
            score = 0
            if r["Papel"] in {"ATIVO", "PASSIVO"}:
                score += 3
            if val != "A confirmar":
                score += 2
            if val not in {"A confirmar", "0,00"}:
                score += 1
            if val not in {"A confirmar", "0,00"}:
                if val_num >= 1000:
                    score += 2
                if val_num >= 10000:
                    score += 1
            score += r.get("_val_score", 0)
            score += min(len(r["Assunto"]), 80) / 40.0

            if val_num > best_val:
                best_val = val_num
                best_score = score
                best = r
                continue
            if val_num == best_val and score > best_score:
                best_val = val_num
                best_score = score
                best = r

        if best:
            final_rows.append({k: v for k, v in best.items() if not k.startswith("_")})

    df = pd.DataFrame(final_rows, columns=["Recebido", "Assunto", "Valor (R$)", "Papel", "status"])
    df = df.drop_duplicates(subset=["Recebido", "Assunto", "Valor (R$)", "Papel"], keep="first")
    return df


def extract_ultimos_processos_from_text(text_ultimos: str) -> pd.DataFrame:
    return _select_best_ultimos(_parse_ultimos_rows(text_ultimos))


def extract_ultimos_processos_from_texts(texts: List[str]) -> pd.DataFrame:
    rows: List[Dict[str, str]] = []
    for text in texts:
        rows.extend(_parse_ultimos_rows(text))
    return _select_best_ultimos(rows)


def extract_ultimos_processos_from_image(image_path: Path) -> pd.DataFrame:
    img = Image.open(image_path)
    best_rows: List[Dict[str, str]] = []
    best_score = -1

    for variant in _prep_ocr_variants(img):
        for psm in (6, 11, 4):
            data = pytesseract.image_to_data(
                variant,
                lang="por+eng",
                config=f"--oem 3 --psm {psm}",
                output_type=Output.DICT,
            )
            words: List[Dict[str, object]] = []
            n = len(data.get("text", []))
            for i in range(n):
                text = (data["text"][i] or "").strip()
                if not text:
                    continue
                try:
                    conf = float(data.get("conf", ["0"])[i])
                except ValueError:
                    conf = 0.0
                if conf < 15 and not re.search(r"\d", text):
                    continue
                words.append({
                    "text": text,
                    "left": int(data["left"][i]),
                    "top": int(data["top"][i]),
                    "width": int(data["width"][i]),
                    "height": int(data["height"][i]),
                    "line_id": (data["block_num"][i], data["par_num"][i], data["line_num"][i]),
                })

            if not words:
                continue

            header_candidates = [w for w in words if w["top"] < int(variant.height * 0.35)]
            header_tokens: Dict[str, int] = {}
            if header_candidates:
                for key, pat in {
                    "recebido": r"RECEBIDO",
                    "classe": r"CLASSE|ASSUNTO",
                    "valor": r"VALOR",
                    "pesquisa": r"PESQUISA",
                    "tipo": r"TIPO",
                    "uf": r"\bUF\b",
                    "status": r"STATUS",
                }.items():
                    hits = [w for w in header_candidates if re.search(pat, _no_accents_upper(str(w["text"])))]
                    if not hits:
                        header_tokens = {}
                        break
                    header_tokens[key] = min(int(h["left"]) for h in hits)

            if header_tokens:
                ordered = sorted(header_tokens.items(), key=lambda kv: kv[1])
                cutpoints = []
                for i in range(len(ordered) - 1):
                    cutpoints.append((ordered[i][0], ordered[i + 1][0], (ordered[i][1] + ordered[i + 1][1]) / 2))
                ranges: Dict[str, Tuple[int, int]] = {}
                left = -10**9
                for i, (k, _) in enumerate(ordered):
                    right = 10**9 if i == len(ordered) - 1 else int(cutpoints[i][2])
                    ranges[k] = (left, right)
                    left = right
                header_bottom = max(int(w["top"]) + int(w["height"]) for w in header_candidates)
                data_words = [w for w in words if int(w["top"]) > header_bottom + 2]
            else:
                w = variant.width
                ranges = {
                    "recebido": (0, int(w * 0.12)),
                    "classe": (int(w * 0.12), int(w * 0.33)),
                    "valor": (int(w * 0.33), int(w * 0.45)),
                    "pesquisa": (int(w * 0.45), int(w * 0.62)),
                    "tipo": (int(w * 0.62), int(w * 0.72)),
                    "uf": (int(w * 0.72), int(w * 0.78)),
                    "status": (int(w * 0.78), int(w * 0.99)),
                }
                data_words = [w for w in words if int(w["top"]) > int(variant.height * 0.05)]

            lines: Dict[Tuple[int, int, int], List[Dict[str, object]]] = {}
            for w in data_words:
                lines.setdefault(w["line_id"], []).append(w)

            def _line_texts(line_words: List[Dict[str, object]]) -> Dict[str, str]:
                line_words = sorted(line_words, key=lambda w: int(w["left"]))
                col_text: Dict[str, List[str]] = {k: [] for k in ranges.keys()}
                for w in line_words:
                    x = int(w["left"]) + int(w["width"]) // 2
                    for col, (lo, hi) in ranges.items():
                        if lo <= x < hi:
                            col_text[col].append(str(w["text"]))
                            break
                return {k: _strip(" ".join(v)) for k, v in col_text.items()}

            rows: List[Dict[str, str]] = []
            current: Optional[Dict[str, object]] = None

            for _, line_words in sorted(lines.items(), key=lambda kv: min(int(w["top"]) for w in kv[1])):
                texts = _line_texts(line_words)
                recebido_text = texts.get("recebido", "")
                date_norm = _find_date_any(recebido_text)
                if not date_norm:
                    full_line = " ".join(v for v in texts.values() if v)
                    date_norm = _find_date_any(full_line)
                processo_text = texts.get("classe", "")
                has_processo = any(re.search(pat, processo_text) for pat in CNJ_PATTERNS)
                if date_norm or has_processo:
                    if current:
                        rows.append(current)
                    current = {
                        "Recebido": date_norm if date_norm else "A confirmar",
                        "Assunto_parts": [],
                        "Valor (R$)": "A confirmar",
                        "Valor_parts": [],
                        "Papel": "A confirmar",
                        "status_parts": [],
                    }

                if not current:
                    continue

                assunto_raw = texts.get("classe", "")
                if assunto_raw:
                    current["Assunto_parts"].append(assunto_raw)

                valor_raw = texts.get("valor", "")
                if valor_raw:
                    current["Valor_parts"].append(valor_raw)

                tipo_raw = texts.get("tipo", "")
                if tipo_raw:
                    tipo_up = _no_accents_upper(tipo_raw)
                    if "PASSIV" in tipo_up:
                        current["Papel"] = "PASSIVO"
                    elif "ATIV" in tipo_up:
                        current["Papel"] = "ATIVO"

                status_raw = texts.get("status", "")
                if status_raw:
                    current["status_parts"].append(status_raw)

            if current:
                rows.append(current)

            final_rows: List[Dict[str, str]] = []
            for r in rows:
                assunto = _assunto_only(" ".join(r.get("Assunto_parts", [])))
                valor_raw = " ".join(r.get("Valor_parts", []))
                valor_norm = valor_raw.replace("RS", "R$")
                valor = r.get("Valor (R$)", "A confirmar")
                val = _max_money_in_text(valor_norm) or _ptbr_money_from_any(valor_norm)
                if val:
                    valor = val
                elif re.search(r"\b[-–—]\b", valor_norm) or "0,00" in valor_norm:
                    valor = "0,00"

                status_raw = " ".join(r.get("status_parts", []))
                status_up = _no_accents_upper(status_raw)
                status = "A confirmar"
                if "TRAMIT" in status_up:
                    status = "EM TRAMITACAO"
                elif "ARQUIV" in status_up:
                    status = "ARQUIVAMENTO DEFINITIVO"
                elif "DEFINITIVO" in status_up:
                    status = "ARQUIVAMENTO DEFINITIVO"
                elif "SUSPENS" in status_up:
                    status = "SUSPENSO"
                elif "FINALIZ" in status_up:
                    status = "FINALIZADO"
                elif status_raw:
                    status = status_up

                final_rows.append({
                    "Recebido": r.get("Recebido", "A confirmar"),
                    "Assunto": assunto,
                    "Valor (R$)": valor,
                    "Papel": r.get("Papel", "A confirmar"),
                    "status": status,
                })

            if final_rows:
                if all(r["Recebido"] == "A confirmar" for r in final_rows):
                    rx0, rx1 = ranges["recebido"]
                    dlist = _extract_dates_from_column(variant, rx0, rx1)
                    if not dlist:
                        bx0 = int(img.width * (rx0 / max(1, variant.width)))
                        bx1 = int(img.width * (rx1 / max(1, variant.width)))
                        dlist = _extract_dates_from_column(img, bx0, bx1)
                    for i, d in enumerate(dlist):
                        if i < len(final_rows):
                            final_rows[i]["Recebido"] = d
                if all(r["Valor (R$)"] == "A confirmar" for r in final_rows):
                    rx0, rx1 = ranges["valor"]
                    vlist = _extract_values_from_column(variant, rx0, rx1)
                    if not vlist:
                        bx0 = int(img.width * (rx0 / max(1, variant.width)))
                        bx1 = int(img.width * (rx1 / max(1, variant.width)))
                        vlist = _extract_values_from_column(img, bx0, bx1)
                    for i, v in enumerate(vlist):
                        if i < len(final_rows):
                            final_rows[i]["Valor (R$)"] = v

            score = sum(1 for r in final_rows if r["Recebido"] != "A confirmar")
            score += sum(1 for r in final_rows if r["Valor (R$)"] != "A confirmar")
            if len(final_rows) > len(best_rows) or (len(final_rows) == len(best_rows) and score > best_score):
                best_rows = final_rows
                best_score = score

    if not best_rows:
        return pd.DataFrame(columns=["Recebido", "Assunto", "Valor (R$)", "Papel", "status"])

    return _select_best_ultimos(best_rows)


def _score_cards(cards: Dict[str, str]) -> int:
    score = 0
    for key in PROCESSOS_FIELDS:
        if key == "Valor total em protesto":
            continue
        val = cards.get(key)
        if val and val not in {"-", "A confirmar"}:
            score += 1
    return score


def _select_best_processos_text(texts: List[str], text_protestos: str) -> Tuple[str, Dict[str, str]]:
    best_text = ""
    best_cards: Dict[str, str] = {}
    best_score = -1
    for text in texts:
        cards = extract_cards_processos(text_processos_areas=text, text_protestos=text_protestos)
        score = _score_cards(cards)
        if score > best_score:
            best_score = score
            best_text = text
            best_cards = cards
    if not best_text and texts:
        best_text = texts[0]
        best_cards = extract_cards_processos(text_processos_areas=best_text, text_protestos=text_protestos)
    return best_text, best_cards


def _select_best_areas_df(texts: List[str]) -> pd.DataFrame:
    best_df = pd.DataFrame(columns=["Área", "Qtde"])
    best_len = -1
    for text in texts:
        df = extract_areas_from_text(text)
        if len(df.index) > best_len:
            best_len = len(df.index)
            best_df = df
    return best_df

# =========================
# Validacoes STRICT
# =========================


def _validate_money_text(s: str) -> bool:
    if s in {"-", "A confirmar", "a confirmar", "0,00"}:
        return True
    return bool(re.fullmatch(r"\d{1,3}(?:\.\d{3})*,\d{2}", s))


def _validate_headers(df: pd.DataFrame, expected: List[str]) -> None:
    if list(df.columns) != expected:
        raise RuntimeError("ERRO: Cabecalho fora da ordem especificada.")


def _normalize_money_columns(df: pd.DataFrame, col: str) -> None:
    def _norm_one(v: str) -> str:
        s = str(v).strip()
        if _is_blank(s):
            return "A confirmar"
        if s in {"-", "--"}:
            return "-"
        if s.lower() == "a confirmar":
            return "A confirmar"
        if _validate_money_text(s):
            return s
        norm = _ptbr_money_from_any(s)
        return norm if norm is not None else "A confirmar"

    df[col] = df[col].astype(str).map(_norm_one)


def _validate_money_columns(df: pd.DataFrame, col: str) -> None:
    for v in df[col].astype(str).tolist():
        if not _validate_money_text(v):
            raise RuntimeError(f"ERRO: Valor monetario com separador incorreto: {v}")


# =========================
# Montagem XLSX
# =========================


def build_xlsx(
    out_xlsx: Path,
    cnpj: str,
    cards: Dict[str, str],
    df_areas: pd.DataFrame,
    df_ultimos: pd.DataFrame,
) -> None:
    processos_rows = []
    for campo in PROCESSOS_FIELDS:
        val = cards.get(campo, "-")
        processos_rows.append({"Campo": campo, "Valor": val, "cnpj": cnpj})
    df_processos = pd.DataFrame(processos_rows, columns=["Campo", "Valor", "cnpj"])

    if df_areas is None or df_areas.empty:
        df_areas = pd.DataFrame(columns=["Área", "Qtde"])
    df_areas = df_areas.copy()
    if "Area" in df_areas.columns and "Área" not in df_areas.columns:
        df_areas = df_areas.rename(columns={"Area": "Área"})
    df_areas["cnpj"] = cnpj
    df_areas = df_areas[["Área", "Qtde", "cnpj"]]

    if df_ultimos is None or df_ultimos.empty:
        df_ultimos = pd.DataFrame(columns=["Recebido", "Assunto", "Valor (R$)", "Papel", "status"])
    df_ultimos = df_ultimos.copy()
    df_ultimos["cnpj"] = cnpj
    df_ultimos = df_ultimos[["Recebido", "Assunto", "Valor (R$)", "Papel", "status", "cnpj"]]

    _validate_headers(df_processos, ["Campo", "Valor", "cnpj"])
    _validate_headers(df_areas, ["Área", "Qtde", "cnpj"])
    _validate_headers(df_ultimos, ["Recebido", "Assunto", "Valor (R$)", "Papel", "status", "cnpj"])

    _normalize_money_columns(df_processos, "Valor")
    if not df_ultimos.empty:
        _normalize_money_columns(df_ultimos, "Valor (R$)")

    _validate_money_columns(df_processos, "Valor")
    if not df_ultimos.empty:
        _validate_money_columns(df_ultimos, "Valor (R$)")

    for df in (df_processos, df_areas, df_ultimos):
        if df is not None and not df.empty:
            df.replace({"--": "-"}, inplace=True)

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df_processos.to_excel(writer, sheet_name="Processos", index=False)
        df_areas.to_excel(writer, sheet_name="Areas", index=False)
        df_ultimos.to_excel(writer, sheet_name="UltimosProcessos", index=False)

        wb = writer.book
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    cell.number_format = "@"


# =========================
# Orquestracao
# =========================


def run_company(
    company_key: str,
    protestos_img: Optional[Path],
    processos_areas_img: Optional[Path],
    ultimos_imgs: Optional[List[Path]],
    cnpj_img: Optional[Path],
    outdir: Path,
    cnpj_user: Optional[str],
) -> Path:
    prot_texts = ocr_image_texts(protestos_img, psms=[6, 11, 12]) if protestos_img else []
    text_prot = "\n".join(prot_texts)

    proc_texts = ocr_image_texts(processos_areas_img, psms=[6, 11, 12]) if processos_areas_img else []
    best_proc_text, cards = _select_best_processos_text(proc_texts, text_prot)
    if not cards:
        cards = extract_cards_processos(text_processos_areas=best_proc_text, text_protestos=text_prot)

    df_areas = _select_best_areas_df(proc_texts)
    if df_areas.empty and best_proc_text:
        df_areas = extract_areas_from_text(best_proc_text)

    df_ultimos = pd.DataFrame(columns=["Recebido", "Assunto", "Valor (R$)", "Papel", "status"])
    ult_texts: List[str] = []
    if ultimos_imgs:
        for img_path in ultimos_imgs:
            df_img = extract_ultimos_processos_from_image(img_path)
            if df_img is not None and not df_img.empty:
                df_ultimos = pd.concat([df_ultimos, df_img], ignore_index=True)
            ult_texts.extend(ocr_image_texts_variants(img_path, psms=[11, 12, 6, 4]))
        if ult_texts:
            df_fallback = extract_ultimos_processos_from_texts(ult_texts)
            if not df_fallback.empty:
                def _score_ultimos_df(df: pd.DataFrame) -> Tuple[int, int]:
                    if df is None or df.empty:
                        return (-1, -1)
                    filled = int((df["Recebido"] != "A confirmar").sum())
                    filled += int((df["Valor (R$)"] != "A confirmar").sum())
                    return (filled, len(df.index))

                score_img = _score_ultimos_df(df_ultimos)
                score_fb = _score_ultimos_df(df_fallback)
                if score_fb >= score_img:
                    df_ultimos = df_fallback
                else:
                    # Preenche campos faltantes por ordem das linhas
                    lim = min(len(df_ultimos.index), len(df_fallback.index))
                    for i in range(lim):
                        for col in ["Recebido", "Assunto", "Valor (R$)", "Papel", "status"]:
                            v = str(df_ultimos.at[i, col]) if col in df_ultimos.columns else "A confirmar"
                            if v in {"A confirmar", "", "nan"}:
                                fb = str(df_fallback.at[i, col]) if col in df_fallback.columns else ""
                                if fb and fb not in {"A confirmar", "nan"}:
                                    df_ultimos.at[i, col] = fb
        if not df_ultimos.empty:
            df_ultimos = df_ultimos.drop_duplicates(subset=["Recebido", "Assunto", "Valor (R$)", "Papel"], keep="first")

    cnpj_texts = ocr_image_texts(cnpj_img, psms=[6, 11, 12]) if cnpj_img else []
    text_cnpj = "\n".join(cnpj_texts)

    empresa = company_key
    cnpj = extract_cnpj(text_cnpj, cnpj_user=cnpj_user)
    if cnpj == "A confirmar":
        cnpj = extract_cnpj("\n".join([text_prot, best_proc_text, "\n".join(ult_texts)]), cnpj_user=cnpj_user)

    safe_name = re.sub(r"[^\w\s\-\.]", "", empresa, flags=re.UNICODE)
    safe_name = _strip(safe_name).replace(" ", "_")
    out_xlsx = outdir / f"VADU_{safe_name}.xlsx"

    build_xlsx(
        out_xlsx=out_xlsx,
        cnpj=cnpj,
        cards=cards,
        df_areas=df_areas,
        df_ultimos=df_ultimos,
    )
    return out_xlsx


# =========================
# CLI
# =========================


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input-dir", required=True, help="Pasta com imagens .png do VADU (protestos/processos_areas/ultimosprocessos/cnpj).")
    ap.add_argument("--outdir", required=True, help="Pasta de saida.")
    ap.add_argument("--cnpj", default="", help="CNPJ opcional (se informado, vira definitivo).")
    ap.add_argument("--tesseract", default="", help="Caminho do tesseract.exe (opcional).")
    args = ap.parse_args()

    input_dir = Path(args.input_dir)
    outdir = Path(args.outdir)
    cnpj_user = args.cnpj if args.cnpj else None

    if args.tesseract:
        pytesseract.pytesseract.tesseract_cmd = args.tesseract
    _check_tesseract()

    companies = collect_inputs(input_dir)
    out_files = []
    for company_key, info in companies.items():
        out_xlsx = run_company(
            company_key=company_key,
            protestos_img=info.get("protestos"),
            processos_areas_img=info.get("processos_areas"),
            ultimos_imgs=info.get("ultimos_processos"),
            cnpj_img=info.get("cnpj"),
            outdir=outdir,
            cnpj_user=cnpj_user,
        )
        out_files.append(out_xlsx)

    for f in out_files:
        print("Arquivo gerado.")
        print(str(f))


if __name__ == "__main__":
    main()
