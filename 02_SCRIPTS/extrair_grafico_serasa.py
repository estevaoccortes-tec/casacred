#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Extrai série mensal de um gráfico de barras (Serasa / "Consultas nos últimos 13 meses") a partir de PDF.

Dependências:
  python -m pip install -U pdf2image pillow pytesseract opencv-python numpy

Requisitos do SO:
  - Poppler (pdfinfo.exe e pdftoppm.exe) -> você já tem
  - Tesseract OCR instalado (opcional informar caminho via --tesseract)

Exemplo:
  .\.venv\Scripts\python.exe .\02_SCRIPTS\extrair_grafico_serasa.py ^
    "C:/Users/Usuario/Desktop/DADOS_BI/01_INPUT/Global_Tabacos/serasa_cedente_global.pdf" ^
    --outdir "C:/Users/Usuario/Desktop/DADOS_BI/03_OUTPUT/grafico_final" ^
    --dpi 200 ^
    --poppler "C:/Users/Usuario/Desktop/poppler-25.12.0/Library/bin" ^
    --start-label "Nov/2024" ^
    --expected-bars 13
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Tuple

import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path


# ---------------------------
# Meses PT-BR
# ---------------------------
PT_MONTHS = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12
}
INV_PT = {v: k for k, v in PT_MONTHS.items()}

def parse_start_label(label: str) -> datetime:
    """
    Espera algo como "Nov/2024".
    """
    s = (label or "").strip().lower()
    m = re.search(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\s*[/\-]\s*(20\d{2})", s)
    if not m:
        raise ValueError(f'--start-label inválido: "{label}". Use tipo "Nov/2024".')
    mon = PT_MONTHS[m.group(1)]
    year = int(m.group(2))
    return datetime(year, mon, 1)

def add_months(dt: datetime, n: int) -> datetime:
    # sem dateutil pra evitar dependência
    y = dt.year
    m = dt.month + n
    y += (m - 1) // 12
    m = ((m - 1) % 12) + 1
    return datetime(y, m, 1)

def fmt_label(dt: datetime) -> str:
    return f"{INV_PT[dt.month].capitalize()}/{dt.year}"

def mes_ref(dt: datetime) -> str:
    return dt.strftime("%Y-%m-01")


# ---------------------------
# Estruturas
# ---------------------------
@dataclass
class Bar:
    x: int
    y: int
    w: int
    h: int

    @property
    def cx(self) -> float:
        return self.x + self.w / 2.0


# ---------------------------
# Imagem: helpers
# ---------------------------
def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)

def bgr(img_rgb: np.ndarray) -> np.ndarray:
    return cv2.cvtColor(img_rgb, cv2.COLOR_RGB2BGR)

def rgb(img_bgr: np.ndarray) -> np.ndarray:
    return cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGB)

def save_png(path: Path, img_bgr: np.ndarray) -> None:
    cv2.imwrite(str(path), img_bgr)

def clamp(v: int, lo: int, hi: int) -> int:
    return max(lo, min(hi, v))


# ---------------------------
# Máscara do azul (barras)
# ---------------------------
def blue_mask_hsv(img_bgr: np.ndarray) -> np.ndarray:
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)

    # Ajuste fino (esse range costuma pegar bem o azul Serasa)
    lower = np.array([90, 40, 40], dtype=np.uint8)
    upper = np.array([150, 255, 255], dtype=np.uint8)
    mask = cv2.inRange(hsv, lower, upper)

    # limpa ruído
    k1 = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    k2 = cv2.getStructuringElement(cv2.MORPH_RECT, (9, 9))
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, k1, iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, k2, iterations=2)
    return mask

def remove_blue(img_bgr: np.ndarray) -> np.ndarray:
    mask = blue_mask_hsv(img_bgr)
    out = img_bgr.copy()
    out[mask > 0] = (255, 255, 255)
    return out


# ---------------------------
# PDF -> páginas (RGB)
# ---------------------------
def render_pages(pdf_path: Path, dpi: int, poppler_path: Path) -> List[np.ndarray]:
    pages = convert_from_path(
        str(pdf_path),
        dpi=dpi,
        poppler_path=str(poppler_path),
        fmt="png"
    )
    imgs = []
    for p in pages:
        imgs.append(np.array(p))  # RGB
    return imgs

def pick_page_by_blue_score(pages_rgb: List[np.ndarray], debug_dir: Path) -> int:
    scores = []
    for i, img_rgb in enumerate(pages_rgb, start=1):
        img_bgr = bgr(img_rgb)
        mask = blue_mask_hsv(img_bgr)
        score = int(mask.sum())
        scores.append((score, i))
        save_png(debug_dir / f"debug_mask_blue_p{i}.png", mask)

    scores_sorted = sorted(scores, key=lambda t: t[0], reverse=True)
    # página é 1-based aqui
    best = scores_sorted[0][1]
    (debug_dir / "debug_scores.txt").write_text(
        "\n".join([f"page={p} score={s}" for s, p in scores_sorted]),
        encoding="utf-8"
    )
    return best


# ---------------------------
# ROI do gráfico
# ---------------------------
def detect_bars_in_image(img_bgr: np.ndarray) -> List[Bar]:
    mask = blue_mask_hsv(img_bgr)
    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    bars: List[Bar] = []
    H, W = img_bgr.shape[:2]

    for c in cnts:
        x, y, w, h = cv2.boundingRect(c)

        # Heurística bem permissiva (pra não matar barra pequena tipo "3")
        if w < max(6, W // 300):   # muito fino = ruído
            continue
        if h < max(12, H // 300):  # muito baixo = ruído
            continue

        # Evita pegar linha do eixo etc
        if h > int(H * 0.9):
            continue

        bars.append(Bar(x, y, w, h))

    # Ordena por X
    bars.sort(key=lambda b: b.x)

    # Merge simples de barras fragmentadas (quando a máscara quebra a barra em 2)
    merged: List[Bar] = []
    for b in bars:
        if not merged:
            merged.append(b)
            continue
        prev = merged[-1]
        # se está muito perto em X e sobrepõe em Y, une
        if abs(b.x - prev.x) < 10 and abs((b.y + b.h) - (prev.y + prev.h)) < 40:
            x1 = min(prev.x, b.x)
            y1 = min(prev.y, b.y)
            x2 = max(prev.x + prev.w, b.x + b.w)
            y2 = max(prev.y + prev.h, b.y + b.h)
            merged[-1] = Bar(x1, y1, x2 - x1, y2 - y1)
        else:
            merged.append(b)

    return merged

def crop_chart_roi(img_bgr: np.ndarray, bars: List[Bar]) -> Tuple[np.ndarray, Tuple[int, int, int, int]]:
    """
    Retorna:
      - roi_bgr
      - roi_bbox_page: (x, y, w, h) em coordenadas da página
    """
    H, W = img_bgr.shape[:2]
    xs = [b.x for b in bars]
    xe = [b.x + b.w for b in bars]
    ys = [b.y for b in bars]
    ye = [b.y + b.h for b in bars]

    x1 = clamp(min(xs) - 160, 0, W - 1)
    x2 = clamp(max(xe) + 160, 0, W)
    y1 = clamp(min(ys) - 260, 0, H - 1)   # espaço números
    y2 = clamp(max(ye) + 260, 0, H)       # espaço meses

    roi = img_bgr[y1:y2, x1:x2].copy()
    return roi, (x1, y1, x2 - x1, y2 - y1)


# ---------------------------
# OCR números (multi-pass forte)
# ---------------------------
def preprocess_for_digits(roi_bgr: np.ndarray, mode: str) -> np.ndarray:
    gray = cv2.cvtColor(roi_bgr, cv2.COLOR_BGR2GRAY)

    # upscale para ajudar o tesseract
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
        # fallback: simples
        _, th = cv2.threshold(gray, 170, 255, cv2.THRESH_BINARY)

    return th

def tighten_binary(th_img: np.ndarray) -> np.ndarray:
    # Tight crop around ink to reduce whitespace before OCR.
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

def largest_component(th_img: np.ndarray) -> np.ndarray:
    # Keep only the largest ink component to reduce noise for OCR.
    inv = (th_img < 128).astype(np.uint8)
    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(inv, connectivity=8)
    if num_labels <= 1:
        return th_img
    idx = 1 + int(np.argmax(stats[1:, cv2.CC_STAT_AREA]))
    mask = (labels == idx).astype(np.uint8)
    out = np.full(th_img.shape, 255, dtype=np.uint8)
    out[mask > 0] = 0
    return out

def ocr_digits(th_img: np.ndarray, psm: int) -> str:
    cfg = f"--oem 3 --psm {psm} -c tessedit_char_whitelist=0123456789"
    txt = pytesseract.image_to_string(th_img, config=cfg) or ""
    txt = re.sub(r"\D", "", txt)
    return txt

def best_digit_ocr(roi_bgr: np.ndarray, debug_paths: List[Tuple[str, Path]]) -> Tuple[int | None, str | None, str | None]:
    """
    Retorna:
      consultas (int | None), raw_ocr (str|None), ocr_mode (str|None)
    """
    # remove azul antes de OCR
    roi_nb = remove_blue(roi_bgr)

    tries = [
        ("otsu", 7),
        ("adapt", 7),
        ("adapt_inv", 7),
        ("otsu", 10),      # psm 10 (single char) ajuda muito no "9"
        ("adapt", 10),
        ("otsu", 8),
        ("adapt", 8),
        ("otsu", 13),
    ]

    best_txt = ""
    best_mode = None

    for mode, psm in tries:
        th = preprocess_for_digits(roi_nb, mode=mode)

        # pequena morfologia pra reforçar traço do dígito
        k = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        th2 = cv2.morphologyEx(th, cv2.MORPH_CLOSE, k, iterations=1)
        th3 = tighten_binary(th2)
        txt = ocr_digits(th3, psm=psm)
        if not txt:
            th4 = largest_component(th2)
            th3 = tighten_binary(th4)
            txt = ocr_digits(th3, psm=psm)

        # debug (salva o threshold que tentou)
        for tag, path in debug_paths:
            if tag == f"{mode}_psm{psm}":
                save_png(path, th3)

        # critério: preferir resultado maior (ex: "13" > "1")
        if len(txt) > len(best_txt):
            best_txt = txt
            best_mode = f"{mode}_psm{psm}"
        elif len(txt) == len(best_txt) and txt > best_txt:
            best_txt = txt
            best_mode = f"{mode}_psm{psm}"

        # se já achou 2 dígitos tipo 10/12/13, ótimo
        if len(best_txt) >= 2:
            break

    if not best_txt:
        # Fallback: stronger upscale + inverted Otsu on raw ROI.
        gray = cv2.cvtColor(roi_nb, cv2.COLOR_BGR2GRAY)
        gray = cv2.resize(gray, None, fx=6, fy=6, interpolation=cv2.INTER_CUBIC)
        _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        k = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        th = cv2.morphologyEx(th, cv2.MORPH_CLOSE, k, iterations=1)
        th = 255 - th
        th = largest_component(th)
        th = tighten_binary(th)

        for psm in (10, 8, 7):
            txt = ocr_digits(th, psm=psm)
            if txt:
                best_txt = txt
                best_mode = f"fallback_psm{psm}"
                break

    if not best_txt:
        return None, None, None

    try:
        val = int(best_txt)
        # sanity: se vier algo absurdo, rejeita
        if val < 0 or val > 500:
            return None, best_txt, best_mode
        return val, best_txt, best_mode
    except ValueError:
        return None, best_txt, best_mode


# ---------------------------
# OCR meses (opcional) + fallback sequência
# ---------------------------
def normalize_month_text(s: str) -> str:
    s = (s or "").strip()
    s = s.replace("\\", "/").replace("|", "/")
    s = re.sub(r"\s+", "", s)
    return s

def parse_month_label(s: str) -> datetime | None:
    s2 = normalize_month_text(s).lower()
    m = re.search(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez).*(20\d{2})", s2)
    if not m:
        return None
    mon = PT_MONTHS[m.group(1)]
    year = int(m.group(2))
    return datetime(year, mon, 1)

def ocr_month_under_bar(chart_nb_bgr: np.ndarray, bar: Bar) -> str | None:
    """
    Tenta ler label abaixo de cada barra.
    """
    H, W = chart_nb_bgr.shape[:2]
    # região abaixo (ajuste fino)
    x1 = clamp(bar.x - 40, 0, W - 1)
    x2 = clamp(bar.x + bar.w + 40, 0, W)
    y1 = clamp(bar.y + bar.h + 25, 0, H - 1)
    y2 = clamp(bar.y + bar.h + 140, 0, H)

    roi = chart_nb_bgr[y1:y2, x1:x2]
    if roi.size == 0:
        return None

    gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
    gray = cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    th = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 41, 15)

    cfg = "--oem 3 --psm 7"
    txt = pytesseract.image_to_string(th, config=cfg) or ""
    txt = normalize_month_text(txt)
    if len(txt) < 5:
        return None
    return txt




def infer_start_month(ocr_months: List[datetime | None], fallback_start: datetime) -> datetime:
    # Choose the start month that best matches OCR labels in order.
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
        start = add_months(dt, -i)
        score = 0
        exact = 0
        for j, dtj in enumerate(ocr_months):
            if dtj is None:
                continue
            exp = add_months(start, j)
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

def extract_one_pdf(
    pdf_path: str | Path,
    outdir: str | Path,
    dpi: int,
    poppler: str | Path,
    page: int,
    expected_bars: int,
    start_label: str,
    override: str,
    tesseract: str,
) -> dict:
    pdf_path = Path(pdf_path)
    outdir = Path(outdir)
    poppler_path = Path(poppler)

    debug_dir = outdir / "debug"
    ensure_dir(outdir)
    ensure_dir(debug_dir)

    if tesseract:
        pytesseract.pytesseract.tesseract_cmd = tesseract

    pages_rgb = render_pages(pdf_path, dpi=dpi, poppler_path=poppler_path)

    if page and page > 0:
        page_selected = page
    else:
        page_selected = pick_page_by_blue_score(pages_rgb, debug_dir=debug_dir)

    # page_selected é 1-based
    page_rgb = pages_rgb[page_selected - 1]
    page_bgr = bgr(page_rgb)
    save_png(debug_dir / f"debug_page_p{page_selected}.png", page_bgr)

    # detecta barras na página inteira (pra achar ROI)
    bars_page = detect_bars_in_image(page_bgr)
    if len(bars_page) < 5:
        raise RuntimeError(f"Poucas barras detectadas na página {page_selected}. Ajuste DPI/HSV.")

    # recorta ROI do gráfico
    roi_bgr, roi_bbox = crop_chart_roi(page_bgr, bars_page)
    (rx, ry, rw, rh) = roi_bbox
    save_png(debug_dir / f"debug_roi_p{page_selected}.png", roi_bgr)

    # detecta barras dentro do ROI (mais confiável)
    roi_mask = blue_mask_hsv(roi_bgr)
    save_png(debug_dir / f"debug_mask_blue_p{page_selected}_roi.png", roi_mask)

    # barras no ROI
    bars_roi = detect_bars_in_image(roi_bgr)

    # em alguns PDFs, o topo do relatório tem azul; aqui filtramos barras pela “faixa do gráfico”
    # regra: manter barras cuja base (y+h) esteja na parte inferior do ROI
    base_ys = np.array([b.y + b.h for b in bars_roi], dtype=np.int32)
    if len(base_ys) > 0:
        base_ref = int(np.median(base_ys))
        bars_roi = [b for b in bars_roi if abs((b.y + b.h) - base_ref) < int(roi_bgr.shape[0] * 0.25)]

    bars_roi.sort(key=lambda b: b.x)

    # se ainda vier barulho, keep só os N mais “altos”
    if len(bars_roi) > expected_bars:
        bars_roi = sorted(bars_roi, key=lambda b: b.h, reverse=True)[:expected_bars]
        bars_roi.sort(key=lambda b: b.x)

    # debug bbox
    dbg = roi_bgr.copy()
    for b in bars_roi:
        cv2.rectangle(dbg, (b.x, b.y), (b.x + b.w, b.y + b.h), (0, 0, 255), 2)
    save_png(debug_dir / "bars_bbox.png", dbg)

    # OCR números
    chart_nb = remove_blue(roi_bgr)
    save_png(debug_dir / "roi_no_blue.png", chart_nb)

    results = []
    ocr_months: List[datetime | None] = []
    for idx, b in enumerate(bars_roi):
        # ROI acima da barra: (muito importante pro "9")
        H, W = chart_nb.shape[:2]
        x1 = clamp(b.x - 55, 0, W - 1)
        x2 = clamp(b.x + b.w + 55, 0, W)

        # altura do "teto" acima: proporcional ao ROI, mas com piso mínimo
        up = max(140, int(H * 0.22))
        y1 = clamp(b.y - up, 0, H - 1)
        y2 = clamp(b.y - 5, 0, H)

        num_roi = chart_nb[y1:y2, x1:x2].copy()

        # salva num roi bruto
        save_png(debug_dir / f"num_raw_{idx:02d}.png", num_roi)

        # prepara lista de debug thresholds por tentativa
        debug_paths = [
            (f"otsu_psm7", debug_dir / f"num_{idx:02d}_otsu_psm7.png"),
            (f"adapt_psm7", debug_dir / f"num_{idx:02d}_adapt_psm7.png"),
            (f"adapt_inv_psm7", debug_dir / f"num_{idx:02d}_adaptinv_psm7.png"),
            (f"otsu_psm10", debug_dir / f"num_{idx:02d}_otsu_psm10.png"),
            (f"adapt_psm10", debug_dir / f"num_{idx:02d}_adapt_psm10.png"),
        ]

        val, raw, mode = best_digit_ocr(num_roi, debug_paths=debug_paths)

        # tenta OCR do mês (opcional)
        mtxt = ocr_month_under_bar(chart_nb, b)
        mdt = parse_month_label(mtxt) if mtxt else None
        ocr_months.append(mdt)

        results.append({
            "ordem": idx,
            "pagina": page_selected,
            "mes": fmt_label(mdt) if mdt else None,
            "mes_ref": mes_ref(mdt) if mdt else None,
            "consultas": val,
            "raw_ocr": raw,
            "ocr_mode": mode,
            "bar_bbox_roi": [b.x, b.y, b.w, b.h],
            "roi_bbox_page": [rx, ry, rw, rh],
        })

    # meses: usa OCR para inferir inicio e gera sequencia consistente
    dt_fallback = parse_start_label(start_label)
    dt0 = infer_start_month(ocr_months, fallback_start=dt_fallback)

    # Se quantidade de barras diferente do esperado, a gente continua mas preenche meses pelo que saiu
    n = len(results)
    months = [add_months(dt0, i) for i in range(n)]
    for i in range(n):
        results[i]["mes"] = fmt_label(months[i])
        results[i]["mes_ref"] = mes_ref(months[i])

    # Override manual por indice (ordem)
    if override:
        for item in override.split(","):
            item = item.strip()
            if not item:
                continue
            if "=" not in item:
                raise ValueError('Formato de --override invalido. Use "3=9,10=8".')
            idx_str, val_str = item.split("=", 1)
            idx = int(idx_str.strip())
            val = int(val_str.strip())
            if idx < 0 or idx >= len(results):
                raise ValueError(f"Indice de override fora do range: {idx}")
            results[idx]["consultas"] = val
            results[idx]["raw_ocr"] = str(val)
            results[idx]["ocr_mode"] = "manual"

    # Export JSON + CSV
    payload = {
        "pdf": str(pdf_path).replace("\\", "/"),
        "dpi": dpi,
        "page_selected": page_selected,
        "bars_detected": len(results),
        "roi_bbox_page": [rx, ry, rw, rh],
        "data": results,
        "debug_dir": str(debug_dir).replace("\\", "/"),
    }

    json_path = outdir / "consultas_grafico_full.json"
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    csv_path = outdir / "consultas_grafico.csv"
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["pagina", "mes_ref", "mes", "consultas"])
        w.writeheader()
        for r in results:
            w.writerow({
                "pagina": r["pagina"],
                "mes_ref": r["mes_ref"],
                "mes": r["mes"],
                "consultas": r["consultas"] if r["consultas"] is not None else "",
            })

    return payload

# ---------------------------
# Pipeline principal
# ---------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", help="Caminho do PDF")
    ap.add_argument("--outdir", default="./03_OUTPUT/grafico_final", help="Pasta de saída")
    ap.add_argument("--dpi", type=int, default=200, help="DPI (200-450). Comece com 200.")
    ap.add_argument("--poppler", required=True, help="Pasta do Poppler bin (onde tem pdfinfo.exe)")
    ap.add_argument("--page", type=int, default=0, help="Página 1-based (0 = auto pelo score azul)")
    ap.add_argument("--tesseract", default="", help="Caminho do tesseract.exe (opcional)")
    ap.add_argument("--expected-bars", type=int, default=13, help="Quantas barras espera (13 meses)")
    ap.add_argument("--start-label", default="Nov/2024", help='Fallback de sequência, ex: "Nov/2024"')
    ap.add_argument("--override", default="", help='Override manual por indice: "3=9,10=8"')
    args = ap.parse_args()

    payload = extract_one_pdf(
        pdf_path=args.pdf,
        outdir=args.outdir,
        dpi=args.dpi,
        poppler=args.poppler,
        page=args.page,
        expected_bars=args.expected_bars,
        start_label=args.start_label,
        override=args.override,
        tesseract=args.tesseract,
    )

    print(f"OK Página escolhida: {payload['page_selected']}")
    print(f"OK Barras detectadas: {payload['bars_detected']} (esperado: {args.expected_bars})")
    print("OK Salvei:")
    print(f" - {str(Path(args.outdir) / 'consultas_grafico_full.json')}")
    print(f" - {str(Path(args.outdir) / 'consultas_grafico.csv')}")
    print(f"DIR Debugs em: {str(Path(args.outdir) / 'debug')}")


if __name__ == "__main__":
    main()



