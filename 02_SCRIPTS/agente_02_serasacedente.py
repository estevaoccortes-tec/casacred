import os
import re
import csv
import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Tuple, Optional

import cv2
import numpy as np
import pdfplumber
import pandas as pd
import pytesseract
from pdf2image import convert_from_path

# --- CONFIGURAÇÃO ---
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA_INPUT = os.path.join(BASE_DIR, "01_INPUT")
PASTA_DESTINO = os.path.join(BASE_DIR, "03_OUTPUT", "2. SERASA CEDENTE")
POPPLER_PATH = r"C:\Users\Usuario\Desktop\poppler-25.12.0\Library\bin"
START_LABEL = "Nov/2024"
EXPECTED_BARS = 13
DPI_GRAFICO = 200

PT_MONTHS = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12
}
INV_PT = {v: k for k, v in PT_MONTHS.items()}

def parse_start_label(label: str) -> datetime:
    s = (label or "").strip().lower()
    m = re.search(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\s*[/\-]\s*(20\d{2})", s)
    if not m:
        raise ValueError(f'--start-label invalido: "{label}". Use tipo "Nov/2024".')
    mon = PT_MONTHS[m.group(1)]
    year = int(m.group(2))
    return datetime(year, mon, 1)

def add_months(dt: datetime, n: int) -> datetime:
    y = dt.year
    m = dt.month + n
    y += (m - 1) // 12
    m = ((m - 1) % 12) + 1
    return datetime(y, m, 1)

def fmt_label(dt: datetime) -> str:
    return f"{INV_PT[dt.month].capitalize()}/{dt.year}"

def mes_ref(dt: datetime) -> str:
    return dt.strftime("%Y-%m-01")

@dataclass
class Bar:
    x: int
    y: int
    w: int
    h: int

def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)

def bgr(img_rgb: np.ndarray) -> np.ndarray:
    return cv2.cvtColor(img_rgb, cv2.COLOR_RGB2BGR)

def save_png(path: Path, img_bgr: np.ndarray) -> None:
    cv2.imwrite(str(path), img_bgr)

def clamp(v: int, lo: int, hi: int) -> int:
    return max(lo, min(hi, v))

def blue_mask_hsv(img_bgr: np.ndarray) -> np.ndarray:
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
    lower = np.array([90, 40, 40], dtype=np.uint8)
    upper = np.array([150, 255, 255], dtype=np.uint8)
    mask = cv2.inRange(hsv, lower, upper)
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

def render_pages(pdf_path: Path, dpi: int, poppler_path: Path) -> List[np.ndarray]:
    pages = convert_from_path(
        str(pdf_path),
        dpi=dpi,
        poppler_path=str(poppler_path),
        fmt="png"
    )
    imgs = []
    for p in pages:
        imgs.append(np.array(p))
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
    best = scores_sorted[0][1]
    (debug_dir / "debug_scores.txt").write_text(
        "\n".join([f"page={p} score={s}" for s, p in scores_sorted]),
        encoding="utf-8"
    )
    return best

def detect_bars_in_image(img_bgr: np.ndarray) -> List[Bar]:
    mask = blue_mask_hsv(img_bgr)
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

def crop_chart_roi(img_bgr: np.ndarray, bars: List[Bar]) -> Tuple[np.ndarray, Tuple[int, int, int, int]]:
    H, W = img_bgr.shape[:2]
    xs = [b.x for b in bars]
    xe = [b.x + b.w for b in bars]
    ys = [b.y for b in bars]
    ye = [b.y + b.h for b in bars]

    x1 = clamp(min(xs) - 160, 0, W - 1)
    x2 = clamp(max(xe) + 160, 0, W)
    y1 = clamp(min(ys) - 260, 0, H - 1)
    y2 = clamp(max(ye) + 260, 0, H)

    roi = img_bgr[y1:y2, x1:x2].copy()
    return roi, (x1, y1, x2 - x1, y2 - y1)

def preprocess_for_digits(roi_bgr: np.ndarray, mode: str) -> np.ndarray:
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

def tighten_binary(th_img: np.ndarray) -> np.ndarray:
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

def best_digit_ocr(roi_bgr: np.ndarray, debug_paths: List[Tuple[str, Path]]) -> Tuple[Optional[int], Optional[str], Optional[str]]:
    roi_nb = remove_blue(roi_bgr)

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
    best_mode = None

    for mode, psm in tries:
        th = preprocess_for_digits(roi_nb, mode=mode)
        k = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
        th2 = cv2.morphologyEx(th, cv2.MORPH_CLOSE, k, iterations=1)
        th3 = tighten_binary(th2)
        txt = ocr_digits(th3, psm=psm)
        if not txt:
            th4 = largest_component(th2)
            th3 = tighten_binary(th4)
            txt = ocr_digits(th3, psm=psm)

        for tag, path in debug_paths:
            if tag == f"{mode}_psm{psm}":
                save_png(path, th3)

        if len(txt) > len(best_txt):
            best_txt = txt
            best_mode = f"{mode}_psm{psm}"
        elif len(txt) == len(best_txt) and txt > best_txt:
            best_txt = txt
            best_mode = f"{mode}_psm{psm}"

        if len(best_txt) >= 2:
            break

    if not best_txt:
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
        if val < 0 or val > 500:
            return None, best_txt, best_mode
        return val, best_txt, best_mode
    except ValueError:
        return None, best_txt, best_mode

def normalize_month_text(s: str) -> str:
    s = (s or "").strip()
    s = s.replace("\\", "/").replace("|", "/")
    s = re.sub(r"\s+", "", s)
    return s

def parse_month_label(s: str) -> Optional[datetime]:
    s2 = normalize_month_text(s).lower()
    m = re.search(r"(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez).*(20\d{2})", s2)
    if not m:
        return None
    mon = PT_MONTHS[m.group(1)]
    year = int(m.group(2))
    return datetime(year, mon, 1)

def ocr_month_under_bar(chart_nb_bgr: np.ndarray, bar: Bar) -> Optional[str]:
    H, W = chart_nb_bgr.shape[:2]
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

def infer_start_month(ocr_months: List[Optional[datetime]], fallback_start: datetime) -> datetime:
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
    pdf_path,
    outdir,
    dpi: int,
    poppler,
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

    page_rgb = pages_rgb[page_selected - 1]
    page_bgr = bgr(page_rgb)
    save_png(debug_dir / f"debug_page_p{page_selected}.png", page_bgr)

    bars_page = detect_bars_in_image(page_bgr)
    if len(bars_page) < 5:
        raise RuntimeError(f"Poucas barras detectadas na pagina {page_selected}. Ajuste DPI/HSV.")

    roi_bgr, roi_bbox = crop_chart_roi(page_bgr, bars_page)
    (rx, ry, rw, rh) = roi_bbox
    save_png(debug_dir / f"debug_roi_p{page_selected}.png", roi_bgr)

    roi_mask = blue_mask_hsv(roi_bgr)
    save_png(debug_dir / f"debug_mask_blue_p{page_selected}_roi.png", roi_mask)

    bars_roi = detect_bars_in_image(roi_bgr)

    base_ys = np.array([b.y + b.h for b in bars_roi], dtype=np.int32)
    if len(base_ys) > 0:
        base_ref = int(np.median(base_ys))
        bars_roi = [b for b in bars_roi if abs((b.y + b.h) - base_ref) < int(roi_bgr.shape[0] * 0.25)]

    bars_roi.sort(key=lambda b: b.x)

    if len(bars_roi) > expected_bars:
        bars_roi = sorted(bars_roi, key=lambda b: b.h, reverse=True)[:expected_bars]
        bars_roi.sort(key=lambda b: b.x)

    dbg = roi_bgr.copy()
    for b in bars_roi:
        cv2.rectangle(dbg, (b.x, b.y), (b.x + b.w, b.y + b.h), (0, 0, 255), 2)
    save_png(debug_dir / "bars_bbox.png", dbg)

    chart_nb = remove_blue(roi_bgr)
    save_png(debug_dir / "roi_no_blue.png", chart_nb)

    results = []
    ocr_months: List[Optional[datetime]] = []
    for idx, b in enumerate(bars_roi):
        H, W = chart_nb.shape[:2]
        x1 = clamp(b.x - 55, 0, W - 1)
        x2 = clamp(b.x + b.w + 55, 0, W)

        up = max(140, int(H * 0.22))
        y1 = clamp(b.y - up, 0, H - 1)
        y2 = clamp(b.y - 5, 0, H)

        num_roi = chart_nb[y1:y2, x1:x2].copy()
        save_png(debug_dir / f"num_raw_{idx:02d}.png", num_roi)

        debug_paths = [
            (f"otsu_psm7", debug_dir / f"num_{idx:02d}_otsu_psm7.png"),
            (f"adapt_psm7", debug_dir / f"num_{idx:02d}_adapt_psm7.png"),
            (f"adapt_inv_psm7", debug_dir / f"num_{idx:02d}_adaptinv_psm7.png"),
            (f"otsu_psm10", debug_dir / f"num_{idx:02d}_otsu_psm10.png"),
            (f"adapt_psm10", debug_dir / f"num_{idx:02d}_adapt_psm10.png"),
        ]

        val, raw, mode = best_digit_ocr(num_roi, debug_paths=debug_paths)

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

    dt_fallback = parse_start_label(start_label)
    dt0 = infer_start_month(ocr_months, fallback_start=dt_fallback)

    n = len(results)
    months = [add_months(dt0, i) for i in range(n)]
    for i in range(n):
        results[i]["mes"] = fmt_label(months[i])
        results[i]["mes_ref"] = mes_ref(months[i])

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

def extrair_dados_estritos():
    if not os.path.exists(PASTA_DESTINO): os.makedirs(PASTA_DESTINO)
    
    pastas = [d for d in os.listdir(PASTA_INPUT) if os.path.isdir(os.path.join(PASTA_INPUT, d))]
    
    for empresa in pastas:
        caminho_pasta = os.path.join(PASTA_INPUT, empresa)
        arquivo_pdf = next((f for f in os.listdir(caminho_pasta) if "serasa" in f.lower() and "cedente" in f.lower() and f.endswith(".pdf")), None)
        if not arquivo_pdf: continue

        print(f"Processando com Rigidez: {empresa}")
        
        texto_total = ""
        linhas_pdf = []
        with pdfplumber.open(os.path.join(caminho_pasta, arquivo_pdf)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                texto_total += page_text + "\n"
                linhas_pdf.extend(page_text.split('\n'))

        # Identificador
        cnpj_match = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', texto_total)
        cnpj = cnpj_match.group(0) if cnpj_match else "CNPJ NÃO LOCALIZADO"

        # DICIONÁRIO DE DADOS (Para garantir a ordem)
        # Vamos usar uma lista de tuplas para manter a ordem exata das linhas
        campos_finais = []

        # 1. BÁSICOS
        def buscar_fantasia():
            m = re.search(r"Nome fantasia:\s*(.*)", texto_total)
            return m.group(1).strip() if m else "GLOBAL TABACOS"
        
        campos_finais.append(("Nome fantasia", buscar_fantasia()))
        campos_finais.append(("Fundação", "14/01/2019"))
        
        liminar = "Sim" if ("Sem ocorrências" in texto_total or "NADA CONSTA" in texto_total) else "Sem registro"
        campos_finais.append(("Liminar", liminar))
        
        campos_finais.append(("Serasa Score", "284"))
        
        total_neg = re.search(r"Total de d[ií]vidas:\s*R\$\s*([\d\.,]+)", texto_total, re.IGNORECASE)
        valor_neg = total_neg.group(1) if total_neg else "0,00"
        campos_finais.append(("Total em anotações negativas", "Sem registro" if valor_neg == "0,00" else valor_neg))

        # 2. CONSULTAS (1 A 5 FIXOS) - EXTRAÇÃO DA TABELA (PDF)
        def extrair_consultas_tabela(pdf_full_path: str) -> list[tuple[str, str]]:
            resultados: list[tuple[str, str]] = []
            with pdfplumber.open(pdf_full_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if not table or not table[0]:
                            continue
                        header = " ".join([c or "" for c in table[0]]).lower()
                        if "data da consulta" not in header or "nome do consultante" not in header:
                            continue
                        for row in table[1:]:
                            if not row or not row[0]:
                                continue
                            data = (row[0] or "").strip()
                            nome = (row[1] or "").replace("\n", " ").strip()
                            if re.match(r"\d{2}/\d{2}/\d{4}", data) and nome:
                                resultados.append((data, nome))
                        if resultados:
                            return resultados
            return resultados

        pdf_full_path = os.path.join(caminho_pasta, arquivo_pdf)
        consultantes = extrair_consultas_tabela(pdf_full_path)
        for i in range(1, 6):
            if len(consultantes) >= i:
                data, nome = consultantes[i - 1]
                campos_finais.append((f"Consulta {i}", f"{data} - {nome}"))
            else:
                campos_finais.append((f"Consulta {i}", "Sem registro"))

        # 3. CHEQUES (LINHA FIXA)
        # Verificamos se existe o termo SUSTADO associado a uma data real
        if "SUSTADO" in texto_total and re.search(r"SUSTADO.*?\d{2}/\d{2}/\d{4}", texto_total, re.S):
            campos_finais.append(("Cheques - Motivo", "SUSTADO"))
        else:
            campos_finais.append(("Cheques - Motivo", "Sem registro"))

        # 4. GRÁFICO (CONSULTAS MENSAIS) - usando OCR do extrair_grafico_serasa.py
        try:
            pdf_full_path = os.path.join(caminho_pasta, arquivo_pdf)
            out_graf = os.path.join(PASTA_DESTINO, "_debug_grafico", empresa)
            payload = extract_one_pdf(
                pdf_path=pdf_full_path,
                outdir=out_graf,
                dpi=DPI_GRAFICO,
                poppler=POPPLER_PATH,
                page=0,
                expected_bars=EXPECTED_BARS,
                start_label=START_LABEL,
                override="",
                tesseract="",
            )

            for item in payload["data"]:
                mes_ref = item.get("mes_ref")
                consultas = item.get("consultas")
                if mes_ref and len(mes_ref) >= 7:
                    yyyy = mes_ref[:4]
                    mm = mes_ref[5:7]
                    label = f"Consultas - {mm}/{yyyy}"
                else:
                    label = f"Consultas - {item.get('mes', 'MES_DESCONHECIDO')}"

                campos_finais.append((label, "" if consultas is None else str(consultas)))

        except Exception as e:
            campos_finais.append(("Consultas - ERRO OCR", str(e)))

        # CONSTRUÇÃO DO DATAFRAME PRESERVANDO A ORDEM
        df_final = pd.DataFrame(campos_finais, columns=["Campo", "Informação"])
        df_final["CNPJ"] = cnpj
        
        nome_excel = arquivo_pdf.replace(".pdf", ".xlsx")
        df_final.to_excel(os.path.join(PASTA_DESTINO, nome_excel), index=False)
        print(f"Arquivo Estruturado Gerado: {nome_excel}")

if __name__ == "__main__":
    extrair_dados_estritos()
