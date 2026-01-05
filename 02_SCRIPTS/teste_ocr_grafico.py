import os
import re
import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path
import pandas as pd

# ========= CONFIG =========
PDF_PATH = r"C:\Users\Usuario\Desktop\DADOS_BI\01_INPUT\Global_Tabacos\serasa_cedente_global.pdf"
POPPLER_PATH = r"C:\Users\Usuario\Desktop\poppler-25.12.0\Library\bin"
DPI = 200
OUT_DIR = r"C:\Users\Usuario\Desktop\DADOS_BI\03_OUTPUT\debug_grafico_final"

# Se precisar (seu tesseract estiver fora do PATH), descomente e ajuste:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


# ========= HELPERS =========
def ensure_dir(p):
    os.makedirs(p, exist_ok=True)

def blue_mask(bgr):
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
    # Faixa de azul típica do gráfico
    lower = np.array([90, 40, 40])
    upper = np.array([140, 255, 255])
    return cv2.inRange(hsv, lower, upper)

def pick_page_with_most_blue(pages_bgr):
    scores = []
    for i, bgr in enumerate(pages_bgr, start=1):
        m = blue_mask(bgr)
        scores.append((int(m.sum()), i))
    scores.sort(reverse=True)
    return scores[0][1], scores

def find_chart_roi(page_bgr, debug_dir):
    """
    Acha uma ROI aproximada do gráfico olhando onde está o 'miolo' de azul.
    A gente ignora topo (títulos/links) e pega a região mais provável do gráfico.
    """
    h, w = page_bgr.shape[:2]
    m = blue_mask(page_bgr)

    # ignora faixa do topo (muita coisa azul fora do gráfico)
    top_cut = int(h * 0.25)
    m2 = m.copy()
    m2[:top_cut, :] = 0

    ys, xs = np.where(m2 > 0)
    if len(xs) == 0:
        return None

    x1, x2 = xs.min(), xs.max()
    y1, y2 = ys.min(), ys.max()

    # dá uma folga para pegar números e meses
    pad_x = int(0.05 * w)
    pad_top = int(0.08 * h)
    pad_bottom = int(0.12 * h)

    rx1 = max(x1 - pad_x, 0)
    rx2 = min(x2 + pad_x, w)
    ry1 = max(y1 - pad_top, 0)
    ry2 = min(y2 + pad_bottom, h)

    roi = page_bgr[ry1:ry2, rx1:rx2].copy()

    cv2.imwrite(os.path.join(debug_dir, "debug_roi.png"), roi)
    return (rx1, ry1, rx2, ry2)

def segments_from_projection(mask, min_width=18):
    """
    Converte o 'perfil' em X da máscara azul em segmentos contínuos = barras.
    Isso NÃO perde barra pequena (tipo Mai/2025).
    """
    col = mask.sum(axis=0).astype(np.float32)
    if col.max() <= 0:
        return []
    col = col / col.max()

    # threshold suave: qualquer coluna com um pouco de azul entra
    on = col > 0.08

    segs = []
    i = 0
    n = len(on)
    while i < n:
        if not on[i]:
            i += 1
            continue
        j = i
        while j < n and on[j]:
            j += 1
        if (j - i) >= min_width:
            segs.append((i, j))  # [i, j)
        i = j
    return segs

def bar_bbox_from_segment(mask, x1, x2):
    """
    Dado um segmento em X, acha o y top/bottom onde tem azul.
    """
    slice_ = mask[:, x1:x2]
    ys, xs = np.where(slice_ > 0)
    if len(xs) == 0:
        return None
    y1 = ys.min()
    y2 = ys.max()
    return (x1, y1, x2 - x1, y2 - y1)

def sane_num(n):
    if n is None:
        return None
    # Correção comum: "30" ao invés de "3", "130" ao invés de "13"
    if n >= 20 and n % 10 == 0 and (n // 10) <= 50:
        n = n // 10

    # rejeita lixo
    if n < 0 or n > 50:
        return None
    return n

def tesseract_digits(img_bin, psm):
    cfg = f"--oem 3 --psm {psm} -c tessedit_char_whitelist=0123456789"
    txt = pytesseract.image_to_string(img_bin, config=cfg) or ""
    txt = re.sub(r"\D", "", txt)
    if txt == "":
        return None
    try:
        return int(txt)
    except:
        return None

def ocr_number_above_bar(roi_bgr, bar_bbox, debug_dir, idx):
    """
    Lê o número acima da barra:
    - recorta uma faixa acima do topo da barra
    - binariza
    - usa contornos pra recortar SÓ o texto
    - OCR
    """
    x, y, w, h = bar_bbox

    # faixa acima do topo da barra (ajuste fino se quiser)
    y1 = max(y - 90, 0)
    y2 = max(y - 5, 0)
    x1 = max(x - 12, 0)
    x2 = min(x + w + 12, roi_bgr.shape[1])

    crop = roi_bgr[y1:y2, x1:x2].copy()
    if crop.size == 0:
        return None

    gray = cv2.cvtColor(crop, cv2.COLOR_BGR2GRAY)
    gray = cv2.resize(gray, None, fx=5, fy=5, interpolation=cv2.INTER_CUBIC)
    gray = cv2.copyMakeBorder(gray, 25, 25, 25, 25, cv2.BORDER_CONSTANT, value=255)

    # 1) OTSU
    _, th = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # garantir “dígito escuro em fundo claro”
    # se ficou invertido, inverte
    if th.mean() < 127:
        th = 255 - th

    # contornos do que é "escuro" (dígitos)
    inv = 255 - th
    cnts, _ = cv2.findContours(inv, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # filtra contornos plausíveis
    boxes = []
    for c in cnts:
        xx, yy, ww, hh = cv2.boundingRect(c)
        area = ww * hh
        if area < 80:
            continue
        if hh < 12 or ww < 6:
            continue
        boxes.append((xx, yy, ww, hh))

    if boxes:
        # une tudo (bom pra "10", "12", "13")
        xs = [b[0] for b in boxes]
        ys = [b[1] for b in boxes]
        xe = [b[0] + b[2] for b in boxes]
        ye = [b[1] + b[3] for b in boxes]
        bx1 = max(min(xs) - 10, 0)
        by1 = max(min(ys) - 10, 0)
        bx2 = min(max(xe) + 10, th.shape[1])
        by2 = min(max(ye) + 10, th.shape[0])
        th2 = th[by1:by2, bx1:bx2]
    else:
        th2 = th

    cv2.imwrite(os.path.join(debug_dir, f"num_{idx:02d}.png"), th2)

    # tenta OCR (psm 7 lê 1 ou 2 dígitos bem)
    n = tesseract_digits(th2, psm=7)
    n = sane_num(n)
    if n is not None:
        return n

    # 2) fallback: adaptive + OCR de novo
    ad = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 7
    )
    if ad.mean() < 127:
        ad = 255 - ad
    cv2.imwrite(os.path.join(debug_dir, f"num_{idx:02d}_adapt.png"), ad)

    n = tesseract_digits(ad, psm=7)
    return sane_num(n)

def meses_13():
    # seu gráfico é fixo: Nov/2024 ... Nov/2025 (13 meses)
    return ["Nov/2024","Dez/2024","Jan/2025","Fev/2025","Mar/2025","Abr/2025",
            "Mai/2025","Jun/2025","Jul/2025","Ago/2025","Set/2025","Out/2025","Nov/2025"]


# ========= MAIN =========
def extrair_grafico(pdf_path):
    ensure_dir(OUT_DIR)

    pages = convert_from_path(pdf_path, dpi=DPI, poppler_path=POPPLER_PATH)
    pages_bgr = [cv2.cvtColor(np.array(p), cv2.COLOR_RGB2BGR) for p in pages]

    pagina_escolhida, scores = pick_page_with_most_blue(pages_bgr)
    print("Páginas (score azul):", scores)
    print("Página escolhida:", pagina_escolhida)

    page_bgr = pages_bgr[pagina_escolhida - 1]

    roi_bbox = find_chart_roi(page_bgr, OUT_DIR)
    if roi_bbox is None:
        raise RuntimeError("Não encontrei ROI do gráfico (sem azul detectável).")

    rx1, ry1, rx2, ry2 = roi_bbox
    roi = page_bgr[ry1:ry2, rx1:rx2].copy()

    # máscara azul na ROI
    m = blue_mask(roi)

    # dilata para unir possíveis falhas e NÃO perder barra pequena
    k = np.ones((5, 5), np.uint8)
    m = cv2.dilate(m, k, iterations=1)

    cv2.imwrite(os.path.join(OUT_DIR, "mask_blue.png"), m)

    # pega somente a faixa vertical onde as barras estão (evita capturar “azuis” do resto)
    # ajuste fino: geralmente barras ficam na metade superior da ROI
    h, w = m.shape[:2]
    y_top = int(h * 0.10)
    y_bot = int(h * 0.72)
    m_bars = np.zeros_like(m)
    m_bars[y_top:y_bot, :] = m[y_top:y_bot, :]

    cv2.imwrite(os.path.join(OUT_DIR, "mask_blue_bars.png"), m_bars)

    segs = segments_from_projection(m_bars, min_width=18)
    bar_bboxes = []
    for (sx1, sx2) in segs:
        bb = bar_bbox_from_segment(m_bars, sx1, sx2)
        if bb is None:
            continue
        x, y, bw, bh = bb
        # filtros leves: largura típica de barra
        if bw < 25 or bw > 160:
            continue
        if bh < 20:  # barra muito pequena ainda conta
            continue
        bar_bboxes.append(bb)

    bar_bboxes.sort(key=lambda t: t[0])
    print("Barras detectadas:", len(bar_bboxes))

    # desenha debug das barras
    dbg = roi.copy()
    for i, (x, y, bw, bh) in enumerate(bar_bboxes):
        cv2.rectangle(dbg, (x, y), (x + bw, y + bh), (0, 0, 255), 2)
        cv2.putText(dbg, str(i), (x, max(y-10, 0)), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,0,255), 2)
    cv2.imwrite(os.path.join(OUT_DIR, "bars_bbox.png"), dbg)

    # meses
    meses = meses_13()
    if len(bar_bboxes) != len(meses):
        print("⚠️ Atenção: barras != 13. Vou mapear pelo que tiver, mas o ideal é bater 13.")

    # OCR números
    out = []
    for i, bb in enumerate(bar_bboxes[:len(meses)]):
        n = ocr_number_above_bar(roi, bb, OUT_DIR, i)
        out.append({"ordem": i, "mes": meses[i], "consultas": n, "bar_bbox_roi": list(bb), "roi_bbox_page": [rx1, ry1, rx2-rx1, ry2-ry1]})

    df = pd.DataFrame(out)
    df.to_csv(os.path.join(OUT_DIR, "consultas_grafico.csv"), index=False, encoding="utf-8-sig")
    with open(os.path.join(OUT_DIR, "consultas_grafico_full.json"), "w", encoding="utf-8") as f:
        import json
        json.dump(out, f, ensure_ascii=False, indent=2)

    print("✅ Salvei:")
    print(" -", os.path.join(OUT_DIR, "consultas_grafico_full.json"))
    print(" -", os.path.join(OUT_DIR, "consultas_grafico.csv"))
    print("📁 Debugs em:", OUT_DIR)

    return out


if __name__ == "__main__":
    extrair_grafico(PDF_PATH)

























