import argparse
import pdfplumber
import pandas as pd
import os
import re
import unicodedata
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pdf2image import convert_from_path
import pytesseract

# -----------------------------
# Utils
# -----------------------------
# --- CONFIGURAÇÃO DE PASTAS ---
# BASE_DIR é C:\Users\Usuario\Desktop\DADOS_BI
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA_INPUT = os.path.join(BASE_DIR, "01_INPUT")
# Salva direto na sua pasta "1. Relatório de Visita" dentro da 03_OUTPUT
PASTA_DESTINO = os.path.join(BASE_DIR, "03_OUTPUT", "1. Relatório de Visita")
POPPLER_PATH = r"C:\Users\Usuario\Desktop\poppler-25.12.0\Library\bin"

UF_SET = {
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG",
    "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO",
}

def formatar_moeda(v):
    if not v:
        return "R$ 0,00"
    s = re.sub(r"[R\$\s]", "", str(v))
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "." in s and "," not in s:
        if re.fullmatch(r"\d{1,3}(?:\.\d{3})+\.\d{2}", s):
            parts = s.split(".")
            s = "".join(parts[:-1]) + "." + parts[-1]
        else:
            s = s.replace(".", "")
    else:
        s = s.replace(".", "").replace(",", ".")
    try:
        return f"{float(s):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "a preencher"

def extrair_campo_gpo(texto, rotulo):
    padrao = rf"{rotulo}[:\s]*(.*?)(?=\n\s*[A-Z][a-zç]+:|\n\n|$)"
    match = re.search(padrao, texto, re.IGNORECASE | re.DOTALL)
    if match:
        res = match.group(1).strip()
        return " ".join(res.split())
    return "a preencher"


def _read_pdf_text(path_file: str) -> str:
    with pdfplumber.open(path_file) as pdf:
        texto = "\n".join([p.extract_text() or "" for p in pdf.pages])
    if len(texto.strip()) >= 40:
        return texto
    try:
        images = convert_from_path(path_file, dpi=300, poppler_path=POPPLER_PATH, fmt="png")
    except Exception:
        return texto
    ocr_parts = []
    for img in images:
        txt = pytesseract.image_to_string(img, lang="por+eng", config="--oem 3 --psm 6") or ""
        if txt:
            ocr_parts.append(txt)
    ocr_text = "\n".join(ocr_parts)
    return ocr_text if len(ocr_text.strip()) > len(texto.strip()) else texto


def _find_date_after_labels(texto: str, labels: list[str]) -> str:
    if not texto:
        return ""
    lines = [l.strip() for l in texto.splitlines() if l.strip()]
    for lab in labels:
        m = re.search(rf"{re.escape(lab)}\s*[:\-]?\s*(\d{{2}}/\d{{2}}/\d{{4}})", texto, flags=re.I)
        if m:
            return m.group(1)
        for i, line in enumerate(lines):
            if re.search(re.escape(lab), line, flags=re.I):
                m2 = re.search(r"\b\d{2}/\d{2}/\d{4}\b", line)
                if m2:
                    return m2.group(0)
                for j in range(i + 1, min(i + 3, len(lines))):
                    m3 = re.search(r"\b\d{2}/\d{2}/\d{4}\b", lines[j])
                    if m3:
                        return m3.group(0)
    return ""


def _extract_localizacao(texto: str) -> str:
    if not texto:
        return ""
    lines = [l.strip() for l in texto.splitlines() if l.strip()]
    for i, line in enumerate(lines):
        up = unicodedata.normalize("NFD", line).encode("ascii", "ignore").decode("ascii").upper()
        if "MUNICIPIO" in up and "UF" in up:
            m = re.search(r"([A-ZÁÉÍÓÚÂÊÔÃÕÇ ]+)\s*/\s*([A-Z]{2})", line, flags=re.I)
            if m and m.group(2).strip().upper() in UF_SET:
                return f"{m.group(1).strip()}/{m.group(2).strip()}"
            for j in range(i + 1, min(i + 4, len(lines))):
                m = re.search(r"([A-ZÁÉÍÓÚÂÊÔÃÕÇ ]+)\s*/\s*([A-Z]{2})", lines[j], flags=re.I)
                if m:
                    return f"{m.group(1).strip()}/{m.group(2).strip()}"
                tokens = lines[j].split()
                if tokens and tokens[-1] in UF_SET:
                    m2 = re.findall(r"\b([A-ZÁÉÍÓÚÂÊÔÃÕÇ]+)\s+([A-Z]{2})\b", lines[j])
                    if m2:
                        city, uf = m2[-1]
                        return f"{city}/{uf}"
                    return f"{tokens[-2]}/{tokens[-1]}" if len(tokens) >= 2 else ""
    m = re.findall(r"\b([A-ZÁÉÍÓÚÂÊÔÃÕÇ]+)\s+([A-Z]{2})\b", texto)
    if m:
        for city, uf in reversed(m):
            if uf.strip().upper() in UF_SET:
                return f"{city.strip()}/{uf.strip()}"
    return ""

def _month_token_to_num(token):
    if not token:
        return None
    s = unicodedata.normalize("NFD", token)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"[^A-Za-z]", "", s).lower()
    if len(s) >= 3:
        s = s[:3]
    meses = {"jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
             "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12}
    return meses.get(s)

def _parse_faturamento(texto):
    fat_map = {}
    if not texto:
        return fat_map

    padrao_nome = re.compile(r"([A-Za-z]{3,})\s*/\s*(\d{4})\s+R?\$?\s*([\d\.,]+)")
    padrao_num = re.compile(r"(\d{2})\s*/\s*(\d{4})\s+R?\$?\s*([\d\.,]+)")

    for m, a, v in padrao_nome.findall(texto):
        mon = _month_token_to_num(m)
        if not mon:
            continue
        try:
            dt = datetime(int(a), mon, 1)
        except ValueError:
            continue
        fat_map[dt] = formatar_moeda(v)

    for m, a, v in padrao_num.findall(texto):
        try:
            mon = int(m)
            dt = datetime(int(a), mon, 1)
        except ValueError:
            continue
        fat_map[dt] = formatar_moeda(v)

    return fat_map


def _unique_keep(seq):
    seen = set()
    out = []
    for item in seq:
        if not item or item in seen:
            continue
        seen.add(item)
        out.append(item)
    return out


def _extract_doc_id(pdf_path, prefer_cpf=False):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join([p.extract_text() or "" for p in pdf.pages])
    except Exception:
        return ""
    cnpj = re.search(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b", text)
    cpf = re.search(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b", text)
    if prefer_cpf:
        if cpf:
            return cpf.group(0)
        if cnpj:
            return cnpj.group(0)
    else:
        if cnpj:
            return cnpj.group(0)
        if cpf:
            return cpf.group(0)
    return ""

    cnpj = re.search(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b", text)
    if cnpj:
        return cnpj.group(0)
    cpf = re.search(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b", text)
    if cpf:
        return cpf.group(0)
    return ""

def _resolve_empresas(input_dir: str):
    subdirs = [d for d in os.listdir(input_dir) if os.path.isdir(os.path.join(input_dir, d))]
    if subdirs:
        return [(d, os.path.join(input_dir, d)) for d in subdirs]
    return [(os.path.basename(input_dir), input_dir)]


def processar_bi_agente_1(input_dir: str, out_dir: str):
    if not os.path.exists(out_dir):
        os.makedirs(out_dir)

    empresas = _resolve_empresas(input_dir)

    for empresa, caminho_empresa in empresas:
        print(f"Lendo arquivos da {empresa}...")
        texto_gpo, texto_contabil, texto_gerencial, texto_cartao = "", "", "", ""
        
        for f in os.listdir(caminho_empresa):
            path_file = os.path.join(caminho_empresa, f)
            try:
                nome = f.lower()
                ext = os.path.splitext(nome)[1]
                is_pdf = nome.endswith('.pdf')
                is_text = ext in {".txt", ".csv"}
                if not (is_pdf or is_text):
                    continue

                if is_pdf:
                    texto = _read_pdf_text(path_file)
                else:
                    with open(path_file, 'r', encoding='utf-8') as file:
                        texto = file.read()

                is_contabil = "contabil" in nome or "cont bil" in nome
                is_gerencial = "gerencial" in nome
                is_gpo = "gpo" in nome or "relatorio" in nome or "visita" in nome
                is_cartao = "cnpj" in nome and ("cartao" in nome or "cartão" in nome or "comprovante" in nome)

                if is_contabil:
                    texto_contabil += texto
                elif is_gerencial:
                    texto_gerencial += texto
                elif is_gpo:
                    texto_gpo += texto
                elif is_cartao:
                    texto_cartao += texto
                else:
                    continue
            except Exception as e:
                print(f"Erro ao ler {f}: {e}")

        # --- EXTRAÇÃO (REGRAS V6) ---
        cnpj_match = re.search(r'\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}', texto_contabil)
        cnpj_final = cnpj_match.group(0) if cnpj_match else "32.434.675/0001-35"

        linhas = []
        def add(c, i): linhas.append({"Campo": c, "Informação": i, "CNPJ": cnpj_final})

        mapa = {
            "Atividade Economica": "Atividade Econômica",
            "Tipo de Sacados": "Tipo de Sacados",
            "Perfil de Sacados": "Perfil de Sacados",
            "Tipo de Faturamento": "Tipo de Faturamento",
            "Faturamento Descontável": "Faturamento Descontável",
            "Tipo de Desconto": "Tipo de Desconto",
            "Região de Vendas (UFs)": "Região de Vendas",
            "Ticket Médio": "Ticket Médio",
            "Prazos de Vendas": "Prazos de Vendas",
            "Performance": "Performance",
            "Lastro": "Lastro",
            "Transportes": "Transportes",
            "Prazos de Entrega": "Prazos de Entrega",
            "Limite solicitado": "Limite Solicitado",
        }

        for c in ["Data Comitê", "AGENTE COMERCIAL", "G.P.O", "Origem"]: add(c, "a preencher")
        for c, busca in mapa.items():
            add(c, extrair_campo_gpo(texto_gpo, busca))
        for c in ["Rede Social", "Site", "Imagem"]: add(c, "a preencher")

        ctx = re.search(r"Contexto:\s*(.*?)(?=\nsócio|CPF|$)", texto_gpo, re.S | re.I)
        add("Contexto", " ".join(ctx.group(1).split()).strip() if ctx else "a preencher")

        data_abertura = _find_date_after_labels(texto_cartao, ["Data de abertura"])
        data_atividade = _find_date_after_labels(
            texto_cartao,
            ["Data de início de atividade", "Data de inicio de atividade", "Data da situação cadastral", "Data da situacao cadastral"],
        )
        localizacao = _extract_localizacao(texto_cartao)
        add("Tempo de abertura", data_abertura or "a preencher")
        add("Tempo de atividade", data_atividade or data_abertura or "a preencher")
        add("Localização", localizacao or "a preencher")

        fat_map_contabil = _parse_faturamento(texto_contabil)
        fat_map_gerencial = _parse_faturamento(texto_gerencial)

        todas_chaves = list(fat_map_contabil.keys()) + list(fat_map_gerencial.keys())
        ultimo_mes = max(todas_chaves) if todas_chaves else datetime.now()
        janela = [ultimo_mes - relativedelta(months=i) for i in range(11, -1, -1)]

        for dt in janela:
            add(f"Faturamento CONTÁBIL - {dt.strftime('01/%m/%Y')}", fat_map_contabil.get(dt, "A confirmar"))
        for dt in janela:
            add(f"Faturamento GERENCIAL - {dt.strftime('01/%m/%Y')}", fat_map_gerencial.get(dt, "R$ 0,00"))

        files = os.listdir(caminho_empresa)
        serasa_socio_pdfs = sorted([
            os.path.join(caminho_empresa, f)
            for f in files
            if f.lower().endswith('.pdf') and 'serasa' in f.lower() and 'socio' in f.lower()
        ])
        socio_ids = [_extract_doc_id(p, prefer_cpf=True) for p in serasa_socio_pdfs]
        socio_ids = _unique_keep([x for x in socio_ids if x])
        if not socio_ids:
            socio_ids = _unique_keep(re.findall(r'(\d{3}\.?\d{3}\.?\d{3}-?\d{2})', texto_gpo))
        for i, socio_id in enumerate(socio_ids, start=1):
            add(f"Sócio {i}", socio_id)

        serasa_sacado_pdfs = sorted([
            os.path.join(caminho_empresa, f)
            for f in files
            if f.lower().endswith('.pdf') and 'serasa' in f.lower() and 'sacado' in f.lower()
        ])
        sacado_ids = [_extract_doc_id(p) for p in serasa_sacado_pdfs]
        sacado_ids = _unique_keep([x for x in sacado_ids if x])

        add("Demais empresas do sócio 1", "a preencher")
        add("Cônjuge", "a preencher")
        add("Link Curva ABC", "a preencher")
        add("Curva ABC", "a preencher")
        for i in range(1, 6):
            add(f"CURVA ABC - SACADO {i}", sacado_ids[i-1] if len(sacado_ids) >= i else "a preencher")
        add("Link Imposto de renda", "a preencher")
        add("Nome do responsável", "a preencher")
        add("Função do responsável", "a preencher")

        # --- SALVA NA PASTA 1 ---
        nome_arquivo = f"{empresa}.xlsx"
        df_final = pd.DataFrame(linhas)
        df_final.to_excel(os.path.join(out_dir, nome_arquivo), index=False)
        print(f"Planilha da {empresa} gerada na pasta '1. Relatório de Visita'")

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", default=PASTA_INPUT, help="Pasta base 01_INPUT (ou pasta da empresa).")
    ap.add_argument("--outdir", default=PASTA_DESTINO, help="Pasta de saida.")
    args = ap.parse_args()
    processar_bi_agente_1(args.input, args.outdir)
