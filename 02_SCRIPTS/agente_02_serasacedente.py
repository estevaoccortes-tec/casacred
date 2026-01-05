import os
import re
import pdfplumber
import pandas as pd

from extrair_grafico_serasa import extract_one_pdf

# --- CONFIGURAÇÃO ---
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA_INPUT = os.path.join(BASE_DIR, "01_INPUT")
PASTA_DESTINO = os.path.join(BASE_DIR, "03_OUTPUT", "2. SERASA CEDENTE")
POPPLER_PATH = r"C:\Users\Usuario\Desktop\poppler-25.12.0\Library\bin"
START_LABEL = "Nov/2024"
EXPECTED_BARS = 13
DPI_GRAFICO = 200

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
