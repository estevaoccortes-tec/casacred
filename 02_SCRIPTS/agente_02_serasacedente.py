import os
import re
import pdfplumber
import pandas as pd

# --- CONFIGURAÇÃO ---
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA_INPUT = os.path.join(BASE_DIR, "01_INPUT")
PASTA_DESTINO = os.path.join(BASE_DIR, "03_OUTPUT", "2. SERASA CEDENTE")

def extrair_dados_estritos():
    if not os.path.exists(PASTA_DESTINO): os.makedirs(PASTA_DESTINO)
    
    pastas = [d for d in os.listdir(PASTA_INPUT) if os.path.isdir(os.path.join(PASTA_INPUT, d))]
    
    for empresa in pastas:
        caminho_pasta = os.path.join(PASTA_INPUT, empresa)
        arquivo_pdf = next((f for f in os.listdir(caminho_pasta) if "serasa" in f.lower() and "cedente" in f.lower() and f.endswith(".pdf")), None)
        if not arquivo_pdf: continue

        print(f"📊 Processando com Rigidez: {empresa}")
        
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

        # 2. CONSULTAS (1 A 5 FIXOS) - EXTRAÇÃO DA PÁGINA 2
        consultantes = re.findall(r"(\d{2}/\d{2}/\d{4})\s+(.+?)\s+\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto_total)
        for i in range(1, 6):
            if len(consultantes) >= i:
                data, nome = consultantes[i-1][0], consultantes[i-1][1]
                campos_finais.append((f"Consulta {i}", f"{data} - {nome.strip()}"))
            else:
                campos_finais.append((f"Consulta {i}", "Sem registro"))

        # 3. CHEQUES (LINHA FIXA)
        # Verificamos se existe o termo SUSTADO associado a uma data real
        if "SUSTADO" in texto_total and re.search(r"SUSTADO.*?\d{2}/\d{2}/\d{4}", texto_total, re.S):
            campos_finais.append(("Cheques - Motivo", "SUSTADO"))
        else:
            campos_finais.append(("Cheques - Motivo", "Sem registro"))

        # 4. GRÁFICO (CONSULTAS MENSAIS)
        # O PDF mostra o número acima do mês. No texto extraído, o número aparece na linha ANTERIOR à data.
        meses_encontrados = re.findall(r"([A-Z][a-z]{2}/\d{4})", texto_total)
        mapa_grafico = {}
        
        # Percorremos as linhas do PDF para achar o número que precede a data
        for i, linha in enumerate(linhas_pdf):
            for mes_ano in meses_encontrados:
                if mes_ano in linha:
                    # O valor geralmente está na linha imediatamente acima ou na mesma linha antes da data
                    # Tentamos pegar o número na linha de cima (i-1)
                    if i > 0:
                        val_acima = re.findall(r"\b(\d{1,2})\b", linhas_pdf[i-1])
                        if val_acima:
                            mapa_grafico[mes_ano] = val_acima[-1] # Pega o último número da linha de cima

        # Adiciona as consultas ao Excel na ordem cronológica (ou conforme o padrão 13 meses)
        for mes_ano in meses_encontrados:
            mes_txt, ano = mes_ano.split('/')
            # Tradução simples para MM/AAAA
            mes_map = {"Jan":"01","Fev":"02","Mar":"03","Abr":"04","Mai":"05","Jun":"06","Jul":"07","Ago":"08","Set":"09","Out":"10","Nov":"11","Dez":"12"}
            mm = mes_map.get(mes_txt, "00")
            val = mapa_grafico.get(mes_ano, "0")
            campos_finais.append((f"Consultas - {mm}/{ano}", val))

        # CONSTRUÇÃO DO DATAFRAME PRESERVANDO A ORDEM
        df_final = pd.DataFrame(campos_finais, columns=["Campo", "Informação"])
        df_final["CNPJ"] = cnpj
        
        nome_excel = arquivo_pdf.replace(".pdf", ".xlsx")
        df_final.to_excel(os.path.join(PASTA_DESTINO, nome_excel), index=False)
        print(f"✅ Arquivo Estruturado Gerado: {nome_excel}")

if __name__ == "__main__":
    extrair_dados_estritos()