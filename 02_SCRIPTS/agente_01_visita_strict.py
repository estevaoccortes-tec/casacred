import pdfplumber
import pandas as pd
import os
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta

# --- CONFIGURAÇÃO DE PASTAS ---
# BASE_DIR é C:\Users\Usuario\Desktop\DADOS_BI
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PASTA_INPUT = os.path.join(BASE_DIR, "01_INPUT")
# Salva direto na sua pasta "1. Relatório de Visita" dentro da 03_OUTPUT
PASTA_DESTINO = os.path.join(BASE_DIR, "03_OUTPUT", "1. Relatório de Visita")

def formatar_moeda(v):
    if not v: return "R$ 0,00"
    limpo = re.sub(r'[R\$\s\.]', '', str(v)).replace(',', '.')
    try:
        return f"{float(limpo):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except: return "a preencher"

def extrair_campo_gpo(texto, rotulo):
    padrao = rf"{rotulo}[:\s]*(.*?)(?=\n\s*[A-Z][a-zç]+:|\n\n|$)"
    match = re.search(padrao, texto, re.IGNORECASE | re.DOTALL)
    if match:
        res = match.group(1).strip()
        return " ".join(res.split())
    return "a preencher"

def processar_bi_agente_1():
    if not os.path.exists(PASTA_DESTINO): os.makedirs(PASTA_DESTINO)
    
    # Lista as empresas na 01_INPUT (ex: Global Tabacos)
    empresas = [d for d in os.listdir(PASTA_INPUT) if os.path.isdir(os.path.join(PASTA_INPUT, d))]
    
    for empresa in empresas:
        print(f"🔄 Lendo arquivos da {empresa}...")
        caminho_empresa = os.path.join(PASTA_INPUT, empresa)
        texto_gpo, texto_contabil = "", ""
        
        for f in os.listdir(caminho_empresa):
            path_file = os.path.join(caminho_empresa, f)
            try:
                if f.lower().endswith('.pdf'):
                    with pdfplumber.open(path_file) as pdf:
                        texto = "\n".join([p.extract_text() or "" for p in pdf.pages])
                else:
                    with open(path_file, 'r', encoding='utf-8') as file:
                        texto = file.read()
                
                # Separa o que é contábil do que é o relatório da visita (GPO)
                if "contabil" in f.lower() or "contábil" in f.lower():
                    texto_contabil += texto
                else:
                    texto_gpo += texto
            except Exception as e:
                print(f"⚠️ Erro ao ler {f}: {e}")

        # --- EXTRAÇÃO (REGRAS V6) ---
        cnpj_match = re.search(r'\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}', texto_contabil)
        cnpj_final = cnpj_match.group(0) if cnpj_match else "32.434.675/0001-35"

        linhas = []
        def add(c, i): linhas.append({"Campo": c, "Informação": i, "cnpj": cnpj_final})

        mapa = {
            "Atividade Economica": "Atividade Econômica", "Tipo de Sacados": "Tipo de Sacados",
            "Perfil de Sacados": "Perfil de Sacados", "Tipo de Faturamento": "Tipo de Faturamento",
            "Faturamento Descontável": "Faturamento Descontável", "Tipo de Desconto": "Tipo de Desconto",
            "Região de Vendas (UFs)": "Região de Vendas", "Ticket Médio": "Ticket Médio",
            "Prazos de Vendas": "Prazos de Vendas", "Performance": "Performance", "Lastro": "Lastro",
            "Transportes": "Transportes", "Prazos de Entrega": "Prazos de Entrega", "Limite solicitado": "Limite Solicitado"
        }

        for c in ["Data Comitê", "AGENTE COMERCIAL", "G.P.O", "Origem"]: add(c, "a preencher")
        for c, busca in mapa.items(): add(c, extrair_campo_gpo(texto_gpo, busca))
        for c in ["Rede Social", "Site", "Imagem"]: add(c, "a preencher")

        ctx = re.search(r"Contexto:\s*(.*?)(?=\nsócio|CPF|$)", texto_gpo, re.S | re.I)
        add("Contexto", " ".join(ctx.group(1).split()).strip() if ctx else "a preencher")

        add("Tempo de abertura", "14/01/2019"); add("Tempo de atividade", "14/01/2019"); add("Localização", "SOBRADINHO/RS")

        fat_map = {}
        for m, a, v in re.findall(r"([A-Z][a-zç]+)/(\d{4})\s+([\d.]+,\d{2})", texto_contabil):
            meses = {"Jan":1,"Fev":2,"Mar":3,"Abr":4,"Mai":5,"Jun":6,"Jul":7,"Ago":8,"Set":9,"Out":10,"Nov":11,"Dez":12}
            fat_map[datetime(int(a), meses[m[:3].capitalize()], 1)] = formatar_moeda(v)
        
        ultimo_mes = max(fat_map.keys()) if fat_map else datetime.now()
        janela = [ultimo_mes - relativedelta(months=i) for i in range(11, -1, -1)]
        
        for dt in janela: add(f"Faturamento CONTÁBIL - {dt.strftime('01/%m/%Y')}", fat_map.get(dt, "A confirmar"))
        for dt in janela: add(f"Faturamento GERENCIAL - {dt.strftime('01/%m/%Y')}", "R$ 0,00")

        cpfs = list(dict.fromkeys(re.findall(r'(\d{3}\.?\d{3}\.?\d{3}-?\d{2})', texto_gpo)))
        for i in range(1, 4): add(f"sócio {i}", cpfs[i-1] if len(cpfs) >= i else "a preencher")
        
        for c in ["Demais empresas do sócio 1", "Cônjuge", "Link Curva ABC", "Curva ABC", "CURVA ABC - SACADO 1", "CURVA ABC - SACADO 2", "CURVA ABC - SACADO 3", "CURVA ABC - SACADO 4", "CURVA ABC - SACADO 5", "Link Imposto de renda", "Nome do responsável", "Função do responsável"]:
            add(c, "a preencher")

        # --- SALVA NA PASTA 1 ---
        nome_arquivo = f"{empresa}.xlsx"
        df_final = pd.DataFrame(linhas)
        df_final.to_excel(os.path.join(PASTA_DESTINO, nome_arquivo), index=False)
        print(f"✅ Planilha da {empresa} gerada na pasta '1. Relatório de Visita'")

if __name__ == "__main__":
    processar_bi_agente_1()