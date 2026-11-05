Python 3.13.14 (tags/v3.13.14:fd17997, Jun 10 2026, 13:03:48) [MSC v.1944 64 bit (AMD64)] on win32
Enter "help" below or click "Help" above for more information.
import os
import re
import pandas as pd
from pypdf import PdfReader

# ================= CONFIGURAÇÕES =================
# Cole o caminho da pasta ou do servidor onde estão os seus PDFs (ex: r"D:\MeusPDFs")
PASTA_ALVO = r"Z:\INDEXAÇÃO\CONVERTIDOS"

# Defina aqui os filtros que você deseja buscar (deixe em branco "" para trazer tudo)
FILTRO_DATA = ""          # Ex: "12/10/2025" ou deixe ""
FILTRO_VALOR = "R$1.330,00"         # Ex: "1.500,00" ou deixe ""
FILTRO_CNPJ = "10.759.706.0001-95"          # Ex: "12.345.678/0001-99" ou deixe ""
TERMO_LIVRE = ""          # Ex: "Nome do Cliente" ou deixe ""
# ==================================================

... def extrair_dados_pdf(caminho_pdf):
...     dados = {
...         "arquivo": os.path.basename(caminho_pdf),
...         "caminho_completo": os.path.abspath(caminho_pdf), 
...         "datas_encontradas": [],
...         "valores_encontrados": [],
...         "cnpjs_encontrados": [],
...         "status": "Sucesso"
...     }
...     
...     texto_completo = ""
...     try:
...         reader = PdfReader(caminho_pdf)
...         for pagina in reader.pages:
...             t = pagina.extract_text()
...             if t:
...                 texto_completo += t + "\n"
...     except Exception as e:
...         dados["status"] = f"Erro ao ler PDF: {str(e)}"
...         return dados
... 
...    
...     padrao_data = r"\b\d{2}/\d{2}/\d{4}\b|\b\d{4}-\d{2}-\d{2}\b"
...     padrao_valor = r"R\$\s*[\d\.]+,\d{2}|\b\d{1,3}(?:\.\d{3})*,\d{2}\b"
...     padrao_cnpj = r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b|\b\d{14}\b"
... 
...     dados["datas_encontradas"] = list(set(re.findall(padrao_data, texto_completo)))
...     dados["valores_encontrados"] = list(set(re.findall(padrao_valor, texto_completo)))
...     dados["cnpjs_encontrados"] = list(set(re.findall(padrao_cnpj, texto_completo)))
...     dados["texto_bruto"] = texto_completo
... 
...     return dados
... 
... def varrer_pasta():
...     resultados = []
...     print(f"Iniciando varredura na pasta: {PASTA_ALVO}...\n")
...     
...     total_lidos = 0
...     for raiz, _, arquivos in os.walk(PASTA_ALVO):
...         for arquivo in arquivos:
...             if arquivo.lower().endswith(".pdf"):
                total_lidos += 1
                if total_lidos % 100 == 0:
                    print(f"Processados {total_lidos} arquivos até agora...")
                
                caminho_completo = os.path.join(raiz, arquivo)
                res = extrair_dados_pdf(caminho_completo)
                
                if res["status"] != "Sucesso":
                    continue
                
                
                atende_filtros = True
                if FILTRO_DATA and not any(FILTRO_DATA in d for d in res["datas_encontradas"]):
                    atende_filtros = False
                if FILTRO_VALOR and not any(FILTRO_VALOR in v for v in res["valores_encontrados"]):
                    atende_filtros = False
                if FILTRO_CNPJ and not any(FILTRO_CNPJ in c for c in res["cnpjs_encontrados"]):
                    atende_filtros = False
                if TERMO_LIVRE and TERMO_LIVRE.lower() not in res["texto_bruto"].lower():
                    atende_filtros = False
                
                
                if atende_filtros or (not FILTRO_DATA and not FILTRO_VALOR and not FILTRO_CNPJ and not TERMO_LIVRE):
                   
                    link_excel = f'=HYPERLINK("{res["caminho_completo"]}"; "{res["arquivo"]}")'
                    
                    resultados.append({
                        "Abrir Arquivo": link_excel,
                        "Caminho Direto": res["caminho_completo"],
                        "Datas": ", ".join(res["datas_encontradas"]),
                        "Valores": ", ".join(res["valores_encontrados"]),
                        "CNPJs": ", ".join(res["cnpjs_encontrados"])
                    })

    print(f"\nVarredura concluída! Total de {total_lidos} PDFs analisados.")
    return resultados

if __name__ == "__main__":
    if not os.path.exists(PASTA_ALVO):
        print(f"Erro: A pasta '{PASTA_ALVO}' não foi encontrada. Verifique o caminho.")
    else:
        dados_encontrados = varrer_pasta()
        if dados_encontrados:
            df = pd.DataFrame(dados_encontrados)
            nome_saida = "resultado_busca_pdfs.xlsx"
            df.to_excel(nome_saida, index=False)
            print(f"Sucesso! Encontrados {len(dados_encontrados)} arquivos correspondentes.")
            print(f"Planilha gerada com sucesso: {os.path.abspath(nome_saida)}")
        else:
