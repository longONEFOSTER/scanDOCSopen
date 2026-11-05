
import os
import re
import pandas as pd
from pypdf import PdfReader
from pathlib import Path

# =========================================================
# CONFIGURAÇÕES
# =========================================================

PASTA_ALVO = Path(r"Z:\INDEXAÇÃO\CONVERTIDOS")

FILTRO_DATA = ""       # Ex: 12/10/2025
FILTRO_VALOR = ""      # Ex: 1.500,00
FILTRO_CNPJ = ""
TERMO_LIVRE = ""       # Ex: Nome do Cliente

NOME_SAIDA = "resultado_busca_pdfs.xlsx"
NOME_ERROS = "erros_leitura_pdfs.xlsx"



PADRAO_DATA = r"\b\d{2}/\d{2}/\d{4}\b|\b\d{4}-\d{2}-\d{2}\b"

PADRAO_VALOR = (
    r"R\$\s*[\d.]+,\d{2}"
    r"|\b\d{1,3}(?:\.\d{3})*,\d{2}\b"
)

PADRAO_CNPJ = (
    r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b"
    r"|\b\d{14}\b"
)



def normalizar_cnpj(valor):
    return re.sub(r"\D", "", str(valor))


def normalizar_valor(valor):
    return re.sub(r"[^\d]", "", str(valor))


def remover_duplicados(lista):
    return list(dict.fromkeys(lista))


def extrair_dados_pdf(caminho_pdf):

    dados = {
        "arquivo": caminho_pdf.name,
        "caminho_completo": str(caminho_pdf.resolve()),
        "datas_encontradas": [],
        "valores_encontrados": [],
        "cnpjs_encontrados": [],
        "texto_bruto": "",
        "status": "Sucesso"
    }

    texto_completo = []

    try:

        reader = PdfReader(str(caminho_pdf), strict=False)

        if reader.is_encrypted:
            try:
                resultado = reader.decrypt("")
                if not resultado:
                    raise ValueError("PDF protegido por senha.")
            except Exception:
                raise ValueError("PDF criptografado ou protegido.")

        for pagina in reader.pages:

            try:
                texto = pagina.extract_text()

                if texto:
                    texto_completo.append(texto)

            except Exception:
                continue

        texto_completo = "\n".join(texto_completo)

        dados["texto_bruto"] = texto_completo

        dados["datas_encontradas"] = remover_duplicados(
            re.findall(PADRAO_DATA, texto_completo)
        )

        dados["valores_encontrados"] = remover_duplicados(
            re.findall(PADRAO_VALOR, texto_completo)
        )

        dados["cnpjs_encontrados"] = remover_duplicados(
            re.findall(PADRAO_CNPJ, texto_completo)
        )

    except Exception as erro:

        dados["status"] = str(erro)

    return dados


def verificar_filtros(dados):

    texto = dados["texto_bruto"]

    if FILTRO_DATA:
        if FILTRO_DATA not in dados["datas_encontradas"]:
            return False

    if FILTRO_VALOR:

        filtro = normalizar_valor(FILTRO_VALOR)

        encontrados = [
            normalizar_valor(v)
            for v in dados["valores_encontrados"]
        ]

        if filtro not in encontrados:
            return False

    if FILTRO_CNPJ:

        filtro = normalizar_cnpj(FILTRO_CNPJ)

        encontrados = [
            normalizar_cnpj(c)
            for c in dados["cnpjs_encontrados"]
        ]

        if filtro not in encontrados:
            return False

    if TERMO_LIVRE:

        if TERMO_LIVRE.casefold() not in texto.casefold():
            return False

    return True


def varrer_pasta():

    resultados = []
    erros = []

    total_lidos = 0
    total_encontrados = 0

    print("\nIniciando varredura...")
    print(f"Pasta: {PASTA_ALVO}\n")

    for raiz, _, arquivos in os.walk(PASTA_ALVO):

        for arquivo in arquivos:

            if not arquivo.lower().endswith(".pdf"):
                continue

            total_lidos += 1

            caminho = Path(raiz) / arquivo

            if total_lidos % 100 == 0:
                print(f"Processados: {total_lidos} PDFs")

            dados = extrair_dados_pdf(caminho)

            # Registra arquivos com problemas
            if dados["status"] != "Sucesso":

                erros.append({
                    "Arquivo": arquivo,
                    "Caminho": str(caminho.resolve()),
                    "Erro": dados["status"]
                })

                continue

            # Verifica filtros
            if not verificar_filtros(dados):
                continue

            caminho_excel = dados["caminho_completo"]

            # Escapa aspas para evitar erro na fórmula Excel
            caminho_excel = caminho_excel.replace('"', '""')

            nome_excel = dados["arquivo"].replace('"', '""')

            link_excel = (
                f'=HYPERLINK("{caminho_excel}";"{nome_excel}")'
            )

            resultados.append({

                "Abrir Arquivo": link_excel,

                "Caminho Direto": dados["caminho_completo"],

                "Datas": ", ".join(
                    dados["datas_encontradas"]
                ),

                "Valores": ", ".join(
                    dados["valores_encontrados"]
                ),

                "CNPJs": ", ".join(
                    dados["cnpjs_encontrados"]
                )

            })

            total_encontrados += 1

    print("\n===================================")
    print("VARREDURA CONCLUÍDA")
    print("===================================")

    print(f"PDFs analisados: {total_lidos}")
    print(f"Arquivos encontrados: {total_encontrados}")
    print(f"Arquivos com erro: {len(erros)}")

    return resultados, erros



def gerar_excel(resultados, erros):

    if resultados:

        df = pd.DataFrame(resultados)

        df.to_excel(
            NOME_SAIDA,
            index=False,
            engine="openpyxl"
        )

        print(f"\nPlanilha gerada: {os.path.abspath(NOME_SAIDA)}")

    else:

        print("\nNenhum arquivo correspondeu aos filtros.")

    if erros:

        df_erros = pd.DataFrame(erros)

        df_erros.to_excel(
            NOME_ERROS,
            index=False,
            engine="openpyxl"
        )

        print(f"Relatório de erros: {os.path.abspath(NOME_ERROS)}")


if __name__ == "__main__":

    if not PASTA_ALVO.exists():

        print(f"ERRO: Pasta não encontrada: {PASTA_ALVO}")

    elif not PASTA_ALVO.is_dir():

        print("ERRO: O caminho informado não é uma pasta.")

    else:

        resultados, erros = varrer_pasta()

        gerar_excel(resultados, erros)

        print("\nProcessamento finalizado.")
