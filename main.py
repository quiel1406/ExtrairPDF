import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime


# Pasta raiz onde estão os anos
base_path = r"C:\Projetos\ExtrairPDF\docs\2020"


# Regex antigo: "dd/MM descrição valor"
padrao_antigo = re.compile(r"^(\d{2}/\d{2})\s+(.*?)\s+(-?\s?\d{1,3}(?:\.\d{3})*,\d{2})$")

# Regex novo: "dd/MM/yyyy descrição valor"
padrao_novo = re.compile(r"^(\d{2}/\d{2}/\d{4})\s+(.*?)\s+(-?\s?\d{1,3}(?:\.\d{3})*,\d{2})$")

def converte_valor(valor_str):
    """Converte '1.234,56' ou '-1.234,56' para float"""
    return float(valor_str.replace(".", "").replace(",", ".").replace(" ", ""))

def processar_pdf(pdf_path):
    """Lê um PDF de extrato e retorna um DataFrame"""
    dados = []
    ano = os.path.basename(os.path.dirname(os.path.dirname(pdf_path)))  # ex: "2025"
    mes = os.path.basename(os.path.dirname(pdf_path))  # ex: "Março"

    linha_index = 1  # contador de leitura
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            linhas = pagina.extract_text().split("\n")
            for linha in linhas:
                linha = linha.strip()

                # Caso especial: linhas com saldo
                if ("SALDO ANTERIOR" in linha or "SALDO" in linha or "SDO" in linha) and ("," in linha):
                    partes = linha.split()
                    data =  partes[0]
                    descricao = " ".join(partes[1:-1])
                    valor = converte_valor(partes[-1])
                    dados.append([linha_index,data, descricao, None, valor])
                else:
                    m1 = padrao_antigo.match(linha)
                    m2 = padrao_novo.match(linha)
                   
                    if m1:
                        data = m1.group(1) + "/" + ano
                        descricao = m1.group(2).strip()
                        valor = converte_valor(m1.group(3))
                        dados.append([linha_index,data, descricao, valor, None])
                    elif m2:

                        data = m2.group(1)  # já vem dd/MM/yyyy
                        descricao = m2.group(2).strip()
                        valor = converte_valor(m2.group(3))
                        dados.append([linha_index, data, descricao, valor, None])
                linha_index +=1

    df = pd.DataFrame(dados, columns=["Index_Leitura", "Data", "Descrição", "Valor (R$)", "Saldo (R$)"])
    df["Ano"] = ano
    df["Mês"] = mes

    return df
# Lista para consolidar todos os DataFrames
todos_dfs = []

# Percorrer todas as pastas dentro de base_path
for root, dirs, files in os.walk(base_path):
    for file in files:
        #if file.startswith("Extrato") and file.endswith(".pdf"):
        if file.startswith("Extrato") and file.endswith(".pdf") and "Poupança" not in file:
            pdf_file = os.path.join(root, file)
            print(f"🔎 Processando {pdf_file} ...")
            df = processar_pdf(pdf_file)
            todos_dfs.append(df)

# Consolidar tudo em um único DataFrame
df_final = pd.concat(todos_dfs, ignore_index=True)
ordem_meses = {
    "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
    "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
}
df_final["Mês_num"] = df_final["Mês"].map(ordem_meses)

# Ordenar por ano, mês e data
# (Se os meses forem nomes, podemos depois criar um dicionário para ordenar corretamente)
df_final.sort_values(by=["Ano", "Mês_num", "Index_Leitura"], inplace=True)
#df_final.sort_values(by=["Data"], inplace=True)


# Exportar para Excel consolidado
output_path = os.path.join(base_path, "extrato_consolidado.xlsx")
df_final.to_excel(output_path, index=False)

print(f"\n✅ Consolidado salvo em: {output_path}")
