import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime


# Pasta raiz onde estão os anos
base_path = r"C:\Projetos\ExtrairPDF\docs\2025\Janeiro"

# Regex ajustada para capturar Data, Descrição e Valor
padrao = re.compile(r"^(\d{2}/\d{2})\s+(.*?)\s+(-?\s?\d{1,3}(?:\.\d{3})*,\d{2})$")

def converte_valor(valor_str):
    """Converte '1.234,56' ou '-1.234,56' para float"""
    return float(valor_str.replace(".", "").replace(",", ".").replace(" ", ""))

def processar_pdf(pdf_path):
    """Lê um PDF de extrato e retorna um DataFrame"""
    dados = []
    ano = os.path.basename(os.path.dirname(os.path.dirname(pdf_path)))  # ex: "2025"
    mes = os.path.basename(os.path.dirname(pdf_path))  # ex: "Março"
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            linhas = pagina.extract_text().split("\n")
            for linha in linhas:
                linha = linha.strip()

                # Caso especial: linhas com saldo
                if ("SALDO ANTERIOR" in linha or "SALDO" in linha or "SDO" in linha) and ("," in linha):
                    partes = linha.split()
                    data_str = partes[0] + "/" + ano   # vira dd/MM/YYYY
                    data = datetime.strptime(data_str, "%d/%m/%Y")
                    descricao = " ".join(partes[1:-1])
                    valor = converte_valor(partes[-1])
                    dados.append([data, descricao, None, valor])
                else:
                    m = padrao.match(linha)
                    if m:
                        data_str = m.group(1) + "/" + ano   # vira dd/MM/YYYY
                        data = datetime.strptime(data_str, "%d/%m/%Y")
                        descricao = m.group(2).strip()
                        valor = converte_valor(m.group(3))
                        dados.append([data, descricao, valor, None])

    df = pd.DataFrame(dados, columns=["Data", "Descrição", "Valor (R$)", "Saldo (R$)"])
    
    # Adiciona colunas auxiliares para facilitar o consolidado
    
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

# Ordenar por ano, mês e data
# (Se os meses forem nomes, podemos depois criar um dicionário para ordenar corretamente)
#df_final.sort_values(by=["Ano", "Mês", "Data"], inplace=True)
#df_final.sort_values(by=["Data"], inplace=True)

# Converter Data para string no formato dd/MM/yyyy
df_final["Data"] = df_final["Data"].dt.strftime("%d/%m/%Y")
# Exportar para Excel consolidado
output_path = os.path.join(base_path, "extrato_consolidado.xlsx")
df_final.to_excel(output_path, index=False)

print(f"\n✅ Consolidado salvo em: {output_path}")
