import re
import pdfplumber
import pandas as pd

# Caminho do PDF
pdf_path = "docs//Extrato Fevereiro 2025.pdf"

dados = []

# Regex para capturar: Data, Descrição, Valor, Saldo
padrao = re.compile(r"^(\d{2}/\d{2})\s+(.*?)\s+(-?\d{1,3}(?:\.\d{3})*,\d{2})$")

with pdfplumber.open(pdf_path) as pdf:
    for pagina in pdf.pages:
        linhas = pagina.extract_text().split("\n")
        for linha in linhas:
            # Caso especial: linhas com saldo final
            if "SALDO ANTERIOR" in linha or "SALDO" in linha and linha.strip().endswith(",00"):
                partes = linha.split()
                data = partes[0]
                descricao = " ".join(partes[1:-1])
                valor = partes[-1]
                dados.append([data, descricao, valor, None])
            else:
                m = padrao.match(linha.strip())
                if m:
                    data = m.group(1)
                    descricao = m.group(2)
                    valor = m.group(3)
                    dados.append([data, descricao, valor, None])

# Criar DataFrame
df = pd.DataFrame(dados, columns=["Data", "Descrição", "Valor (R$)", "Saldo (R$)"])

# Exportar para Excel
df.to_excel("extrato_fevereiro_2025.xlsx", index=False)

print("Extrato convertido para Excel com sucesso!")
