import pandas as pd

# Carregar o arquivo Excel, pulando as duas primeiras linhas
file_path = 'ponto.xlsx'  # Substitua pelo caminho do seu arquivo
df = pd.read_excel(file_path, skiprows=1)

# Definir os cabeçalhos
cabecalhos = [
    "Data", "1ª Entrada", "1ª Saída", "2ª Entrada", "2ª Saída", 
    "3ª Entrada", "3ª Saída", "Crédito", "Débito", "H. intervalo", 
    "Horas normais", "Horas extras fator 1 (50%)", 
    "Horas extras fator 2 (100%)", "Adicional noturno", 
    "Saldo", "Motivo/Observação"
]

# Atribuir os cabeçalhos ao DataFrame
df.columns = cabecalhos

# Preencher campos vazios com "-"
df.fillna("-", inplace=True)

colaborador_atual = None
nomes = []

for index, row in df.iterrows():
    if "Colaborador" in str(row["Data"]):
        colaborador_atual = row["1ª Entrada"]  # Armazena o nome encontrado
    nomes.append(colaborador_atual if colaborador_atual else "-")  # Preenche a lista com o nome atual ou "-"

df["Nome"] = nomes 

df = df[~df["Data"].isin(["Colaborador", "Data", "TOTAIS"])]

colunas_reordenadas = ["Nome"] + [col for col in df.columns if col != "Nome"]
df = df[colunas_reordenadas]

# Salvar o DataFrame modificado em um novo arquivo Excel
output_path = 'arquivo_modificado.xlsx'  # Nome do novo arquivo
df.to_excel(output_path, index=False)

print(f"Arquivo salvo como: {output_path}")
