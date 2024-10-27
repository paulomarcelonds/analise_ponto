import pandas as pd

# Carregar o arquivo Excel, pulando as duas primeiras linhas
file_path = 'ponto.xlsx'
df = pd.read_excel(file_path, skiprows=2, header=None)

# Definir o cabeçalho conforme solicitado
df.columns = [
    "Data", "1ª Entrada", "1ª Saída", "2ª Entrada", "2ª Saída", "3ª Entrada", 
    "Crédito", "Débito", "H. intervalo", "Horas normais", "Horas extras fator 1 (50%)", 
    "Horas extras fator 2 (100%)", "Adicional noturno", "Saldo", "Motivo/Observação"
]

# Lista para armazenar o nome do colaborador para cada linha
colaborador_column = []

# Variável para guardar o nome atual do colaborador
current_colaborador = None

# Iterar pelas linhas do DataFrame para detectar e preencher o nome do colaborador
for index, row in df.iterrows():
    # Verificar se a linha contém "Colaborador" na coluna "Data"
    if row["Data"] == "Colaborador":
        # Definir o nome do colaborador a partir da coluna seguinte
        current_colaborador = row[1]
    # Adicionar o nome do colaborador atual à lista
    colaborador_column.append(current_colaborador)

# Adicionar a coluna "Nome" com os dados dos colaboradores
df.insert(0, "Nome", colaborador_column)

# Excluir as linhas que contêm a palavra "TOTAIS" na coluna "Data"
df = df[df["Data"] != "TOTAIS"]

# Excluir linhas onde a coluna "Data" tem o valor "Colaborador" (linhas de cabeçalho intermediárias)
df = df[df["Data"] != "Colaborador"]

# Excluir todas as linhas onde a coluna "Data" contém a palavra "Data"
df = df[~df["Data"].astype(str).str.contains("Data", case=False, na=False)]

# Salvar o DataFrame modificado em um novo arquivo Excel
output_path = 'colaboradores.xlsx'
df.to_excel(output_path, index=False)

print(f"Arquivo salvo em: {output_path}")
