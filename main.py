import pandas as pd

# Definir os nomes das colunas
colunas = [
    "Data", "1ª Entrada", "1ª Saída", "2ª Entrada", "2ª Saída", "3ª Entrada",
    "Crédito", "Débito", "H. intervalo", "Horas normais", "Horas extras fator 1 (50%)",
    "Horas extras fator 2 (100%)", "Adicional noturno", "Saldo", "Motivo/Observação"
]

# Carregar o arquivo Excel
file_path = r'C:\Users\anne_marcelo_bento\Documents\analise_ponto\ponto.xlsx'  # Substitua pelo caminho do seu arquivo
df = pd.read_excel(file_path, header=None)

# Definir as colunas usando os nomes fornecidos
df.columns = colunas

# Inicializar variáveis
dados_tecnicos = []
tecnico_atual = None
dados_atual = []

# Percorrer todas as linhas
for index, row in df.iterrows():
    # Verificar se a linha contém "Colaborador"
    if 'Colaborador' in row.astype(str).values:
        if tecnico_atual and dados_atual:
            # Salvar dados do técnico anterior
            dados_tecnicos.append([tecnico_atual] + dados_atual)
        # Capturar o nome do técnico
        tecnico_atual = row.dropna().values[-1]  # Último valor não nulo é o nome do técnico
        dados_atual = []  # Reiniciar dados para o novo técnico
    
    # Verificar se a linha contém "TOTAIS"
    elif 'TOTAIS' in row.astype(str).values:
        # Ignorar linhas de totais
        continue
    
    else:
        # Adicionar dados das outras linhas
        dados_atual.append(row.dropna().values.tolist())

# Adicionar o último técnico ao final da lista
if tecnico_atual and dados_atual:
    dados_tecnicos.append([tecnico_atual] + dados_atual)

# Converter os dados em um DataFrame
linhas_formatadas = []
for tecnico_dados in dados_tecnicos:
    tecnico = tecnico_dados[0]
    for dados in tecnico_dados[1:]:
        linha = [tecnico] + dados
        linhas_formatadas.append(linha)

# Criar DataFrame com os dados processados
df_final = pd.DataFrame(linhas_formatadas, columns=["Técnico"] + colunas)

# Salvar em um novo arquivo Excel
output_path = 'tecnicos_dados_formatados.xlsx'
df_final.to_excel(output_path, index=False)

print(f"Dados organizados e salvos em: {output_path}")
