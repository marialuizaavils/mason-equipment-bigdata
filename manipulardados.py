import pandas as pd

# Carregar a planilha
file_path = 'Programa Excelencia Komatsu 2024.xlsx'
df = pd.read_excel(file_path, index_col=0)

# Exibir os dados
print("Dados originais:")
print(df)

# Adicionar uma coluna com o total de pontos obtídos 
df['Total de pontos obtídos'] = df.sum(axis=1)

# Exibir os dados com a nova coluna
print("\nDados com a coluna 'Total de pontos obtídos':")
print(df)

# Filtrar a quantidade de pontos positivos
filtered_df = df[df['Total de ponto obtídos'] > 160]

# Filtrar a quantidade de pontos pendentes
filtered_df = df[df['Total de ponto obtídos'] > 102]

# Filtrar a quantidade de pontos negativos
filtered_df = df[df['Total de ponto obtídos'] > 90]

# Exibir pontos totais
print("\nPontos positivos 160:")
print(filtered_df)

# Salvar os dados manipulados em um novo arquivo Excel
output_path = 'programakomatsu.xlsx'
df.to_excel(output_path)

print(f"\nDados manipulados foram salvos em: {output_path}")
