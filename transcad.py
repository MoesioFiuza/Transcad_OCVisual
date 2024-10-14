import pandas as pd

# Carregar o arquivo Excel principal
file_path = r'C:\Users\moesios\Desktop\Material Tóxico - não mexer\2 -NDOM_NMOT_unificada_transformada.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Garantir que todas as zonas sejam strings
df['ZONA_ORIGEM'] = df['ZONA_ORIGEM'].astype(str)
df['ZONA_DESTINO'] = df['ZONA_DESTINO'].astype(str)

# Identificar as zonas únicas
zonas = pd.unique(df[['ZONA_ORIGEM', 'ZONA_DESTINO']].values.ravel('K'))

# Carregar o arquivo Excel com os fatores horários
planilha_tempo = r'C:\Users\moesios\Desktop\Material Tóxico - não mexer\tempo.xlsx'
df_tempo = pd.read_excel(planilha_tempo, sheet_name='Planilha1')

# Definir o caminho de saída
output_dir = r'C:\Users\moesios\Desktop\Material Tóxico - não mexer\sa9ida'

# Ocupação média
ocupacao_media = 1

# Converter a coluna "VALOR" para float
df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce')

# Criar uma matriz inicial com os valores da coluna VALOR, preenchendo valores faltantes com 0
matriz = pd.pivot_table(df, index='ZONA_ORIGEM', columns='ZONA_DESTINO', values='VALOR', aggfunc='sum', fill_value=0)

# Adicionar linhas e colunas com zeros para tornar a matriz quadrada
for zona in zonas:
    if zona not in matriz.index:
        matriz.loc[zona] = 0
    if zona not in matriz.columns:
        matriz[zona] = 0

# Reordenar as linhas e colunas para que sejam consistentes
matriz = matriz.reindex(index=zonas, columns=zonas).fillna(0)

# Gerar uma planilha para cada fator horário
for index, row in df_tempo.iterrows():
    hora = row['HORA']
    fator = row['Não-domiciliar']
    
    # Aplicar o fator horário e a ocupação média à matriz
    matriz_ajustada = matriz * fator * ocupacao_media
    
    # Salvar a nova planilha
    output_file = f'{output_dir}/NDOM_NMOT{hora}.xlsx'
    matriz_ajustada.to_excel(output_file)

print("Planilhas geradas com sucesso!")
