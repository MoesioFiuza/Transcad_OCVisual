import pandas as pd

file_path = r'planilha matriz'
df = pd.read_excel(file_path, sheet_name='Sheet1')

df['ZONA_ORIGEM'] = df['ZONA_ORIGEM'].astype(str)
df['ZONA_DESTINO'] = df['ZONA_DESTINO'].astype(str)

zonas = pd.unique(df[['ZONA_ORIGEM', 'ZONA_DESTINO']].values.ravel('K'))

planilha_tempo = r'C:\Users\moesios\Desktop\Material Tóxico - não mexer\tempo.xlsx'
df_tempo = pd.read_excel(planilha_tempo, sheet_name='Planilha1')

output_dir = r'caminho de saída'

ocupacao_media = 1

df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce')

matriz = pd.pivot_table(df, index='ZONA_ORIGEM', columns='ZONA_DESTINO', values='VALOR', aggfunc='sum', fill_value=0)

for zona in zonas:
    if zona not in matriz.index:
        matriz.loc[zona] = 0
    if zona not in matriz.columns:
        matriz[zona] = 0

matriz = matriz.reindex(index=zonas, columns=zonas).fillna(0)

for index, row in df_tempo.iterrows():
    hora = row['HORA']
    fator = row['Não-domiciliar']
    
    matriz_ajustada = matriz * fator * ocupacao_media

    output_file = f'{output_dir}/NDOM_NMOT{hora}.xlsx'
    matriz_ajustada.to_excel(output_file)

