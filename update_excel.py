import pandas as pd
import os

archivo = "base_matriz_mce.xlsx"
df = pd.read_excel(archivo)

updates = {
    77: 'evidencias/SCA-077_generar_aplicacion_para_incomming_de_materia_prima_20260629.png',
    80: 'evidencias/SCA-080_actualizacion_de_app_mejoras_con_imagen_corporativ_20260629.png',
    81: 'evidencias/SCA-081_actualizacion_de_app_mejoras_con_imagen_corporativ_20260629.png'
}

for num, path in updates.items():
    idx = df[df['No'] == num].index
    if len(idx) > 0:
        df.loc[idx[0], 'Evidencia'] = path
        df.loc[idx[0], '% Avance'] = 100
        print(f"Updated {num}")

df.to_excel(archivo, index=False)
print("Excel saved.")
