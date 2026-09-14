import pandas as pd
import os

archivo = "base_matriz_mce.xlsx"
df = pd.read_excel(archivo)

# Buscar la actividad 81
idx = df[df['No'] == 81].index
if len(idx) > 0:
    df.loc[idx[0], 'Evidencia'] = 'evidencias/SCA-081_actualizacion_de_app_mejoras_con_imagen_corporativ_20260629.png'
    df.loc[idx[0], '% Avance'] = 100
    df.to_excel(archivo, index=False)
    print("Excel actualizado localmente.")
else:
    print("Actividad 81 no encontrada.")
