import streamlit as st
import pandas as pd
import os
import sys 
sys.path.append(r"C:\Users\IVI6131\OneDrive - MDLZ\Consolidación Información OSA WACAM - General\Manual Files\2025\Proceso\Funciones")

# 👉 Importa tu función de validación
from Funciones_validacion_inventory import validar_reglas_manual_file_inventory_prueba

st.title("Validador Automático de Archivos Manual file Inventory")
archivo = st.file_uploader("📂 Carga tu archivo Excel", type=["xlsx"])
if archivo:
   df = pd.read_excel(archivo, sheet_name="Plantilla", dtype=str)
   st.success("✅ Archivo cargado correctamente")
   st.write("Vista previa del archivo:")
   st.dataframe(df.head())
   # 👉 Validar
   resultado = validar_reglas_manual_file_inventory_prueba(df, archivo.name)
   st.write("### ✅ Resultado de la validación")
   st.dataframe(resultado)
   # 👉 Descargar resultado
   if st.button("Descargar resultado en Excel"):
       resultado.to_excel("resultado_validacion.xlsx", index=False)
       with open("resultado_validacion.xlsx", "rb") as f:
           st.download_button(
               label="📥 Descargar archivo validado",
               data=f,
               file_name="resultado_validacion.xlsx",
               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
           )