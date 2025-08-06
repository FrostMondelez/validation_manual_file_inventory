#-----------------------------------IMPORTACION DE LIBRERIAS-----------------------------------#
import shutil
import pandas as pd 
import numpy as np
import re
# import matplotlib.pyplot as plt
from datetime import datetime
import os
import win32com.client as win32

# Funcion para leer el manual file a elegir 
def obtener_archivo_mas_reciente(Anio, periodo, tipo_manual_file, wacam, patron):
   """
   Construye la ruta y retorna el path completo del archivo más reciente
   que contiene un patrón opcional en su nombre.

   Parámetros:
   - Anio (str) : establecer el año de validacion (2025)
   - periodo (str) : establecer el periodo de validacion (P6)
   - tipo_manual_file (str): es la carpeta que debo establecer (Sell_out_igual_sell_in)
   - wacam (str): es la carpeta de consolidacion wacam.
   - patron: str (parte del nombre del archivo a filtrar, ej: 'sell_out')

   Retorna:
   - Ruta completa del archivo más reciente o None si no encuentra archivos.
   """
   
   base = r'C:\\Users\\IVI6131\\OneDrive - MDLZ\\Consolidación Información OSA WACAM - General\\Manual Files'
   ruta = os.path.join(base, str(Anio), periodo, 'Input', tipo_manual_file, wacam)
   if not os.path.exists(ruta):
       print(f"Ruta no encontrada: {ruta}")
       return None
   archivos = [f for f in os.listdir(ruta) if patron.lower() in f.lower()]
   if not archivos:
       print(f"No se encontraron archivos con el patrón '{patron}' en: {ruta}")
       return None
   # Ordenar por fecha de modificación más reciente
   archivos.sort(key=lambda f: os.path.getmtime(os.path.join(ruta, f)), reverse=True)
   archivo_mas_reciente = archivos[0]
   return os.path.join(ruta, archivo_mas_reciente)

def validar_reglas_manual_file_inventory_prueba(df, nombre_archivo):
   """
   Validar las reglas de negocio y estructura del manual file inventory.
   Parámetros:
       df (pandas.DataFrame): El manual file de inventory.
       nombre_archivo (str): Nombre del archivo que se valida.
   Retorna:
       pd.DataFrame con resultados de la validación (Manual_file, Regla, Indicador, Resultado, Hallazgo).
   """
   resultados = []
   errores = 0
   columnas_requeridas = [
       'Country_Key', 'Year', 'Period', 'SI_Sub_Channel', 'Customer_SI', 'SKU', 'Inventory_Tons'
   ]
   valores_country_key = {"AE", "BO", "CL", "PE", "CO", "EC", "NI", "HN", "SV", "CR", "PA", "GT", "PR", "DO"}
   def add(indicador, resultado, hallazgo, regla='Reglas de estructura'):
       resultados.append({
           'Manual_file': nombre_archivo,
           'Regla': regla,
           'Indicador': indicador,
           'Resultado': resultado,
           'Hallazgo': hallazgo
       })
   # === 1. Estructura ===
   faltantes = [c for c in columnas_requeridas if c not in df.columns]
   if faltantes:
       errores += 1
       add('Estructura', 'ERROR', 'Faltan columnas: ' + ', '.join(faltantes))
   else:
       add('Estructura', 'OK', 'Todas las columnas requeridas están presentes')

   # === 3. Duplicados lógicos (Customer_SI + SKU) ===
   if all(col in df.columns for col in ["Country_Key","Customer_SI", "SKU"]):
       conteos = df.groupby(["Country_Key","Customer_SI", "SKU"])["SKU"].transform("count")
       duplicados_logicos = conteos > 1
       if duplicados_logicos.any():
           errores += 1
           filas = (df[duplicados_logicos].index + 2).tolist()
           combinaciones = (
               df.loc[duplicados_logicos, ["Country_Key","Customer_SI", "SKU"]]
               .astype(str)
               .agg(" - ".join, axis=1)
               .drop_duplicates()
               .head(10)
               .tolist()
           )
           add(
               "Duplicados lógicos",
               "ERROR",
               f"{len(filas)} fila(s) con combinaciones duplicadas (Country_Key + Customer_SI + SKU). Ejemplos: {combinaciones} → Filas: {filas[:10]}"
           )
       else:
           add("Duplicados lógicos", "OK", "No hay duplicidad lógica en combinación Country_Key + Customer_SI + SKU")
   # === 4. Nulos ===
   nulos_total = 0
   columnas_revision_nulos = ['Country_Key', 'Year', 'Period', 'Customer_SI', 'SKU', 'Inventory_Tons']
   for col in columnas_revision_nulos:
       if col in df.columns:
           nulos = df[df[col].isnull()]
           nulos_total += nulos.shape[0]
           if not nulos.empty:
               errores += 1
               filas = (nulos.index + 2).tolist()
               add('Nulos', 'ERROR', f'Nulos en {col} → Filas: {filas[:10]}')
   if nulos_total == 0:
       add('Nulos', 'OK', 'No hay nulos en columnas requeridas')
 
   # === 5. Tipo de dato ===
   tipo_error = False
   for col in columnas_requeridas:
       if col in df.columns:
           no_string = df[~df[col].apply(lambda x: isinstance(x, str))]
           if not no_string.empty and col in ['Customer_SI', 'SKU']:
               tipo_error = True
               filas = (no_string.index + 2).tolist()
               add('Tipo de dato', 'ERROR', f'{col} no es string → Filas: {filas[:10]}')
   if not tipo_error:
       add('Tipo de dato', 'OK', 'Todas las columnas de texto tienen el tipo correcto')
   # === 6. Validación Country_Key ===
   if 'Country_Key' in df.columns:
       no_validos = df[~df['Country_Key'].isin(valores_country_key)]
       if not no_validos.empty:
           errores += 1
           filas = (no_validos.index + 2).tolist()
           valores = no_validos["Country_Key"].unique().tolist()
           add("Country_Key", "ERROR", f"Valores inválidos: {valores} → Filas: {filas[:10]}")
       else:
           add("Country_Key", "OK", "Todos los valores válidos")
           
   # === 7. Validaccion Longitud SKU > 10 ===
   if "SKU" in df.columns:
       # Convertimos todo a string y validamos longitud 
       sku_invalidos = df[df["SKU"].astype(str).str.len() <=10]

       if not sku_invalidos.empty:
           errores += 1
           filas =[sku_invalidos.index + 2].tolist()
           add("SKUS > 10 Digitos", "ERROR", f"SKU con 10 digitos o menos -> Filas: {filas[:10]}")
       else:
           add("SKUS > 10 Digitos","OK","Todos los Skus tinen más de 10 digitos")

#    # === 8. Indicador por país del estado de los SKU ===
#    if 'STATUS_UNICO' in df.columns and 'PAIS' in df.columns:
#     resumen_estado = (
#         df.groupby(['PAIS', 'STATUS_UNICO'])
#         .size()
#         .unstack(fill_value=0))

#     resumen_estado['TOTAL'] = resumen_estado.sum(axis=1)

#     for estado in ['03', 'Z4', 'NO_ENCONTRO_PDR']:
#         resumen_estado[f'%_{estado}'] = (
#             resumen_estado.get(estado, 0) / resumen_estado['TOTAL'] * 100).round(2)

#     # Crear fila por cada país
#     for pais in resumen_estado.index:
#         fila = resumen_estado.loc[pais]
#         add(f'SKU Status por {pais} PDR', 'Adventencia', f"{pais} → 03: {fila.get('%_03', 0)}%, Z4: {fila.get('%_Z4', 0)}%, Descontinuados: {fila.get('%_NO_ENCONTRO_PDR', 0)}%")
#    else:
#         add(f'SKU Status por {pais} PDR', 'ERROR', 'Faltan columnas STATUS_UNICO o PAIS')
#    # === 9. Indicador por país del estado de los SKU ===
#    if "SKU" in df.columns:
#        # Detectamos Skus que no sean numéricos.
#        sku_alfanumericos = df[~df["SKU"].astype(str).str.isnumeric()]
#        if not sku_alfanumericos.empty:
#            errores += 1
#            filas = (sku_alfanumericos.index + 2).tolist()
#            add("SKU Alfanuméricos", "ERROR", f"Skus alfanuméricos rencontrados -> Filas:{filas[:10]}")
#        else:
#            add("SKU Alfanuméricos", "OK","Todos los Skus son Númericos")
   # === Resultado general ===
   estado = 'Archivo conforme' if errores == 0 else 'Archivo con errores'
   add('Resultado general', estado, None, regla='Consolidado')
   return pd.DataFrame(resultados)

# def enviar_validacion_inventory(df, destionarios,nombre_archivo=None, archivo_adjunto="tabla_indicadores_inventory.png",cc = None):
#    """
#    Enviar correo de los resultados de validacion a los responsable del manual file de SI=SO.

#    Parámetros 
#    df (pandas) :es el manual file de SI=SO que se va a realizar su respectiva validacion.
#    nombre_archivo (str) :es el nombre del archvio al que se realizo la validacion. 
#    destionarios (list str) :es el nombre del responsable del manual file.
#    archivo_adjunto (png) : son los resultados en una imagen de los diferentes indicadores de SI=SO.
#    cc (list str) : es una lista de correos electronicos que se van a copiar los resultados 

#    Retorna
#    Genera un Mail donde tiene los resultados de la revision adjunto y el resultado final si puede subirse o no. 

#    """
#    # Inicializar COM
#    pythoncom.CoInitialize()
#    # Generar imagen del resultado de validación
#    fig, ax = plt.subplots(figsize=(12, 2 + 0.4 * len(df)))
#    ax.axis('off')
#    tabla = ax.table(
#        cellText=df.values,
#        colLabels=df.columns,
#        cellLoc='center',
#        loc='center'
#    )
#    tabla.auto_set_font_size(False)
#    tabla.set_fontsize(10)
#    tabla.scale(1.2, 1.2)
#    plt.savefig(archivo_adjunto, bbox_inches='tight')
#    plt.close()
#    # Evaluar si el archivo pasó todas las validaciones (basado en Resultado general)
#    try:
#        resultado_final = df[df["Indicador"] == "Resultado general"]["Resultado"].iloc[0]
#        archivo_valido = resultado_final.strip().lower() == "archivo conforme"
#    except:
#        archivo_valido = False
#    # Definir el mensaje
#    if archivo_valido:
#        cuerpo_html = (
#            "<b>✅ El archivo pasó todas las validaciones. Puede montarse.</b><br><br>"
#            "<p>Hola,</p>"
#            "<p>Se realizó la validación automática del archivo manual.</p>"
#            "<p>Conclusión: <b style='color:green;'>Archivo conforme</b>.</p>"
#        )
#    else:
#        errores = df[(df["Indicador"] != "Resultado general") & (df["Resultado"].str.lower() != "ok")]["Indicador"].tolist()
#        errores_texto = ", ".join(errores)
#        cuerpo_html = (
#            f"<b>ERROR El archivo {nombre_archivo} tiene observaciones. No se debe montar.</b><br><br>"
#            "<p>Hola,</p>"
#            "<p>Se realizó la validación automática del archivo manual.</p>"
#            f"<p>Conclusión: <b style='color:red;'>Archivo con errores</b>.</p>"
#            f"<p>Errores detectados en: <b>{errores_texto}</b></p>"
#        )
#    # Enviar correo
#    outlook = win32.Dispatch("Outlook.Application")
#    mail = outlook.CreateItem(0)
#    # Convetimos listas en cadenas separadas por punto y coma par mandar correos 
#    if isinstance(destionarios, list):
#        mail.To = ";".join(destionarios)
#    else:
#        mail.To = destionarios
#    if cc:
#        if isinstance(cc, list):
#            mail.cc = ";".join(cc)
#        else:
#            mail.cc = cc        
#    mail.Subject = f"📊 Validación {nombre_archivo}"
#    mail.HTMLBody = (
#        f"{cuerpo_html}"
#        "<p>Se adjunta imagen con los indicadores de validación.</p>"
#        "<p>Saludos,</p>"
#    )
#    # Adjuntar la imagen
#    attachment_path = os.path.abspath(archivo_adjunto)
#    mail.Attachments.Add(attachment_path)
#    # Mostrar el correo (para revisión antes de enviar)
#    mail.Display()

# def enviar_validacion_inventory_prueba(df, destinatarios,nombre_archivo=None, archivo_adjunto= None,cc = None):
#    """
#    Enviar correo de los resultados de validacion a los responsable del manual file de SI=SO.

#    Parámetros 
#    df (pandas) :es el manual file de SI=SO que se va a realizar su respectiva validacion.
#    nombre_archivo (str) :es el nombre del archvio al que se realizo la validacion. 
#    destinatarios (list str) :es el nombre del responsable del manual file.
#    archivo_adjunto (str) : son los resultados de las validaciones del manual file inventory 
#    cc (list str) : es una lista de correos electronicos que se van a copiar los resultados 

#    Retorna
#    Genera un Mail donde tiene los resultados de la revision adjunto y el resultado final si puede subirse o no. 

#    """
#    # Establecer la ruta 
#    carpeta_output = r'C:\Users\IVI6131\OneDrive - MDLZ\Consolidación Información OSA WACAM - General\Manual Files\2025\P6\Output'
#    timestamp = datetime.now().strftime("%Y%m%d")
#    if not archivo_adjunto:
#        archivo_adjunto = f"Resultados_validacion_inventory_{timestamp}.xlsx"
#    ruta_completa_excel = os.path.join(carpeta_output,archivo_adjunto)

#    # Guardar el DataFrame como archivo Excel
#    df.to_excel(ruta_completa_excel,index=False)
#    # Evaluar si el archivo pasó todas las validaciones
#    resultado_final = df[df["Indicador"] == "Resultado general"]["Resultado"].iloc[0]
#    archivo_valido = resultado_final.strip().lower() == "archivo conforme"
#    # Definir el cuerpo del mensaje
#    if archivo_valido:
#        cuerpo_html = f"""
#     <p><b>✅ El archivo pasó todas las validaciones. Puede montarse.</b></p><br>
#     <p>Hola,</p>
#     <p>Se realizó la validación automática del archivo manual.</p>
#     <p>Conclusión: <b style='color:green;'>Archivo conforme</b>.</p>
#        """
#    else:
#        errores = df[
#            (df["Indicador"] != "Resultado general") &
#            (df["Resultado"].str.lower() != "ok")
#        ]["Hallazgo"].tolist()
#        errores_texto = "<br>• " + "<br>• ".join(errores)
#        cuerpo_html = f"""
#     <p><b>ERROR El archivo ({nombre_archivo}) tiene observaciones. No se debe montar.</b></p>
#     <p>Hola,</p>
#     <p>Se realizó la validación automática del archivo manual.</p>
#     <p>Conclusión: <b style='color:red;'>Archivo con errores</b>.</p>
#     <p>Errores detectados:</p>
#        {errores_texto}
#        """
#    # Crear correo en Outlook
#    outlook = win32.Dispatch("Outlook.Application")
#    mail = outlook.CreateItem(0)
#    # Convertir listas en texto separado por punto y coma
#    if isinstance(destinatarios, list):
#        mail.To = ";".join(destinatarios)
#    else:
#        mail.To = destinatarios
#    if cc:
#        if isinstance(cc, list):
#            mail.CC = ";".join(cc)
#        else:
#            mail.CC = cc
#    mail.Subject = f"📋 Validación {nombre_archivo}"
#    mail.HTMLBody = cuerpo_html
#    # Adjuntar el archivo Excel
#    if not os.path.exists(ruta_completa_excel):
#        raise FileNotFoundError(f"No se encontró el archivo para adjuntar:{ruta_completa_excel}")
#    attachment_path = os.path.abspath(ruta_completa_excel)
#    mail.Attachments.Add(attachment_path)
#    # Mostrar correo (revisión manual)

#    mail.Display()

