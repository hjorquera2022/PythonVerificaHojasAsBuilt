#PythonVerificaHojasAsBuilt.py
          
        #ASBUILT

        #ACTUALIZA DOC VIG
        #ACTUALIZA REV LETRA PARCI APRO
        #ACTUALIZA REV NUM PARCI

import pandas as pd
import os

# Ruta base donde se deben verificar los subdirectorios
ruta_base = 'R:\\01 PARCIALIDADES\\'   

# Nombre del archivo de log
archivo_log = ruta_base + '0000-00 ADMINISTRACION\\BAT\\log_ValidarHojasParcialidadesAsBuilt.txt'

# Planilla con la lista de parcialidades
archivo_excel = ruta_base + 'Listado de Parcialidades_AsBuilt.xlsx'

# Carga el archivo Excel en un DataFrame Hoja de Parcialidades.
df = pd.read_excel(archivo_excel, sheet_name='PARCIALIDADES')

# Filtra el DataFrame para considerar solo parcialidades a 'PROCESAR' igual a 'S'
df_parcialidades = df[df['PROCESAR'] == 'S']

# Abre el archivo de log en modo de escritura
with open(archivo_log, 'w') as log_file:

    # Itera a través de cada parcialidad y la procesa
    for parcialidad in df_parcialidades['PARCIALIDAD']:
        log_file.write(f'Parcialidad ASBUILT: {parcialidad}\n')

        #******* 
        #******* RECORRER TODAS LAS PARCIALIDADES VERIFICANDO SI LOS ARCHIVOS DE INGENIERIA CONTIENEN LAS 3 HOJAS DE LA PLANILLA
        #******* 
        #ASBUILT

        #ACTUALIZA DOC VIG
        #ACTUALIZA REV LETRA PARCI APRO
        #ACTUALIZA REV NUM PARCI
         
        
      
        #******* Abrir Planilla CONTROL DOCUMENTOS ING DEF Pxxxx-xx con las 8 hojas para traspasar a BAT
        archivo_parcialidad = ruta_base + parcialidad + '\\CONTROL DOCUMENTOS AS-BUILT P' + parcialidad[0:7] + '.xlsx'
        
        if not os.path.exists(archivo_parcialidad):
              log_file.write(f'Parcialidad ASBUILT: {parcialidad} SIN ARCHIVO DE INGENIERIA {archivo_parcialidad}\n')
        else:
                print(f'Procesando Parcialidad ASBUILT: {parcialidad} ARCHIVO:  {archivo_parcialidad}')
                log_file.write(f'Parcialidad ASBUILT: {parcialidad} Archivo {archivo_parcialidad}\n')

                # Lee el archivo Excel para obtener los nombres de las hojas
                xl = pd.ExcelFile(archivo_parcialidad)
                nombres_hojas = xl.sheet_names

                # Verifica si existen las hojas específicas
                hojas_a_verificar = ['ACTUALIZA DOC VIG', 'ACTUALIZA REV LETRA PARCI APRO', 'ACTUALIZA REV NUM PARCI']
                for hoja in hojas_a_verificar:
                    if not hoja in nombres_hojas:
                        print(f'La hoja ASBUILT "{hoja}" no existe en {archivo_excel}.')
                        log_file.write(f'Parcialidad ASBUILT: {parcialidad} Archivo {archivo_parcialidad}\n')
                
print("Validacion ASBUILD finalizada. Los resultados se han guardado en R:\01 PARCIALIDADES\0000-00 ADMINISTRACION\BAT en el archivo de log_ValidarHojasParcialidadesAsBuilt.")
log_file.close
