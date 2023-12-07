# --------------------------------------------------------------------------------------------------------------
# BABOT - Asistente virtual creado en Python por EXCEL-ENTE 2023
# Desarrollador : Kevin Turkienich
# Contacto : Kevin_turkienich@outlook.com
# --------------------------------------------------------------------------------------------------------------
    # En esta seccion se programan todos los pasos del asistente
# --------------------------------------------------------------------------------------------------------------
# Importacion de modulos externos
# --------------------------------------------------------------------------------------------------------------
import logging
import pandas as pd
# --------------------------------------------------------------------------------------------------------------
# Importacion de modulos funciones
# --------------------------------------------------------------------------------------------------------------
from Functions.funcionesBabot import *
from Functions.funcionesCustom import *

# --------------------------------------------------------------------------------------------------------------

# --------------------------------------------------------------------------------------------------------------
# Funcion de inicio
# --------------------------------------------------------------------------------------------------------------

def IniciarProceso(config_bot):

    Path_Pasos_bot = config_bot['pathSteps']
    Sheet_Pasos_bot = config_bot['nameSheetSteps']

    pasos_bot = {}

# --------------------------------------------------------------------------------------------------------------
# Lectura de archivo de pasos (si no existe finalizar)
# --------------------------------------------------------------------------------------------------------------
    print(f'Leyendo Pasos...')
    try:
        archivo_excel = pd.ExcelFile(Path_Pasos_bot)
        df = archivo_excel.parse(Sheet_Pasos_bot)

        for index, row in df.iterrows():
            key = row['PASO']
            value = row['FUNCION']
            pasos_bot[key] = value
        print(f'Lectura Exitosa.')

    except Exception as e:
        logging.info(f'    No se pudo leer pasos del asistente. Detalle del error: {e}')
        print(f'No se pudo leer pasos del asistente. Detalle del error: {e}')
        print(f'Finalizando Pasos...')
        return None
# --------------------------------------------------------------------------------------------------------------
#  BUSQUEDA DE PASOS
# --------------------------------------------------------------------------------------------------------------

    for i in range(10):
        
        paso = pasos_bot[f'PASO_{i}']

        logging.info(f'')

# --------------------------------------------------------------------------------------------------------------
#  EJECUCION DE PASOS DEL ASISTENTE
# --------------------------------------------------------------------------------------------------------------
            
        # Ejecutar la función correspondiente al paso
        funcion = globals().get(paso) or locals().get(paso)
        if funcion is not None and callable(funcion):
            funcion(pasos_bot)
        else:
            if str(paso) != "SIN_ASIGNAR":
                logging.warning(f'    No se encontró la función {paso}')
# --------------------------------------------------------------------------------------------------------------
#  FUNCIONES DE EJECUCION DE PASOS DEL ASISTENTE
# --------------------------------------------------------------------------------------------------------------
    pasos_asignados = sum(1 for paso in pasos_bot.values() if paso != "SIN_ASIGNAR")


    logging.info(f'    Pasos asignados: {pasos_asignados}')

