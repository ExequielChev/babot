import os
import pandas as pd
import logging


def Unificador(path_entrada,path_salida):

    logging.info('Sub-Proceso: UnificadorDF.Unificador Ha comenzado')

    # Establecer la ruta de la carpeta donde se encuentran los archivos Excel
    ruta_carpeta = path_entrada

    if ruta_carpeta:
        logging.info(f'Carpeta de entrada: {ruta_carpeta}')
    else:
        logging.warning(f'No se ha encontrado la carpeta de entrada. Path: {ruta_carpeta}')


    # Obtener la lista de archivos Excel en la carpeta
    archivos_excel = [archivo for archivo in os.listdir(ruta_carpeta) if archivo.endswith('.xlsx')]

    if archivos_excel:
        logging.info(f'Archivos a procesar: {archivos_excel}')
        validator = True
    else:
        validator = False


    if validator == True:
        logging.info(f'Creando data frame en pandas.')
        # Crear un DataFrame vacío donde se unirán todos los datos de Excel
        df_completo = pd.DataFrame()

        # Recorrer todos los archivos Excel en la carpeta
        for archivo in archivos_excel:

            logging.info(f'Archivo a procesar: {archivo}')
            
            try:
                # Leer el archivo Excel en un DataFrame de pandas
                df_excel = pd.read_excel(os.path.join(ruta_carpeta, archivo))

                # Excluir las últimas dos líneas del archivo Excel
                #df_excel = df_excel.iloc[:-2]

                # Unir los datos del archivo Excel al DataFrame completo
                df_completo = pd.concat([df_completo, df_excel])

                logging.info(f'Procesado exitosamente: {archivo}')

            except Exception as e:
                logging.warning(f'No se pudo procesar el archivo: {archivo}')
                print(f'Error al leer el archivo "{archivo}": {e}')

        logging.info(f'Todos los archivos fueron procesados.')

        # Guardar el DataFrame completo en un nuevo archivo Excel
        ruta_archivo_salida = f'{path_salida}'

        logging.info(f'Creando archivo unificado.')

        try:
            df_completo.to_excel(ruta_archivo_salida, index=False)
            logging.info(f'Archivo unificado creado en ruta: {ruta_archivo_salida}')
        except Exception as e:
            logging.warning(f'No se pudo crear el archivo: {ruta_archivo_salida} - error: {e}')

        logging.info(f'Todos los archivos Excel en "{ruta_carpeta}" se han unido en "{ruta_archivo_salida}".')

        return df_completo
    
    else:

        logging.info(f'No se ha generado el DataFrame debido a que en el path: {ruta_carpeta} No se han encontrado archivos ".xlsx".')
        return False
