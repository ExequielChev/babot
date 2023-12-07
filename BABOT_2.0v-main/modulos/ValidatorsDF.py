
# --------------------------------------------------------------------------------------------------------------
# BABOT - Asistente virtual creado en Python por EXCEL-ENTE 2023
# Desarrollador : Kevin Turkienich
# Contacto : Kevin_turkienich@outlook.com
# --------------------------------------------------------------------------------------------------------------

# --------------------------------------------------------------------------------------------------------------
# Importacion de modulos
# --------------------------------------------------------------------------------------------------------------

import os
import logging
import pandas as pd

# --------------------------------------------------------------------------------------------------------------
# Funcion: ValidacionNulos()
# --------------------------------------------------------------------------------------------------------------

def ValidacionDatos(Export,DataFrame,path_salida,Requerido_0,Requerido_1,Requerido_2,Requerido_3,Requerido_4,Requerido_5,Requerido_6,Requerido_7,Requerido_8,Requerido_9,CleanDF):

    logging.info('Sub-Proceso: Validacion de datos ha comenzado.')

    if DataFrame is None or DataFrame.empty:
        logging.warning('El DataFrame que intenta validar es nulo o vacío. Saliendo del Sub-Proceso: Validacion de datos.')
        return DataFrame

    logging.info('Test:')
    logging.info(f'Columna requerida 0:{Requerido_0}')
    logging.info(f'Columna requerida 1:{Requerido_1}')
    logging.info(f'Columna requerida 2:{Requerido_2}')
    logging.info(f'Columna requerida 3:{Requerido_3}')
    logging.info(f'Columna requerida 4:{Requerido_4}')
    logging.info(f'Columna requerida 5:{Requerido_5}')
    logging.info(f'Columna requerida 6:{Requerido_6}')
    logging.info(f'Columna requerida 7:{Requerido_7}')
    logging.info(f'Columna requerida 8:{Requerido_8}')
    logging.info(f'Columna requerida 9:{Requerido_9}')
    logging.info('Fin Test')
#-----------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------
# Validando datos de entrada - Busca entre las variables Requerido_0,Requerido_1,Requerido_2,etc. del Config, y captura todos los valores de los Headers del DataFrame a evaluar
# En caso de que difa "False" no se contemplará ese header.

    logging.info('Validando datos de entrada...')

    requeridos = [valor for nombre, valor in locals().items() if nombre.startswith('Requerido_') and valor != "False"]

    columnas_requeridas = len(requeridos)

    if columnas_requeridas > 0:
        logging.info(f'Se detectaron {columnas_requeridas} columnas requeridas, recuerde que puede configurar las columnas requeridas en el archivo de configuracion > apartado "Columnas requeridas"')
    else:
        logging.info(f'Se detectaron {columnas_requeridas} columnas requeridas, Sub-Proceso Validacion terminado debido a que no existen columnas requeridas en Config diferente a "False".')
        return
    
    # Validar la existencia de headers requeridos
    headers_faltantes = [requerido for requerido in requeridos if requerido not in DataFrame.columns]

    if headers_faltantes:
        logging.info(f'')
        logging.info(f'**********************************************************************************************************************')
        logging.info(f'ADVERTENCIA: Los siguientes headers indicados en el archivo Config.xslx, no se encuentran en el DataFrame de entrada')
        logging.info(f'Headers:')
        logging.info(f'{headers_faltantes}')
        logging.info(f'')
        logging.info(f'Los headers indicados en la advertencia no podrán ser procesados, se procede a evaluar los headers coincidentes.')
        logging.info(f'**********************************************************************************************************************')
        logging.info(f'')
    else:
        logging.info(f'Todos los Headers requeridos existen en el DataFrame de entrada.')

    logging.info(f'Se evaluarán las celdas vacías de las columnas cuyos headers son: {requeridos}')

#-----------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------
# Validar que exista columna de Comentarios en el DataFrame, caso que no exista, se crea al final del DF.

    logging.info(f'En caso que el DataFrame no contenga la columna "Comentarios", se procede a crearla al final..')
    # Verificar si la columna "Comentarios" no existe en el DataFrame
    if 'Comentarios' not in DataFrame.columns:
    # Crear la columna "Comentarios" con valor inicial vacío si no existe
        DataFrame = DataFrame.assign(Comentarios='')

    Validador = True
#-----------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------
# Validar registros vacíos en las columnas indicadas en columnas_requeridas

    requeridos = [valor for nombre, valor in locals().items() if nombre.startswith('Requerido_') and valor != "False" and valor in DataFrame.columns]

    for requerido in requeridos:
        logging.info(f'Método de Validacion: Columna con encabezado: "{requerido}" Valores a buscar: "Vacios"')
        registros_vacios = DataFrame[DataFrame[requerido].isnull() | (DataFrame[requerido] == '')]
        if not registros_vacios.empty:
            logging.info(f'Se encontraron registros vacíos en la columna "{requerido}".')
            DataFrame.loc[registros_vacios.index, 'Comentarios'] += f'| Valor esperado en Columna: "{requerido}" |'

            Validador = False
        else:
            logging.info(f'Método de Validacion: Columna con encabezado: "{requerido}" Valores a buscar: "Vacios" | Correcta')
            logging.info(f'No se encontraron registros vacíos en la columna "{requerido}".')

#-----------------------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------
# Validar ClieanDF : Si ClieanDF se setea como True en Config, se exportara el DataFrame con comentarios.

    logging.info(f'Resultado Sub-Proceso: Validacion de datos')
    if Validador == True:
        logging.info(f'Todos los campos fueron validados correctamente.')
    else:
        ruta_archivo_salida = os.path.join(path_salida, "Registros_con_comentarios.xlsx")
        logging.info(f'Las validaciones no fueron satisfactorias, por favor revise el archivo >{ruta_archivo_salida}<, en la columna "Comentarios" quedaran registrados los campos requeridos.')


    if Export.lower() =="true":


    
        if CleanDF=="False":
            if Validador == True:
                df_validado = DataFrame[DataFrame['Comentarios'] == '']
                ruta_archivo_salida = os.path.join(path_salida, "Archivo_Validado.xlsx")
                df_validado.to_excel(ruta_archivo_salida, index=False)
                logging.info(f'Se generó el archivo "{ruta_archivo_salida}" con los registros validados.')
                logging.info('Sub-Proceso: Validacion de datos ha finalizado.')
            else:
                ruta_archivo_salida = os.path.join(path_salida, "Registros_con_comentarios.xlsx")
                DataFrame.to_excel(ruta_archivo_salida, index=False)
                logging.info(f'Se generó el archivo "{ruta_archivo_salida}" con los registros validados y comentarios.')
                logging.info('Sub-Proceso: Validacion de datos ha finalizado.')
                return DataFrame
        else:
            # Filtrar DataFrame por registros con Comentarios vacíos
            df_validado = DataFrame[DataFrame['Comentarios'] == '']
            ruta_archivo_salida = os.path.join(path_salida, "Archivo_Validado.xlsx")
            df_validado.to_excel(ruta_archivo_salida, index=False)
            logging.info(f'Se generó el archivo "{ruta_archivo_salida}" con los registros validados.')
            logging.info('Sub-Proceso: Validacion de datos ha finalizado.')
            return df_validado
    else:
        if CleanDF=="False":
            logging.info('Proceso finalizado, se retorna el DataFrame.')
            return DataFrame
            logging.info('Sub-Proceso: Validacion de datos ha finalizado.')
        else:
            # Filtrar DataFrame por registros con Comentarios vacíos
            df_validado = DataFrame[DataFrame['Comentarios'] == '']
            logging.info('Proceso finalizado, se retorna el DataFrame validado (Con los registros validados).')
            logging.info('Sub-Proceso: Validacion de datos ha finalizado.')
            return df_validado



# --------------------------------------------------------------------------------------------------------------
# Funcion: ValidacionDatosExactos()
# --------------------------------------------------------------------------------------------------------------

def ValidacionIguales(
        Export,
        DataFrame,
        path_salida,
        CleanDF,
        **kwargs
):      
    
    logging.info('Sub-Proceso: Validacion de datos exactos ha comenzado.')
    
    if DataFrame is None:
        logging.info('El DataFrame proporcionado no tiene registros para procesar.')
        logging.info('Se finaliza el Sub-Proceso: Validacion de datos exactos')
        return False

    if len(DataFrame) > 0:
        # El DataFrame tiene registros
        logging.info('')
    else:
        # El DataFrame está vacío
        logging.info('El DataFrame proporcionado no tiene registros para procesar.')
        logging.info('Se finaliza el Sub-Proceso: Validacion de datos exactos')
        return False
    
    

    logging.info('Test:')
    for i in range(10):
        requerido = kwargs.get(f'Requerido_{i}', None)
        valor_requerido = kwargs.get(f'Valor_Requerido_{i}', None)
        logging.info(f'Columna requerida {i}: {requerido}')
        logging.info(f'Valor esperado para la Columna requerida {i}: {valor_requerido}')
    logging.info('Fin Test')

    # ----------------------------------------------------------------------------------------------
    # Validando datos de entrada - Busca las variables Requerido_0, Requerido_1, Requerido_2, etc.
    # del Config y captura todos los valores de los Headers del DataFrame a evaluar.
    # En caso de que diga "False" no se contemplará esa validación.
    # ----------------------------------------------------------------------------------------------

    logging.info('Validando datos de entrada...')

    requeridos = [valor for nombre, valor in kwargs.items() if nombre.startswith('Requerido_') and valor != "False"]

    if requeridos is not None:
        columnas_requeridas = len(requeridos)
    else:
        logging.error('No se encontraron columnas requeridas en los datos de entrada.')
        return DataFrame
        

    logging.info(f'Se encontraron {columnas_requeridas} columnas requeridas.')


    # ----------------------------------------------------------------------------------------------
    # Validar que exista columna de Comentarios en el DataFrame, caso que no exista, se crea al final del DF.
    # ----------------------------------------------------------------------------------------------

    logging.info(f'En caso que el DataFrame no contenga la columna "Comentarios", se procede a crearla al final..')
    # Verificar si la columna "Comentarios" no existe en el DataFrame
    if 'Comentarios' not in DataFrame.columns:
    # Crear la columna "Comentarios" con valor inicial vacío si no existe
       DataFrame = DataFrame.assign(Comentarios='')

    # ----------------------------------------------------------------------------------------------
    # Validando que los Headers requeridos estén presentes en el DataFrame
    # ----------------------------------------------------------------------------------------------

    logging.info('Validando Headers requeridos en DataFrame...')

    headers_requeridos = []

    for requerido in requeridos:
        if requerido in DataFrame.columns:
            headers_requeridos.append(requerido)
        else:
            logging.info(f'Header requerido "{requerido}" no encontrado en DataFrame.')



    logging.info(f'Se procede a validar los datos de los headers encontrados. Headers encontrados: {headers_requeridos}')

    # ----------------------------------------------------------------------------------------------
    # Realizando la validación de datos exactos
    # ----------------------------------------------------------------------------------------------



    
    logging.info('Iniciando validación de datos exactos...')

    for requerido, valor_requerido in zip(headers_requeridos, [kwargs.get(f'Valor_Requerido_{i}', None) for i in range(columnas_requeridas)]):

        if CleanDF:
            DataFrame[requerido] = DataFrame[requerido].astype(str).str.strip()

        validacion_igualdad = DataFrame[requerido] == str(valor_requerido)

        try:
            # Validar si el valor es un número
            if str(valor_requerido).isdigit():
                validacion_igualdad |= DataFrame[requerido].astype(str).str.isdigit() & (DataFrame[requerido].astype(str) == str(valor_requerido))
        except ValueError:
            pass  # Ignorar el error si el valor no se puede convertir a número

        if not validacion_igualdad.all():
            registros_incorrectos = DataFrame.loc[~validacion_igualdad]

            logging.error(f'Error de validación de datos exactos en columna "{requerido}".')
            logging.error(f'Registros incorrectos encontrados:\n{registros_incorrectos}')

            # Agregar comentario en la columna "Comentarios"
            DataFrame.loc[~validacion_igualdad, 'Comentarios'] += f'| Valor esperado en Columna: "{requerido}"  Valor:"{valor_requerido}" |'

    logging.info('Validación de datos exactos completada con éxito.')
            

    # ----------------------------------------------------------------------------------------------
    # Guardando el DataFrame validado en el archivo de salida
    # ----------------------------------------------------------------------------------------------
    
    # La exportacion de datos del DataFrame puede configurarse desde el archivo Config.xlsx. Parametros True/False

    if Export.lower() =="true":

        # Validar CleanDF_Obligatory: Si CleanDF_Obligatory se establece como True, se exportará el DataFrame con comentarios.
        logging.info(f'Resultado Sub-Proceso: Validacion de datos exactos')
        if str(CleanDF)=="False":
            ruta_archivo_salida = os.path.join(path_salida, "Registros_Validacion_Exactos_con_comentarios.xlsx")
            DataFrame.to_excel(ruta_archivo_salida, index=False)
            logging.info(f'Se generó el archivo "{ruta_archivo_salida}" con los registros y sus comentarios.')
            logging.info('Sub-Proceso: Validacion de datos ha finalizado.')
            return DataFrame
        else:
            # Filtrar DataFrame por registros con Comentarios vacíos para limpiar el DataFrame
            df_validado = DataFrame[DataFrame['Comentarios'] == '']
            ruta_archivo_salida = os.path.join(path_salida, "Archivo_Validado_Exactos.xlsx")
            df_validado.to_excel(ruta_archivo_salida, index=False)
            logging.info(f'Se generó el archivo "{ruta_archivo_salida}" con los registros exactos validados.')
            logging.info('Sub-Proceso: Validacion de datos exactos ha finalizado.')
            return df_validado
    else:

        logging.info(f'Resultado Sub-Proceso: Validacion de datos exactos')
        if CleanDF=="False":
            logging.info('Sub-Proceso: Validacion de datos ha finalizado.')
            return DataFrame
        else:
            logging.info('Sub-Proceso: Validacion de datos exactos ha finalizado.')
            return df_validado