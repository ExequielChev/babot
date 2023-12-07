import os
import re
import logging

def sanitize_filename(filename):
    # Expresión regular para eliminar caracteres no deseados
    invalid_chars = r'[<>:"/\\|?*\x00-\x1F]'
    
    # Eliminar caracteres no deseados y reemplazarlos por "_"
    sanitized_filename = re.sub(invalid_chars, "_", filename)
    
    return sanitized_filename

def FiltroExternos(DataFrame, path_salida, Requerido_1, Requerido_1_Parameter,CleanDF):

    logging.info('Sub-Proceso: Filtro Externos.')

    # Filtrado 1 -----------
    
    logging.info(f'Método de filtrado: Columna con encabezado: "{Requerido_1}" Valores a tomar: "{Requerido_1_Parameter}"')
    
    Parameter = Requerido_1_Parameter.lower()

    logging.info(f'Validar existencia de columna: {Requerido_1}')

    if Requerido_1 not in DataFrame.columns:
        logging.info(f'La columna con encabezado "{Requerido_1}" no existe en el DataFrame.')
        return None
    else: 
        logging.info(f'La columna con encabezado "{Requerido_1}" existe en el DataFrame.')

    df_filtrado = DataFrame[DataFrame[Requerido_1].str.lower() == Parameter]  # Filtrar registros externos

    if df_filtrado.shape[0] > 0:
        # Nombre de excel con registros a corregir
        ruta_archivo_salida = os.path.join(path_salida, f"{sanitize_filename(Requerido_1)}_{sanitize_filename(Parameter)}_filtrados.xlsx")
        df_filtrado.to_excel(ruta_archivo_salida, index=False)
        logging.info('"Externos_filtrados.xlsx" generado exitosamente.')

    logging.info(f'Registros en el DataFrame original: {DataFrame.shape[0]}')
    logging.info(f'Registros en el DataFrame filtrado: {df_filtrado.shape[0]}')


    return df_filtrado
