

def CleanPathOutput(PathOutput,Check):

# --------------------------------------------------------------------------------------------------------------
#  Output Folder Cleanup Process. 
#  
#       Verifica la existencia de una variable en config llamada "CleanPathOutput".
#       Si tiene como valor "TRUE", eliminará todos los archivos de la carpeta antes de empezar el proceso.
# --------------------------------------------------------------------------------------------------------------

    import logging
    import os

# --------------------------------------------------------------------------------------------------------------
#  Example of use. 
#  
#       PathOutput = "C:\Python\BABOT1.0v\Outputs"  <----- folder where the files will be deleted
#       Check = "true"   <----- process activated
#
#       CleanPathOutput(PathOutput=PathOutput,Check=Check) <----- Example of use
# --------------------------------------------------------------------------------------------------------------

    #  Captura de variable
    check = Check.upper()


    if check == 'TRUE':
        
        logging.info(f'Limpieza de carpeta de Output en proceso...')

        pathOutput = PathOutput

        logging.info(f'Se eliminarán todos los archivos del path: {pathOutput}')

        if pathOutput:
            
            conteo = 0
            
            for filename in os.listdir(pathOutput):
                file_path = os.path.join(pathOutput, filename)
                try:
                    os.remove(file_path)
                    logging.info(f'Archivo {file_path} eliminado correctamente.')
                    conteo += 1
                except Exception as e:
                    logging.error(f'Error al eliminar el archivo {file_path}: {e}')
                    conteo += 1
            if conteo == 0:

                logging.info(f'No se han encontrado archivos en la carpeta {pathOutput}')
    else:
        logging.info(f'Para eliminar los archivos debe configurar (KEY: CleanPathOuput, value: true) en el archivo de configuraciones.')


