# --------------------------------------------------------------------------------------------------------------
# BABOT - Asistente virtual creado en Python por EXCEL-ENTE 2023
# Desarrollador : Kevin Turkienich
# Contacto : Kevin_turkienich@outlook.com
# --------------------------------------------------------------------------------------------------------------

# --------------------------------------------------------------------------------------------------------------
# Importacion de modulos externos
# --------------------------------------------------------------------------------------------------------------

import logging
import os
import sys
import tkinter as tk
import time
import pandas as pd
from datetime import datetime
from tkinter import messagebox
import customtkinter
from PIL import Image, ImageTk
import re

# Fin de importaciones
# --------------------------------------------------------------------------------------------------------------

# Leer la versión desde el archivo setup.py
with open('Setup.py', 'r') as file:
    setup_contents = file.read()
    version_match = re.search(r"version = '([^']*)'", setup_contents)
    version = version_match.group(1) if version_match else 'N/A'

# Capturar version del paquete
# --------------------------------------------------------------------------------------------------------------

config_bot = {}

def current_date_format(date):
    months = ("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
    day = date.day
    month = months[date.month - 1]
    year = date.year
    message = "{} de {} del {}".format(day, month, year)
    return message

LecturaConfig = False
ProcesoIniciado = False


PathSettings = os.path.join(os.getcwd(), 'config', 'settings.xlsx')

print(f'Path Config: {PathSettings}')

try:
    print(f'Leyendo archivo Config...')
    archivo_excel = pd.ExcelFile(PathSettings)
    df = archivo_excel.parse('Config')
    print(f'Leyendo Variables...')
    for index, row in df.iterrows():
        key = row['KEY']
        value = row['VALUE']
        config_bot[key] = value

    LecturaConfig = True
    print(f'Lectura de Config finalizada correctamente.')

except Exception as e:
    LecturaConfig = False
    print(f'Hubo un error al leer el Config: {e}')
    
# --------------------------------------------------------------------------------------------------------------
# Configuracion de Logs
# --------------------------------------------------------------------------------------------------------------

nameProcess = config_bot['NameRobot']
detailProcess = config_bot['DetailRobot']
userProcess = config_bot['UserRobot']

inicioRobot = time.time()

fechaActual = datetime.now()

PathLogs = os.path.join(config_bot['pathLogs'], str(current_date_format(fechaActual)) + '.log')

FormatLogs = '%(message)s'

logging.basicConfig(level=logging.INFO, format=FormatLogs, filename=PathLogs, filemode='a')

class LogHandler(logging.Handler):
    def __init__(self, log_text):
        super().__init__()
        self.log_text = log_text

    def emit(self, record):
        log_msg = self.format(record)
        self.log_text.insert("end", log_msg + "\n")
        self.log_text.see("end")

# --------------------------------------------------------------------------------------------------------------
# Finalizacion de proceso si no existe SETTINGS
# --------------------------------------------------------------------------------------------------------------

if LecturaConfig == False:
    messagebox.showinfo(title="ERROR DE CONFIG", message='El archivo Settings no existe, por favor verificar la ruta: ' + PathSettings)
    sys.exit()

# --------------------------------------------------------------------------------------------------------------
# Inicio de precesamiento del Robot
# --------------------------------------------------------------------------------------------------------------

def EJECUTAR():

    logging.info(f'  ¡ Bienvenido ! - BABOT 2.0v')
    logging.info(f'')
    logging.info(f'  Comenzando proceso....')
    logging.info(f'')
    logging.info(f'  Proceso: {nameProcess} ')
    logging.info(f'')
    logging.info(f'  Fecha: {fechaActual}')
    logging.info(f'')
    logging.info(f'  Logs: {PathLogs}')
    logging.info(f'')
    logging.info(f'  Settings: {PathSettings}')
    logging.info(f'')

    inicioRobot = time.time()

    logging.info(f'')
    logging.info(f'    ----------------------------------------')
    logging.info(f'    Inicio de proceso')
    logging.info(f'    ----------------------------------------')
# --------------------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------------------
#  Inicio de proceso. 
#  
#       Comienzo de ejecucion de pasos automatizados
# --------------------------------------------------------------------------------------------------------------
    logging.info(f'    Importando modulo Steps..')
    from StepsApp import IniciarProceso
    try:
        IniciarProceso(config_bot)
    except Exception as e:
        logging.info(f'    No se pudo ejecutar la funcion IniciarProceso. Detalle: {e}')

# --------------------------------------------------------------------------------------------------------------
# --------------------------------------------------------------------------------------------------------------
    #  Registro de cierre de robot. 
    logging.info(f'    ----------------------------------------')
    logging.info(f'    Fin de proceso')
    logging.info(f'    ----------------------------------------')

    finRobot = time.time()
    tiempoEjecucion = round( - 1 * (inicioRobot - finRobot), 0)
    if tiempoEjecucion > 60:
        tiempoEjecucionMinutos = tiempoEjecucion/60
        logging.info(f'    Tiempo de ejecución: {tiempoEjecucion} Minutos')
    else:
        logging.info(f'    Tiempo de ejecución: {tiempoEjecucion} Segundos')


    # ------- Pregunta para cerrar proceso o dejar abierta la consola  ----------
    respuesta = messagebox.askyesno("Proceso finalizado correctamente.", "¿Desea cerrar el sistema?")

    if respuesta:
        FIN()
#       Fin del proceso
# --------------------------------------------------------------------------------------------------------------

# ------- Cierre de proceso  ---------------------------------------------------------
def FIN():
    app.quit()

# ------------------------------------------------------------------------------------

# --------------------------------------------------------------------------------------------------------------
# Creacion de ventana de INICIO
# --------------------------------------------------------------------------------------------------------------

def toggle_log_window(switch_var, log_window):
     
    if switch_var.get() == "on":
        log_window.deiconify()
    else:
        log_window.withdraw()

class MyScrollableButtonFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, title, items):
        super().__init__(master, label_text=title)
        self.grid_columnconfigure(0, weight=1)
        self.items = items

        for i, (item, description) in enumerate(self.items.items()):
            button = customtkinter.CTkButton(self, text=item)
            button.grid(row=i, column=0, padx=10, pady=(10, 0), sticky="w")
            button.configure(command=lambda desc=description: messagebox.showinfo("Descripción", desc))

class App(customtkinter.CTk):

    # Diccionario de pasos del proceso
    pasos = {
            "Paso 1": 'Descripcion paso 1',
            "Paso 2": 'Descripcion paso 2',
            "Paso 3": 'Descripcion paso 3',
            "Paso 4": 'Descripcion paso 4'
        }
    

    def __init__(self):
        super().__init__()

        self.title(f"BABOT | v{version}")
        self.geometry("800x600")
        self.resizable(False, False)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)

# -------------------------------------------------------------------------------------------
#   Ventana de Log
# -------------------------------------------------------------------------------------------

        log_window = customtkinter.CTkToplevel()
        log_window.title("Logs del Proceso")
        log_window.geometry("500x400")
        log_window.configure(bg="black")
        log_window.protocol("WM_DELETE_WINDOW", lambda: None)

        # Obtener las dimensiones de la pantalla
        screen_width = log_window.winfo_screenwidth()
        screen_height = log_window.winfo_screenheight()

        # Calcular las coordenadas para ubicar la ventana en la parte derecha
        window_width = 500
        window_height = 400
        x = ((screen_width - window_width) // 2 ) + 500
        y = ((screen_height - window_height) // 2 ) 

        log_text = tk.Text(log_window, bg="black", fg="green")
        log_text.pack(fill="both", expand=True)

        log_handler = LogHandler(log_text)
        logging.getLogger().addHandler(log_handler)

        log_window.withdraw()

# -------------------------------------------------------------------------------------------
#   Componentes de UI
# -------------------------------------------------------------------------------------------

        # Construir la ruta completa del archivo de icono
        ruta_icono = os.path.join(os.getcwd(),"static", "img", "icono-app.ico")

        # Establecer el icono de la aplicación principal
        self.iconbitmap(ruta_icono)

        # Establecer el icono de la ventana de registro
        log_window.iconbitmap(ruta_icono)

        frameLogo = customtkinter.CTkFrame(self, width=100)
        frameLogo.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        label = customtkinter.CTkLabel(frameLogo, text=f"BABOT | Asistente virtual", width=10, font=("Arial", 16, "bold"))
        label.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="w")

        label2 = customtkinter.CTkLabel(frameLogo, text="Detalle de proceso:", width=10, font=("Arial", 12, "italic"))
        label2.grid(row=2, column=0, columnspan=2, padx=10, pady=1, sticky="w")

        label2 = customtkinter.CTkLabel(frameLogo, text=f"{detailProcess}", width=10, font=("Arial", 12), wraplength=500, justify="left")
        label2.grid(row=3, column=0, columnspan=1, padx=10, pady=3, sticky="w")

        ruta_imagen = os.path.join(os.getcwd(),"static", "img", "icono.png")
        logo_image = Image.open(ruta_imagen)
        logo_image = logo_image.resize((200, 200),)
        logo = ImageTk.PhotoImage(logo_image)

        logo_label = tk.Label(self, image=logo, text="",bg="#0097B2")
        logo_label.image = logo
        logo_label.grid(row=0, column=2, padx=15, pady=15)

        frame = customtkinter.CTkFrame(self)
        frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        label = customtkinter.CTkLabel(frame, text="Variables definidas en archivo de configuracion",font=("Arial", 12,"italic"))
        label.grid(row=0, column=0, padx=10, pady=0, sticky="w")

        label1 = customtkinter.CTkLabel(frame, text="📂 Input",font=("Arial", 11,"bold"))
        label1.grid(row=1, column=0, padx=10, pady=0, sticky="w")

        textbox1 = customtkinter.CTkTextbox(frame, width=500, height=20)
        textbox1.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        textbox1.insert(tk.END, config_bot.get('pathInput', ''))
        textbox1.configure(state="disabled",)

        label2 = customtkinter.CTkLabel(frame, text="📂 Output",font=("Arial", 11,"bold"))
        label2.grid(row=3, column=0, padx=10, pady=0, sticky="w")

        textbox2 = customtkinter.CTkTextbox(frame, width=500, height=20)
        textbox2.grid(row=4, column=0, padx=10, pady=5, sticky="w")
        textbox2.insert(tk.END, config_bot.get('pathOutput', ''))
        textbox2.configure(state="disabled")

        label3 = customtkinter.CTkLabel(frame, text="📂 Logs",font=("Arial", 11,"bold"))
        label3.grid(row=5, column=0, padx=10, pady=0, sticky="w")

        textbox3 = customtkinter.CTkTextbox(frame, width=500, height=20)
        textbox3.grid(row=6, column=0, padx=10, pady=5, sticky="w")
        textbox3.insert(tk.END, config_bot.get('pathLogs', ''))
        textbox3.configure(state="disabled")

        switch_var = customtkinter.StringVar(value="off")
        switch = customtkinter.CTkSwitch(frame, text="Ver venana de Logs", variable=switch_var, onvalue="on", offvalue="off")
        switch.grid(row=7, column=0, padx=10, pady=20, sticky="se")
        switch_var.trace("w", lambda *args: toggle_log_window(switch_var, log_window))
                
        

        scrollable_button_frame = MyScrollableButtonFrame(self, title="Detalle del proceso", items=self.pasos)
        scrollable_button_frame.grid(row=1, column=2, padx=10, pady=10, sticky="nsew")

        procesar_button = customtkinter.CTkButton(self, text="Iniciar",hover_color="green",width=200,height=40, font=("Arial", 16, "bold"),command=EJECUTAR)
        procesar_button.grid(row=2, column=0, padx=20, pady=10, sticky="w")

        salir_button = customtkinter.CTkButton(self, text="Salir",hover_color="red",width=200,height=40,font=("Arial", 16, "bold"),command=FIN)
        salir_button.grid(row=2, column=2, padx=20, pady=10, sticky="e")


app = App()
app.mainloop()

# -------------------------------------------------------------------------------------------
#   Cierre App()
# -------------------------------------------------------------------------------------------
