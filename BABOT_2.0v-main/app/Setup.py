from setuptools import setup, find_packages
import subprocess
import os

version = '1.0'

def build_executable():
    ruta_static = os.path.join(os.getcwd(), "Static")
    ruta_icono = os.path.join(ruta_static, "img", "icono-app.ico")
    ruta_imagen = os.path.join(ruta_static, "img", "icono.png")
    
    subprocess.check_call([
        'pyinstaller', 
        '--name', f'BABOT_{version}', 
        '--add-data', 'Config;Config', 
        '--add-data', 'Modulos;Modulos', 
        f'--add-data={ruta_icono};Static/img',
        f'--add-data={ruta_imagen};Static/img',
        '--windowed',  # Esta opción hace que no se muestre la ventana de la línea de comandos
        f'BABOT_{version}.py'
    ])


def main():
    build_executable()

setup(
    name='BABOT',
    version=version,
    author='Kevin Turkienich',
    author_email='kevin_turkienich@outlook.com',
    description='Asistente virtual | BABOT 2023.',
    packages=find_packages(),
    install_requires=[
        "webencodings==0.5.1",
        "websocket-client==1.5.2",
        "widgetsnbextension==4.0.7",
        "yarl==1.9.2",
        "customtkinter==5.1.3",
        "darkdetect==0.8.0",
        "et-xmlfile==1.1.0",
        "numpy==1.24.3",
        "openpyxl==3.1.2",
        "pandas==2.0.2",
        "Pillow==9.5.0",
        "python-dateutil==2.8.2",
        "pytz==2023.3",
        "six==1.16.0",
        "tzdata==2023.3"
    ],
    entry_points={
        'console_scripts': [
            f'BABOT_{version}=BABOT_{version}:main'
        ]
    }
)

if __name__ == '__main__':
    main()
