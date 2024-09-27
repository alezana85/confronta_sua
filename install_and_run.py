import subprocess
import sys
import os

# Verificar si el archivo requirements.txt existe
requirements_path = os.path.join(os.path.dirname(__file__), 'requirements.txt')
if not os.path.isfile(requirements_path):
    print(f"Error: No se encontró el archivo {requirements_path}")
    sys.exit(1)

# Función para instalar paquetes
def install(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
    except subprocess.CalledProcessError as e:
        print(f"Error al instalar el paquete {package}: {e}")
        sys.exit(1)

# Leer el archivo requirements.txt e instalar cada paquete si no está instalado
try:
    with open(requirements_path, 'r', encoding='utf-8') as f:
        packages = f.read().splitlines()
except Exception as e:
    print(f"Error al leer el archivo {requirements_path}: {e}")
    sys.exit(1)

for package in packages:
    try:
        __import__(package.split('==')[0])
    except ImportError:
        install(package)