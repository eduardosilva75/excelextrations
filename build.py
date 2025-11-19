# build.py
import subprocess
import sys
import os

def install_requirements():
    """Instala as dependências necessárias"""
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])

def create_executable():
    """Cria o executável usando PyInstaller"""
    # Instala o PyInstaller se não estiver instalado
    try:
        import PyInstaller
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Comando para criar o executável
    cmd = [
        'pyinstaller',
        '--name=AnaliseVendas',  # Nome do executável
        '--onefile',             # Arquivo único
        '--windowed',            # Sem console (se for aplicação gráfica)
        '--add-data=requirements.txt;.',  # Inclui requirements
        '--icon=icon.ico',       # Ícone (opcional)
        'main.py'                # Script principal
    ]
    
    subprocess.check_call(cmd)

if __name__ == "__main__":
    install_requirements()
    create_executable()
    print("Build concluído! O executável está na pasta 'dist'")