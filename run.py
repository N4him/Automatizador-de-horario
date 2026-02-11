"""
Punto de entrada para Inno Setup
"""
import sys
import os

# Añadir la carpeta src al path de Python
current_dir = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(current_dir, 'src')
sys.path.insert(0, src_path)

# Ahora importar y ejecutar main
if __name__ == "__main__":
    from src.main import main
    main()