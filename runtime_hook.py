# runtime_hook.py — Se ejecuta antes de que arranque la aplicación cuando
# está empaquetada con PyInstaller. Resuelve problemas de rutas en modo frozen.

import sys
import os

if getattr(sys, 'frozen', False):
    # Directorio donde PyInstaller extrae todos los archivos
    _base = sys._MEIPASS

    # tkinterdnd2 necesita saber dónde están sus DLLs nativas
    os.environ.setdefault('TKDND_LIBRARY', _base)

    # Garantizar que el directorio base esté en PATH para DLLs de Windows
    os.environ['PATH'] = _base + os.pathsep + os.environ.get('PATH', '')
