import sys
import os

# Agregar la carpeta src al PYTHONPATH para permitir la importación
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.join(current_dir, 'src')
sys.path.insert(0, src_dir)

from gui.app import Application

if __name__ == '__main__':
    app = Application()
    app.mainloop()
