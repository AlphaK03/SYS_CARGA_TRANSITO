# src/main.py
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from gui.app import Application

if __name__ == "__main__":
    app = Application()
    app.mainloop()
