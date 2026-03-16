import ctypes
import os
from gui.app import ExcelXliffConverter

def main():
    if os.name == 'nt':
        myappid = 'roki.excel2xliff.v1'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    app = ExcelXliffConverter()
    app.mainloop()
    
if __name__ == "__main__":
    main()