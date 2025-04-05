import time
import win32com.client
import pandas as pd
import os
from datetime import datetime

# Abrir SAP Logon
path_saplogon = r"C:\ProgramData\...\"Ejecutable SAP.exe""
os.startfile(path_saplogon)
time.sleep(1)

# Clase de Conexión SAP
class cls_SAP_Gui_Scripting:
    def __init__(self, connection_name):
        sap_gui = win32com.client.GetObject("SAPGUI")
        self.sap_app = sap_gui.GetScriptingEngine
        self.connection = self.sap_app.OpenConnection(connection_name, True)
        self.session = self.connection.Children(0)

# Conexiones
PRODUCCION = "NOMBRE CONEXION SAP PRODUCCION"
PRUEBAS = "NOMBRE CONEXION SAP PRUEBAS"

# Conectar a SAP
sap_connection_name = PRODUCCION
MySAPGUI = cls_SAP_Gui_Scripting(sap_connection_name)
MySAPGUI.session.findById("wnd[0]").maximize()
session = MySAPGUI.session

# Insertar contenido modificado grabadora de SAP: 

#Ejemplo
# session.findById("wnd[0]/tbar[0]/okcd").text = "fbl1n"
# ...

print("El proceso ha terminado. Si es necesario, cierre SAP manualmente.")