import time
import win32com.client
import pandas as pd
import os

# Leer el archivo Excel
input_file_path = r"C:\...\InputMiro.xlsx"
input_data = pd.read_excel(input_file_path)
Sociedad_SAP = "XX"

# ------------------ Abrir SAP Logon ------------------ #

#...


# Configuración inicial en SAP
MySAPGUI.session.findById("wnd[0]/tbar[0]/okcd").text = "/nMIRO"
MySAPGUI.session.findById("wnd[0]").sendVKey(0)
MySAPGUI.session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").text = Sociedad_SAP
MySAPGUI.session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").caretPosition = 3
MySAPGUI.session.findById("wnd[1]").sendVKey(0)

# Iterar sobre las filas del DataFrame
for index, row in input_data.iterrows():
    # Completar (acciones para contabilizar una factura)
    MySAPGUI.session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-BLDAT").text = row["Fecha"]
    MySAPGUI.session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-XBLNR").text = row["Factura"]
    MySAPGUI.session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/txtINVFO-WRBTR").text = str(row["Importe"])
    MySAPGUI.session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-WAERS").text = row["Moneda"]
    MySAPGUI.session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_TOTAL/ssubHEADER_SCREEN:SAPLFDCB:0010/ctxtINVFO-MWSKZ").text = row["Impuesto"]
    MySAPGUI.session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI").select()

    # Espera y acción adicional al seleccionar la solapa "Detalle"
    time.sleep(2)  # Espera 2 segundos
    MySAPGUI.session.findById("wnd[0]").sendVKey(0)  # Enviar tecla Enter
    time.sleep(1)  # Espera 1 segundo adicional

    # Continuar con el resto de las acciones
    MySAPGUI.session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/tabsHEADER/tabpHEADER_FI/ssubHEADER_SCREEN:SAPLFDCB:0150/cmbINVFO-BLART").key = row["Clase"]
    MySAPGUI.session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subREFERENZBELEG:SAPLMR1M:6214/ctxtRM08M-LBLNI").text = row["HES"]
    MySAPGUI.session.findById("wnd[0]/usr/subHEADER_AND_ITEMS:SAPLMR1M:6005/subITEMS:SAPLMR1M:6010/tabsITEMTAB/tabpITEMS_PO/ssubTABS:SAPLMR1M:6020/subITEM:SAPLMR1M:6310/txtRM08M-SEARCH_STRING").text = ""
    MySAPGUI.session.findById("wnd[0]").sendVKey(0)
    MySAPGUI.session.findById("wnd[0]/tbar[0]/btn[11]").press()

# ------------------ Cerrar SAP ------------------ #
print("El proceso ha terminado. Si es necesario, cierre SAP manualmente.")

# Limpiar variables
MySAPGUI = None
del input_data, index, row