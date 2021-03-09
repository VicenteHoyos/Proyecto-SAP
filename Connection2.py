import win32com.client
import sys
import subprocess
import time
from tkinter import *
from framework import Runnable
from framework import Transaction
import win32clipboard
import win32com.client
import sys
import subprocess
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


##This function will Login to SAP from the SAP Logon window

class Connect:
    '''The intention of this code is to create or connect
    to a SAP connection and return a list of the present sessions.'''

    def __init__(self):
        self.process = None
        self.sapguiauto = None
        self.application = None
        self.connection = None
        self.session = []
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"  # ***The full path should be placed here
        self.user = "ITCO25708"
        self.password = "951211hoyoscastaño26304"
        self.language = None
        self.connection = None
        self.connect_to = '2 - QH8 - ERP Calidad isaeccqh8'

    def open_sap(self, user, password, language, connect_to='2 - QH8 - ERP Calidad isaeccqh8'):
        self.user = user
        self.password = password
        self.language = language
        self.connect_to = connect_to
        self.process = subprocess.Popen(self.path)
        time.sleep(4)  # time sleep so the computer has time to open SAP.

        # Connecting to the SAP API
        self.sapguiauto = win32com.client.GetObject('SAPGUI')
        if not type(self.sapguiauto) == win32com.client.CDispatch:
            return
        self.application = self.sapguiauto.GetScriptingEngine
        if not type(self.application) == win32com.client.CDispatch:
            self.sapguiauto = None
            return

        # Checks if user is already logged in SAP in the computer used. If so, it uses the current connection
        if len(self.application.Children) > 0:
            for con in range(0, len(self.application.Children)):
                self.connection = self.application.Children(con)
                if not type(self.connection) == win32com.client.CDispatch:
                    self.application = None
                    self.sapguiauto = None
                    return
                if self.connection.Children(0).Info.User == self.user:
                    for i in range(0, len(self.connection.Children)):
                        self.session.append(self.connection.Children(i))
                    return
            self.login()
            return

        # If not, creates a new login session
        else:
            self.login()
            return

    def login(self):
        self.connection = self.application.Openconnection(self.connect_to, True)

        if not type(self.connection) == win32com.client.CDispatch:
            self.application = None
            self.sapguiauto = None
            return

        self.session.append(self.connection.Children(0))
        self.session[0].findById("wnd[0]/usr/txtRSYST-BNAME").text = self.user
        self.session[0].findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password
        self.session[0].findById("wnd[0]/usr/txtRSYST-LANGU").text = self.language
        self.session[0].findById("wnd[0]").sendVKey(0)
        while len(self.session[0].Children) == 2:
            try:
                self.session[0].findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select()
            except:
                pass
            self.session[0].findById("wnd[1]").sendVKey(0)

    # Verifies the amount of open sessions, open a new one and appends it to the list os sessions.
    def new_session(self):

        open_sessions = len(self.connection.Children)
        self.session[0].createsession()
        time.sleep(1)
        self.session.append(self.connection.Children(open_sessions))
        return

    def disconnect(self):
        self.connection.CloseConnection()

    # Forces entry when we are facing warning messages while trying to go to the next step.
    def force_entry(self, session_num):
        while self.session[session_num].findById("wnd[0]/sbar/").text != '':
            self.session[session_num].findById("wnd[0]").sendVKey(0)

    # Forces entry on possible warning Popup screens.
    def force_popup(self, session_num):
        while True:
            try:
                self.session[session_num].findById("wnd[1]").sendVKey(0)
            except Exception:
                break
    
    def logoff(self):
        self.session[0].findById("wnd[0]/tbar[0]/btn[15]").press()
        self.session[0].findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    def transaccion_zpm_pt_2(self): 
        
        self.session[0].findById("wnd[0]").maximize
        self.session[0].findById("wnd[0]/tbar[0]/okcd").text = "ZPM_PT"
        self.session[0].findById("wnd[0]").sendVKey (0)
        #session.findById("wnd[0]/tbar[0]/btn[0]").press
        self.session[0].findById("wnd[0]/usr/chkP_EST02").selected = True
        self.session[0].findById("wnd[0]/usr/chkP_EST03").selected = True
        self.session[0].findById("wnd[0]/usr/chkP_EST04").selected = True
        self.session[0].findById("wnd[0]/usr/chkP_EST05").selected = True
        self.session[0].findById("wnd[0]/usr/ctxtP_IWERK").text = "ITCO"
        self.session[0].findById("wnd[0]/usr/txtS_ZWEEKR-LOW").text = "1"
        self.session[0].findById("wnd[0]/usr/txtS_ZWEEKR-HIGH").text = "52"
        self.session[0].findById("wnd[0]/usr/txtS_ZZANO-LOW").text = "2020"
        self.session[0].findById("wnd[0]/usr/ctxtS_ZZTCON-LOW").text = "L"
        self.session[0].findById("wnd[0]/usr/ctxtS_ZZTCON-HIGH").text = "N"
        self.session[0].findById("wnd[0]/usr/ctxtS_ZZTCON-HIGH").setFocus
        self.session[0].findById("wnd[0]/usr/ctxtS_ZZTCON-HIGH").caretPosition = 1
        self.session[0].findById("wnd[0]/usr/btn%_S_INGOPE_%_APP_%-VALU_PUSH").press()
        self.session[0].findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "10008"
        self.session[0].findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "3831"
        self.session[0].findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "3871"
        self.session[0].findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").SetFocus
        self.session[0].findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
        self.session[0].findById("wnd[1]/tbar[0]/btn[8]").press()
        self.session[0].findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").setCurrentCell (-1,"ZZCONSIGE")
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").selectColumn("ZZCONSIGE")
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").contextMenu()
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").selectContextMenuItem ("&FILTER")
        self.session[0].findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "*A*"
        self.session[0].findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 3
        self.session[0].findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").pressToolbarContextButton ("&MB_EXPORT")
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").selectContextMenuItem ("&XXL")
        self.session[0].findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session[0].findById("wnd[1]/tbar[0]/btn[11]").press()
        self.session[0].findById("wnd[0]/tbar[0]/btn[15]").press()

def my_sap_script():
    connect = Connect()
    session = connect.open_sap("ITCO25708","951211hoyoscastaño26304","ES")

def transaccion_zpm_pt():
    connect = Connect()
    session = connect.open_sap("ITCO25708","951211hoyoscastaño26304","ES")
    connect.transaccion_zpm_pt_2()

def logoff_SAP():
    connect = Connect()
    session = connect.open_sap("ITCO25708","951211hoyoscastaño26304","ES")
    connect.logoff()

def File_dialog():
    """Esta función abrirá el explorador de archivos y asignará la ruta de archivo elegida a label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Seleccione un Archivo",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file["text"] = filename
    
    return None

def Load_excel_data():
    """Si el archivo seleccionado es válido, se cargará el archivo en la Treeview."""
    file_path = label_file["text"]
    
    excel_filename = r"{}".format(file_path)
    
    file_path = pd.read_excel(excel_filename)

    file_path.to_csv (r'C:\Temp\EXPORT.csv', index = None, header=True)

    file_path = r'C:\Temp\EXPORT.csv'
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

    except ValueError:
        tk.messagebox.showerror("Información", "El archivo que ha elegido no es válido")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Información", f"No existe un archivo como {file_path}")
        return None

    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # Encabezado de la columna = column name

    df_rows = df.to_numpy().tolist() # convierte el marco de datos (data frrame) en una lista de listas
    for row in df_rows:
        tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
    return None

def clear_data():
    tv1.delete(*tv1.get_children())
    return None



def load_clipboard():
    try:

        cols = ["Planes Trabajo",
          'Número Consigna',
          'Denominación Ubicación Técnica',
          'Clasificación',
          'Denominación de la revisión',
          'Estado',
          'Descripción Estado',
          'Fecha Inicio',
          'Hora Inicio',
          'Fecha Fin',
          'Hora Fin',
          'Jefe de Trabajo',
          'Tipo Origen',
          'Origen'                             
          ]

        clipboardData = pd.read_clipboard(sep=' ') #pd.read_clipboard(sep='\s+', index_col=[0], header=[0,1])
        
        tv2["columns"] = cols
        tv2["show"] = "headings"

        for column in tv2["columns"]:            
            tv2.heading(column, text=column) # Encabezado de la columna = column name
        
        cols = list(clipboardData.columns)
        tv2.insert ("", END,text = 0 , values = cols)

        for index, row in clipboardData.iterrows():
            tv2.insert("",END,text=index,values=list(row))
        
        
        #print(clipboardData)
        #print(cols)
        
        time.sleep(1)
    except:
        selection = None
    return None
    
def Clear_Estadosoperativos():
    tv2.delete(*tv2.get_children())
    return None 
    

# initalise the tkinter GUI
root = tk.Tk()
root.title("Cortes Visibles") #Cambiar el nombre de la ventana
root.geometry("500x500") # Establecer dimenciones
root.wm_state('zoomed')
root.pack_propagate(False) # le dice a la ventana que no permita que los widgets de su interior determinen su tamaño.
root.resizable(0, 0) # hace que la ventana raíz tenga un tamaño fijo.

# Frame para planes de trabajo TreeView
frame1 = tk.LabelFrame(root, text="Planes de trabajo")
frame1.place(height=250, width=600)

# Frame para planes de trabajo TreeView
frame2 = tk.LabelFrame(root, text="Estados operativos")
frame2.place(height=250, width=730, x=615, y =0)

# Frame para planes de trabajo TreeView
frame3 = tk.LabelFrame(root, text="Crear Cortes Visibles")
frame3.place(height=250, width=1335, x=10, y =350)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Abrir Archivo")
file_frame.place(height=80, width=570, rely=0.35, relx=0.01)

# Frame
file_frame2 = tk.LabelFrame(root, text="Copiar Estados Operativos")
file_frame2.place(height=80, width=700, rely=0.35, relx=0.01, x=615, y =0)

# Frame 
file_frame3 = tk.LabelFrame(root, text="Cortes Visibles")
file_frame3.place(height=80, width=700, rely=0.35, relx=0.01, x=0.5, y =345)

# Buttons
button1 = tk.Button(file_frame, text="Buscar un Archivo", command=lambda: File_dialog())
button1.place(rely=0.45, relx=0.55)

button2 = tk.Button(file_frame, text="Cargar Archivo", command=lambda: Load_excel_data())
button2.place(rely=0.45, relx=0.20)

# The file/file path text
label_file = ttk.Label(file_frame, text="Ningún archivo seleccionado")
label_file.place(rely=0, relx=0)



# Buttons

button3 = tk.Button(file_frame2, text="Cargar ", command=lambda: load_clipboard())
button3.place(rely=0.45, relx=0.20, height=25, width=120)

button4 = tk.Button(file_frame2, text="Limpiar Estados Operativos", command=lambda: Clear_Estadosoperativos())
button4.place(rely=0.45, relx=0.55, height=25, width=180)


# The file/file path text
#label_file2 = ttk.Label(file_frame2, text="Ningún archivo seleccionado")
#label_file2.place(rely=0, relx=0)

## Treeview Widget
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1) #establece la altura y el ancho del widget al 100% de su contenedor (frame1).

treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview) 
treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview) 
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set) # asignar las barras de desplazamiento al widget Treeview
treescrollx.pack(side="bottom", fill="x") # hacer que la barra de desplazamiento llene el eje x del widget Treeview
treescrolly.pack(side="right", fill="y") # hacer que la barra de desplazamiento llene el eje y del widget Treeview

tv2 = ttk.Treeview(frame2)
tv2.place(relheight=1, relwidth=1) #establece la altura y el ancho del widget al 100% de su contenedor (frame1).

treescrolly2 = tk.Scrollbar(frame2, orient="vertical", command=tv2.yview) 
treescrollx2 = tk.Scrollbar(frame2, orient="horizontal", command=tv2.xview) 
tv2.configure(xscrollcommand=treescrollx2.set, yscrollcommand=treescrolly2.set) # asignar las barras de desplazamiento al widget Treeview
treescrollx2.pack(side="bottom", fill="x") # hacer que la barra de desplazamiento llene el eje x del widget Treeview
treescrolly2.pack(side="right", fill="y") # hacer que la barra de desplazamiento llene el eje y del widget Treeview


"""scrollbary = Scrollbar(frame2, orient = VERTICAL)
scrollbarx = Scrollbar(frame2, orient = HORIZONTAL)

listbox = Listbox(root, yscrollcommand=scrollbary.set, xscrollcommand=scrollbarx.set,width=116, height = 13)

listbox.config(yscrollcommand=scrollbary.set)
scrollbary.config(command=listbox.yview)
scrollbary.place(x=705, y =0, width=15, height = 215)

listbox.config(xscrollcommand=scrollbarx.set)
scrollbarx.config(command=listbox.xview)
scrollbarx.place(x=0.5, y =215,  width=720, height = 15 )

listbox.place(x=618, y =16)"""

Button1 = tk.Button(root,text="Script Iniciar sección",command=my_sap_script).place(x=800, y =640) #.grid(pady=5, row=0, column=3)
Button2 = tk.Button(root,text="Script transacción ZPM_PT",command=transaccion_zpm_pt).place(x=950, y =640) #.grid(pady=5, row=3, column=3)
Button3 = tk.Button(root,text="Cerrar sección SAP",command=logoff_SAP).place(x=1150, y =640) #.grid(pady=5, row=3, column=3)

    
root.mainloop()
