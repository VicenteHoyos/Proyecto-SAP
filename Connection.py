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
        self.semana = None
        self.year = None
        self.cols = None
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
        self.connect_to = 'QE8- ERP Calidad isaeccqe8'

    def open_sap(self, user, password, language, connect_to='QE8- ERP Calidad isaeccqe8'):
        self.user = user
        self.password = password
        self.language = language
        self.connect_to = connect_to
        self.process = subprocess.Popen(self.path)
        time.sleep(4)  # tiempo de reposo para que la computadora tenga tiempo de abrir SAP.

        # Connecting to the SAP API
        self.sapguiauto = win32com.client.GetObject('SAPGUI')
        if not type(self.sapguiauto) == win32com.client.CDispatch:
            return
        self.application = self.sapguiauto.GetScriptingEngine
        if not type(self.application) == win32com.client.CDispatch:
            self.sapguiauto = None
            return

        # Comprueba si el usuario ya ha iniciado sesión en SAP en la computadora utilizada. Si es así, usa la conexión actual.
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

    def transaccion_zpm_pt_2(self, semana, year):         
        self.session[0].findById("wnd[0]").maximize
        self.session[0].findById("wnd[0]/tbar[0]/okcd").text = "ZPM_PT"
        self.session[0].findById("wnd[0]").sendVKey (0)
        #session.findById("wnd[0]/tbar[0]/btn[0]").press
        self.session[0].findById("wnd[0]/usr/chkP_EST01").selected = True
        self.session[0].findById("wnd[0]/usr/chkP_EST02").selected = True
        self.session[0].findById("wnd[0]/usr/chkP_EST03").selected = True
        self.session[0].findById("wnd[0]/usr/chkP_EST04").selected = True
        self.session[0].findById("wnd[0]/usr/chkP_EST05").selected = True
        self.session[0].findById("wnd[0]/usr/ctxtP_IWERK").text = "ITCO"
        self.session[0].findById("wnd[0]/usr/txtS_ZWEEKR-LOW").text = semana
        #self.session[0].findById("wnd[0]/usr/txtS_ZWEEKR-HIGH").text = "52"
        self.session[0].findById("wnd[0]/usr/txtS_ZZANO-LOW").text = year
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
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").setCurrentCell (-1,"")
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").selectAll()
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").contextMenu()
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").selectContextMenuItemByPosition ("0")
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").pressToolbarContextButton ("&MB_EXPORT")
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").selectContextMenuItem ("&XXL")
        self.session[0].findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session[0].findById("wnd[1]/tbar[0]/btn[11]").press()
        self.session[0].findById("wnd[0]/tbar[0]/btn[15]").press()
    
    def InformeMedidasOperativas(self, cols):
        self.session[0].findById("wnd[0]").maximize
        self.session[0].findById("wnd[0]/tbar[0]/okcd").text = "ZPM_PT"
        self.session[0].findById("wnd[0]/tbar[0]/btn[0]").press()
        self.session[0].findById("wnd[0]/usr/ctxtP_IWERK").text = "ITCO"
        self.session[0].findById("wnd[0]/usr/ctxtP_IWERK").caretPosition = 4
        self.session[0].findById("wnd[0]/mbar/menu[2]/menu[3]").select()
        self.session[0].findById("wnd[0]/usr/txtS_REVNR-LOW").text = str(cols[0])
        self.session[0].findById("wnd[0]/usr/txtS_REVNR-LOW").setFocus
        self.session[0].findById("wnd[0]/usr/txtS_REVNR-LOW").caretPosition = 8
        self.session[0].findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").setCurrentCell (-1,"")
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").selectAll()
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").contextMenu()
        self.session[0].findById("wnd[0]/usr/cntlCC_9000_CONT1/shellcont/shell").selectContextMenuItemByPosition ("0")
       
    def RegresarHome(self):
        self.session[0].findById("wnd[0]/tbar[0]/btn[15]").press()
        self.session[0].findById("wnd[0]/tbar[0]/btn[15]").press()
        self.session[0].findById("wnd[0]/tbar[0]/btn[15]").press()

class LoginVentanaEmergente():

    def __init__(self, root, connect):
        self.connect = connect
        self.root = root 
        self.ventana = tk.Toplevel()
        self.ventana.grab_set()
        self.ventana.title('Transacción ZPM_PT')
        self.ventana.geometry('300x200')

        self.Inicializar_VentanaEmergente()

    def Inicializar_VentanaEmergente(self):
        lbl_semana = tk.Label(self.ventana , text= 'Semana: ')
        lbl_semana.place(x =20 , y =20)
        self.txt_semana = tk.Entry(self.ventana)
        self.txt_semana.place(x=80, y=20)
        self.txt_semana.focus()

        lbl_year = tk.Label(self.ventana , text= 'Año: ')
        lbl_year.place(x =20 , y =50)
        self.txt_year = tk.Entry(self.ventana)
        self.txt_year.place(x=80, y=50)

        btn_aceptar = tk.Button(self.ventana, text = 'Aceptar')
        btn_aceptar.place( x = 80 , y = 80 )
        btn_aceptar['command']= self.Aceptar 
    
    def Aceptar(self):
        semana = self.txt_semana.get()
        year  = self.txt_year.get()

        
        if len(semana)  == 0:
            messagebox.showwarning('Mensaje', 'El campo semana es obligatorio.')
            return
        
        
        if len(year)  == 0:
            messagebox.showwarning('Mensaje', 'El campo año es obligatorio.')
            return 
           
        self.connect.transaccion_zpm_pt_2(semana, year)
        
        self.ventana.destroy()
       
def my_sap_script():
    connect = Connect()
    session = connect.open_sap("ITCO25708","951211hoyoscastaño26304","ES")

def transaccion_zpm_pt():
    connect = Connect()
    session = connect.open_sap("ITCO25708","951211hoyoscastaño26304","ES")

    login_ventana = LoginVentanaEmergente(root, connect)
    root.wait_window(login_ventana.ventana)

    clear_data()

    try:

        cols = ['Planes Trab.',	
                'Núm.Consignación',
                'Denominación Ub.Técnica',
                'Clasificación De La Consignación.',
                'Denominación de la revisión',
                'Estado',
                'Desc. Estado',
                'Fecha inic. revisión',
                'Hora inic.revisión',
                'Fecha fin revisión',
                'Hora fin revisión',
                'Desc. Jefe Trab.',
                'Tipo Ingreso',
                'Codigo Origen de Mantenimiento'                            
          ]

        clipboardData = pd.read_clipboard(sep=' ') #pd.read_clipboard(sep='\s+', index_col=[0], header=[0,1])
        
        tv1["columns"] = cols
        tv1["show"] = "headings"

        for column in tv1["columns"]:            
            tv1.heading(column, text=column) # Encabezado de la columna = column name
        
        cols = list(clipboardData.columns)
        #print(cols)
        tv1.insert ("", END,text = 0 , values = cols)

        for index, row in clipboardData.iterrows():
            tv1.insert("",END,text=index,values=list(row))
        
        
        #print(clipboardData)
        #print(cols)
        
        time.sleep(1)
    except:
        selection = None
    return None
    

def logoff_SAP():
    connect = Connect()
    session = connect.open_sap("ITCO25708","951211hoyoscastaño26304","ES")
    connect.logoff()

def File_dialog():
    """Esta función abrirá el explorador de archivos y asignará la ruta de archivo elegida a label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Seleccione un Archivo",
                                          filetype=( ("All Files", "*.*"),("xlsx files", "*.xlsx")))
    label_file["text"] = filename
    return None

def Load_excel_data():
    """Si el archivo seleccionado es válido, se cargará el archivo en la Treeview."""

    file_path = label_file["text"]
    
    excel_filename = r"{}".format(file_path)

    print( excel_filename)

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

        cols = ["Revisión",
        	"Ubicacion Tecnica",
            "Denominación de la ubicación técnica",
            "Texto código medidas",
            "Fecha inic.prevista",
            "Hora inic.prevista",
            "Fin planificado",
            "Hora fin prevista",
            "Clas. UT ( Activo - Equipo)",
            "Detalle",
            "Tipo Reg Diario"                            
          ]

        clipboardData = pd.read_clipboard(sep=' ') #pd.read_clipboard(sep='\s+', index_col=[0], header=[0,1])
        
        tv2["columns"] = cols
        tv2["show"] = "headings"

        for column in tv2["columns"]:            
            tv2.heading(column, text=column) # Encabezado de la columna = column name
        
        cols = list(clipboardData.columns)
        print(cols)
        tv2.insert ("", END,text = 0 , values = cols)

        for index, row in clipboardData.iterrows():
            tv2.insert("",END,text=index,values=list(row))
        
        
        #print(clipboardData)
        #print(cols)
        
        time.sleep(1)
    except:
        selection = None
    return None

def append_select():
    
    Encabezados = ['Planes Trab.',	
                    'Núm.Consignación',
                    'Denominación Ub.Técnica',
                    'Clasificación De La Consignación.',
                    'Denominación de la revisión',
                    'Estado',
                    'Desc. Estado',
                    'Fecha inic. revisión',
                    'Hora inic.revisión',
                    'Fecha fin revisión',
                    'Hora fin revisión',
                    'Desc. Jefe Trab.',
                    'Tipo Ingreso',
                    'Codigo Origen de Mantenimiento'
                ]

    tv3["columns"] = Encabezados
    tv3["show"] = "headings"

    for column in tv3["columns"]:            
        tv3.heading(column, text=column) # Encabezado de la columna = column name
    
    cur_id = tv1.focus()
    cols = list(tv1.item(cur_id)["values"])
    print(cols[2])
        
    if cur_id:   # do nothing if there"s no selection
        tv3.insert ("", END,text = 0 , values = cols)

def load_EstadosoperativosPT():
    if tv1.focus():
        cur_id = tv1.focus()
        cols = list(tv1.item(cur_id)["values"])
        print(cols[0])
        connect = Connect()
        session = connect.open_sap("ITCO25708","951211hoyoscastaño26304","ES")
        
        connect.InformeMedidasOperativas(cols)

        print(pd.read_clipboard(sep=' '))
        print("--------------------------------------------------------------")
        try:
            cols = ["Revisión",
                "Ubicacion Tecnica",
                "Denominación de la ubicación técnica",
                "Texto código medidas",
                "Fecha inic.prevista",
                "Hora inic.prevista",
                "Fin planificado",
                "Hora fin prevista",
                "Clas. UT ( Activo - Equipo)",
                "Detalle",
                "Tipo Reg Diario"                            
            ]

            clipboardData = pd.read_clipboard(sep=' ') #pd.read_clipboard(sep='\s+', index_col=[0], header=[0,1])
            
            tv2["columns"] = cols
            tv2["show"] = "headings"

            for column in tv2["columns"]:            
                tv2.heading(column, text=column) # Encabezado de la columna = column name
            
            cols = list(clipboardData.columns)
            print(cols)
            tv2.insert ("", END,text = 0 , values = cols)

            for index, row in clipboardData.iterrows():
                tv2.insert("",END,text=index,values=list(row))
            
            
            #print(clipboardData)
            #print(cols)
            
            time.sleep(1)
        except:
            selection = None
        #return None
        print("*******************")
        connect.RegresarHome()

    else:
        tk.messagebox.showerror("Información", f"No ha seleccionado Plan de Trabajo")
        return None

          
def Clear_Estadosoperativos():
    tv2.delete(*tv2.get_children())
    return None 
    
def Clear_Cortesvisibles():
    tv3.delete(*tv3.get_children())
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
button3.place(rely=0.45, relx=0.10, height=25, width=120)

button4 = tk.Button(file_frame2, text="Cargar Estados Operativos PT ", command=lambda: load_EstadosoperativosPT())
button4.place(rely=0.45, relx=0.35, height=25, width=180)

button5 = tk.Button(file_frame2, text="Limpiar Estados Operativos", command=lambda: Clear_Estadosoperativos())
button5.place(rely=0.45, relx=0.70, height=25, width=180)

button6 = tk.Button(file_frame3, text="Prueba select item ", command=lambda: append_select())
button6.place(rely=0.45, relx=0.20, height=25, width=120)

button7 = tk.Button(file_frame3, text="Limpiar Estados Operativos", command=lambda: Clear_Cortesvisibles())
button7.place(rely=0.45, relx=0.55, height=25, width=180)

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

tv3 = ttk.Treeview(frame3)
tv3.place(relheight=1, relwidth=1) #establece la altura y el ancho del widget al 100% de su contenedor (frame1).

treescrolly3 = tk.Scrollbar(frame3, orient="vertical", command=tv3.yview) 
treescrollx3 = tk.Scrollbar(frame3, orient="horizontal", command=tv3.xview) 
tv3.configure(xscrollcommand=treescrollx3.set, yscrollcommand=treescrolly3.set) # asignar las barras de desplazamiento al widget Treeview
treescrollx3.pack(side="bottom", fill="x") # hacer que la barra de desplazamiento llene el eje x del widget Treeview
treescrolly3.pack(side="right", fill="y") # hacer que la barra de desplazamiento llene el eje y del widget Treeview

Button1 = tk.Button(root,text="Script Iniciar sección",command=my_sap_script).place(x=800, y =640) #.grid(pady=5, row=0, column=3)
Button2 = tk.Button(root,text="Script transacción ZPM_PT",command=transaccion_zpm_pt).place(x=950, y =640) #.grid(pady=5, row=3, column=3)
Button3 = tk.Button(root,text="Cerrar sección SAP",command=logoff_SAP).place(x=1150, y =640) #.grid(pady=5, row=3, column=3)

    
root.mainloop()
