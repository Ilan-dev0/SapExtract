import win32com.client
import subprocess
import time
from tkinter import *
from tkinter import ttk, messagebox
from ttkthemes import ThemedStyle

class SapGui():
    def __init__(self, sap_system):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(5)

        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return

        application = self.SapGuiAuto.GetScriptingEngine
        self.connection = application.OpenConnection(sap_system, True)
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize()

    def sapLogin(self):
        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "100"
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "2160028386"
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "MxXN8955342*"
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
            self.session.findById("wnd[0]").sendVKey(0)
        except:
            print(sys.exc_info()[0])
            messagebox.showerror('Erro', 'Falha ao fazer login.')
            return False

        messagebox.showinfo('showinfo', 'Logado com Sucesso!')
        return True

    def run_script_fbl3n(self):
        # Configura o código da transação FBL3N
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "FBL3N"
        # Pressiona o botão OK para abrir a transação
        self.session.findById("wnd[0]/tbar[0]/btn[0]").press()

    def run_script_zvv_tco_005(self):
        # Configura o código da transação ZVVTCO005
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "ZVVTCO005"
        # Pressiona o botão OK para abrir a transação
        self.session.findById("wnd[0]/tbar[0]/btn[0]").press()

# Variável global para a instância de SapGui
sap_gui_instance = None

def login_sap():
    global sap_gui_instance
    sap_system = entry_sap_system.get()

    if not sap_system:
        messagebox.showerror('Erro', 'Por favor, insira um nome de Sistema SAP válido.')
    else:
        sap_gui_instance = SapGui(sap_system)
        if sap_gui_instance.sapLogin():
            btn_fbl3n.config(state="normal")
            btn_zvv_tco_005.config(state="normal")

def run_script_fbl3n():
    if sap_gui_instance:
        sap_gui_instance.run_script_fbl3n()

def run_script_zvv_tco_005():
    if sap_gui_instance:
        sap_gui_instance.run_script_zvv_tco_005()

if __name__ == '__main__':
    window = Tk()
    window.title("Sistema de automação de extração de relatórios")
    window.geometry('400x200')

    style = ThemedStyle(window)
    style.set_theme("clam")

    window.configure(bg="#2e2e2e")
    frame = ttk.Frame(window, padding="20", style="TFrame")
    frame.grid(row=0, column=0, sticky=(N, S, E, W))

    label_sap_system = ttk.Label(frame, text="Nome do Sistema SAP:", style="TLabel")
    label_sap_system.grid(row=0, column=0, pady=10, sticky=W)

    entry_sap_system = ttk.Entry(frame, style="TEntry")
    entry_sap_system.grid(row=0, column=1, pady=10, padx=10, sticky=(N, E, S, W))

    btn_login_sap = ttk.Button(frame, text="Login SAP", command=login_sap, style="TButton")
    btn_login_sap.grid(row=1, column=0, columnspan=2, pady=20)

    btn_fbl3n = ttk.Button(frame, text="Extrair T FBL3N", command=run_script_fbl3n, style="TButton", state="disabled")
    btn_fbl3n.grid(row=2, column=0, columnspan=2, pady=10)

    btn_zvv_tco_005 = ttk.Button(frame, text="Extrair T. ZVVTCO005", command=run_script_zvv_tco_005, style="TButton", state="disabled")
    btn_zvv_tco_005.grid(row=3, column=0, columnspan=2, pady=10)

    label_signature = ttk.Label(frame, text="Desenvolvido por: Ilan Costa", style="TLabel")
    label_signature.grid(row=4, column=0, columnspan=2, pady=10)

    window.grid_rowconfigure(0, weight=1)
    window.grid_columnconfigure(0, weight=1)

    window.mainloop()
