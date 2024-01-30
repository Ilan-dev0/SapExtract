import win32com.client

# Inicializa o objeto SapGuiAuto
SapGuiAuto = win32com.client.GetObject("SAPGUI")
if not type(SapGuiAuto) == win32com.client.CDispatch:
    SapGuiAuto = win32com.client.Dispatch("SAPGUI")

# Obtém o objeto Application
application = SapGuiAuto.GetScriptingEngine

# Obtém o objeto Connection
connection = application.Children(0)
if not type(connection) == win32com.client.CDispatch:
    connection = connection.NewConnection()

# Obtém o objeto Session
session = connection.Children(0)
if not type(session) == win32com.client.CDispatch:
    session = connection.NewSession()

# Maximiza a janela da sessão
session.findById("wnd[0]").maximize()

# Configura o código da transação
session.findById("wnd[0]/tbar[0]/okcd").text = "ZVVTCO005"

# Pressiona o botão OK para abrir a transação
session.findById("wnd[0]/tbar[0]/btn[0]").press()
