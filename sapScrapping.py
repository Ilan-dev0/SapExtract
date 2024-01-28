from pywinauto.application import Application
import time

def abrir_sap_logon():
    try:
        saplogon_path = r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
        app = Application().start(saplogon_path)
        return app
    except Exception as e:
        print(f"Erro ao abrir o SAP Logon: {e}")
        return None

def encontrar_janela_sap_logon(app):
    try:
        for _ in range(20):  # Tente por até 20 segundos
            janelas_sap = app.windows(title_re=".*SAP Logon 740.*")
            if janelas_sap and len(janelas_sap) > 0:
                return janelas_sap[0]
            time.sleep(5)
        raise Exception("Janela do SAP Logon não encontrada.")
    except Exception as e:
        print(f"Erro ao encontrar a janela do SAP Logon: {e}")
        return None

def pressionar_seta_para_baixo_e_enter(janela_sap):
    try:
        # Pressiona a seta para baixo três vezes
        janela_sap.type_keys("{DOWN 3}")

        # Pressiona Enter
        janela_sap.type_keys("{ENTER}")

        print("Sucesso: Setas para baixo e Enter pressionados.")
    except Exception as e:
        print(f"Erro ao pressionar setas para baixo e Enter: {e}")

def aguardar_nova_janela_apos_enter(janela_sap):
    try:
        # Aguarde até 20 segundos para a nova janela aparecer após a primeira pressionada de Enter
        nova_janela = None
        for _ in range(20):
            for child in janela_sap.children():
                if "SAP" in child.window_text():
                    nova_janela = child
                    break

            if nova_janela:
                break

            time.sleep(1)

        if nova_janela:
            nova_janela.set_focus()
            print("Sucesso: Nova janela detectada após a primeira pressionada de Enter.")
            return nova_janela
        else:
            raise Exception("Não foi possível encontrar a nova janela após a primeira pressionada de Enter.")
    except Exception as e:
        print(f"Erro ao aguardar a nova janela após a primeira pressionada de Enter: {e}")
        return None

def preencher_usuario_senha(nova_janela, usuario, senha):
    try:
        # Aguarde até 10 segundos para garantir que a nova janela esteja pronta para receber entrada
        time.sleep(10)

        # Preenche o campo de usuário
        campo_usuario = nova_janela.child_window(title="Usuário", control_type="Edit")
        if campo_usuario:
            campo_usuario.set_edit_text(usuario)

        # Preenche a senha
        campo_senha = nova_janela.child_window(title="Senha", control_type="Edit")
        if campo_senha:
            campo_senha.set_edit_text(senha)

        # Pressiona Enter
        campo_senha.type_keys("{ENTER}")

        print("Sucesso: Usuário, senha preenchidos e Enter pressionado.")
    except Exception as e:
        print(f"Erro ao preencher usuário e senha: {e}")

def fechar_sap_logon_e_encerrar(janela_sap):
    try:
        # Aguarde alguns segundos
        time.sleep(5)

        # Feche a janela do SAP Logon
        janela_sap.close()

        print("Sucesso: Janela do SAP Logon fechada.")
    except Exception as e:
        print(f"Erro ao fechar a janela do SAP Logon: {e}")
    finally:
        # Encerre a execução do script
        exit()

def main():
    # Substitua 'seu_usuario' e 'sua_senha' pelos dados de login reais
    usuario = 'seu_usuario'
    senha = 'sua_senha'

    # Abre o SAP Logon
    app_sap = abrir_sap_logon()

    if app_sap:
        # Encontra a janela do SAP Logon
        janela_sap = encontrar_janela_sap_logon(app_sap)

        if janela_sap:
            # Foca na janela do SAP Logon
            janela_sap.set_focus()

            # Pressiona as teclas
            pressionar_seta_para_baixo_e_enter(janela_sap)

            # Aguarda a nova janela abrir após a primeira pressionada de Enter
            nova_janela = aguardar_nova_janela_apos_enter(janela_sap)

            if nova_janela:
                # Preenche usuário e senha e pressiona Enter
                preencher_usuario_senha(nova_janela, usuario, senha)

                # Fecha o SAP Logon e encerra a execução
                fechar_sap_logon_e_encerrar(janela_sap)

if __name__ == "__main__":
    main()
