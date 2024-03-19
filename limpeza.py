#invoke the needed modules
import os
import glob
import wmi
import subprocess

#Wipe-clean base dirs

def wipe_directories_clean():
    # List of directories to wipe clean
    directories_to_wipe = ['C:\\Temp', 'C:\\Windows\\Prefetch', 'C:\\Windows\\Temp']

    # Iterate through the list of directories
    for directory in directories_to_wipe:
        try:
            # Check if the directory exists
            if os.path.exists(directory) and os.path.isdir(directory):
                # Iterate over the files and subdirectories in the directory
                for root, dirs, files in os.walk(directory):
                    # Remove files in the directory
                    for file in files:
                        file_path = os.path.join(root, file)
                        os.remove(file_path)
                        print(f"{file_path} removido.")
                        
                    # Remove subdirectories in the directory
                    for dir in dirs:
                        dir_path = os.path.join(root, dir)
                        os.rmdir(dir_path)
                        print(f"{dir_path} removido.")

                print("")
                print("-"*60)
                print("")
                print(f"{directory} teve seu conteúdo excluído. Pressione ENTER.")
                input()
            else:
                print(f"Diretório {directory} não existe ou não é um diretório. Pressione ENTER.")
                input("")
        except Exception as e:
            print(f"Erro na limpeza de {directory}: {str(e)}")
            print("Pressione ENTER para voltar ao menu.")
            input("")
    
#Wipe-clean base dirs in auto mode

def wipe_directories_clean_auto():
    # List of directories to wipe clean
    directories_to_wipe = ['C:\\Temp', 'C:\\Windows\\Prefetch', 'C:\\Windows\\Temp']

    # Iterate through the list of directories
    for directory in directories_to_wipe:
        try:
            # Check if the directory exists
            if os.path.exists(directory) and os.path.isdir(directory):
                # Iterate over the files and subdirectories in the directory
                for root, dirs, files in os.walk(directory):
                    # Remove files in the directory
                    for file in files:
                        file_path = os.path.join(root, file)
                        os.remove(file_path)
                    # Remove subdirectories in the directory
                    for dir in dirs:
                        dir_path = os.path.join(root, dir)
                        os.rmdir(dir_path)

                print(f"{directory} teve seu conteúdo excluído.")
            else:
                print(f"Diretório {directory} não existe ou não é um diretório.")
        except Exception as e:
            print(f"Erro na limpeza de {directory}: {str(e)}")
    

# Iterate through Users folders in two languages to wipe-clean cache folders
def wipe_user_folders():
    # List of possible user directories
    user_directories = ['C:\\Users\\*', 'C:\\Usuários\\*', 'C:\\Usuarios\\*']

    # Subdirectories to wipe clean for each user
    subdirectories_to_wipe = [
        'AppData\\Local\\Temp', 
        'AppData\\Local\\Microsoft\\Outlook',
        'AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache',
        'AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Cache'
    ]

    # Force-quit Edge and Chrome
    processes_to_quit = ['chrome.exe', 'msedge.exe']
    for process in processes_to_quit:
        try:
            subprocess.run(["taskkill", "/f", "/im", process], check=True)
        except subprocess.CalledProcessError:
            continue

    # Get a list of user profile directories
    for user_directory in user_directories:
        user_profiles = glob.glob(user_directory)
        for profile in user_profiles:
            for subdirectory in subdirectories_to_wipe:
                path_to_clean = os.path.join(profile, subdirectory)
                if os.path.exists(path_to_clean) and os.path.isdir(path_to_clean):
                    try:
                        # Wipe the subdirectory clean
                        for root, dirs, files in os.walk(path_to_clean):
                            for file in files:
                                # Exclude .pst files in the Outlook directory
                                if subdirectory.lower().endswith('outlook') and file.lower().endswith('.pst'):
                                    continue
                                file_path = os.path.join(root, file)
                                os.remove(file_path)
                                print(f"{file_path} removido.")
                            for dir in dirs:
                                dir_path = os.path.join(root, dir)
                                os.rmdir(dir_path)
                                print(f"{dir_path} removido.")

                        print(f"{path_to_clean} teve seu conteúdo excluído. Pressione ENTER para retornar ao menu.")
                        input("")
                    except Exception as e:
                        print(f"Erro: {e}. {path_to_clean} não pôde ser limpo. Proceda manualmente. Pressione ENTER para retornar ao menu.")
                        input("")
                else:
                    print(f"Diretório {path_to_clean} não existe ou não é um diretório. Pressione ENTER para retornar ao menu.")
                    input("")

# Iterate through Users folders in two languages to wipe-clean cache folders in auto mode
def wipe_user_folders_auto():
    # List of possible user directories
    user_directories = ['C:\\Users\\*', 'C:\\Usuários\\*', 'C:\\Usuarios\\*']

    # Subdirectories to wipe clean for each user
    subdirectories_to_wipe = [
        'AppData\\Local\\Temp', 
        'AppData\\Local\\Microsoft\\Outlook',
        'AppData\\Local\\Google\\Chrome\\User Data\\Default\\Cache',
        'AppData\\Local\\Microsoft\\Edge\\User Data\\Default\\Cache'
    ]

    # Force-quit Edge and Chrome
    processes_to_quit = ['chrome.exe', 'msedge.exe']
    for process in processes_to_quit:
        try:
            subprocess.run(["taskkill", "/f", "/im", process], check=True)
        except subprocess.CalledProcessError:
            continue

    # Get a list of user profile directories
    for user_directory in user_directories:
        user_profiles = glob.glob(user_directory)
        for profile in user_profiles:
            for subdirectory in subdirectories_to_wipe:
                path_to_clean = os.path.join(profile, subdirectory)
                if os.path.exists(path_to_clean) and os.path.isdir(path_to_clean):
                    try:
                        # Wipe the subdirectory clean
                        for root, dirs, files in os.walk(path_to_clean):
                            for file in files:
                                # Exclude .pst files in the Outlook directory
                                if subdirectory.lower().endswith('outlook') and file.lower().endswith('.pst'):
                                    continue
                                file_path = os.path.join(root, file)
                                os.remove(file_path)
                            for dir in dirs:
                                dir_path = os.path.join(root, dir)
                                os.rmdir(dir_path)

                        print(f"{path_to_clean} teve seu conteúdo excluído.")
                    except Exception as e:
                        print(f"Erro: {e}. {path_to_clean} não pôde ser limpo. Proceda manualmente.")
                else:
                    print(f"Diretório {path_to_clean} não existe ou não é um diretório.")

#Windows clean-up
def run_system_commands():
    # Command to run DISM and SFC
    command = 'DISM /ONLINE /CLEANUP-IMAGE /RESTOREHEALTH & SFC /SCANNOW & GPUPDATE /FORCE'

    try:
        # Execute the command
        os.system(command)
        print("Integridade do Windows e seu registro verificadas com sucesso. Políticas de grupo e usuário atualizadas com sucesso.")
        print("Pressione ENTER para voltar ao menu principal.")
        input("")
    except Exception as e:
        print(f"Erro na verificação de integridade do Windows e seu registro ou com a atualização de políticas. Proceda manualmente. Código de erro: {e}")
        print("Pressione ENTER para voltar ao menu principal.")
        input("")
    
#Windows clean-up in auto mode
def run_system_commands_auto():
    # Command to run DISM and SFC
    command = 'DISM /ONLINE /CLEANUP-IMAGE /RESTOREHEALTH & SFC /SCANNOW & GPUPDATE /FORCE'

    try:
        # Execute the command
        os.system(command)
        print("Integridade do Windows e seu registro verificadas com sucesso. Políticas de grupo e usuário atualizadas com sucesso.")
    except Exception as e:
        print(f"Erro na verificação de integridade do Windows e seu registro ou com a atualização de políticas. Proceda manualmente. Código de erro: {e}")


#Check whether ethernet and Wi-Fi adapters are off and turns them back on
def enable_network_adapters():
    try:
        c = wmi.WMI()
        for adapter in c.Win32_NetworkAdapter():
            if "Wireless" in adapter.Name or "Ethernet" in adapter.Name:
                adapter.Enable()  # Enable the adapter
        print(f"{adapter.Name} habilitado com sucesso.")
        print("")
        print("Pressione ENTER para voltar ao menu.")
        input("")
        
    except Exception as e:
        print(f"Um erro ocorreu: {str(e)}")
        print("")
        print("Este erro ocorre em função do módulo WMI. Verifique em Adaptadores de Rede para ver Ethernet e Wi-Fi habilitados.")
        print("Pressione ENTER para voltar ao menu.")
        input("")

#Check whether ethernet and Wi-Fi adapters are off and turns them back on in auto mode
def enable_network_adapters_auto():
    try:
        c = wmi.WMI()
        for adapter in c.Win32_NetworkAdapter():
            if "Wireless" in adapter.Name or "Ethernet" in adapter.Name:
                adapter.Enable()  # Enable the adapter
        print("Adaptadores habilitados com sucesso.")
        
    except Exception as e:
        print(f"Um erro ocorreu: {str(e)}")

        

#Checks whether the spooler service is up and running. This is primordial for the future correction of the spooler bug.
def check_and_restart_spooler():
    try:
        # Check the status of the spooler service
        spooler_status = subprocess.run(['sc', 'query', 'spooler'], capture_output=True, text=True, check=True)
        
        if "RUNNING" not in spooler_status.stdout:
            # Start the spooler service if it's not running
            subprocess.run(['sc', 'start', 'spooler'], check=True)
            print("Spooler iniciado.")
            print("Pressione ENTER para voltar ao menu")
            input("")
        else:
            print("Spooler já está rodando.")
            print("Pressione ENTER para voltar ao menu")
            input("")
    #This snippet will be called in case of the spooler bug        
    except subprocess.CalledProcessError as e:
        print(f"Um erro ocorreu ao tentar acionar o spoolsv.exe")
        print("")
        print("-"*60)
        print(f"Cógido do erro: {e}")
        print("")
        print("-"*60)
        print("Abra um chamado para FIELD SUPPORT PRINTERS LATAM pedindo verificação do bug de spooler.")
        print("Pressione ENTER para voltar ao menu")
        input("")



#In auto-mode, checks whether the spooler service is up and running. This is primordial for the future correction of the spooler bug.
def check_and_restart_spooler_auto():
    
        # Check the status of the spooler service
        spooler_status = subprocess.run(['sc', 'query', 'spooler'], capture_output=True, text=True, check=True)
        if "RUNNING" not in spooler_status.stdout:
            # Start the spooler service if it's not running
            subprocess.run(['sc', 'start', 'spooler'], check=True)
            print("Spooler iniciado.")
        else:
            print("Spooler já está rodando.")

    
        
    
#Auto mode
def allAndReboot():
    wipe_directories_clean_auto()
    wipe_user_folders_auto()
    check_and_restart_spooler_auto()
    enable_network_adapters_auto()
    run_system_commands_auto()
    
    # Reboot the system
    commandToReboot = 'shutdown /r /f /t 0'
    os.system(commandToReboot)

#sets up the menu
def criarMenu():
    #clears screen
    os.system("cls")

    #a line
    linha = "+" + "-" * 68 + "+"
        
    #Menu's array
    opcoes = [
        "| 1. Limpar arquivos temporários de uso geral                        |",
        "| 2. Limpar caches locais de todos os usuários                       |",
        "| 3. Serviços de Windows e seu registro e de políticas de AD         |",
        "| 4. Habilitar adaptadores de rede                                   |",
        "| 5. Verificar bug do spooler                                        |",
        "| 6. Executar todas as funções automaticamente e reiniciar ao final  |",
        "| 7. Sair                                                            |"
    ]
    
    #show-time
    print(linha)
    print("|                                                                    |")
    print("|               Programa de manutenção de PCs Windows                |")
    print("|                                                                    |")
    print("|                                                                    |")
    print("|2024-03-19 --- Versão 1.1 --- Eder Castro: design, pesquisa e código|")
    print("|                                                                    |")
    print(linha)
    print("|                                                                    |")
    print(opcoes[0])
    print(opcoes[1])
    print(opcoes[2])
    print(opcoes[3])
    print(opcoes[4])
    print(opcoes[5])
    print(opcoes[6])
    print(linha)




def main():
    while True:
        criarMenu()
        opcao = input("Digite o número da opção desejada: ")

        if opcao == '1':
            wipe_directories_clean()
            
        elif opcao == '2':
            wipe_user_folders()
        
        elif opcao == '3':
            run_system_commands()
        
        elif opcao == '4':
            enable_network_adapters()

        elif opcao == '5':
            check_and_restart_spooler()

        elif opcao == '6':
            allAndReboot()
                            
        elif opcao == '7':
            break
        else:
            print("Opção inválida. Por favor, escolha uma opção válida.")
            print("")
            input("Pressione ENTER para retornar ao menu")
            

if __name__ == "__main__":                                                          # necessário para o interpretador entender de que se trata de um script primário
    main()                                                                          # loop inicial
