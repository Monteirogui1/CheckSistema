import winreg
import configparser

FILE_PROGRAMS = 'programs.ini'
config = configparser.ConfigParser()
config.read(FILE_PROGRAMS)

path = winreg.HKEY_LOCAL_MACHINE
path_user = winreg.HKEY_CURRENT_USER
programs = config['Programs']['installed'].split(', ')


def verifica_instalacao(nome, key_path, sub_path):
    try:
        program_key = winreg.OpenKey(key_path, sub_path)
        winreg.CloseKey(program_key)
        return True

    except FileNotFoundError:
        return False


# for program_name in programs:
#     installed = (
#             verifica_instalacao(program_name, path, f'SOFTWARE\\{program_name}') or
#             verifica_instalacao(program_name, path_user, f'Software\\Microsoft\\Office\\{program_name}') or
#             verifica_instalacao(program_name, path_user, f'SOFTWARE\\{program_name}')
#     )
#
#     if installed:
#         print(f'{program_name} está instalado.')
#     else:
#         print(f'{program_name} não está instalado.')

def verifica_ativação():
    try:
        path = winreg.HKEY_LOCAL_MACHINE
        ativo = winreg.OpenKey(path,"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\SoftwareProtectionPlatform")

        valor, tipo = winreg.QueryValueEx(ativo, "BackupProductKeyDefault")

        if valor:
            print('Ativo')
        else:
            print('Desativado')

    except FileNotFoundError:
        print("Chave de registro não encontrada. O Windows pode não estar ativado ou a versão pode ser incompatível.")
    except Exception as e:
        print(f"Erro ao verificar ativação: {e}")

# verifica_ativação()

def buscador_edge():
    pagina_inicial = config['Home']['pagina_inicial']
    try:
        path = winreg.HKEY_LOCAL_MACHINE

        chave = winreg.OpenKey(path, "SOFTWARE\\Policies\\Microsoft\\Edge")


        winreg.SetValueEx(chave, "DefaultSearchProviderEnabled", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(chave, "DefaultSearchProviderName", 0, winreg.REG_SZ, "Google")
        winreg.SetValueEx(chave, "DefaultSearchProviderSearchURL", 0, winreg.REG_SZ,
                          "https://www.google.com/search?q={searchTerms}")


        winreg.SetValueEx(chave, "Homepage", 0, winreg.REG_SZ, f"{pagina_inicial}")
        winreg.SetValueEx(chave, "HomepageIsNewTabPage", 0, winreg.REG_DWORD, 0)

        winreg.CloseKey(chave)

    except PermissionError:
        print("Erro: Execute o script como administrador.")
    except Exception as e:
        print(f"Erro ao modificar o registro: {e}")
