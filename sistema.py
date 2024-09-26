import customtkinter as ctk
import winreg
import configparser
from tkinter import messagebox


FILE_PROGRAMS = 'programs.ini'
config = configparser.ConfigParser()
config.read(FILE_PROGRAMS)

path = winreg.HKEY_LOCAL_MACHINE
path_user = winreg.HKEY_CURRENT_USER
programs = config['Programs']['installed'].split(', ')


class App():
    def __init__(self):
        super().__init__()
        self.sistema = ctk.CTk()
        self.configuracoes_da_janela_inicial()
        self.tema()
        self.tela_principal()

    def configuracoes_da_janela_inicial(self):
        self.sistema.geometry("575x500")
        ctk.deactivate_automatic_dpi_awareness()
        self.sistema.title(f"Sistema de Verificação")
        self.sistema.resizable(False, False)

    def tema(self):
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

    def tela_principal(self):
        self.frame_installed = ctk.CTkFrame(self.sistema, width=275, height=250)
        self.frame_installed.place(x=5, y=5)

        self.frame_no_installed = ctk.CTkFrame(self.sistema, width=275, height=250)
        self.frame_no_installed.place(x=300, y=5)

        self.frame_footer = ctk.CTkFrame(self.sistema, width=275, height=50)
        self.frame_footer.place(x=300, y=260)

        self.frame_footer_info = ctk.CTkFrame(self.sistema, width=270, height=75)
        self.frame_footer_info.place(x=5, y=260)

        self.lbl_installed = ctk.CTkLabel(self.frame_installed, text="Softwares Instalados:")
        self.lbl_installed.grid(row=0, column=0, padx=10, pady=10)

        self.text_installed = ctk.CTkTextbox(self.frame_installed, width=250, height=180)
        self.text_installed.grid(row=1, column=0, padx=10, pady=10)


        self.lbl_no_installed = ctk.CTkLabel(self.frame_no_installed, text="Softwares Não Instalados:")
        self.lbl_no_installed.grid(row=0, column=0, padx=10, pady=10)

        self.text_no_installed = ctk.CTkTextbox(self.frame_no_installed, width=250, height=180)
        self.text_no_installed.grid(row=1, column=0, padx=10, pady=10)


        self.lbl_infos = ctk.CTkLabel(self.frame_footer_info, text="Informações do Sistema:")
        self.lbl_infos.grid(row=0, column=0, padx=10, pady=10)

        self.text_infos = ctk.CTkTextbox(self.frame_footer_info, width=250, height=150)
        self.text_infos.grid(row=1, column=0, padx=10, pady=10)


        self.btn_verifica = ctk.CTkButton(self.frame_footer, width=100, height=50, text='Verificar',
                                          command=self.serviços)
        self.btn_verifica.grid(row=0, column=0, padx=10, pady=10)


    def serviços(self):
        self.verifica_ativacao()
        self.check()
        self.edge()
        self.text_installed.configure(state='disabled')
        self.text_no_installed.configure(state='disabled')
        self.text_infos.configure(state='disabled')


    def verifica_instalacao(self, nome, key_path, sub_path):
        try:
            program_key = winreg.OpenKey(key_path, sub_path)
            winreg.CloseKey(program_key)
            return True

        except FileNotFoundError:
            return False

    def check(self):
        for program_name in programs:
            installed = (
                    self.verifica_instalacao(program_name, path, f'SOFTWARE\\{program_name}') or
                    self.verifica_instalacao(program_name, path_user, f'Software\\Microsoft\\Office\\{program_name}') or
                    self.verifica_instalacao(program_name, path_user, f'SOFTWARE\\{program_name}')
            )

            if installed:
                self.text_installed.insert("end", f"{program_name} está instalado.\n")

            else:
                self.text_no_installed.insert("end", f"{program_name} não está instalado.\n")



    def verifica_ativacao(self):
        try:
            path = winreg.HKEY_LOCAL_MACHINE
            ativo = winreg.OpenKey(path, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\SoftwareProtectionPlatform")

            valor, tipo = winreg.QueryValueEx(ativo, "BackupProductKeyDefault")

            if valor:
                self.text_infos.insert("end", f"O Windows está ATIVO\n")
            else:
                self.text_infos.insert("end", f"O Windows está DESATIVADO\n")

        except FileNotFoundError:
            messagebox.showerror('Error', 'Chave de registro não encontrada. '
                                 'O Windows pode não estar ativado ou a versão pode ser incompatível.')
        except Exception as e:
            messagebox.showerror('Error', f'Erro ao verificar ativação: {e}')


    def edge(self):
        pagina_inicial = config['Navegador']['pagina_inicial']
        buscador = config['Navegador']['buscador']

        try:
            path = winreg.HKEY_LOCAL_MACHINE

            chave = winreg.OpenKey(path, "SOFTWARE\\Policies\\Microsoft\\Edge")

            winreg.SetValueEx(chave, "DefaultSearchProviderEnabled", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(chave, "DefaultSearchProviderName", 0, winreg.REG_SZ, f"{buscador}")
            winreg.SetValueEx(chave, "DefaultSearchProviderSearchURL", 0, winreg.REG_SZ,
                              "https://www.google.com/search?q={searchTerms}")

            winreg.SetValueEx(chave, "Homepage", 0, winreg.REG_SZ, f"{pagina_inicial}")
            winreg.SetValueEx(chave, "HomepageIsNewTabPage", 0, winreg.REG_DWORD, 0)

            winreg.CloseKey(chave)

            self.text_infos.insert("end", f"Buscador {buscador} definido com padrão!\n")
            self.text_infos.insert("end", f"Home Page padrão definida {pagina_inicial} !\n")


        except PermissionError:
            messagebox.showerror('Error', f'Execute como administrador para definir Buscador e Home Page.')
        except Exception as e:
            messagebox.showerror('Error', f'Erro ao modificar o registro: {e}')



app = App()
app.sistema.mainloop()