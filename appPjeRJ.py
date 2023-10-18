# Importação das Bibliotecas
from selenium.webdriver.common.by import By
from selenium import webdriver as wb
from openpyxl import load_workbook
from tkinter import filedialog
from tkinter import messagebox
from datetime import datetime
import tkinter as tk
import pyautogui
import warnings
import time
import os
warnings.filterwarnings('ignore')

# Classe para a interface gráfica do tkinter
class Aplicativo(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Download de Processos PJE - RJ")
        self.geometry("400x300")
        self.arquivo_excel = None
        
        self.label_usuario = tk.Label(self, text="Usuário:")
        self.label_usuario.pack()
        self.entry_usuario = tk.Entry(self)
        self.entry_usuario.pack()
        
        self.label_senha = tk.Label(self, text="Senha:")
        self.label_senha.pack()
        self.entry_senha = tk.Entry(self, show="*")
        self.entry_senha.pack()
        
        self.label_arquivo = tk.Label(self, text="Arquivo Excel:")
        self.label_arquivo.pack()
        self.button_arquivo = tk.Button(self, text="Selecionar Arquivo", command=self.selecionar_arquivo)
        self.button_arquivo.pack()
        
        self.label_numero_processo = tk.Label(self, text="Número do Processo:")
        self.label_numero_processo.pack()
        self.entry_numero_processo = tk.Entry(self)
        self.entry_numero_processo.pack()
        
        self.button_iniciar = tk.Button(self, text="Iniciar Download", command=self.iniciar_download)
        self.button_iniciar.pack()

    # Função para selecionar o arquivo Excel
    def selecionar_arquivo(self):
        self.arquivo_excel = filedialog.askopenfilename(initialdir="/", title="Selecione o arquivo Excel",
                                                    filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))

    # Função para iniciar os downloads 
    def iniciar_download(self):
        usuario = self.entry_usuario.get()
        senha = self.entry_senha.get()
        numero_processo = self.entry_numero_processo.get()
        
        if not usuario or not senha:
            messagebox.showerror("Erro", "Por favor, preencha os campos de usuário e senha.")
        elif not (self.arquivo_excel or numero_processo):
            messagebox.showerror("Erro", "Selecione um arquivo Excel ou digite o número do processo.")
        else:
            if self.arquivo_excel:
                # Se um arquivo Excel foi selecionado, use o arquivo Excel
                self.download_files_from_excel(usuario, senha, arquivo_excel=self.arquivo_excel)
            else:
                # Se não, use o número do processo digitado
                self.download_files(usuario, senha, numero_processo=numero_processo)

    # Função de download pela planilha Excel
    def download_files_from_excel(self, usuario, senha, arquivo_excel):

        # Especificações do arquivo Excel
        workbook = load_workbook(arquivo_excel)
        sheet = workbook.active
        coluna = sheet['A']
        # Configurar as opções do Chrome
        chrome_options = wb.ChromeOptions()
        
        # Obter o diretório inicial de downloads do sistema operacional
        diretorio_inicial = os.path.join(os.path.expanduser("~"), "Downloads")
        data_atual = datetime.now().strftime("%Y-%m-%d")
        diretorio_download = os.path.join(diretorio_inicial, data_atual)
        if not os.path.exists(diretorio_download):
            os.makedirs(diretorio_download)

        # Configurar a preferência de download
        prefs = {
            "download.default_directory": diretorio_download,
            "download.prompt_for_download": False,
        }
        chrome_options.add_experimental_option("prefs", prefs)
         
        # Instância do WebDriver
        driver = wb.Chrome(options=chrome_options)
        driver.implicitly_wait(30)
        driver.get('https://tjrj.pje.jus.br/1g/login.seam?loginComCertificado=false')
        time.sleep(1)

        # Clica no campo CPF/CNPJ*
        iframe = driver.find_element(By.XPATH, "//iframe[@id='ssoFrame']")
        driver.switch_to.frame(iframe)
        botaoCpf = driver.find_element(By.XPATH, "//*[@id='username']")
        botaoCpf.click()
        time.sleep(1)

        # Digita o Usuário
        botaoCpf.send_keys(usuario)
        time.sleep(1)

        # Clica no campo Senha
        botaoPassword = driver.find_element(By.XPATH, "//*[@id='password']")
        botaoPassword.click()
        time.sleep(0.5)

        # Digita a Senha 
        botaoPassword.send_keys(senha)
        time.sleep(0.5)

        # Clica no Botão de Login
        driver.find_element(By.XPATH, "//*[@id='kc-login']").click()
        time.sleep(0.5)

        # Clica na Barra Menu
        driver.find_element(By.XPATH, "//*[@id='barraSuperiorPrincipal']/div/div[1]/ul/li/a").click()
        time.sleep(1)

        # Clica em Processo
        driver.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/a").click()
        time.sleep(1)

        # Clica em Pesquisar
        driver.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/div/ul/li[4]").click()
        time.sleep(1)

        # Clica em Processo
        driver.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/div/ul/li[4]/div/ul/li[1]").click()
        time.sleep(1)

        # Iteração sobre as células da coluna e preenchimento do campo
        for cell in coluna:

            # Aplica a função para formatar o número do processo
            valor_limpo = limpar_valor(str(cell.value))

            # Econtra o campo do numero do processo
            numeroProc = driver.find_element(By.XPATH, "//*[@id='fPP:numeroProcesso:numeroSequencial']")

            # Preenchimento do campo com o valor da célula atual
            numeroProc.send_keys(valor_limpo)
            time.sleep(2)

            # Clica no botao pesquisar
            driver.find_element(By.XPATH, "//*[@id='fPP:searchProcessos']").click()
            time.sleep(2)

            # Clica no número do processo
            driver.find_element(By.XPATH, "/html/body/div[5]/div/div/div/div[2]/form/div[2]/div/table/tbody/tr/td[2]/a").click()
            time.sleep(5)

            # Clicar no Alert
            pyautogui.press('enter')
            time.sleep(6)

            # Todas as alças de janelas abertas (janelas e abas)
            abas = driver.window_handles
            time.sleep(1)
            # Alterar para a nova janela (a segunda na lista de alças)
            driver.switch_to.window(abas[1])
            time.sleep(2)

            # Clicar no Botão de Download
            driver.find_element(By.XPATH, "//*[@id='navbar:ajaxPanelAlerts']/ul[2]/li[5]/a").click()
            time.sleep(4)

            # Clica em Download
            driver.find_element(By.XPATH, "//*[@id='navbar:downloadProcesso']").click()
            time.sleep(4)

            # Clicar no Alert
            pyautogui.press('enter')
            time.sleep(4)

            # Fechar a aba[1]
            time.sleep(12)
            driver.close()

            # Alterar para a aba[0]
            abas = driver.window_handles
            driver.switch_to.window(abas[0])
            time.sleep(3)

            # Clicar no botão Limpar 
            driver.find_element(By.XPATH, '//*[@id="fPP:clearButtonProcessos"]').click()
            time.sleep(5)

            # Clica no campo do processo
            time.sleep(3)
            driver.find_element(By.XPATH, "/html/body/div[5]/div/div/div/div[2]/form/div[1]/div/div/div[5]/div/div/div[2]/input[1]").click()
            time.sleep(8)

        # Sair do Navegador
        driver.quit()    

        # Mensagem após a conclusão dos Downloads     
        messagebox.showinfo("Concluído", "Downloads concluídos com sucesso!")

  
    # Função de download pelo número de processo digitado        
    def download_files(self, usuario, senha, numero_processo=None):
        # Configurar as opções do Chrome
        chrome_options = wb.ChromeOptions()
        
        # Obter o diretório inicial de downloads do sistema operacional
        diretorio_inicial = os.path.join(os.path.expanduser("~"), "Downloads")
        data_atual = datetime.now().strftime("%Y-%m-%d")
        diretorio_download = os.path.join(diretorio_inicial, data_atual)
        if not os.path.exists(diretorio_download):
            os.makedirs(diretorio_download)

        # Configurar a preferência de download
        prefs = {
            "download.default_directory": diretorio_download,
            "download.prompt_for_download": False,
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
            
        # Instância do WebDriver
        driver = wb.Chrome(options=chrome_options)
        driver.implicitly_wait(30)
        driver.get('https://tjrj.pje.jus.br/1g/login.seam?loginComCertificado=false')
        time.sleep(1)

        # Clica no campo CPF/CNPJ*
        iframe = driver.find_element(By.XPATH, "//iframe[@id='ssoFrame']")
        driver.switch_to.frame(iframe)
        botaoCpf = driver.find_element(By.XPATH, "//*[@id='username']")
        botaoCpf.click()
        time.sleep(1)

        # Digita o Usuário
        botaoCpf.send_keys(usuario)
        time.sleep(1)

        # Clica no campo Senha
        botaoPassword = driver.find_element(By.XPATH, "//*[@id='password']")
        botaoPassword.click()
        time.sleep(0.5)

        # Digita a Senha 
        botaoPassword.send_keys(senha)
        time.sleep(0.5)

        # Clica no Botão de Login
        driver.find_element(By.XPATH, "//*[@id='kc-login']").click()
        time.sleep(0.5)

        # Clica na Barra Menu
        driver.find_element(By.XPATH, "//*[@id='barraSuperiorPrincipal']/div/div[1]/ul/li/a").click()
        time.sleep(1)

        # Clica em Processo
        driver.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/a").click()
        time.sleep(1)

        # Clica em Pesquisar
        driver.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/div/ul/li[4]").click()
        time.sleep(1)

        # Clica em Processo
        driver.find_element(By.XPATH, "//*[@id='menu']/div[2]/ul/li[2]/div/ul/li[4]/div/ul/li[1]").click()
        time.sleep(1)

        # Aplica a função para formatar o número do processo
        valor_limpo = limpar_valor(numero_processo)
        numeroProc = driver.find_element(By.XPATH, "//*[@id='fPP:numeroProcesso:numeroSequencial']")
        numeroProc.send_keys(valor_limpo)
        time.sleep(2)

        # Clica no botao pesquisar
        driver.find_element(By.XPATH, "//*[@id='fPP:searchProcessos']").click()
        time.sleep(2)

        # Clica no número do processo
        driver.find_element(By.XPATH, "/html/body/div[5]/div/div/div/div[2]/form/div[2]/div/table/tbody/tr/td[2]/a").click()
        time.sleep(5)

        # Clicar no alert
        pyautogui.press('enter')
        time.sleep(6)

        # Obter todas as alças de janelas abertas
        abas = driver.window_handles
        time.sleep(1)
        # Altera para a nova janela
        driver.switch_to.window(abas[1])
        time.sleep(2)

        # Clicar no Botão de Download
        driver.find_element(By.XPATH, "//*[@id='navbar:ajaxPanelAlerts']/ul[2]/li[5]/a").click()
        time.sleep(4)

        # Clica em Download
        driver.find_element(By.XPATH, "//*[@id='navbar:downloadProcesso']").click()
        time.sleep(4)

        # Clicar no Alert
        pyautogui.press('enter')
        time.sleep(4)

        # Fechar a aba[1]
        time.sleep(12)
        driver.close()
        time.sleep(3)

        # Sair do Navegador
        driver.quit()

        # Mensagem após a conclusão dos Downloads    
        messagebox.showinfo("Concluído", "Downloads concluídos com sucesso!")

        
# Função para formatar o número do processo        
def limpar_valor(valor):
    valor = valor.replace('-', '')  # Remove os "-"
    valor = valor.replace('8.19', '')  # Remove os "819"
    valor = valor.replace('.', '')  # Remove os "."
    return valor

# Especificações da execução do código
if __name__ == "__main__":
    app = Aplicativo()
    app.mainloop()