# # Criar ambiente Virtual, substituir meu_ambiente_virtual pelo nome do projeto
'''
python -m venv meu_ambiente_virtual
'''

# # Ativar ambiente virtual
'''
projeto_sisreg/Scripts/activate
'''


import json
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import Canvas, Scrollbar
from tkinter import Entry, Button
from tkinter import Label
from tkinter import Listbox
import time
import threading
import os
import re
from tqdm import tqdm
import pandas as pd
from datetime import date, timedelta
from time import sleep


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ChromeOptions
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException

#Chrome
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
opts = ChromeOptions()
#esta opcao serve para nao fechar o navegador apos a execucao do script
opts.add_experimental_option("detach", True)
# opts.add_experimental_option('excludeSwitches', ['enable-logging']) # Remove erros de LOG do terminal
#opts.add_argument("--headless")
# os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3' # desativa todos os INFOS, WARNING e ERROS do tensorflow

driver = webdriver.Chrome(service=ChromeService(),options=opts)

def main():
    


    codigo_procedimentos = []
    nome_procedimentos_geral_inicio = []

    nome_procedimentos_geral = []
    codigo_procedimentos_geral = []

    
    nome_procedimentos_mulher = []
    codigo_procedimentos_mulher = []

    nome_procedimentos_idoso = []
    codigo_procedimentos_idoso = []

    nome_procedimentos_crianca = []
    codigo_procedimentos_crianca = []

    vagas_geral = [] #lista que irá receber os valores de vagas disponiveis Geral Unissex
    vagas_mulher = [] #lista que irá receber os valores de vagas disponiveis Geral Feminino
    vagas_crianca = [] #lista que irá receber os valores de vagas disponiveis Criança
    vagas_idoso = []

    def carregar_site():
        
            
        def load_site():
            tempo_carregar_site = time
            driver.get('https://sisregiii.saude.gov.br/')
            while True:
                
                try:                
                    driver.find_element(By.ID, 'usuario').is_displayed() and driver.find_element(By.ID, 'senha').is_displayed()
                    
                except:
                    label_status.config(text="Carregando Site.")

                else:
                    progress_bar["value"] = 100
                    label_status.config(text="Site carregado")
                    icon_status.config()
                    break
                
            habilitar_campos_login()
                    
        def habilitar_campos_login():
            usuario_entry.config(state="normal")
            senha_entry.config(state="normal")
            conectar_button.config(state="normal")
        
        def atualizar_status_login(texto, sucesso=False):
            label_status_login.config(text=texto)
            if sucesso:
                icon_status.config()
            else:
                icon_status.config()
        
        def Login_sisreg():

            usuario = usuario_entry.get()  # Obtém o valor do campo de usuário
            senha = senha_entry.get()  # Obtém o valor do campo de senha

                        
            CampoNome = driver.find_element(By.ID, 'usuario')  # Procura o elemento
            CampoNome.send_keys(usuario)  # Envia dados ao elemento

            CampoSenha = driver.find_element(By.ID, 'senha')  # Procura o elemento
            CampoSenha.send_keys(senha)  # Envia dados ao elemento

           
            driver.find_element(By.NAME, 'entrar').click()  # Procura o elemento entrar e clica nele
                        

            # Tentar localizar a mensagem de erro
            try:
                error_message = driver.find_element(By.XPATH, '//*[@id="mensagem"]/center/font/b')
                if error_message.text == "Login ou senha incorreto(s).":
                    atualizar_status_login("Login ou Senha incorretos!!!")
                
            except NoSuchElementException:
                               
                driver.get("https://sisregiii.saude.gov.br/cgi-bin/cadweb50?url=/cgi-bin/marcar")
                
                sleep(1)
                atualizar_status_login("Acessando Procedimentos...", sucesso=True)
                # Força a atualização da interface gráfica
                Tela_Login.update_idletasks()  

                df = pd.read_excel ('Banco_de_Dados.xlsx', sheet_name= 0)
                for index, row in df.iterrows():
                    CampoCnes = driver.find_element(By.NAME,'nu_cns')
                    cnes_f = (str(row['cnes masculino']))
                    CampoCnes.send_keys(cnes_f)  #Cnes masculino Adulto 


                driver.find_element(By.NAME,'btn_pesquisar').click()
                sleep(1)
                driver.find_element(By.NAME,'btn_continuar').click()
                
                dropbox_select = driver.find_element(By.NAME,"pa")
                opcoes = dropbox_select.find_elements(By.TAG_NAME,"option")
                for option in opcoes[1:]:
                    nome_procedimentos_geral_inicio.append(option.text)
                    codigo_procedimentos.append(option.get_attribute("value"))

                atualizar_status_login("Procedimentos Extraidos!", sucesso=True)
                # Força a atualização da interface gráfica
                Tela_Login.update_idletasks()    
                Tela_Login.withdraw()
                
                abrir_extrair_procedimentos()
                   

            
        
        # Configuração da janela principal
        Tela_Login = tk.Tk()
        Tela_Login.title("Carregador de Site")
        Tela_Login.geometry("400x300")  # Aumentei a altura para acomodar o botão "Abrir Tela de Login"
     

        # Frame para conter o label e o ícone
        status_frame = tk.Frame(Tela_Login)
        status_frame.pack(pady=10)

        # Label de Status
        label_status = tk.Label(status_frame, text="", font=("Arial", 12))
        label_status.pack(side="left")  # Alinha à esquerda


        # Ícone de Status
        icon_status = tk.Label(status_frame)
        icon_status.pack(side="right")  # Alinha à direita

        # Botão de Recarregar Site
        reload_button = tk.Button(Tela_Login, text="Recarregar Site", command=load_site)
        reload_button.pack(pady=10)

        # Barra de Progresso
        progress_bar = ttk.Progressbar(Tela_Login, length=300, mode="determinate")
        progress_bar.pack()

        # Campos de Login e Senha (inicialmente desabilitados)
        usuario_label = tk.Label(Tela_Login, text="Usuário:")
        usuario_label.pack(pady=5)
        usuario_entry = tk.Entry(Tela_Login, state="disabled")
        usuario_entry.pack()

        senha_label = tk.Label(Tela_Login, text="Senha:")
        senha_label.pack(pady=5)
        senha_entry = tk.Entry(Tela_Login, show="*", state="disabled")  # Para ocultar a senha
        senha_entry.pack()

        # Botão "Conectar" (inicialmente desabilitado)
        conectar_button = tk.Button(Tela_Login, text="Conectar", state="disabled", command=Login_sisreg)
        conectar_button.pack(pady=10)

        # Label de Status para o Login
        label_status_login = tk.Label(Tela_Login, text="", font=("Arial", 12))
        label_status_login.pack(side="top")  # Alinha acima do botão


        progress_bar["value"] = 0
        label_status.config(text="Carregando site...")
       
        Tela_Login.mainloop()


    def abrir_extrair_procedimentos():
        
        # Função para criar os Checkbuttons
        def criar_checkbuttons():
            for col_idx, coluna in enumerate(colunas):
                for row_idx, procedimento in enumerate(coluna):
                    if 'grupo' in procedimento.lower() or procedimento.count('x') >= 3:
                        # Configurar a cor do texto para azul se o procedimento atender aos critérios
                        texto_cor = "blue"
                        # Crie o Checkbutton correspondente ao procedimento e vincule a função abrir_tela_grupo
                        checkbox = tk.Checkbutton(checkbox_canvas_frame, text=procedimento, variable=procedimento_vars_por_coluna[col_idx][row_idx], anchor="w", padx=0, pady=0, fg=texto_cor)
                        checkbox.grid(row=row_idx, column=col_idx, padx=0, pady=0, sticky="w")
                        checkbox.bind("<Button-1>", lambda event, p=procedimento: abrir_tela_grupo(p))
                        checkbuttons.append(checkbox)  # Adicione o Checkbutton à lista
                    else:
                        # Se não atender aos critérios, crie o Checkbutton normalmente, mas sem a função abrir_tela_grupo
                        checkbox = tk.Checkbutton(checkbox_canvas_frame, text=procedimento, variable=procedimento_vars_por_coluna[col_idx][row_idx], anchor="w", padx=0, pady=0)
                        checkbox.grid(row=row_idx, column=col_idx, padx=0, pady=0, sticky="w")
                        checkbuttons.append(checkbox)  # Adicione o Checkbutton à lista

        # Função para pesquisar os Checkbuttons em tempo real
        def pesquisar_checkbuttons(*args):
            termo_pesquisa = caixa_pesquisa.get().lower()  # Obtém o termo de pesquisa em minúsculas
            for checkbox in checkbuttons:
                texto_checkbox = checkbox.cget("text").lower()  # Obtém o texto do Checkbutton em minúsculas
                if termo_pesquisa == "" or termo_pesquisa in texto_checkbox:
                    checkbox.grid()  # Exibe o Checkbutton se corresponder à pesquisa
                else:
                    checkbox.grid_remove()  # Oculta o Checkbutton se não corresponder à pesquisa

        # Função para limpar a pesquisa
        def limpar_pesquisa():
            caixa_pesquisa.delete(0, tk.END)  # Limpa o conteúdo da caixa de pesquisa
            pesquisar_checkbuttons()  # Mostra todos os procedimentos novamente

        # Função para salvar procedimentos selecionados
        def salvar_selecionados():
            procedimentos_selecionados = []
            for col_idx, coluna in enumerate(colunas):
                for row_idx, procedimento in enumerate(coluna):
                    if procedimento_vars_por_coluna[col_idx][row_idx].get() == 1:
                        procedimentos_selecionados.append(nome_procedimentos_geral_inicio[col_idx * num_colunas + row_idx])

            if not procedimentos_selecionados:
                messagebox.showerror("Erro", "Nenhum procedimento foi selecionado. Selecione pelo menos 1 procedimento na lista.")
            else:
                with open('procedimentos_selecionados.json', 'w') as file:
                    json.dump(procedimentos_selecionados, file)
                messagebox.showinfo("Informação", "Procedimentos selecionados foram salvos com sucesso!")

        # Função para carregar procedimentos selecionados
        def carregar_selecionados():
            try:
                with open('procedimentos_selecionados.json', 'r') as file:
                    procedimentos_selecionados = json.load(file)

                procedimentos_nao_encontrados = []

                for col_idx, coluna in enumerate(colunas):
                    for row_idx, procedimento in enumerate(coluna):
                        var = procedimento_vars_por_coluna[col_idx][row_idx]
                        nome_procedimento = nome_procedimentos_geral_inicio[col_idx * num_colunas + row_idx]
                        if nome_procedimento in procedimentos_selecionados:
                            var.set(1)
                        else:
                            var.set(0)

                for procedimento in procedimentos_selecionados:
                    encontrado = False
                    for col_idx, coluna in enumerate(colunas):
                        for row_idx, nome_procedimento in enumerate(coluna):
                            if nome_procedimento == procedimento:
                                encontrado = True
                                break
                        if encontrado:
                            break
                    if not encontrado:
                        procedimentos_nao_encontrados.append(procedimento)

                if procedimentos_nao_encontrados:
                    mensagem = "Os seguintes procedimentos do arquivo não foram encontrados no tkinter:\n"
                    mensagem += "\n".join(procedimentos_nao_encontrados)
                    messagebox.showwarning("Aviso", mensagem)
                else:
                    messagebox.showinfo("Informação", "Procedimentos selecionados foram carregados com sucesso!")

            except FileNotFoundError:
                messagebox.showerror("Erro", "Arquivo 'procedimentos_selecionados.json' não encontrado.")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar procedimentos: {str(e)}")

        # Função para verificar a seleção
        def verificar_selecao():
            selecionado = False
            codigo_procedimentos_geral.clear()
            for col_idx, coluna in enumerate(colunas):
                for row_idx, procedimento in enumerate(coluna):
                    if procedimento_vars_por_coluna[col_idx][row_idx].get() == 1:
                        selecionado = True
                        codigo_procedimentos_geral.append(codigo_procedimentos[col_idx * num_colunas + row_idx])
            if not selecionado:
                messagebox.showerror("Erro", "Escolha pelo menos 1 procedimento na lista")
            else:
                start_extraction()
        
        

        def extraindo_vagas():
            

            Barra_de_progresso = tk.Toplevel()
            Barra_de_progresso.title("Barra de Progresso")

            
            progress_label = ttk.Label(Barra_de_progresso, text="Progresso: 0%")
            progress_label.pack()

            progress_bar = ttk.Progressbar(Barra_de_progresso, length=300, mode="determinate")
            progress_bar.pack()

            # Impede a interação com a tela anterior
            Barra_de_progresso.grab_set()
            
            # Determina o número total de iterações em todos os loops combinados
            
            
            driver.get("https://sisregiii.saude.gov.br/cgi-bin/cadweb50?url=/cgi-bin/marcar")

            sleep(1)

            df = pd.read_excel ('Banco_de_Dados.xlsx', sheet_name= 0)
            for index, row in df.iterrows():
                CampoCnes = driver.find_element(By.NAME,'nu_cns')
                cnes_f = (str(row['cnes masculino']))
                CampoCnes.send_keys(cnes_f)  #Cnes masculino Adulto 

    

            driver.find_element(By.NAME,'btn_pesquisar').click()
            sleep(1)
            driver.find_element(By.NAME,'btn_continuar').click()


            sleep(1)
            CampoCid = driver.find_element(By.NAME,'cid10')
            CampoCid.send_keys('z000')  #cid10

            SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
            SelectMedico.select_by_value('00000000000') #Profissional não listado
            sleep(1)
            CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
            CampoMedico.send_keys('Medico')  #Nome da Doutora

            # Crie uma função para atualizar a barra de progresso
            def update_progress(current_progress, text):
                progress_label.config(text=f"Progresso: {current_progress}% - {text}")
                progress_bar["value"] = current_progress
                Barra_de_progresso.update()
            

            data = {'CODIGO':codigo_procedimentos_geral}
            dfnew = pd.DataFrame(data)
            total_1 = dfnew.shape[0]
            #dfnew.to_excel('Lista de Procedimentos.xlsx', sheet_name='Geral_Unisex', index=False)
            
            
            with tqdm(total=total_1) as pbar:
                for index, row in dfnew.iterrows():
                    text = f"Iteração {index + 1} de {total_1} - Consultando Prodedimentos Gerais"     

                    select_procedimento = Select(driver.find_element(By.NAME,"pa"))
                    select_procedimento.select_by_value(row["CODIGO"])
                    elementos_select = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/table[2]/tbody/tr[6]/td[2]/select')
                    elementos_option = elementos_select.find_elements(By.TAG_NAME, 'option')

                    quantidade_de_elementos = len(elementos_option)

                    if quantidade_de_elementos > 1:
                        SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
                        SelectMedico.select_by_index(0) #Profissional não listado
                        sleep(1)
                        SelectMedico.select_by_value('00000000000') #Profissional não listado
                        sleep(1)
                        CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
                        CampoMedico.send_keys('Medico')  #Nome da Doutora
                        central_executante  = Select(driver.find_element(By.NAME, 'centralexec'))
                        central_executante.select_by_index(1)
                    

                    driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center/input').click() # Procura elemento e clica
                
                    
                    try:
                        driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr[2]/td/font/center').text
                        erro = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr[2]/td/font/center').text
                        erro2 = driver.find_element(By.XPATH,'//*[@id="main_div"]/center/table/tbody/tr[3]/td[2]').text
                    
                        #O codigo numbers a seguir procura por numeros dentro de um texto e o codigo "\d+" significa
                        #que ele deve procurar por um numero inteiro
                    
                        numbers = re.findall(r'\d+|\d+', erro2) 
                        idade_menor = int(numbers[0])
                        idade_maior = int(numbers[1])
                        idoso = idade_menor >= 40 and idade_maior < 500
                        pediatrico = idade_menor >= 0 and idade_menor < 18     
                        if erro == 'Sexo do usuário incompatível com o procedimento.':
                            codigo_procedimentos_mulher.append(row["CODIGO"])
                        elif erro == 'Idade do usuário incompatível com o procedimento.' and idoso:
                            codigo_procedimentos_idoso.append(row["CODIGO"])
                        elif erro == 'Idade do usuário incompatível com o procedimento.' and pediatrico:
                            codigo_procedimentos_crianca.append(row["CODIGO"])
                    except:
                        try: # Se aparecer o botão btnSolicitar então recebe 0 caso contrario recebe Vagas Encontradas
                            resultado = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr/td/center/b').text
                            if resultado == 'PROCEDIMENTO REGULADO':
                                pass
                            elif resultado == 'VAGAS DISPONÍVEIS':
                                #codigo_procedimentos_geral.append(row["CODIGO"])
                                nome_geral = driver.find_element(By.XPATH,'//*[@id="main_div"]/center/table/tbody/tr[2]/td[2]/font').text
                                nome_procedimentos_geral.append(nome_geral)
                                Tem_vaga = "Vagas Encontradas"
                                vagas_geral.append(Tem_vaga) # Adiciona o valor na lista
                            elif resultado == 'NENHUMA VAGA ENCONTRADA':
                                pass
                        except:
                            driver.find_element(By.NAME,'max_count')
                        
                            driver.find_element(By.XPATH,'//*[@id="main_div"]/center/center/form/center/table[1]/tbody/tr[2]/td[1]/input').click()
                            sleep(1)                        
                            driver.find_element(By.NAME,'btnConfirmar').click()

                            resultado2 = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr/td/center/b').text

                            if resultado2 == 'PROCEDIMENTO REGULADO':
                                pass
                            elif resultado2 == 'VAGAS DISPONÍVEIS':
                                #codigo_procedimentos_geral.append(row["CODIGO"])
                                nome_geral = driver.find_element(By.XPATH,'//*[@id="main_div"]/center/table/tbody/tr[2]/td[2]/font').text
                                nome_procedimentos_geral.append(nome_geral)
                                Tem_vaga = "Vagas Encontradas"
                                vagas_geral.append(Tem_vaga) # Adiciona o valor na lista
                            elif resultado2 == 'NENHUMA VAGA ENCONTRADA':
                                pass
                            driver.find_element(By.NAME,'btnVoltar').click() # Procura elemento voltar e clica
                    
                    # Atualize a barra de progresso após cada iteração
                    current_progress = int((index + 1) / total_1 * 100)
                    update_progress(current_progress, text)
                    pbar.update(1)
                    driver.find_element(By.NAME,'btnVoltar').click() # Procura elemento voltar e clica
            
                    
            

            def extracao_mulher_geral():

                data2 = {'CODIGO': codigo_procedimentos_mulher} 
                df_mulher = pd.DataFrame(data2)
                total_2 = df_mulher.shape[0]

                

                driver.get("https://sisregiii.saude.gov.br/cgi-bin/cadweb50?url=/cgi-bin/marcar")

                for index, row in df.iterrows():
                    CampoCnes = driver.find_element(By.NAME,'nu_cns')
                    cnes_m = (str(row['cnes feminino']))
                    CampoCnes.send_keys(cnes_m)  #Cnes feminino 


                driver.find_element(By.NAME,'btn_pesquisar').click()

                sleep(1)

                driver.find_element(By.NAME,'btn_continuar').click()

                sleep(3)
                CampoCid = driver.find_element(By.NAME,'cid10')
                CampoCid.send_keys('z000')  #cid10

                SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
                SelectMedico.select_by_value('00000000000') #Profissional não listado
                sleep(1)
                CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
                CampoMedico.send_keys('Medico')  #Nome da Doutora
                sleep(1)
                with tqdm(total=total_2) as pbar:
                    for index, row in df_mulher.iterrows():
                        text = f"Iteração {index + 1} de {total_1} - Consultando Prodedimentos Femininos"
                        select_procedimento = Select(driver.find_element(By.NAME,"pa"))
                        select_procedimento.select_by_value(row["CODIGO"])
                        elementos_select = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/table[2]/tbody/tr[6]/td[2]/select')
                        elementos_option = elementos_select.find_elements(By.TAG_NAME, 'option')
                        quantidade_de_elementos = len(elementos_option)

                        if quantidade_de_elementos > 1:
                            SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
                            SelectMedico.select_by_index(0) #Profissional não listado
                            sleep(1)
                            SelectMedico.select_by_value('00000000000') #Profissional não listado
                            sleep(1)
                            CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
                            CampoMedico.send_keys('Medico')  #Nome da Doutora
                            central_executante  = Select(driver.find_element(By.NAME, 'centralexec'))
                            central_executante.select_by_index(1)

                        driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center/input').click() # Procura elemento e clica
                        sleep(1)
                        try: # Se aparecer o botão btnSolicitar então recebe 0 caso contrario recebe Vagas Encontradas
                            resultado = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr/td/center/b').text
                            if resultado == 'PROCEDIMENTO REGULADO':
                                pass
                            elif resultado == 'VAGAS DISPONÍVEIS':
                                nome_mulher = driver.find_element(By.XPATH,'//*[@id="main_div"]/center/table/tbody/tr[2]/td[2]/font').text
                                nome_procedimentos_mulher.append(nome_mulher)
                                Tem_vaga = "Vagas Encontradas"
                                vagas_mulher.append(Tem_vaga) # Adiciona o valor na lista
                            elif resultado == 'NENHUMA VAGA ENCONTRADA':
                                pass
                        except:
                            pass
                        # Atualize a barra de progresso após cada iteração
                        current_progress = int((index + 1) / total_2 * 100)
                        update_progress(current_progress, text)
                        pbar.update(1)       
                    
                        driver.find_element(By.NAME,'btnVoltar').click() # Procura elemento voltar e clica'
            
            if  len(codigo_procedimentos_mulher) <= 0:
                pass
            else:
                extracao_mulher_geral()

            

            def extracao_crianca_geral():

                data3 = {'CODIGO': codigo_procedimentos_crianca} 
                df_crianca = pd.DataFrame(data3)
                total_3 = df_crianca.shape[0]

                

                driver.get("https://sisregiii.saude.gov.br/cgi-bin/cadweb50?url=/cgi-bin/marcar")

                for index, row in df.iterrows():
                    CampoCnes = driver.find_element(By.NAME,'nu_cns')
                    cnes_c = (str(row['cnes crianca']))
                    CampoCnes.send_keys(cnes_c)  #Cnes criança


                driver.find_element(By.NAME,'btn_pesquisar').click()

                sleep(1)

                driver.find_element(By.NAME,'btn_continuar').click()

                sleep(1)
                CampoCid = driver.find_element(By.NAME,'cid10')
                CampoCid.send_keys('z000')  #cid10

                SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
                SelectMedico.select_by_value('00000000000') #Profissional não listado
                sleep(1)
                CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
                CampoMedico.send_keys('Medico')  #Nome da Doutora

                with tqdm(total=total_3) as pbar:
                    for index, row in df_crianca.iterrows():
                        text = f"Iteração {index + 1} de {total_1} - Consultando Prodedimentos Pediatricos"
                        select_procedimento = Select(driver.find_element(By.NAME,"pa"))
                        select_procedimento.select_by_value(row["CODIGO"])
                        elementos_select = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/table[2]/tbody/tr[6]/td[2]/select')
                        elementos_option = elementos_select.find_elements(By.TAG_NAME, 'option')
                        quantidade_de_elementos = len(elementos_option)

                        if quantidade_de_elementos > 1:
                            SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
                            SelectMedico.select_by_index(0) #Profissional não listado
                            sleep(1)
                            SelectMedico.select_by_value('00000000000') #Profissional não listado
                            sleep(1)
                            CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
                            CampoMedico.send_keys('Medico')  #Nome da Doutora
                            central_executante  = Select(driver.find_element(By.NAME, 'centralexec'))
                            central_executante.select_by_index(1)
                        driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center/input').click() # Procura elemento e clica
                        sleep(1)
                        try: # Se aparecer o botão btnSolicitar então recebe 0 caso contrario recebe Vagas Encontradas
                            resultado = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr/td/center/b').text
                            if resultado == 'PROCEDIMENTO REGULADO':
                                pass
                            elif resultado == 'VAGAS DISPONÍVEIS':
                                nome_crianca = driver.find_element(By.XPATH,'//*[@id="main_div"]/center/table/tbody/tr[2]/td[2]/font').text
                                nome_procedimentos_crianca.append(nome_crianca)
                                Tem_vaga = "Vagas Encontradas"
                                vagas_crianca.append(Tem_vaga) # Adiciona o valor na lista
                            elif resultado == 'NENHUMA VAGA ENCONTRADA':
                                pass
                        except:
                            pass

                        # Atualize a barra de progresso após cada iteração
                        current_progress = int((index + 1) / total_3 * 100)
                        update_progress(current_progress,text)
                        pbar.update(1)
                        driver.find_element(By.NAME,'btnVoltar').click() # Procura elemento voltar e clica
            if  len(codigo_procedimentos_crianca) <= 0:
                pass
            else:
                extracao_crianca_geral()

            

            def extracao_idoso_geral():

                data4 = {'CODIGO': codigo_procedimentos_idoso} 
                df_idoso = pd.DataFrame(data4)
                total_4 = df_idoso.shape[0]

                

                driver.get("https://sisregiii.saude.gov.br/cgi-bin/cadweb50?url=/cgi-bin/marcar")

                for index, row in df.iterrows():
                    CampoCnes = driver.find_element(By.NAME,'nu_cns')
                    cnes_i = (str(row['cnes idoso']))
                    CampoCnes.send_keys(cnes_i)  #Cnes idoso 


                driver.find_element(By.NAME,'btn_pesquisar').click()

                sleep(1)

                driver.find_element(By.NAME,'btn_continuar').click()

                sleep(1)
                CampoCid = driver.find_element(By.NAME,'cid10')
                CampoCid.send_keys('z000')  #cid10

                SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
                SelectMedico.select_by_value('00000000000') #Profissional não listado
                sleep(1)
                CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
                CampoMedico.send_keys('Medico')  #Nome da Doutora

                with tqdm(total=total_1) as pbar:
                    for index, row in df_idoso.iterrows():
                        text = f"Iteração {index + 1} de {total_1} - Consultando Prodedimentos Geriatricos"
                        select_procedimento = Select(driver.find_element(By.NAME,"pa"))
                        select_procedimento.select_by_value(row["CODIGO"])
                        elementos_select = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/table[2]/tbody/tr[6]/td[2]/select')
                        elementos_option = elementos_select.find_elements(By.TAG_NAME, 'option')
                        quantidade_de_elementos = len(elementos_option)

                        if quantidade_de_elementos > 1:
                            SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
                            SelectMedico.select_by_index(0) #Profissional não listado
                            sleep(1)
                            SelectMedico.select_by_value('00000000000') #Profissional não listado
                            sleep(1)
                            CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
                            CampoMedico.send_keys('Medico')  #Nome da Doutora
                            central_executante  = Select(driver.find_element(By.NAME, 'centralexec'))
                            central_executante.select_by_index(1)
                        driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center/input').click() # Procura elemento e clica
                        sleep(1)
                        try: # Se aparecer o botão btnSolicitar então recebe 0 caso contrario recebe Vagas Encontradas
                            resultado = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr/td/center/b').text
                            if resultado == 'PROCEDIMENTO REGULADO':
                                pass
                            elif resultado == 'VAGAS DISPONÍVEIS':
                                nome_idoso = driver.find_element(By.XPATH,'//*[@id="main_div"]/center/table/tbody/tr[2]/td[2]/font').text
                                nome_procedimentos_idoso.append(nome_idoso)
                                Tem_vaga = "Vagas Encontradas"
                                vagas_idoso.append(Tem_vaga) # Adiciona o valor na lista
                            elif resultado == 'NENHUMA VAGA ENCONTRADA':
                                pass
                        except:
                            pass

                        # Atualize a barra de progresso após cada iteração
                        current_progress = int((index + 1) / total_4 * 100)
                        update_progress(current_progress,text)
                        pbar.update(1)
                        driver.find_element(By.NAME,'btnVoltar').click() # Procura elemento voltar e clica

            if  len(codigo_procedimentos_idoso) <= 0:
                pass
            else:
                extracao_idoso_geral()
            
            procedimentos_concateandos = []
            vagas_concatenadas = []


            for nome_todos, vaga_todos in zip(nome_procedimentos_geral + nome_procedimentos_mulher + nome_procedimentos_crianca + nome_procedimentos_idoso, vagas_geral + vagas_mulher + vagas_crianca + vagas_idoso):
                procedimentos_concateandos.append(nome_todos)
                vagas_concatenadas.append(vaga_todos)

            procedimentos_concateandos.sort()

            data5 = {'Procedimentos': procedimentos_concateandos, 'Vagas': vagas_concatenadas}
            global df_final_geral
            df_final_geral = pd.DataFrame(data5)

            
            Barra_de_progresso.destroy()
            Barra_de_progresso.grab_release()        
            # Agende a chamada para exibir_dataframe() na thread principal
            abrir_extrair_procedimentos.after(0, exibir_dataframe)

        def start_extraction():
            progress_thread = threading.Thread(target=extraindo_vagas)
            progress_thread.start()

        

        # Função para selecionar todos os Checkbuttons
        def selecionar_todos():
            for col_idx, coluna in enumerate(colunas):
                for row_idx, procedimento in enumerate(coluna):
                    procedimento_vars_por_coluna[col_idx][row_idx].set(1)

        # Função para desmarcar todos os Checkbuttons
        def desmarcar_todos():
            for col_idx, coluna in enumerate(colunas):
                for row_idx, procedimento in enumerate(coluna):
                    procedimento_vars_por_coluna[col_idx][row_idx].set(0)

        # Função para desmarcar Checkbuttons com labels azuis
        def desmarcar_checkbuttons_azuis():
            for checkbox in checkbuttons:
                if checkbox.cget("fg") == "blue":
                    checkbox.deselect()



        # Cria a janela principal
        abrir_extrair_procedimentos = tk.Toplevel()
        abrir_extrair_procedimentos.title("Seleção de Procedimentos")

        # Cria uma lista de listas para armazenar as variáveis dos procedimentos em cada coluna
        procedimento_vars_por_coluna = []

        # Divide os procedimentos em colunas de no máximo 
        num_colunas = 500
        colunas = [nome_procedimentos_geral_inicio[i:i + num_colunas] for i in range(0, len(nome_procedimentos_geral_inicio), num_colunas)]

        # Inicializa as variáveis para cada coluna
        for coluna in colunas:
            procedimento_vars = []
            for _ in coluna:
                var = tk.IntVar()
                procedimento_vars.append(var)
            procedimento_vars_por_coluna.append(procedimento_vars)

        # Obtém as dimensões da tela
        largura_tela = 1024
        altura_tela = 768

        # Altura da barra de tarefas (assumindo que a barra de tarefas esteja na parte inferior)
        altura_barra_tarefas = 100  # Ajuste conforme necessário

        # Calcula a altura da janela para que ela não ultrapasse a barra de tarefas
        altura_janela = altura_tela - altura_barra_tarefas

        # Define a geometria da janela
        abrir_extrair_procedimentos.geometry(f"{largura_tela}x{altura_janela}")

        # Cria um frame para os botões e o posiciona na parte superior da janela usando o gerenciador de geometria "grid"
        button_frame = tk.Frame(abrir_extrair_procedimentos)
        button_frame.grid(column=0, row=0, columnspan=6, padx=10, pady=20)

        # Cria um botão com o texto "Extrair vagas novamente" e vincula a função extrair_vagas a ele
        botao_extrair_vaga_novamente = tk.Button(button_frame, text="Atualizar Lista", command=abrir_extrair_procedimentos)
        botao_extrair_vaga_novamente.grid(column=0, row=0, padx=5)

        # Botão "Pesquisar Vagas"
        botao_pesquisar = tk.Button(button_frame, text="Pesquisar Vagas", command=verificar_selecao)
        botao_pesquisar.grid(column=1, row=0, padx=5)

        # Botão "Pesquisar Procedimentos Regulados"
        botao_pesquisar_regulados = tk.Button(button_frame, text="Pesquisar Procedimentos Regulados")
        botao_pesquisar_regulados.grid(column=2, row=0, padx=5)

        # Botão "Pesquisar Procedimentos Devolvidos"
        botao_pesquisar_devolvidos = tk.Button(button_frame, text="Pesquisar Procedimentos Devolvidos")
        botao_pesquisar_devolvidos.grid(column=3, row=0, padx=5)

        # Botão "Salvar Selecionados"
        botao_salvar = tk.Button(button_frame, text="Salvar Selecionados", command=salvar_selecionados)
        botao_salvar.grid(column=0, row=1, padx=5)

        # Botão "Carregar Selecionados"
        botao_carregar = tk.Button(button_frame, text="Carregar Selecionados", command=carregar_selecionados)
        botao_carregar.grid(column=1, row=1, padx=5)

        # Botão "Selecionar Todos"
        botao_selecionar_todos = tk.Button(button_frame, text="Selecionar Todos", command=selecionar_todos)
        botao_selecionar_todos.grid(column=2, row=1, padx=5)

        # Botão "Desmarcar Todos"
        botao_desmarcar_todos = tk.Button(button_frame, text="Desmarcar Todos", command=desmarcar_todos)
        botao_desmarcar_todos.grid(column=3, row=1, padx=5)

        # Botão "Desmarcar Todos"
        botao_desmarcar_todos_grupo = tk.Button(button_frame, text="Desmarcar Todos Grupos", command=desmarcar_checkbuttons_azuis)
        botao_desmarcar_todos_grupo.grid(column=4, row=1, padx=5)
        

        # Cria um frame para os Checkbuttons
        checkbox_frame = tk.Frame(abrir_extrair_procedimentos)
        checkbox_frame.grid(column=0, row=1, columnspan=1, padx=0, pady=0)

        # Função para redimensionar o Canvas quando a tela for redimensionada
        def resize_canvas(event):
            canvas.config(scrollregion=canvas.bbox("all"), width=event.width, height=event.height)

        # Cria um Canvas para os Checkbuttons
        canvas = Canvas(checkbox_frame, width=400, height=600)  # Defina a largura e altura desejadas
        canvas.grid(row=0, column=0, padx=0, pady=0, sticky="nsew")

        # Configure o tamanho máximo do Canvas com base na altura disponível
        canvas_max_height = altura_janela - 150  # Ajuste conforme necessário
        canvas.config(height=canvas_max_height)

        # Cria uma barra de rolagem vertical para o Canvas
        scrollbar_y = Scrollbar(checkbox_frame, orient="vertical", command=canvas.yview)
        scrollbar_y.grid(row=0, column=1, padx=0, pady=0, sticky="ns")

        # Crie uma barra de rolagem horizontal para o Canvas
        scrollbar_x = Scrollbar(checkbox_frame, orient="horizontal", command=canvas.xview)
        scrollbar_x.grid(row=1, column=0, padx=0, pady=0, sticky="ew")

        # Configure as barras de rolagem para controlar o Canvas
        canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # Crie um novo frame dentro do Canvas para os Checkbuttons
        checkbox_canvas_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=checkbox_canvas_frame, anchor="nw")

        # Lista para armazenar os Checkbuttons
        checkbuttons = []

        # Dicionário para associar Checkbuttons aos rótulos
        label_dict = {}

        # Chame a função criar_checkbuttons aqui
        criar_checkbuttons()
        

        # Atualize a visualização do Canvas e configure as barras de rolagem para rolar o Canvas
        checkbox_canvas_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

        # Lista para armazenar os rótulos
        labels = []

        # Inicialize a variável para acompanhar o índice do item selecionado durante a pesquisa
        indice_item_selecionado = -1

        # Crie um frame para a pesquisa
        pesquisa_frame = tk.Frame(checkbox_frame)
        pesquisa_frame.grid(row=0, column=2, padx=10, pady=10, sticky="n")

        # Crie um Label para o texto "Pesquise o procedimento"
        label_pesquisa = tk.Label(pesquisa_frame, text="Pesquise o procedimento", fg="red")
        label_pesquisa.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # Crie uma Entry para a caixa de pesquisa
        caixa_pesquisa = tk.Entry(pesquisa_frame)
        caixa_pesquisa.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        # Vincule a função pesquisar_checkbuttons ao evento de mudança de texto na caixa de pesquisa
        caixa_pesquisa.bind("<KeyRelease>", pesquisar_checkbuttons)

        # Crie um botão "Limpar" para redefinir a pesquisa
        botao_limpar_pesquisa = tk.Button(pesquisa_frame, text="Limpar", command=limpar_pesquisa)
        botao_limpar_pesquisa.grid(row=1, column=1, padx=5, pady=5)

        abrir_extrair_procedimentos.mainloop()

    def abrir_tela_grupo(nome_procedimento):
        procedimento_extraidos_grupo = []
        codigo_procedimento_extraidos_grupo = []
        codigo_procedimento_grupo = []

        nome_procedimento_geral_grupo = []

        nome_procedimentos_mulher_grupo = []
        codigo_procedimentos_mulher_grupo = []

        nome_procedimentos_idoso_grupo = []
        codigo_procedimentos_idoso_grupo = []

        nome_procedimentos_crianca_grupo = []
        codigo_procedimentos_crianca_grupo = []

        vagas_geral_grupo = [] #lista que irá receber os valores de vagas disponiveis Geral Unissex
        vagas_mulher_grupo = [] #lista que irá receber os valores de vagas disponiveis Geral Feminino
        vagas_crianca_grupo = [] #lista que irá receber os valores de vagas disponiveis Criança
        vagas_idoso_grupo = []
        
        variavel_titulo = nome_procedimento
        variavel_codigo_titulo = None

        # Verifique se o valor está presente em nome_procedimentos_geral_inicio
        if variavel_titulo in nome_procedimentos_geral_inicio:
            # Encontre o índice do valor em nome_procedimentos_geral_inicio
            indice = nome_procedimentos_geral_inicio.index(variavel_titulo)
            
            # Use o índice para acessar o valor correspondente em codigo_procedimentos
            variavel_codigo_titulo = codigo_procedimentos[indice]
        else:
            # Se o valor não estiver presente, defina variavel_codigo_titulo como None
            variavel_codigo_titulo = None

        driver.get("https://sisregiii.saude.gov.br/cgi-bin/cadweb50?url=/cgi-bin/marcar")

        sleep(1)

        df = pd.read_excel ('Banco_de_Dados.xlsx', sheet_name= 0)
        for index, row in df.iterrows():
            CampoCnes = driver.find_element(By.NAME,'nu_cns')
            cnes_f = (str(row['cnes masculino']))
            CampoCnes.send_keys(cnes_f)  #Cnes masculino Adulto 

    

        driver.find_element(By.NAME,'btn_pesquisar').click()
        sleep(1)
        driver.find_element(By.NAME,'btn_continuar').click()


        sleep(1)
        CampoCid = driver.find_element(By.NAME,'cid10')
        CampoCid.send_keys('z000')  #cid10

        SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
        SelectMedico.select_by_value('00000000000') #Profissional não listado
        sleep(1)
        CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
        CampoMedico.send_keys('Medico')  #Nome da Doutora

        select_procedimento = Select(driver.find_element(By.NAME,"pa"))
        select_procedimento.select_by_value(variavel_codigo_titulo)
        
        driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center/input').click() # Procura elemento e clica
        

        # Localize a tabela pelo seletor CSS
        tabela = driver.find_element(By.CLASS_NAME, 'table_listagem')

        # Iterar sobre as linhas da tabela
        linhas = tabela.find_elements(By.TAG_NAME,'tr')
        print(linhas)
        # Itere pelas linhas da tabela, ignorando a primeira linha de cabeçalho
        for linha in linhas[1:]:
            # Encontre o primeiro e o segundo elemento (células) dentro da tag 'tr'
            primeira_celula = linha.find_elements(By.TAG_NAME, 'td')[0]
            segunda_celula = linha.find_elements(By.TAG_NAME, 'td')[1]
            
            # Extraia o texto das células
            texto_primeira_celula = primeira_celula.text
            texto_segunda_celula = segunda_celula.text
            
            # Extraia o valor do atributo 'name' do primeiro elemento
            codigo = primeira_celula.find_element(By.TAG_NAME, 'input').get_attribute('name')
            # Extraia o valor do atributo 'name' do segundo elemento
            codigo2 = segunda_celula.find_element(By.TAG_NAME, 'input').get_attribute('name')
            
            # Adicione o texto e o código às listas
            procedimento_extraidos_grupo.append(texto_primeira_celula)
            codigo_procedimento_extraidos_grupo.append(codigo)
            procedimento_extraidos_grupo.append(texto_segunda_celula)
            codigo_procedimento_extraidos_grupo.append(codigo2)
            
        
        
        # Encontre o elemento <script> que contém a variável g_QtdMaxItens
        script_element = driver.find_element(By.XPATH, "//script[contains(., 'g_QtdMaxItens')]")

        # Obtenha o texto dentro do elemento <script>
        script_text = script_element.get_attribute("text")

        # Extrair o valor da variável g_QtdMaxItens usando expressões regulares
        match = re.search(r'g_QtdMaxItens\s*=\s*(\d+);', script_text)

        if match:
            g_QtdMaxItens = int(match.group(1))
                
        else:
            messagebox.showerror("Erro", "A variável g_QtdMaxItens não foi encontrada.")

        # Armazene o valor em uma variável Python
        valor_g_QtdMaxItens = g_QtdMaxItens


        def salvar_selecionados():
            procedimentos_selecionados = []
            for col_idx, coluna in enumerate(colunas):
                for row_idx, procedimento in enumerate(coluna):
                    if procedimento_vars_por_coluna[col_idx][row_idx].get() == 1:
                        procedimentos_selecionados.append(procedimento_extraidos_grupo[col_idx * num_colunas + row_idx])

        # Usar o nome do procedimento como nome do arquivo JSON
            nome_arquivo = f'{nome_procedimento}.json'
            with open(nome_arquivo, 'w') as file:
                json.dump(procedimentos_selecionados, file)

            messagebox.showinfo("Informação", "Procedimentos selecionados foram salvos com sucesso!")

        def carregar_selecionados():
            try:
                # Usar o nome do procedimento como nome do arquivo JSON
                nome_arquivo = f'{variavel_titulo}.json'
                with open(nome_arquivo, 'r') as file:
                    procedimentos_selecionados = json.load(file)

                procedimentos_nao_encontrados = []

                for col_idx, coluna in enumerate(colunas):
                    for row_idx, procedimento in enumerate(coluna):
                        var = procedimento_vars_por_coluna[col_idx][row_idx]
                        nome_procedimento = procedimento_extraidos_grupo[col_idx * num_colunas + row_idx]
                        if nome_procedimento in procedimentos_selecionados:
                            var.set(1)
                        else:
                            var.set(0)

                for procedimento in procedimentos_selecionados:
                    encontrado = False
                    for col_idx, coluna in enumerate(colunas):
                        for row_idx, nome_procedimento in enumerate(coluna):
                            if nome_procedimento == procedimento:
                                encontrado = True
                                break
                        if encontrado:
                            break
                    if not encontrado:
                        procedimentos_nao_encontrados.append(procedimento)

                if procedimentos_nao_encontrados:
                    mensagem = "Os seguintes procedimentos do arquivo não foram encontrados no tkinter:\n"
                    mensagem += "\n".join(procedimentos_nao_encontrados)
                    messagebox.showwarning("Aviso", mensagem)
                else:
                    messagebox.showinfo("Informação", "Procedimentos selecionados foram carregados com sucesso!")

            except FileNotFoundError:
                messagebox.showerror("Erro", f"Arquivo '{nome_arquivo}' não encontrado.")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar procedimentos: {str(e)}")

        
        
        def verificar_selecao():
            selecionado = False
            codigo_procedimento_grupo.clear()
            for col_idx, coluna in enumerate(colunas):
                for row_idx, procedimento in enumerate(coluna):
                    if procedimento_vars_por_coluna[col_idx][row_idx].get() == 1:
                        selecionado = True
                        codigo_procedimento_grupo.append(codigo_procedimento_extraidos_grupo[col_idx * num_colunas + row_idx])
            if not selecionado:
                messagebox.showerror("Erro", "Escolha pelo menos 1 procedimento na lista")
            else:
                start_extraction()  

        def extraindo_vagas_grupo():
            

            data = {'CODIGO':codigo_procedimento_grupo}
            dfnew = pd.DataFrame(data)
            total_1 = dfnew.shape[0]
                    
            

            # Itere pelos códigos e clique nos elementos correspondentes
            for codigo in dfnew['CODIGO']:
                try:
                    # Localize a tabela pelo seletor CSS
                    tabela = driver.find_element(By.CLASS_NAME, 'table_listagem')
                    # Dentro da tabela, localize o elemento pelo atributo "name" igual ao código
                    checkbox_element = tabela.find_element(By.NAME, codigo)

                    # Clique no checkbox
                    checkbox_element.click()
                    
                    driver.find_element(By.NAME, 'btnConfirmar').click() # Procura elemento e clica
                    
                    try:
                        driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr[2]/td/font/center').text
                        erro = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr[2]/td/font/center').text
                        erro2 = driver.find_element(By.XPATH,'//*[@id="main_div"]/center/table/tbody/tr[3]/td[2]').text
                    
                        #O codigo numbers a seguir procura por numeros dentro de um texto e o codigo "\d+" significa
                        #que ele deve procurar por um numero inteiro
                    
                        numbers = re.findall(r'\d+|\d+', erro2) 
                        idade_menor = int(numbers[0])
                        idade_maior = int(numbers[1])
                        idoso = idade_menor >= 40 and idade_maior < 500
                        pediatrico = idade_menor >= 0 and idade_menor < 18     
                        if erro == 'Sexo do usuário incompatível com o procedimento.':
                            codigo_procedimentos_mulher_grupo.append(row["CODIGO"])
                        elif erro == 'Idade do usuário incompatível com o procedimento.' and idoso:
                            codigo_procedimentos_idoso_grupo.append(row["CODIGO"])
                        elif erro == 'Idade do usuário incompatível com o procedimento.' and pediatrico:
                            codigo_procedimentos_crianca_grupo.append(row["CODIGO"])
                    except:
                        try: # Se aparecer o botão btnSolicitar então recebe 0 caso contrario recebe Vagas Encontradas
                            resultado = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr/td/center/b').text
                            if resultado == 'PROCEDIMENTO REGULADO':
                                pass
                            elif resultado == 'VAGAS DISPONÍVEIS':
                                nome_geral = driver.find_element(By.XPATH,'//*[@id="main_div"]/center/table/tbody/tr[3]/td[2]/i').text
                                nome_procedimento_geral_grupo.append(nome_geral)
                                Tem_vaga = "Vagas Encontradas"
                                vagas_geral_grupo.append(Tem_vaga) # Adiciona o valor na lista
                            elif resultado == 'NENHUMA VAGA ENCONTRADA':
                                pass
                        except:
                            pass
                    
                    
                    driver.find_element(By.NAME,'btnVoltar').click() # Procura elemento voltar e clica
                    
                    # Localize a tabela pelo seletor CSS
                    tabela = driver.find_element(By.CLASS_NAME, 'table_listagem')
                    # Dentro da tabela, localize o elemento pelo atributo "name" igual ao código
                    checkbox_element = tabela.find_element(By.NAME, codigo)

                    # Clique no checkbox
                    checkbox_element.click()
                    # Alternar o estado do checkbox
                    
                except Exception as e:
                    print(f"Erro ao clicar no checkbox com código {codigo}: {str(e)}")

            if  len(codigo_procedimentos_mulher_grupo) <= 0:
                pass
            else:
                extracao_mulher()


            def extracao_mulher():

                data2 = {'CODIGO': codigo_procedimentos_mulher_grupo} 
                df_mulher = pd.DataFrame(data2)
                total_2 = df_mulher.shape[0]

                    

                driver.get("https://sisregiii.saude.gov.br/cgi-bin/cadweb50?url=/cgi-bin/marcar")

                for index, row in df.iterrows():
                    CampoCnes = driver.find_element(By.NAME,'nu_cns')
                    cnes_m = (str(row['cnes feminino']))
                    CampoCnes.send_keys(cnes_m)  #Cnes feminino 


                driver.find_element(By.NAME,'btn_pesquisar').click()

                sleep(1)

                driver.find_element(By.NAME,'btn_continuar').click()

                sleep(1)
                CampoCid = driver.find_element(By.NAME,'cid10')
                CampoCid.send_keys('z000')  #cid10

                SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
                SelectMedico.select_by_value('00000000000') #Profissional não listado
                sleep(1)
                CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
                CampoMedico.send_keys('Medico')  #Nome da Doutora
                sleep(1)
                select_procedimento = Select(driver.find_element(By.NAME,"pa"))
                select_procedimento.select_by_value(variavel_codigo_titulo)
            
                driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center/input').click() # Procura elemento e clica
                    
                for codigo in df_mulher['CODIGO']:
                    # Localize a tabela pelo seletor CSS
                    tabela = driver.find_element(By.CLASS_NAME, 'table_listagem')
                    # Dentro da tabela, localize o elemento pelo atributo "name" igual ao código
                    checkbox_element = driver.find_element(By.NAME, codigo)

                    # Clique no checkbox
                    checkbox_element.click()

                    driver.find_element(By.NAME,'btnConfirmar').click() # Procura elemento e clica
                            
                    try: # Se aparecer o botão btnSolicitar então recebe 0 caso contrario recebe Vagas Encontradas
                        resultado = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr/td/center/b').text
                        if resultado == 'PROCEDIMENTO REGULADO':
                            pass
                        elif resultado == 'VAGAS DISPONÍVEIS':
                            nome_mulher = driver.find_element(By.XPATH,'//*[@id="main_div"]/center/table/tbody/tr[3]/td[2]/i').text
                            nome_procedimentos_mulher_grupo.append(nome_mulher)
                            Tem_vaga = "Vagas Encontradas"
                            vagas_mulher_grupo.append(Tem_vaga) # Adiciona o valor na lista
                        elif resultado == 'NENHUMA VAGA ENCONTRADA':
                            pass
                    except:
                        pass
                                
                    sleep(1)    
                    driver.find_element(By.NAME,'btnVoltar').click() # Procura elemento voltar e clica
                    sleep(1)
                    # Localize a tabela pelo seletor CSS
                    tabela = driver.find_element(By.CLASS_NAME, 'table_listagem')
                    # Dentro da tabela, localize o elemento pelo atributo "name" igual ao código
                    checkbox_element = tabela.find_element(By.NAME, codigo)

                    # Clique no checkbox
                    checkbox_element.click()
                    # Alternar o estado do checkbox
                
            if  len(codigo_procedimentos_crianca_grupo) <= 0:
                pass
            else:
                extracao_crianca()    

            def extracao_crianca():

                data3 = {'CODIGO': codigo_procedimentos_crianca_grupo} 
                df_crianca = pd.DataFrame(data3)
                total_3 = df_crianca.shape[0]

                    

                driver.get("https://sisregiii.saude.gov.br/cgi-bin/cadweb50?url=/cgi-bin/marcar")

                for index, row in df.iterrows():
                    CampoCnes = driver.find_element(By.NAME,'nu_cns')
                    cnes_c = (str(row['cnes crianca']))
                    CampoCnes.send_keys(cnes_c)  #Cnes criança


                driver.find_element(By.NAME,'btn_pesquisar').click()

                sleep(1)

                driver.find_element(By.NAME,'btn_continuar').click()

                sleep(1)
                CampoCid = driver.find_element(By.NAME,'cid10')
                CampoCid.send_keys('z000')  #cid10

                SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
                SelectMedico.select_by_value('00000000000') #Profissional não listado
                sleep(1)
                CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
                CampoMedico.send_keys('Medico')  #Nome da Doutora
                sleep(1)
                select_procedimento = Select(driver.find_element(By.NAME,"pa"))
                select_procedimento.select_by_value(variavel_codigo_titulo)
            
                driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center/input').click() # Procura elemento e clica
                    
                for codigo in df_crianca['CODIGO']:
                    # Localize a tabela pelo seletor CSS
                    tabela = driver.find_element(By.CLASS_NAME, 'table_listagem')
                    # Dentro da tabela, localize o elemento pelo atributo "name" igual ao código
                    checkbox_element = driver.find_element(By.NAME, codigo)

                    # Clique no checkbox
                    checkbox_element.click()

                    driver.find_element(By.NAME,'btnConfirmar').click() # Procura elemento e clica
                            
                    sleep(1)
                    try: # Se aparecer o botão btnSolicitar então recebe 0 caso contrario recebe Vagas Encontradas
                        resultado = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr/td/center/b').text
                        if resultado == 'PROCEDIMENTO REGULADO':
                            pass
                        elif resultado == 'VAGAS DISPONÍVEIS':
                            nome_crianca = driver.find_element(By.XPATH,'//*[@id="main_div"]/center/table/tbody/tr[3]/td[2]/i').text
                            nome_procedimentos_crianca_grupo.append(nome_crianca)
                            Tem_vaga = "Vagas Encontradas"
                            vagas_crianca_grupo.append(Tem_vaga) # Adiciona o valor na lista
                        elif resultado == 'NENHUMA VAGA ENCONTRADA':
                            pass
                    except:
                        pass

                    sleep(1)  
                    driver.find_element(By.NAME,'btnVoltar').click() # Procura elemento voltar e clica
                    sleep(1)
                    # Localize a tabela pelo seletor CSS
                    tabela = driver.find_element(By.CLASS_NAME, 'table_listagem')
                    # Dentro da tabela, localize o elemento pelo atributo "name" igual ao código
                    checkbox_element = tabela.find_element(By.NAME, codigo)

                    # Clique no checkbox
                    checkbox_element.click()
                    # Alternar o estado do checkbox

            if  len(codigo_procedimentos_idoso_grupo) <= 0:
                pass
            else:
                extracao_idoso() 

            def extracao_idoso():

                data4 = {'CODIGO': codigo_procedimentos_idoso_grupo} 
                df_idoso = pd.DataFrame(data4)
                total_4 = df_idoso.shape[0]

                    

                driver.get("https://sisregiii.saude.gov.br/cgi-bin/cadweb50?url=/cgi-bin/marcar")

                for index, row in df.iterrows():
                    CampoCnes = driver.find_element(By.NAME,'nu_cns')
                    cnes_i = (str(row['cnes idoso']))
                    CampoCnes.send_keys(cnes_i)  #Cnes idoso 


                driver.find_element(By.NAME,'btn_pesquisar').click()

                sleep(1)

                driver.find_element(By.NAME,'btn_continuar').click()

                sleep(1)
                CampoCid = driver.find_element(By.NAME,'cid10')
                CampoCid.send_keys('z000')  #cid10

                SelectMedico = Select(driver.find_element(By.NAME,'cpfprofsol'))
                SelectMedico.select_by_value('00000000000') #Profissional não listado
                sleep(1)
                CampoMedico = driver.find_element(By.NAME,'nomeprofsol')
                CampoMedico.send_keys('Medico')  #Nome da Doutora

                sleep(1)
                select_procedimento = Select(driver.find_element(By.NAME,"pa"))
                select_procedimento.select_by_value(variavel_codigo_titulo)
            
                driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center/input').click() # Procura elemento e clica

                    
                for codigo in df_idoso['CODIGO']:
                    # Localize a tabela pelo seletor CSS
                    tabela = driver.find_element(By.CLASS_NAME, 'table_listagem')
                    # Dentro da tabela, localize o elemento pelo atributo "name" igual ao código
                    checkbox_element = driver.find_element(By.NAME, codigo)

                    # Clique no checkbox
                    checkbox_element.click()

                    driver.find_element(By.NAME,'btnConfirmar').click() # Procura elemento e clica
                    sleep(1)
                    try: # Se aparecer o botão btnSolicitar então recebe 0 caso contrario recebe Vagas Encontradas
                        resultado = driver.find_element(By.XPATH,'//*[@id="main_div"]/form/center[2]/table/tbody/tr/td/center/b').text
                        if resultado == 'PROCEDIMENTO REGULADO':
                            pass
                        elif resultado == 'VAGAS DISPONÍVEIS':
                            nome_idoso = driver.find_element(By.XPATH,'//*[@id="main_div"]/center/table/tbody/tr[3]/td[2]/i').text
                            nome_procedimentos_idoso_grupo.append(nome_idoso)
                            Tem_vaga = "Vagas Encontradas"
                            vagas_idoso_grupo.append(Tem_vaga) # Adiciona o valor na lista
                        elif resultado == 'NENHUMA VAGA ENCONTRADA':
                            pass
                    except:
                        pass

                    sleep(1)   
                    driver.find_element(By.NAME,'btnVoltar').click() # Procura elemento voltar e clica
                    sleep(1)
                    # Localize a tabela pelo seletor CSS
                    tabela = driver.find_element(By.CLASS_NAME, 'table_listagem')
                    # Dentro da tabela, localize o elemento pelo atributo "name" igual ao código
                    checkbox_element = tabela.find_element(By.NAME, codigo)

                    # Clique no checkbox
                    checkbox_element.click()
                    # Alternar o estado do checkbox
                
        
            
            procedimentos_concateandos = []
            vagas_concatenadas = []


            for nome_todos, vaga_todos in zip(nome_procedimento_geral_grupo + nome_procedimentos_mulher_grupo + nome_procedimentos_crianca_grupo + nome_procedimentos_idoso_grupo, vagas_geral_grupo + vagas_mulher_grupo + vagas_crianca_grupo + vagas_idoso_grupo):
                procedimentos_concateandos.append(nome_todos)
                vagas_concatenadas.append(vaga_todos)

            procedimentos_concateandos.sort()

            data5 = {'Procedimentos': procedimentos_concateandos, 'Vagas': vagas_concatenadas}
            global df_final_geral_grupo
            df_final_geral_grupo = pd.DataFrame(data5)
            # Agende a chamada para exibir_dataframe() na thread principal
            janela_grupo.after(0, exibir_dataframe_grupo)

        
        def start_extraction():
            progress_thread = threading.Thread(target=extraindo_vagas_grupo)
            progress_thread.start()

                    
        def selecionar_todos():
            for col_idx, coluna in enumerate(colunas):
                for row_idx, procedimento in enumerate(coluna):
                    procedimento_vars_por_coluna[col_idx][row_idx].set(1)

        def desmarcar_todos():
            for col_idx, coluna in enumerate(colunas):
                for row_idx, procedimento in enumerate(coluna):
                    procedimento_vars_por_coluna[col_idx][row_idx].set(0)

        
        
        
        def fechar_janela_grupo():
            # Destrua a janela secundária
            janela_grupo.destroy()
            
        # Crie uma nova janela para mostrar os Checkbuttons
        janela_grupo = tk.Toplevel()
        janela_grupo.title(nome_procedimento)

        # Impede que o usuário feche a nova tela antes de voltar
        janela_grupo.grab_set()

        # Obtenha as dimensões da tela
        largura_tela = 1024
        altura_tela = 768

        # Altura da barra de tarefas (assumindo que a barra de tarefas esteja na parte inferior)
        altura_barra_tarefas = 100  # Ajuste conforme necessário

        # Calcule a altura da janela para que ela não ultrapasse a barra de tarefas
        altura_janela = altura_tela - altura_barra_tarefas

        # Defina a geometria da janela
        janela_grupo.geometry(f"{largura_tela}x{altura_janela}")

            
        

        # Cria um frame para os botões e o posiciona na parte superior da janela usando o gerenciador de geometria "grid"
        button_frame = tk.Frame(janela_grupo)
        button_frame.grid(column=0, row=0, columnspan=6, padx= 10, pady=20)

        # Cria um botão com o texto "Extrair vagas novamente" e vincula a função extrair_vagas a ele
        botao_extrair_vaga_novamente = tk.Button(button_frame, text="Atualizar Lista", command=abrir_extrair_procedimentos)
        botao_extrair_vaga_novamente.grid(column=0, row=0, padx=10)

        # Botão "Pesquisar Vagas"
        botao_pesquisar = tk.Button(button_frame, text="Pesquisar Vagas", command=verificar_selecao)
        botao_pesquisar.grid(column=1, row=0, padx=10)

        
        # Botão "Salvar Selecionados"
        botao_salvar = tk.Button(button_frame, text="Salvar Selecionados", command=salvar_selecionados)
        botao_salvar.grid(column=2, row=0, padx=10)

        # Botão "Carregar Selecionados"
        botao_carregar = tk.Button(button_frame, text="Carregar Selecionados", command=carregar_selecionados)
        botao_carregar.grid(column=0, row=1, padx=10)

        # Botão "Selecionar Todos"
        botao_selecionar_todos = tk.Button(button_frame, text="Selecionar Todos", command=selecionar_todos)
        botao_selecionar_todos.grid(column=1, row=1, padx=10)

        # Botão "Desmarcar Todos"
        botao_desmarcar_todos = tk.Button(button_frame, text="Desmarcar Todos", command=desmarcar_todos)
        botao_desmarcar_todos.grid(column=2, row=1, padx=10)

        # Configure um botão "Fechar" na nova tela
        botao_fechar = tk.Button(button_frame, text="Voltar", command=fechar_janela_grupo)
        botao_fechar.grid(column=3, row=1, padx=10)

        # Cria uma lista de listas para armazenar as variáveis dos procedimentos em cada coluna
        procedimento_vars_por_coluna = []

        # Divide os procedimentos em colunas de no máximo 
        num_colunas = 500
        colunas = [procedimento_extraidos_grupo[i:i + num_colunas] for i in range(0, len(procedimento_extraidos_grupo), num_colunas)]

        # Inicializa as variáveis para cada coluna
        for coluna in colunas:
            procedimento_vars = []
            for _ in coluna:
                var = tk.IntVar()
                procedimento_vars.append(var)
            procedimento_vars_por_coluna.append(procedimento_vars)

                
        
        checkbox_frame = tk.Frame(janela_grupo)
        checkbox_frame.grid(column=0, row=1, columnspan=1, padx=0, pady=0)
        
        
        # Crie um Canvas para os Checkbuttons
        canvas = Canvas(checkbox_frame, width=600, height=600)  # Defina a largura e altura desejadas
        canvas.grid(row=0, column=0, padx=0, pady=0, sticky="nsew")

        # Configure o tamanho máximo do Canvas com base na altura disponível
        canvas_max_height = altura_janela - 150  # Ajuste conforme necessário
        canvas.config(height=canvas_max_height)

        # Crie uma barra de rolagem vertical para o Canvas
        scrollbar_y = Scrollbar(checkbox_frame, orient="vertical", command=canvas.yview)
        scrollbar_y.grid(row=0, column=1, padx=0, pady=0, sticky="ns")

        # Crie uma barra de rolagem horizontal para o Canvas
        scrollbar_x = Scrollbar(checkbox_frame, orient="horizontal", command=canvas.xview)
        scrollbar_x.grid(row=1, column=0, padx=0, pady=0, sticky="ew")

        # Configure as barras de rolagem para controlar o Canvas
        canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        # Crie um novo frame dentro do Canvas para os Checkbuttons
        checkbox_canvas_frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=checkbox_canvas_frame, anchor="nw")

        checkbuttons = []  # Lista para armazenar os Checkbuttons

        
        def criar_checkbuttons():
            for col_idx, coluna in enumerate(colunas):
                for row_idx, procedimento in enumerate(coluna):
                    if not ('grupo' in procedimento.lower() or procedimento.count('x') >= 3):
                        # Crie o Checkbutton apenas se o procedimento não atender aos critérios
                        checkbox = tk.Checkbutton(checkbox_canvas_frame, text=procedimento, variable=procedimento_vars_por_coluna[col_idx][row_idx], anchor="w", padx=0, pady=0)
                        checkbox.grid(row=row_idx, column=col_idx, padx=0, pady=0, sticky="w")
                        checkbuttons.append(checkbox)  # Adicione o Checkbutton à lista


        # Chame a função criar_checkbuttons aqui
        criar_checkbuttons()
        
        # Atualize a visualização do Canvas e configure as barras de rolagem para rolar o Canvas
        checkbox_canvas_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        
    
        # Atualize a função pesquisar_checkbuttons para destacar em tempo real
        # Mude a cor da pesquisa para vermelho
        def pesquisar_checkbuttons(*args):
            termo_pesquisa = caixa_pesquisa.get().lower()  # Obtém o termo de pesquisa em minúsculas
            for checkbox in checkbuttons:
                texto_checkbox = checkbox.cget("text").lower()  # Obtém o texto do Checkbutton em minúsculas
                if termo_pesquisa == "" or termo_pesquisa in texto_checkbox:
                    checkbox.grid()  # Exibe o Checkbutton se corresponder à pesquisa
                else:
                    checkbox.grid_remove()  # Oculta o Checkbutton se não corresponder à pesquisa

        def limpar_pesquisa():
            caixa_pesquisa.delete(0, tk.END)  # Limpa o conteúdo da caixa de pesquisa
            pesquisar_checkbuttons()  # Mostra todos os procedimentos novamente
        
        
        
        
        # Crie um frame para a pesquisa
        pesquisa_frame = tk.Frame(checkbox_frame)
        pesquisa_frame.grid(row=0, column=2, padx=10, pady=10, sticky="n")

        # Crie um Label para o texto "Pesquise o procedimento"
        label_pesquisa = tk.Label(pesquisa_frame, text="Pesquise o procedimento", fg="red")
        label_pesquisa.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # Crie uma Entry para a caixa de pesquisa
        caixa_pesquisa = tk.Entry(pesquisa_frame)
        caixa_pesquisa.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        # Vincule a função pesquisar_checkbuttons ao evento de mudança de texto na caixa de pesquisa
        caixa_pesquisa.bind("<KeyRelease>", pesquisar_checkbuttons)

        # Crie um botão "Limpar" para redefinir a pesquisa
        botao_limpar_pesquisa = tk.Button(pesquisa_frame, text="Limpar", command=limpar_pesquisa)
        botao_limpar_pesquisa.grid(row=1, column=1, padx=5, pady=5)
        
        janela_grupo.mainloop()

    def exibir_dataframe_grupo():
            # Crie uma janela secundária para exibir o DataFrame
            janela_dataframe = tk.Toplevel()
            janela_dataframe.title("DataFrame")
            janela_dataframe.grab_set()
            #janela_grupo.withdraw

            # Crie um widget Text para exibir o DataFrame
            text_widget = tk.Text(janela_dataframe, wrap=tk.NONE)
            text_widget.pack(fill=tk.BOTH, expand=True)

            # Converte o DataFrame em uma string formatada
            dataframe_str = df_final_geral_grupo.to_string(index=False)  # Altere df_final_geral para o seu DataFrame real

            # Insira o DataFrame formatado no widget Text
            text_widget.insert(tk.END, dataframe_str)

            # Impede que o usuário edite o widget Text (somente leitura)
            text_widget.config(state=tk.DISABLED)
            janela_dataframe.wait_window(janela_dataframe)

    def exibir_dataframe():
            # Crie uma janela secundária para exibir o DataFrame
            janela_dataframe = tk.Toplevel()
            janela_dataframe.title("DataFrame")

            # Crie um widget Text para exibir o DataFrame
            text_widget = tk.Text(janela_dataframe, wrap=tk.NONE)
            text_widget.pack(fill=tk.BOTH, expand=True)

            # Converte o DataFrame em uma string formatada
            dataframe_str = df_final_geral.to_string(index=False)  # Altere df_final_geral para o seu DataFrame real

            # Insira o DataFrame formatado no widget Text
            text_widget.insert(tk.END, dataframe_str)

            # Impede que o usuário edite o widget Text (somente leitura)
            text_widget.config(state=tk.DISABLED)
            janela_dataframe.mainloop()
        

    carregar_site()
if __name__ == "__main__":
    main()


