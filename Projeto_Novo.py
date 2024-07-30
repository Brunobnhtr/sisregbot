import json
import time
import os
import re
import pandas as pd
from datetime import date, timedelta
import sys

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ChromeOptions
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc



class SisregBot:
    def __init__(self):
        os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
        self.driver = uc.Chrome(headless=False, use_subprocess=False)
    
    def acessar_site_login(self):
        self.driver.get('https://sisregiii.saude.gov.br/')
        
        while True:
            try:
                campo_user = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.ID, 'usuario')))
                
            except:
                self.driver.refresh()
            else:
                usuario = input("Digite o usuário: ")
                campo_user.send_keys(usuario)
                
                try:
                    campo_senha = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.ID, 'senha')))
                
                except:
                    self.driver.refresh()
                else:
                    senha = input("Digite a senha: ")
                    campo_senha.send_keys(senha)
                        
                    try:
                        botao_entrar = WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.NAME, 'entrar')))
                    except:
                        print('Botão entrar não encontrado!')
                        self.driver.refresh()
                    else:
                        botao_entrar.click()

                        try:
                            WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.ID, 'mensagem')))
                            print("Usuário ou senha incorretos")
                            self.driver.refresh()
                        except:
                            self.extrair_procedimentos()

    def extrair_procedimentos(self):
        self.driver.get("https://sisregiii.saude.gov.br/cgi-bin/cadweb50?url=/cgi-bin/marcar")
        

if __name__ == '__main__':
    inicio = SisregBot()
    inicio.acessar_site_login()






