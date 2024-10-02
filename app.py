from selenium import webdriver
from unidecode import unidecode
from fuzzywuzzy import fuzz
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import re
import os
import sys


print('Olá bem vindo vamos iniciar nosso aplicativo')
print('Coloque o nome correto do evento!, pode colocar sem acento ou com!')
print('App criado por Dev Ramon Rodrigues')

nome_evento = input("Digite o nome do Evento: ")

driver = webdriver.Chrome()

url = 'https://mail.hostinger.com/'
driver.get(url)

wait = WebDriverWait(driver, 20)

email_input = wait.until(EC.presence_of_element_located((By.ID, 'rcmloginuser')))
email_input.send_keys('')

password_input = wait.until(EC.presence_of_element_located((By.ID, 'rcmloginpwd')))
password_input.send_keys('')

login_button = driver.find_element(By.ID, 'rcmloginsubmit')
login_button.click()

wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'messagelist')))

email_data = []

def verificar_evento(mensagem, nome_evento):
    mensagem_sem_acento = unidecode(mensagem.lower()) 
    nome_evento_sem_acento = unidecode(nome_evento.lower())


    similaridade = fuzz.partial_ratio(nome_evento_sem_acento, mensagem_sem_acento)
    return similaridade >= 80

def coletar_emails():
    while True:
        emails = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'tr.message')))
        print(f"Encontrados {len(emails)} emails nesta página.")

        for email in emails:
            try:
                email.click()
                    
                iframe = wait.until(EC.presence_of_element_located((By.ID, "messagecontframe")))
                driver.switch_to.frame(iframe)
                                
                wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/div[2]')))
                mensagem = driver.find_element(By.XPATH, "//div[contains(@class, 'message-part')]").text

                if verificar_evento(mensagem, nome_evento):
                    print(f"Mensagem corresponde ao evento: {nome_evento}")

                    nome = re.search(r'Nome:\s*(.+?)\s*$', mensagem, re.MULTILINE)
                    telefone = re.search(r'Telefone:\s*(.+?)\s*$', mensagem, re.MULTILINE)

                    if nome and telefone:
                        telefone_formatado = re.sub(r'[\(\)\-\s]', '', telefone.group(1)) 

                        if not telefone_formatado.startswith('21'):
                            telefone_formatado = '21' + telefone_formatado
                                
                        email_data.append({'Nome': nome.group(1), 'Telefone': telefone_formatado})
                        print(f"Nome: {nome.group(1)}, Telefone: {telefone.group(1)}")
                else:
                    print("Mensagem não corresponde ao evento.")

                driver.switch_to.default_content()
                driver.back()
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'messagelist')))

            except Exception as e:
                print(f"Erro ao processar o email: {e}")
        try:
            next_page = driver.find_element(By.XPATH, '//*[@id="rcmbtn119"]') 
            next_page.click()
            time.sleep(5)
        except Exception as e:
            print("Nenhuma próxima página encontrada. Coleta concluída.")
            break

coletar_emails()

def get_executable_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


executavel_path = get_executable_path()
file_path = os.path.join(executavel_path, 'cadastros.xlsx')

if email_data:
    df = pd.DataFrame(email_data)
    df.to_excel(file_path, index=False)
    print(f"Emails salvos com sucesso em '{file_path}'.")
else:
    print("Nenhum dado para salvar.")

driver.quit()
