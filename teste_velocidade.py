import logging
import logging.config
from selenium import webdriver
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import pandas as pd
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.message import EmailMessage
import os

logging_config = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'standard': {
            'format': '%(asctime)s [%(levelname)s] %(message)s',
        },
    },
    'handlers': {
        'console': {
            'level': 'INFO',
            'class': 'logging.StreamHandler',
            'formatter': 'standard',
        },
        'file': {
            'level': 'INFO',
            'class': 'logging.FileHandler',
            'formatter': 'standard',
            'filename': 'internet_speed_test.log',
            'mode': 'a',
        },
    },
    'root': {
        'level': 'INFO',
        'handlers': ['console', 'file'],
    },
}

logging.config.dictConfig(logging_config)
logger = logging.getLogger()


chrome_options = Options()
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--ignore-ssl-errors')
chrome_options.add_argument('log-level=4')
driver = webdriver.Chrome(options=chrome_options)

def realizar_teste():
    logging.info('Realizando teste de velocidade (aguarde)...')
    driver.get('https://www.speedtest.net/')

    # Clicar no botão de iniciar teste
    start_button = driver.find_element(By.XPATH, '//*[@id="container"]/div/div[3]/div/div/div/div[2]/div[3]/div[1]/a/span[4]')
    start_button.click()

    # Aguarda que o teste seja concluído
    time.sleep(40)

    # Extrai o resultado e armazena em variáveis
    velocidade_download = driver.find_element(By.XPATH, '//*[@id="container"]/div/div[3]/div/div/div/div[2]/div[3]/div[3]/div/div[3]/div/div/div[2]/div[1]/div[1]/div/div[2]/span').text
    velocidade_upload = driver.find_element(By.XPATH, '//*[@id="container"]/div/div[3]/div/div/div/div[2]/div[3]/div[3]/div/div[3]/div/div/div[2]/div[1]/div[2]/div/div[2]/span').text
    ping = driver.find_element(By.XPATH, '//*[@id="container"]/div/div[3]/div/div/div/div[2]/div[3]/div[3]/div/div[3]/div/div/div[2]/div[2]/div/span[3]/span').text

    return velocidade_download, velocidade_upload, ping

def salvar_resultado(download_speed, upload_speed, ping):
    
    logging.info('Salvando o resultado do teste na planilha...')

    data = {
        'Data e Hora': [datetime.datetime.now()],
        'Download (Mbps)': [download_speed],
        'Upload (Mbps)': [upload_speed],
        'Ping (ms)': [ping]
    }

    df = pd.DataFrame(data)
    
    # Adiciona os dados ao CSV, se o arquivo já existir
    df.to_csv('resultado_teste_velocidade.csv', mode='a', header=not pd.io.common.file_exists('resultado_teste_velocidade.csv'), index=False)

def enviar_email():
    logging.info('Enviando o resultado para o seu email...')

    com_anexo = './resultado_teste_velocidade.csv'
    email_de = "andrepinto.cg@gmail.com"
    email_destino = "andrepinto.cg@gmail.com"
    senha = 'jkru nswj iqbu uwsm'

    # Configurar a mensagem
    mensagem = MIMEMultipart()
    mensagem['From'] = email_de
    mensagem['To'] = "andrepinto.cg@gmail.com"
    mensagem['Subject'] = "Resultados dos Testes de Velocidade"

     # Corpo do e-mail
    corpo = "Segue em anexo a planilha com os resultados dos testes de velocidade da internet."
    mensagem.attach(MIMEText(corpo, 'plain'))

    # Anexo
    anexo = open(com_anexo, "rb")

    # Configurar o anexo
    parte = MIMEBase('application', 'octet-stream')
    parte.set_payload(anexo.read())
    encoders.encode_base64(parte)
    parte.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(com_anexo)}')

    mensagem.attach(parte)

     # Configuração do servidor
    servidor = smtplib.SMTP('smtp.gmail.com', 587)
    servidor.starttls()

    # Login no servidor
    servidor.login(email_de, senha)

    # Enviar o e-mail
    texto = mensagem.as_string()
    servidor.sendmail(email_de, email_destino, texto)

    # Encerrar a sessão
    servidor.quit()
    
def main():
    try:
        logging.info('Acessando o site...')
        download_speed, upload_speed, ping = realizar_teste()
        salvar_resultado(download_speed, upload_speed, ping)
        enviar_email()
    finally:
        logging.info('Email enviado.')
        driver.quit()

if __name__ == '__main__':
    main()