import os
import sys, traceback
#Conexão com banco
import cx_Oracle
#Trabalhar com arquivo excel
import pandas as pd
#Modulo define um objeto de sessão do cliente SMTP
import smtplib
#Biblioteca para enviar
import mimetypes
import email
import email.mime.application
#Biblioteca para anexar arquivo no e-mail
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


cx_Oracle.init_oracle_client(lib_dir=r"D:\\instantclient_12_2")
connection = cx_Oracle.connect('usr_ti/senha@ip/schema')

cursor = connection.cursor()

#Tratamento de Erro
def error(emailSQL):
    #Trata erro na consulta do banco
    if emailSQL == 'ConsultaBanco':
        msg = MIMEMultipart()
        message = """Bom dia Pessoal,\n\nOcorreu um possivel erro na geração da consulta.
        \nSegue a consulta:\n
        select
        T.USUARIO,
        T.NOME,
        T.SITUACAO,
        T.EMAIL,
        T.EMPRESA_ADP,
        T.DEPTO_ADP,
        T.GP_CONECTA,
        T.NM_GESTOR,
        T.DT_NASC
        from stg_cibe_prod.DW_STG_FP_AD t
        WHERE gp_conecta = 'S'"""

        #Informacoes do e-mail
        password = "Senha"
        msg['From'] = "matheus.burim@com.br"
        msg['To'] = "matheus.burim@com.br"
        msg['Subject'] = "[ERRO_CONSULTA] - Relação de usuarios no grupo ConectaE"
        
        #Dispara Email
        msg.attach(MIMEText(message, 'plain'))
        server = smtplib.SMTP('smtp.office365.com:porta')
        server.starttls()
        server.login(msg['From'], password)
        server.sendmail(msg['From'], msg['To'], msg.as_string())
        server.quit()
    #Trata erro de envio de e-mail
    elif emailSQL == 'ConsultaEmail':
        msg = MIMEMultipart()
        message = "Bom dia Pessoal,\n\nOcorreu um possivel erro ao enviar o E-mail."

        #Informacoes do e-mail
        password = "Senha"
        msg['From'] = "matheus.burim@com.br"
        msg['To'] = "matheus.burim@com.br"
        msg['Subject'] = "[ERRO_E-MAIL] - Relação de usuarios no grupo ConectaE"
        
        #Dispara Email
        msg.attach(MIMEText(message, 'plain'))
        server = smtplib.SMTP('smtp.office365.com:587')
        server.starttls()
        server.login(msg['From'], password)
        server.sendmail(msg['From'], msg['To'], msg.as_string())
        server.quit()
#Remove arquivo antigo
def removeArquivo():
    anexo = 'D:/Python/teste/ConectaE.xlsx'
    if os.path.exists(anexo) ==  True:
        os.remove(anexo)
        
#Consulta banco SQL Oracle
def consulta():
    removeArquivo()
    try:
        df = pd.read_sql("""
                        select
                        T.USUARIO,
                        T.NOME,
                        T.SITUACAO,
                        T.EMAIL,
                        T.EMPRESA_ADP,
                        T.DEPTO_ADP,
                        T.GP_CONECTA,
                        T.NM_GESTOR,
                        T.DT_NASC
                        from stg_cibe_prod.DW_STG_FP_AD t
                        WHERE gp_conecta = 'S'
                        """, connection)

        df.to_excel("D:/Python/teste/ConectaE.xlsx")
    except:
        errorBanco = 'ConsultaBanco'
        error(errorBanco)
consulta()

#Envio de E-mail
def envia_email():
    try:
        anexo = 'D:/Python/teste/ConectaE.xlsx'
        msg = MIMEMultipart()
        message = "Bom dia Pessoal,\n\nSegue relação dos usuario dentro do grupo ConectaE."

        #Informa
        password = "senha"
        msg['From'] = "matheus.burim@com.br"
        msg['To'] = "infra@com.br"
        msg['Subject'] = "Relação de usuarios no grupo ConectaE"

        #Tratamento do aquivo para anexar no corpo do e-mail
        msg_file = MIMEBase('application', 'octet-stream')
        msg_file.set_payload(open(anexo, 'rb').read())
        encoders.encode_base64(msg_file)
        msg_file.add_header('Content-Disposition', 'attachment', filename='ConectaE.xlsx')
        msg.attach(msg_file)
        
        #Dispara Email
        msg.attach(MIMEText(message, 'plain'))
        server = smtplib.SMTP('smtp.office365.com:porta')
        server.starttls()
        server.login(msg['From'], password)
        server.sendmail(msg['From'], msg['To'], msg.as_string())
        server.quit()
    except:
        errorEmail = 'ConsultaEmail'
        error(errorEmail)
envia_email()
