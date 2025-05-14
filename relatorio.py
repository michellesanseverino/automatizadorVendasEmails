import pandas as pd
import os 
import smtplib
import json

from dotenv import load_dotenv
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from datetime import datetime

#Para iniciar, carregaremos as variáveis do .env
load_dotenv()
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE")
EMAIL_SENHA = os.getenv("EMAIL_SENHA")

#Função para limpar a lista de emails, retirando duplicatas, e salvando apenas a primeira aparição
def salvar_emails(caminho: str) -> dict:
    listaEmailsExcel = "emails_vendedores.xlsx"
    df1 = pd.read_excel(listaEmailsExcel)
    df1.drop_duplicates(subset='Vendedor', keep='first', inplace=True)
    email_dict = df1.set_index('Vendedor')['Email'].to_dict()

    with open ('dicionario_email.json', 'w') as f:
        json.dump(email_dict, f, indent=4)
    return email_dict

#Abrindo o arquivo xlsx
def ler_arquivo(caminho: str) -> pd.DataFrame:
    return pd.read_excel(caminho)

#Limpando os dados dentro do Excel
def limpar_dados(df: pd.DataFrame) -> pd.DataFrame:
    df["Data"] = pd.to_datetime(df["Data"])
    df["Valor Total"] = df["Quantidade"] * df["Preço Unitário"]
    return df.dropna()

#Agrupando os dados por Vendedor
def agrupar_dados(df: pd.DataFrame) -> dict:
    grupos = dict(tuple(df.groupby("Vendedor")))
    return grupos

#Função para gerar os relatórios de vendas de cada vendedor
def gerar_relatorio(df_vendedor: pd.DataFrame, vendedor: str) -> str:
    nome_arquivo = f"relatorios/relatorio_{vendedor.replace(' ', '_')}.xlsx"
    df_vendedor.to_excel(nome_arquivo, index=False)
    return nome_arquivo

#Função para criação do email
def enviar_email(vendedor: str, email_destino: str, anexo_caminho: str):
    msg = MIMEMultipart()
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = email_destino
    msg["Subject"] = f"Relatório de vendas - {vendedor}"
    
    corpoEmail = f"""
    <p>Olá {vendedor},</p>
    <p>Segue em anexo o relatório de vendas atualizado.</p> 
    <p>Atenciosamente, <br>Equipe de Vendas</p>
    """
    
    msg.attach(MIMEText(corpoEmail, "html"))
    
    with open(anexo_caminho, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(anexo_caminho))
        part["Content-Disposition"] = f'attachment; filename="{os.path.basename(anexo_caminho)}"'
    
    #Aqui é realizada a autenticação no servidor da google para que possa enviar os emails    
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_REMETENTE, EMAIL_SENHA)
        smtp.send_message(msg)
        print(f"Email enviado com sucesso para {vendedor} ({email_destino}) ")
    

#Rodar o script 
if __name__ == "__main__":
    os.makedirs("relatorios", exist_ok=True)
    
    emailsList = salvar_emails("emails_vendedores.xlsx") 
    df = ler_arquivo("vendas.xlsx")
    df = limpar_dados(df)
    grupos = agrupar_dados(df)
    
    for vendedor, df_vendedor in grupos.items():
        if vendedor not in emailsList:
            print(f"[AVISO] Vendedor não tem e-mail cadastrado no sistema")
            continue
        
        arquivo = gerar_relatorio(df_vendedor, vendedor)
        enviar_email(vendedor, emailsList[vendedor], arquivo)