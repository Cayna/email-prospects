import pandas as pd
import win32com.client as win32


arquivo_excel = 'seuarquivo.xlsx'


df = pd.read_excel(arquivo_excel)


outlook = win32.Dispatch('outlook.application')


for index, linha in df.iterrows():
    nome = linha['Nome']
    email_destino = linha['Email']

   
    mail = outlook.CreateItem(0)
    mail.To = email_destino
    mail.Subject = "Assunto do E-mail"
    corpo = f"""
    Olá {nome},
    Conteúdo do E-mail
    """
    mail.Body = corpo

    # Envia
    try:
        mail.Send()
        print(f"E-mail enviado para {nome} ({email_destino})")
    except Exception as e:
        print(f"Erro ao enviar para {email_destino}: {e}")
