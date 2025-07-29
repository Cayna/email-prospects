import pandas as pd
import win32com.client as win32

# Caminho da planilha
arquivo_excel = 'planilha-geral-internacional-hubspot.xlsx'

# Lê os dados da planilha
df = pd.read_excel(arquivo_excel)

# Inicia o Outlook
outlook = win32.Dispatch('outlook.application')

# Loop para enviar os e-mails
for index, linha in df.iterrows():
    nome = linha['Nome']
    email_destino = linha['Email']

    # Cria o e-mail
    mail = outlook.CreateItem(0)
    mail.To = email_destino
    mail.Subject = "A Strategic Auto Parts Partner from Brazil"
    corpo = f"""
    Hello {nome},
    I hope you're doing well.
    My name is Felipe, and I represent Gauss, a Brazilian manufacturer of automotive components. All of our products are manufactured in Brazil, and we're currently expanding our international reach through partnerships with eBay resellers.
    We offer over 50 product lines, including:
    Ignition Modules
    Voltage Regulators & Rectifiers
    Sensors, Coils, Relays
    LED Lighting Solutions
    As many sellers seek alternatives to Asian suppliers, partnering with a Latin American manufacturer can be a strategic and timely choice.
    If you're interested, I’d be happy to send our catalog, reseller pricing, and shipping details.
    Best regards,
    Felipe Caynã da Silva
    Gauss – Automotive Technology
    """
    mail.Body = corpo

    # Envia
    try:
        mail.Send()
        print(f"E-mail enviado para {nome} ({email_destino})")
    except Exception as e:
        print(f"Erro ao enviar para {email_destino}: {e}")
