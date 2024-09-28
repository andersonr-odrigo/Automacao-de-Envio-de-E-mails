import win32com.client as win32
import pandas as pd
import os

# Pegando os anexos da pasta de anexos
pasta_anexos = "Anexos_E-mail_A"
arquivos = os.listdir(pasta_anexos)

# Leitura da Planilha
planilha = pd.read_excel("controle_homologação.xlsx", sheet_name="automailA")

# Coleta dos dados necessários para o envio do e-mail
destinatarios = planilha["Destinatários"]
assunto = planilha["Assunto"]
fornecedor = planilha["Fornecedores"]
documentos = planilha["Documentos"]

# Cria a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# configurar as informações do seu e-mail
for i, destinatario in enumerate(planilha["Destinatários"]):

    if destinatario == destinatario:

        # Cria um email
        email = outlook.CreateItem(0)

        # Adiciona os anexos
        for arquivo in arquivos:
            email.Attachments.Add(f"C:/Users/usuario.temporario/Desktop/Email/Anexos_E-mail_A/{arquivo}")

        # Preenche as informações do e-mail
        email.To = destinatario
        print(destinatario)
        email.Subject = f'{assunto.iloc[i]}'
        email.HTMLBody = f"""
        <p>Bom dia, equipe {fornecedor.iloc[i]}</p>
        
        <p>Como um grande parceiro convidamos a sua empresa para o processo de HOMOLOGAÇÃO DE FORNECEDORES TOMBINI. 
        Estamos em constante processo de adequações para atendimento às auditorias de certificação ISO e SASSMAQ e, para isso, precisamos que compartilhem conosco os documentos de sua empresa até o dia 25/09. São eles:
        </p>

        <p>{documentos.iloc[i]}</p><br> 
        """
        email.Send()

        print(f"Email enviado com sucesso para: {fornecedor.iloc[i]}")

    
    else:
        print(f"Destinatário não encontrado para o fornecedor: {fornecedor.iloc[i]}")
        with open('logs_email_A.txt', 'a') as arq:
            arq.write(f"Destinatário não encontrado para o fornecedor: {fornecedor.iloc[i]}\n")

print("Script Finalizado")