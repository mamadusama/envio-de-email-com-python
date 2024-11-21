import os
import win32com.client as win32

# Criar uma nova sessão do Outlook
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Nome ou endereço de e-mail da conta que você quer acessar
target_email = "msamasama.pt@gmail.com"

# Procurar a conta específica entre as contas configuradas no Outlook
target_folder = None
for account in outlook.Folders:
    if account.Name == target_email:
        target_folder = account
        break

if not target_folder:
    print(f"Erro: Conta {target_email} não encontrada.")
else:
    # Acessar a Caixa de Entrada
    inbox = target_folder.Folders["Caixa de Entrada"]  # Use "Inbox" se o idioma for inglês

    # Diretório onde os anexos serão salvos
    save_folder = "C:/Users/mamad/anexos_email"
    os.makedirs(save_folder, exist_ok=True)  # Criar a pasta se ela não existir

    # Limitar a quantidade de e-mails processados
    max_emails = 50
    count = 0

    for message in inbox.Items:
        if count >= max_emails:
            break
        try:
            if message.Attachments.Count > 0:
                print(f"Processando e-mail: {message.Subject}")
                for attachment in message.Attachments:
                    # Nome do anexo
                    attachment_name = attachment.FileName
                    # Caminho completo para salvar o anexo
                    attachment_path = os.path.join(save_folder, attachment_name)

                    # Salvar o anexo
                    attachment.SaveAsFile(attachment_path)
                    print(f"Anexo salvo: {attachment_path}")
            count += 1
        except Exception as e:
            print(f"Erro ao processar e-mail: {e}")
