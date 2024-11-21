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
    # Acessar a Caixa de Entrada da conta
    inbox = target_folder.Folders["Caixa de Entrada"]  # Ou "Inbox" se o idioma for inglês

    # Carregar as mensagens em uma lista estática
    messages = list(inbox.Items)

    # Limitar o processamento a 50 mensagens
    max_messages = 50
    print(f"Total de mensagens na Caixa de Entrada: {len(messages)}")

    for i, message in enumerate(messages):
        if i >= max_messages:
            break  # Parar ao atingir o limite
        try:
            print("---------------------------------------------------")
            print(f"Assunto: {message.Subject}")
            print(f"Remetente: {message.SenderName}")
            print(f"Email do Remetente: {message.SenderEmailAddress}")
            print(f"Data: {message.ReceivedTime}")
            print(f"Corpo: {message.Body[:100]}")  # Apenas os primeiros 100 caracteres
            print("---------------------------------------------------\n")
        except Exception as e:
            print(f"Erro ao processar mensagem {i + 1}: {e}")
