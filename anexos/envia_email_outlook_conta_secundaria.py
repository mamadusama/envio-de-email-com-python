import win32com.client as win32

# Criar uma nova sessão do Outlook
outlook = win32.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace("MAPI")

# Identificar a conta que você deseja usar
contas = outlook.Session.Accounts

target_email = "msamasama.pt@gmail.com"  # A conta de envio desejada
selected_account = None

# Procurar pela conta específica
for conta in contas:
    if conta.SmtpAddress == target_email:
        selected_account = conta
        break

if not selected_account:
    print(f"Erro: A conta {target_email} não foi encontrada.")
else:
    # Criar um novo e-mail
    mail = outlook.CreateItem(0)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, selected_account))  # Vincular à conta secundária

    # Configurar os detalhes do e-mail
    mail.To = 'mamadusama19@gmail.com'  # Substitua pelo destinatário desejado
    mail.Subject = 'Email enviado com Python usando conta secundária'
    mail.HTMLBody = """
    <p>Este é um email enviado com Python usando uma conta secundária.</p>
    <p>Att., Mamadu Sama</p>
    <img src='https://d1muf25xaso8hp.cloudfront.net/https%3A%2F%2Fa6d41686876ceccfc436dd310b9e49aa.cdn.bubble.io%2Ff1658516625802x148010885188176500%2Flogo%2520hash%2520oficial%2520-%2520letra%2520azul.png?w=&h=&auto=compress&dpr=1&fit=max'>
    """

    # Enviar o e-mail
    mail.Send()

    print("E-mail enviado com sucesso pela conta secundária!")
