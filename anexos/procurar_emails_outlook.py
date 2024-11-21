import win32com.client as win32

# Create a new Outlook session
outlook = win32.Dispatch('Outlook.Application')

# criar variavel para caixa de entrada
caixa_email= outlook.GetNamespace('MAPI') #.GetDefaultFolder(6)

#for pasta in caixa_email.Folders:
  #print(pasta.Name) # imprime o nome da pasta

pasta_mamdu1 = caixa_email.Folders.Item(2)

#print(f"a minha pasta 2 é : {pasta_mamdu1}")

#for subpasta in pasta_mamdu1.Folders:
  #print(subpasta) # imprime o nome da subpasta



# Selecionar a pasta desejada (1 = Caixa de Entrada padrão)
caixa_de_entrada = pasta_mamdu1.Folders.Item(1) # 1 = Caixa de Entrada,
caixa_de_saida = pasta_mamdu1.Folders.Item(2) # 2 = Caixa de Saída,(A enviar)
print(caixa_de_entrada)
print(caixa_de_saida)


lista_email = caixa_de_entrada.Items
#print(lista_email.Count) # imprime o número de e-mails na pasta

for email in lista_email:
    if email.To == 'msamasama.pt@gmail.com':
        print(f"Assunto: {email.Subject}")
        print(f"De: {email.SenderName}")
        print(f"emaideremetente: {email.SenderEmailAddress}")
        print(f"Para: {email.To}")
        print(f"Data: {email.ReceivedTime}")
        print(f"Corpo: {email.Body[:100]}")  # Mostrar os primeiros 100 caracteres do corpo

print("fim de codigo")