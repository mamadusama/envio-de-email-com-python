import win32com.client as win32

outlook = win32.Dispatch('Outlook.Application')

print("Contas configuradas no Outlook:")
for account in outlook.Session.Accounts:
    print(account.SmtpAddress)
