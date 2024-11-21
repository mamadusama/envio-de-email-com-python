import win32com.client as win32

# Create a new Outlook session
outlook = win32.Dispatch('Outlook.Application')

# Create a new mail item

mail = outlook.CreateItem(0)

# Set the recipient's email address
mail.To = 'mamadusama19@gmail.com'
# Set the subject lin
mail.Subject = 'Test Email'
# email Cc
mail.Cc = 'mamadusama19@gmail.com'
# Set the body of the email
#mail.Body = 'This is a test email.'

# email html body
mail.HTMLBody = '''<html>
<body>
 <h1>This is a test email.</h1>
 <p>This is a test email.</p>
 <p>Att., Mamadu Sama</p>

</body>

</html>'''

#send
mail.Send()

#outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")