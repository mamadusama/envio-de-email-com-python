import win32com.client as win32

# Create a new Outlook session
outlook = win32.Dispatch('Outlook.Application')
#linkimg = "https://www.google.com/url?sa=i&url=https%3A%2F%2Fportalhashtag.com%2Flogin&psig=AOvVaw3n394OD4-EBUK-iz6davs9&ust=1732228534583000&source=images&cd=vfe&opi=89978449&ved=0CBQQjRxqFwoTCJi3m6b864kDFQAAAAAdAAAAABAE"

# Create a new mail item
mail = outlook.CreateItem(0)
# Set the recipient's email address
mail.To = 'mamadusama19@gmail.com'
# Set the subject of the email
mail.Subject = 'Email enviado com Python'
#set html body
mail.HTMLBody = """
<p>Este é um email enviado com Python</p>
 <p>Att., Mamadu Sama</p>
 <img src='https://d1muf25xaso8hp.cloudfront.net/https%3A%2F%2Fa6d41686876ceccfc436dd310b9e49aa.cdn.bubble.io%2Ff1658516625802x148010885188176500%2Flogo%2520hash%2520oficial%2520-%2520letra%2520azul.png?w=&h=&auto=compress&dpr=1&fit=max'>
 """

#enviando
mail.Send()