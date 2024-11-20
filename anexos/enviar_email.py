# SMTP
from senha import senha_app
import smtplib
import email.message

def enviar_email():
    
    msg = email.message.Message()
    msg["Subject"] = "Email enviado com Python"
    msg["From"] = "mamadusama19@gmail.com"
    msg["To"] = "msamasama.pt@gmail.com"
    msg["Cc"] = "msamasama.pt+copia@gmail.com" #;outroemailcopia@hotmail.com
    msg["Bcc"] = "msamasama.pt@gmail.com"
    
    link_imagem = "https://i.ibb.co/qdPNTvN/flask-django.jpg"

    corpo_email = f"""<p>Boa tarde,</p>
    <p>Esse é meu primeiro email com Python usando smtplib</p>
    <p>Att., Lira</p>
    <img src='{link_imagem}'>"""

    corpo_email = corpo_email.encode("utf-8")

    msg.add_header("Content-Type", "text/html")
    msg.set_payload(corpo_email)

    servidor = smtplib.SMTP("smtp.gmail.com", 587)
    servidor.starttls()
    servidor.login(msg["From"], senha_app)
    servidor.send_message(msg)
    servidor.quit()
    print("Email enviado")


enviar_email()