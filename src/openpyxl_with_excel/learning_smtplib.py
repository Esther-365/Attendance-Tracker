import smtplib
from email.message import EmailMessage

msg = EmailMessage()
sender_email = input("Enter your email address: ")
receiver_email = input("Enter the recipient's email address: ")

msg['subject'] = "Learning smtplib in Python"
msg['From'] = sender_email
msg['To'] = receiver_email
body = "This is a test email sent using smtplib in Python."
msg.set_content(body)

server = smtplib.SMTP('smtp.gmail.com',587)
server.starttls()
server.login(sender_email, "ehwk dedh nhgf mqdb")
server.send_message(msg)

print("Email sent successfully!")

server.quit()


