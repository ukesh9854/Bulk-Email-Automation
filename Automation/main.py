import pandas as p
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
#read the emails from file
data = p.read_excel("Automation sample.xlsx")
#print(type(data))

mail_col = data.get("Mail ")
list_of_mail = list(mail_col)
print(list_of_mail)

try:
    #object of smtp
    server = sm.SMTP("smtp.gmail.com", 587)
    server.starttls()
    #login
    server.login("assignmenthelpernepal@gmail.com", "assignment@@@")
    from_= "assignmenthelpernepal@gmail.com"
    to_= list_of_mail
    message=MIMEMultipart("alternative")
    message['Subject']="This is just testing message"
    message["from"]= "assignmenthelpernepal@gmail.com"
#create the text in the form of html
    html='''
    <html>
    <head>
    </head>
    
    <body>
        <h1>LearnCode</h1>
        <h1>You are great</h2>
        
        
        <button>Verify</button>
    </body>
    
    
    </html>
    
    
    '''

    text=MIMEText(html,"html")

    message.attach(text)

#send the message
    server.sendmail(from_,to_, message.as_string())
    print("message has been sent to the emails. ")

except Exception as e:
    print(e)