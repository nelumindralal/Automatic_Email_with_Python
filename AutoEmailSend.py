import pandas as pd
import smtplib

# change these as per use
your_email = "abc@gmail.com"
your_password = "abc123"

# establishing connection with gmail
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

# reading the spreadsheet
email_list = pd.read_excel('C:/Users/lenovo/Desktop/list.xlsx' , engine='openpyxl')

# getting the names and the emails
names = email_list['NAME']
emails = email_list['EMAIL']

# iterate through the records
for i in range(len(emails)):
    # for every record get the name and the email addresses
    name = names[i]
    email = emails[i]

    # the message to be emailed
    message = "Hello " + name + ". This is a test mail sent from Python by me  "

    # sending the email
    try:
        server.sendmail(your_email, [email], message)
        print('Email to {} successfully sent!\n\n'.format(email))
    except Exception as e:
        print('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))
# close the smtp server
server.close()