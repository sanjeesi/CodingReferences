import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


s = smtplib.SMTP(".com")
s.set_debuglevel(1)
msg = MIMEText(""" This is Test email for load testing 
Testing mail
""")
sender = 'email@pldt.com.ph'
recipients = ['email@amdocs.com']
#msg['Subject'] = "rpa_Testing_phase - Round 1"
msg['From'] = sender
msg['To'] = ", ".join(recipients)

# bodyCounter = 0

for subjectCounter in range(1,2):
    msg['Subject'] = "Testing_phase - round {}".format(subjectCounter)
    s.sendmail(sender, recipients, msg.as_string())

#s.send_message(sender, recipients, msg.as_string())
