# import necessary packages
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import xlrd

from string import Template

def read_file(filename):
    f = open(filename, 'r', encoding='utf-8')
    return f.read()

sender_email = "MIT@dulwich-beijing.cn"
reciever_email = "22IsaacL@dulwich-beijing.cn"

loc = ("MIT newsletter DB.xlsx")
workbook = xlrd.open_workbook(loc)
sheet = workbook.sheet_by_index(0)

s = smtplib.SMTP(host='smtp-mail.outlook.com', port=25) # Change port to 587 if not working
s.starttls()
password = input("Please enter password for the email " + sender_email + ": ")
s.login(sender_email, password)

message = read_file('TestTemplate.txt')

# Create the body of the message (a plain-text and an HTML version).
text = """\
If you're seeing this message, an error occured when trying to load the Newsletter.
Check that your current e-mail viewer supports HTML e-mails and reload the e-mail.
If the problem persists, please try using an alternative e-mail viewer.

Compiments of the the MIT crew."""

html = message.format(Name1=sheet.cell_value(1, 1), Summary1=sheet.cell_value(1, 2), Link1=sheet.cell_value(1, 3), Img1=sheet.cell_value(1, 4), Name2=sheet.cell_value(2, 1), Summary2=sheet.cell_value(2, 2), Link2=sheet.cell_value(2, 3), Img2=sheet.cell_value(2, 4), Name3=sheet.cell_value(3, 1), Summary3=sheet.cell_value(3, 2), Link3=sheet.cell_value(3, 3), Img3=sheet.cell_value(3, 4))

msg = MIMEMultipart("alternative")
# setup the parameters of the message
msg['From'] = sender_email
msg['To'] = reciever_email
msg['Subject'] = "I don't care about your .5"

# Turn these into plain/html MIMEText objects
part1 = MIMEText(text, "plain")
part2 = MIMEText(html, "html")

# Add HTML/plain-text parts to MIMEMultipart message
# The email client will try to render the last part first
msg.attach(part1)
msg.attach(part2)

s.send_message(msg)

del msg
