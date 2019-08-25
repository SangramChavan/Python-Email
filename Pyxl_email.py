import smtplib
import openpyxl as xl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

username = str(input('Your Username:' ))
password = str(input('Your Password:' ))
From = username
Subject = 'Test'

wb = xl.load_workbook(r'emailer.xlsx')
sheet1 = wb.get_sheet_by_name('Sheet1')

names = []
emails = []
content = []

for cell in sheet1['A']:
    emails.append(cell.value)

for cell in sheet1['B']:
    names.append(cell.value)

for cell in sheet1['C']:
    content.append(cell.value)

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(username, password)

for i in range(len(emails)):
    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = names[i]
    msg['Subject'] = Subject
    text = '''
Dear {},
        
Congratulations!! This is a test email by Sangram Chavan
Idea         : {}


Follow us on Social Media @wdevops maybe you get some tips and contest notifications.
        
Keep Coding!
And, off course, don’t forget to have fun! Stay MAD!!
        
Regards,
Sangram Chavan
        '''.format(names[i],content[i])
    msg.attach(MIMEText(text, 'plain'))
    message = msg.as_string()
    server.sendmail(username, emails[i], message)
    print('Mail sent to', emails[i])
server.quit()
print('All emails sent successfully!')
