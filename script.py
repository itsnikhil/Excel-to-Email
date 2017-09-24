import xlrd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import re

workbook  = xlrd.open_workbook('data.xlsx')
sheet = workbook.sheet_by_index(0)

text1 = ''
text2 = ''
ids = ''

for i in range(1,sheet.nrows):
  ids += str(sheet.cell_value(i,0))+','

for i in range(1,sheet.nrows):
  text1 += str(sheet.cell_value(i,1))+'<br><br>'

for i in range(1,sheet.nrows):
  text2 += str(sheet.cell_value(i,2))+'<br><br>'
f = open('1.html','w')
file = open('email.html','rU')
f.write(re.sub('class="put-name">','class="put-name">'+text1,file.read()))
f.close()
file.close()
f = open('final.html','w')
file = open('1.html','rU')
f.write(re.sub('class="put-amount">','class="put-amount">'+text2,file.read()))
f.close()
file.close()
file = open('final.html','rU')
html = file.read()
file.close()
os.remove('1.html')

msg = MIMEMultipart('alternative')
msg['Subject'] = "Python test solution!"
msg['From'] = 'sharma645445@gmail.com'
msg['To'] = ids[:-1]

part = MIMEText(html, 'html')

msg.attach(part)

password = 'aassddffgghhjjkkll'
server = smtplib.SMTP('smtp.gmail.com', 587)  
server.ehlo()
server.starttls()
server.login(msg["From"],password)
#server.sendmail(msg["From"], msg["To"].split(","),msg.as_string())
server.quit()
os.remove('final.html')
print 'Done'