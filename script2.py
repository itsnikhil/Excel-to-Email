import xlrd
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import re

workbook  = xlrd.open_workbook('data.xlsx')
sheet = workbook.sheet_by_index(0)

ids = ''
text1=''
text2=''

for i in range(1,sheet.nrows):
  ids += str(sheet.cell_value(i,0))+','
a = ids[:-1].split(',')
for i in range(1,sheet.nrows):
  text1 += str(sheet.cell_value(i,1))+'<br><br>'
b = text1[:-8].split('<br><br>')
for i in range(1,sheet.nrows):
  text2 += str(sheet.cell_value(i,2))+'<br><br>'
c = text2[:-8].split('<br><br>')

i=0
while i<(sheet.nrows-1):
	file = open('email.html','rU')
	f = open('1.html','w')
	f.write(re.sub('class="put-name">','class="put-name">'+b[i]+'<br><br>',file.read()))
	f.close()
	file = open('1.html','rU')
	f = open('final.html','w')
	f.write(re.sub('class="put-amount">','class="put-amount">'+c[i]+'<br><br>',file.read()))
	f.close()
	file.close()
	file = open('final.html','rU')
	html = file.read()
	file.close()

	msg = MIMEMultipart('alternative')
	msg['Subject'] = "Python test solution :"
	msg['From'] = 'sharma645445@gmail.com'
	msg['To'] = a[i]

	part = MIMEText(html, 'html')

	msg.attach(part)

	password = 'aassddffgghhjjkkll'
	server = smtplib.SMTP('smtp.gmail.com', 587)  
	server.ehlo()
	server.starttls()
	server.login(msg["From"],password)
	server.sendmail(msg["From"], msg["To"],msg.as_string())
	server.quit()
	os.remove('1.html')
	os.remove('final.html')
	i+=1
print 'Done'
