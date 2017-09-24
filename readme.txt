Requirements to run:
	1) pip install workbook
	2) python version 2.7.13 (because i have used print 'hello' which will throw error with python version 3)
	3) for testing purposes I created a gmail account "sharma645445@gmail.com" with password "aassddffgghhjjkkll" and we will use that account
	Note: os, smtplib and re comes pre-installed, for sending mail using smtp server which require us to lower gmail setting (read more about lower gmail setting here https://support.google.com/accounts/answer/6010255?hl=en) 

This folder contains 2 python scripts namely script.py and script2.py

I was a little confused with the question if we want to email Name and Amount of everyperson to everyone or Name and Amount of each person to them only.
So, script.py will read data from data.xlsx file for all the emails, names and accounts and will send mail data of everyperson to everyone!
and script2.py will read data from data.xlsx file for all the emails, names and accounts and will send mail data of each person to them only!

How to run?
	Just the way you run a python program, run file from terminal/cmd or just double click the program and it will execute and close and boom email has been sent 
	to the people who's email is listed in data.xlsx file. Email will include following details: Name and Account, if you run script.py then email will contain Names and Account of all
	listed in that xlsx file, or  if you run script2.py then email will contain Name and Account of just that person only 

How email has been generated?
	I have used a free Email templete as per question and used regular expression to find for class="put-name" and class="put-account" and replaced it with some html code which
	contains all necessary details

How mail is sent?
	The smtplib module defines an SMTP client session object that can be used to send mail to any Internet machine with an SMTP or ESMTP listener daemon. 

**************************************************************************************************************************************************************************************************************
EVERTHING HAS BEEN MADE SOLELY BY: 010011100110100101101011011010000110100101101100 010101000110000101101110011001010110101001100001 
for any more querries, drop me a mail taneja.nikhil03@gmail.com
**************************************************************************************************************************************************************************************************************





