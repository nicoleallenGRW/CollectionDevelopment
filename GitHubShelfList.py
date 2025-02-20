# -*- coding: utf-8 -*-
#!/usr/bin/python2.7
#
# Test being run to export shelf list to collection managers
# Email Excel Spreadhseet to manager and supervisor 
# Use XlsxWriter to create spreadsheet from SQL Query
# 
#

import psycopg2
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from datetime import datetime



excelfile =  'ItemsMarkedMissingPastTwoWeeks.xlsx' #this is the name of the excel file you are creating



#Set variables for email

emailhost = 'xxx.xxx.x.xx' #Update for your email host
emailport = '25' #Update for your port

try:
    conn = psycopg2.connect("dbname=iii user=sqlaccess host=sierra-db port=1032 password=PASSWORD sslmode=require") #update this information for your specific location
except psycopg2.Error as e:
    print ("Unable to connect to database: " + str(e))
    
cursor = conn.cursor()
cursor.execute(open("Art750.sql","r").read()) #this is the name of the SQL you built
rows = cursor.fetchall()
target = cursor.rowcount 
target1 = target - 3062 #sets the target number for items in the collection
emailsubject1 = '[Weeding] Art 750 - 769'
emailsubject2 = '[On Target] Art 750 - 769'
#Creates email message for collections under or over target.
emailmessage1 = '''Hi!

You are above the target size of your collection. Please weed {0} items to meet target size of 3062 items. Attached you will find a shelf list containing circulation data.
If you have any questions, please contact Stephen Schmidt in Resources Management. Currently the total size of the collection is {1}.'''.format((target1),(target)) 
emailmessage2 = '''Hi!

Your collection is on or under target.  Currently the target is 3062 and there are {0} items in this collection.  Thanks!'''.format(target)
emailfrom = 'nallen@greenwichlibrary.org' #this is your email. It means staff can hit reply and contact you with questions.
emailto = ['nallen@greenwichlibrary.org', 'xxxx@greenwichlibrary.org', 'abckd@greenwichlibrary.org'] #this is the list of people you want the email to go to.

workbook = xlsxwriter.Workbook(excelfile, {'remove_timezone': True})
worksheet = workbook.add_worksheet()

worksheet.set_landscape()
worksheet.hide_gridlines(0)
worksheet.repeat_rows(0)
worksheet.set_print_scale(64)
worksheet.set_margins(left=0.2, right=0.2) #created margins and size so staff can print spreadsheet if they need to


eformat= workbook.add_format({'text_wrap': False, 'valign': 'bottom'})
eformat2= workbook.add_format({'text_wrap': False, 'valign': 'bottom', 'num_format': 'mm/dd/yy'}) #basic information that controls the display of your spreadsheet -ex. how dates are formated and whether anything is bolded-
eformatlabel= workbook.add_format({'text_wrap': False, 'valign': 'bottom', 'bold': True})
# sets the width of your columns (notice that in Python spreadsheets start with row and column 0
worksheet.set_column(0,0,14.43)
worksheet.set_column(1,1,62.86)
worksheet.set_column(2,2,21.57)
worksheet.set_column(3,3,8.71)
worksheet.set_column(4,4,13)
worksheet.set_column(5,5,18.57)
worksheet.set_column(6,6,9.29)
worksheet.set_column(7,7,9.29)
worksheet.set_column(8,8,4)
worksheet.set_column(9,9,4)
worksheet.set_column(10,10,11.71)
worksheet.set_column(11,11,15.86)
worksheet.set_column(12,12,8.43)
worksheet.set_column(13,13,31.71)
worksheet.set_column(14,14,34)

worksheet.set_header('&CArt 750-769') #what you want to call your worksheet header
worksheet.set_footer('&CPage &P of &N&R&D') #what you want to call your worksheet footer
#Column headings
worksheet.write(0,0,'Call Number', eformatlabel)
worksheet.write(0,1,'Title', eformatlabel)
worksheet.write(0,2,'Author', eformatlabel)
worksheet.write(0,3,'Pub. Year', eformatlabel)
worksheet.write(0,4,'Item Created', eformatlabel)
worksheet.write(0,5,'Last Checkin', eformatlabel)
worksheet.write(0,6,'Checkouts', eformatlabel)
worksheet.write(0,7,'Renewals', eformatlabel)
worksheet.write(0,8,'YTD', eformatlabel)
worksheet.write(0,9,'LYR', eformatlabel)
worksheet.write(0,10,'Checked Out', eformatlabel)
worksheet.write(0,11,'Barcode', eformatlabel)
worksheet.write(0,12,'Status', eformatlabel)
worksheet.write(0,13,'590', eformatlabel)
worksheet.write(0,14,'695', eformatlabel)
#sets formating for the data returned from your SQl query. 
for rownum, row in enumerate(rows):
    worksheet.write(rownum+1,0,row[0], eformat)
    worksheet.write(rownum+1,1,row[1], eformat)
    worksheet.write(rownum+1,2,row[2], eformat)
    worksheet.write(rownum+1,3,row[3], eformat)
    worksheet.write(rownum+1,4,row[4], eformat2)#Line 65 above refers to how to address dates. This is a date field
    worksheet.write(rownum+1,5,row[5], eformat2)
    worksheet.write(rownum+1,6,row[6], eformat)
    worksheet.write(rownum+1,7,row[7], eformat)
    worksheet.write(rownum+1,8,row[8], eformat)
    worksheet.write(rownum+1,9,row[9], eformat)
    worksheet.write(rownum+1,10,row[10], eformat)
    worksheet.write(rownum+1,11,row[11], eformat)
    worksheet.write(rownum+1,12,row[12], eformat)
    worksheet.write(rownum+1,13,row[13], eformat)
    worksheet.write(rownum+1,14,row[14], eformat)    
    

workbook.close()

#Creating the email message sending shelf list for both under and over target collection size
if target > 3062:
    msg = MIMEMultipart()
    msg['From'] = emailfrom
    if type(emailto) is list:
        msg['To'] = ','.join(emailto)
    else:
        msg['To'] = emailto
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = emailsubject1
    msg.attach (MIMEText(emailmessage1))
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(excelfile,"rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename=%s' % excelfile)
    msg.attach(part)
    smtp = smtplib.SMTP(emailhost, emailport)
    smtp.sendmail(emailfrom, emailto, msg.as_string())

else:
    msg = MIMEMultipart()
    msg['From'] = emailfrom
    if type(emailto) is list:
        msg['To'] = ','.join(emailto)
    else:
        msg['To'] = emailto
    msg['Date'] = formatdate(localtime = True)
    msg['Subject'] = emailsubject2
    msg.attach (MIMEText(emailmessage2))
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(excelfile,"rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename=%s' % excelfile)
    msg.attach(part)
    smtp = smtplib.SMTP(emailhost, emailport)
    smtp.sendmail(emailfrom, emailto, msg.as_string())
    smtp.quit() 


    








