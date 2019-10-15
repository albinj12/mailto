import openpyxl
import smtplib
import getpass

mail_list = []

path = input("Enter path of the excel file: ")

wb_obj = openpyxl.load_workbook(path) 
sheet_obj = wb_obj.active

column = int(input("Enter the column of mail-ID"))

i=1
while True:
    cell_obj = sheet_obj.cell(row = i+1,column = column)
    if(cell_obj.value==None):
        break
    else:
        mail_list.append(cell_obj.value)
    i+=1


s = smtplib.SMTP('smtp.gmail.com',587)

s.starttls()

mailID = input("Enter your mailID")
password = getpass.getpass(prompt ="Enter your mail password: ")

s.login(mailID,password)

message = "This is a test message"

for i in range(len(mail_list)):
    s.sendmail(mailID,mail_list[i],message)

s.quit()