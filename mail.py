import openpyxl
import smtplib

mailID = []

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
        mailID.append(cell_obj.value)
    i+=1


s = smtplib.SMTP('smtp.gmail.com',587)

s.starttls()

s.login("testu7812@gmail.com","7812@testu")

message = "This is a test message"

for i in range(len(mailID)):
    s.sendmail("testu7812@gmail.com",mailID[i],message)

s.quit()