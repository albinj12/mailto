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

i=0
for i in range(len(mailID)):
print(mailID[i])