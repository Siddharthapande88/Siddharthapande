'''
Created on Dec 19, 2017

@author: Sid
'''
# text to multiple worksheet data conversion automation
import xlsxwriter
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

try :
    
    workbook=xlsxwriter.Workbook("/Users/sysadmin/Desktop/foobar.xlsx")
    worksheet=workbook.add_worksheet() #creating a worksheet
    r=0
    sheet=1
    bold = workbook.add_format({'bold': True})  #header format
    
    with open ('/Users/sysadmin/Desktop/test.txt','r') as file_read:
        for i in file_read:  #reading each line of text file 
            if (r>=1000000):      #logic to create new sheet if specified no of rows are over in current sheet
                sheet=sheet+1
                worksheet=workbook.add_worksheet()
                r=0
            c=0
            column_list=i.split('@#!') #splitting one record to create a list of column values
            if (sheet==1 and r==0):
                header_list=column_list #capturing header for re-use, each time a new worksheet is created
            if (sheet>1 and r==0):  #inserting header in new worksheet
                for l,k in enumerate(header_list):
                    if l>0 :
                        value=k.rstrip('\n')
                        worksheet.write_string(r,c,value,bold)
                        c=c+1
                r=r+1
                c=0
            for l,k in enumerate(column_list): # inserting columns values from list into each column of a worksheet row
                if l>0 :
                    value=k.rstrip('\n')
                    if (r==0):
                        worksheet.write_string(r,c,value,bold)
                    else:
                                                
                        worksheet.write_string(r,c,value)
                c=c+1
            r=r+1
            
    workbook.close()
    print("Data is written in the workbook name foobar.xlsx")
    
except UnicodeDecodeError,e:
    print(e)
