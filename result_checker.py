import urllib.request

from xlwt import Workbook 

# Workbook is created 
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('4th Sem')

ini=['Reg_No.','Name','SGPA']
for m in range(3):
    sheet1.write(1, 1+m, ini[m])
    
'''This is list of Registration numbers
if yours is missing simply add it in the end by +[12345678910]
where '12345678910' is your reg no'''
reg_numlist=list(range(18103108001,18103108057))+[19104108905,19103108902]    

count=0
for reg_no in reg_numlist:
    
    text=str(urllib.request.urlopen("http://results.akuexam.net/ResultsBTechBPharm4thSemPub2020.aspx?Sem=IV&RegNo={0}".format(reg_no),timeout=10).read())

    def student_name(text):
        prefix='''<span id="ctl00_ContentPlaceHolder1_DataList1_ctl00_StudentNameLabel" style="font-weight: 700">'''
        n=text.find(prefix)+len(prefix)
        suffix=text.find('''</span>''',n)
        name=''
        for j in range(suffix-n):
            name+=text[n+j]
        return name

    def student_sgpa(text):
        prefix='''<span id="ctl00_ContentPlaceHolder1_DataList5_ctl00_GROSSTHEORYTOTALLabel" style="font-weight: 700">'''
        s=text.find(prefix)+len(prefix)
        suffix=text.find('''</span>''',s)
        sgpa=''
        for i in range(suffix-s):
            sgpa+=text[s+i]
        return (sgpa)
    
    output = [reg_no,student_name(text),student_sgpa(text)]
    count+=1
    # writing to results.xls
    for l in range(3):
        sheet1.write(2+count, 1+l, output[l])

    print(output)

#Saving the workbook
wb.save('result.xls')

print('''\n "result.xls" saved to folder \n C:\\Users\keshav\AppData\Local\ \n Programs\Python\Python39\\aku_result_checker''')
