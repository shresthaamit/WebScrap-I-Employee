import requests
import  xlwt
from xlwt import Workbook
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE,formatdate
USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36'
REQUEST_HEADER ={
    'User-Agent': USER_AGENT,
    'Accept-Language': 'en-US,en;q=0.5',
}
BASE_URL = 'https://dummy.restapiexample.com/api/v1/employees'


def get_employees():
    req =  requests.get(url=BASE_URL,headers=REQUEST_HEADER)
    req.raise_for_status() 
    return req.json().get('data', [])



def convet_to_excel(data):
    # print(data)
    work_book  = Workbook()
    employee_sheet = work_book.add_sheet("Employees")
    header = list(data[0].keys())[:-1]
    for i in range(0,len(header)):
        employee_sheet.write(0, i, header[i])
    for j in range(len(data)):
        employee = data[j]
        values= list(employee.values())[:-1]
        for eachvalue in range( len(values)):
            employee_sheet.write(j+1, eachvalue,values[eachvalue])
    work_book.save("Employees.xls")
    
def send_email(send_from, send_to, send_subject, detail,files=None):
    assert isinstance(send_to,list)
    msg = MIMEMultipart()
    msg['From']=send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date']=formatdate(localtime=True)
    msg['Subject'] =send_subject
    msg.attach(MIMEText(detail))
    for i in files or []:
        with open(i,'rb') as fil:
            part = MIMEApplication(fil.read(), Name = basename(i))
        part['Content-Disposition'] = f'attachment;filename = "{basename(i)}"'
        msg.attach(part)
    smtp = smtplib.SMTP('smtp.gmail.com:587', timeout=60)
    smtp.starttls()
    # smtp.set_debuglevel(1)
    smtp.login(send_from,'hpsl wrwu sqsp yawx')
    smtp.sendmail(send_from,send_to,msg.as_string())
    smtp.close()
    print("Email sent")

if __name__ == '__main__':
    data = get_employees()[1:]
    convet_to_excel(data)
    send_email('sthaamit22fyp@gmail.com',['shresthaamit273@gmail.com'],'Employee Detail','Please find the file of Employee ', 
               files=['Employees.xls'])
    print("Email send")