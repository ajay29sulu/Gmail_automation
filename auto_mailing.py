import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import datetime

today = datetime.datetime.now().strftime("%d-%m")
yearNow = datetime.datetime.now().strftime("%Y")

#you can choose your own excel file.....
employee_data = "xyz.xlsx"
read_data = pd.read_excel(employee_data, sheet_name="Sheet1")
hd = pd.read_excel(employee_data, sheet_name="Sheet2")

from_addr='***************@gmail.com'

to_addr=[]
j=0
for index, item in hd.iterrows():
    holiday = item['Date'].strftime("%d-%m")
    if today == holiday:
        for i in range(3):
            mail = read_data["Email"][i]
            to_addr.append(mail)

        msg = MIMEMultipart()
        msg['From'] = from_addr

        msg['To'] = ",".join(to_addr)
        msg['subject'] = 'NOTICE FOR HOLIDAY'

        body = 'Dear all Staff, I wish to inform u all that today  will be your holiday on the occasion of  ' + hd["Holiday"][j] + ', Stay safe and stay healthy'
        msg.attach(MIMEText(body, 'plain'))

        email = '****************@gmail.com'
        password = '@@@@@@@@@1@'
        mail = smtplib.SMTP('smtp.gmail.com', 587)
        mail.ehlo()
        mail.starttls()
        mail.login(email, password)
        text = msg.as_string()
        mail.sendmail(from_addr, to_addr, text)
        mail.quit()

        break
    j=j+1
