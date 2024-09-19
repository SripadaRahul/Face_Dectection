import smtplib
from email.message import EmailMessage
import glob
from tkinter import messagebox
from datetime import date
import time
import datetime
import pandas as pd

today = date.today()
def mail(check):
    if check:
        print("Automail Time set Done")
        set_time = "12:42:30.000000"  # time at which the mail need to be sent
        a = set_time
        b = str(datetime.datetime.now())
        b = b[11:]
        diff_hr = int(a[:2]) - int(b[:2])
        diff_min = int(a[3:5]) - int(b[3:5])
        diff_sec = float(a[6:]) - float(b[6:])

        diff = (((diff_hr * 60) + (diff_min)) * 60) + (diff_sec)
        # print(diff)
        time.sleep(diff)

    today = date.today()
    msg = EmailMessage()
    tm = str(datetime.datetime.now())[11:16]
    d = datetime.datetime.strptime(tm, "%H:%M")
    tm = d.strftime("%I:%M %p")
    msg['Subject'] = 'Attendance sheet -'+str(today)+" "+str(tm)
    msg['From'] = 'ACE'
    msg['To'] = 'sripadarahul28@gmail.com'

    with open('hello.txt') as myfile:
        data = myfile.read()
        msg.set_content(data)

    for files in glob.glob("data/Attendance_xlsx/*.xlsx",recursive=True):

        with open(files, "rb") as f:
            file_data = f.read()

            file_name = f.name

            msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=file_name)



    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login("sripadarahul428@gmail.com", "rahul@1928")
        server.send_message(msg)

    for sheet in glob.glob("data/Attendance_xlsx/*.xlsx", recursive=True):
        try:
            df=pd.read_excel(sheet)
            df.drop('Name', inplace=True, axis=1)
            df.drop('Time', inplace=True, axis=1)
            df.to_excel(sheet, index=False)
        except:
            pass
    df = pd.read_csv("C:\\Users\\sripa\\PycharmProjects\\Face\\data.csv")
    df.drop('Name', inplace=True, axis=1)
    df.drop('Time', inplace=True, axis=1)
    df.loc[0, 'Name'] =""
    df.loc[1, 'Time'] =""
    df.to_csv("C:\\Users\\sripa\\PycharmProjects\\Face\\data.csv", index=False)



    if check:
        print('Automatic Email sent!!!')
    else:
        print('Manual Email sent!!!')
        messagebox.showinfo("Information", "          Mail sent......!!!        ")
    quit()


