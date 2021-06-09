import requests
import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd

file_name=''
contacts=[]

def sendPostRequest(PhoneNumber,msg):
    return requests.post(f"http://bulksms.alayada.com/vendorsms/pushsms.aspx?user=aryanpatel&password=Aryan@Patel&msisdn={PhoneNumber}&sid=SIHHUB&msg={msg}&fl=0&dc=8")


def validation(x):
    if x.isdigit() and (len(x)==10 or len(x)==12):
        return 'Valid'
    return 'Invalid'


def drawTable(contacts):
    tabel_text.config(state='normal')
    tabel_text.delete(1.0,tk.END)
    tabel_text.insert(1.0, "No   |        Contacts      |   valid/invalid"+ '\n')
    for i in range(len(contacts)):
        string= str(i+1) + (5-(1*len(str(i+1))))*" " + "|       "+ str(contacts[i])+ "     |      "+ validation(str(contacts[i])) + "\n"
        tabel_text.insert(float(i+2),string)
    tabel_text.config(state='disable')

def open_file():
    global file_name,contacts
    file_name = filedialog.askopenfilename(initialdir=os.getcwd(), title='Select File', filetypes =[('Excel Files', '*.xlsx')])
    tabel_text.config(state='normal')
    try:
        file = pd.read_excel(file_name, names=['contact'], header=None)
        contacts= file['contact'].tolist()
        filename_label.config(text=os.path.basename(file_name),font=("Sans-serif", 12))
        drawTable(contacts)
        status.config(text='message will send to only valid numbers')

    except FileNotFoundError:
        return
    except:
        return


def send():
    global contacts,file_name
    if not file_name:
        messagebox.showwarning(title='warning',message='please select the excel file')
    elif text_editor.get(1.0, tk.END)== "\n":
        messagebox.showwarning(title='message field Empty', message='please write message')
    else:
        text= str(text_editor.get(1.0, tk.END))
        text= text.split('\n')
        text = ' %0a '.join(text)
        contacts= [i for i in contacts if validation(str(i))=='Valid']
        for i in contacts:
            response=sendPostRequest(i,text)
            # print(response.text)
        if response.json()['ErrorCode']== '000':
            remain_msg = requests.get("http://bulksms.alayada.com/vendorsms/CheckBalance.aspx?user=aryanpatel&password=Aryan@Patel")
            remain_msg = (remain_msg.text).split('#')[1].split('|')[0].split(':')[1]
            status_text = f"Message sent sucessfully to {str(len(contacts))} contacts (SMS balance = {remain_msg})"
            status.config(text=status_text)
            contacts = []
            filename_label.config(text=" ")
            file_name= ""
        elif response.json()['ErrorCode']== '24':
            status.config(text="Invalid Template, Messages not sent")
        else:
            status.config(text=f"Error{response.json()['ErrorCode']} : {response.json()['ErrorMessage']}")



# window generate-------------------------------------------------
win = tk.Tk()
# win.config(background= 'white')
win.title('Message Sender')
win.geometry('700x630')
win.maxsize(700,630)
f=ttk.Frame(win)
f.pack(side='top')

f2=tk.Frame(win,pady=10)
f2.pack(side='top')

f3=ttk.Frame(win)
f3.pack(side='top')

f4=ttk.Frame(win)
f4.pack(side='bottom')

# main frame ----------------------------------------------------------
title= tk.Label(f,text='Bulk Message Sender ',font=("Sans-serif", 24,'bold'),pady=10)
title.grid(row=0,column=0)

file_label= tk.Label(f,text='Select Excel File : ',font=("Sans-serif", 14))
file_label.grid(row=2,column=0)

filename_label= tk.Label(f,font=("Arial", 14))
filename_label.grid(row=2,column=2)

open_button= ttk.Button(f,text='open',command=open_file)
open_button.grid(row=2,column=1)

message_label= tk.Label(f,text='Write the Message below : ',font=("Sans-serif", 14),pady=4)
message_label.grid(row=4,column=0)

# frame 2 for message writing---------------------------------------------------------
text_editor = tk.Text(f2,height=10)
text_editor.config(wrap='word', relief=tk.FLAT)

scroll_bar = tk.Scrollbar(f2)
text_editor.focus_set()
scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)
text_editor.pack(fill=tk.X)
scroll_bar.config(command=text_editor.yview)
text_editor.config(yscrollcommand=scroll_bar.set)

# frame 3 for contact shownig-----------------------------------------------------------------
contact_label=tk.Label(f3,text='Contacts from the file',font=("Sans-serif", 12))
contact_label.pack(side='top')

tabel_text = tk.Text(f3, height=15,state='disable')
tabel_text.config(wrap='word', relief=tk.FLAT)

scroll = tk.Scrollbar(f3)
tabel_text.focus_set()
scroll.pack(side=tk.RIGHT, fill=tk.Y)
tabel_text.pack(fill=tk.X)
scroll.config(command=tabel_text.yview)
tabel_text.config(yscrollcommand=scroll.set)

# frame 4 for buttones and status------------------------------------------------------
submit=ttk.Button(f4,text='send',command=send)
submit.pack(side='bottom')

status= tk.Label(f4,font=("Arial", 13))
status.pack(side='right')

win.mainloop()













# --------------------------------------------------------------------------------------------------------------------------













# import requests
# param= {
# 'Userid' : 'Demo4Trm',
# 'UserPassword' : 1231993,
# 'PhoneNumber' : 917802938632,
# 'Text' : 'hii',
# 'GSM' : 'SRIMSG'
# }
# url= "http://ip.shreesms.net/smsserver/SMS10N.aspx"
# response= requests.post(url,param)
#
# print(response.text)



import requests
# param= {
# 'Userid' : 'Demo4Trm',
# 'UserPassword' : 1231993,
# 'PhoneNumber' : 917802938632,
# 'Text' : 'hii for check',
# 'GSM' : 'SRIMSG'
# }
#
# response= requests.post(f"http://ip.shreesms.net/smsserver/SMS10N.aspx?Userid={param['Userid']}&UserPassword={param['UserPassword']}&PhoneNumber={param['PhoneNumber']}&Text={param['Text']}&GSM={param['GSM']}")
#
# print(response.text)



# msg= '''હવે સોલાર રૂફટોપ પર મેળવો 1000 થી 2000 રૂપીયા વળતર. આજે જ લગાવો સોલાર રૂફટોપ તમારા ઘરે અને બચાવો 30000/yr સુધી નુ લાઈટબીલ.
#
# 7984102212
# 7990361324
#
# 406, શિવાલય કોમ્પલેક્ષ,  મોવડી ચોકડી, રા જકોટ'''
#
# msg= msg.split('\n')
# f_msg= ' %0a '.join(msg)
# print(f_msg)
#
# re= requests.post(f"http://bulksms.alayada.com/vendorsms/pushsms.aspx?user=aryanpatel&password=Aryan@Patel&msisdn=7802938632&sid=SIHHUB&msg={f_msg}&fl=0&dc=8")
# print(re.text)
