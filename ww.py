import pywhatkit
from tkinter import *
from tkinter.filedialog import askopenfilename
import time 
import csv 
import urllib.parse
import pandas as pd
import openpyxl

filename = ""
filename1 = ""

#GUI Part Starts
window = Tk()
window.title("Welcome to Web_Whatsapp")
window.geometry('350x200')

#Message File lable and button
lbl = Label(window, text="Choose Your Message File ")
lbl.grid(column=0, row=0)
def clicked():
    global filename
    filename = askopenfilename()
    lbl.configure(text="You have choosed "+filename.split('/')[-1])
    print(filename.split('/')[-1])
btn = Button(window, text="Choose File", command=clicked)
btn.grid(column=1, row=0)

#Csv File lable and button
lbl1 = Label(window, text="Choose Your Contact File ")
lbl1.grid(column=0, row=1)
def clicked():
    global filename1
    filename1 = askopenfilename()
    lbl1.configure(text="You have choosed "+filename1.split('/')[-1])
    print(filename1.split('/')[-1])
btn1 = Button(window, text="Choose File", command=clicked)
btn1.grid(column=1, row=1)

#Submit
def clicked():
    print("run")
    window.destroy()  
btn2 = Button(window, text="Start Without Attachment", command=clicked)
btn2.grid(column=0, row=4)
window.mainloop()
#GUI Part Ends

FILE_NAME = filename1.split('/')[-1]
FIELD_NAME = 'Number'
MSG_FILE = filename.split('/')[-1]

with open(FILE_NAME, 'r') as f:
    reader = csv.DictReader(f)
    all_names = [i[FIELD_NAME] for i in reader if i[FIELD_NAME].isdigit()]
    
with open(MSG_FILE, 'r') as f:
    msg = f.read()
    #msg=urllib.parse.quote(msg)
    
lstData=[]
for name in all_names:
        pywhatkit.sendwhatmsg_instantly("+91"+name, msg, 15,True,2)
        time.sleep(3)
        print(name,"Message Sent")
        lstData.append([name,"Message Sent"])
df = pd.DataFrame(lstData, columns=['Number', 'Status'])
df.to_excel('OutputOnlyMessage.xlsx', sheet_name='OutputOnlyMessage')

