import pandas as hi
import smtplib as mail
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pyautogui

def tkinter():
    global screen
    
    screen=Tk()
    screen.geometry("700x700")
    screen.title("PY-Auto Mail")
    screen.resizable(0,0)
    head_label=Label(text="PY-Auto MAIL",bg="grey",fg="white",height= "2").pack()
    Label(text="").pack()
    Label(text="Pls Select the Excel File contains Email list",bg="grey",fg="white").pack()
    Label(screen,text=" ").pack()
    folder_sel_button=Button(text = "Browse Files",
                             command = selectexcel,
                             font=("Comic sans ms",10),
                             relief=GROOVE)
    folder_sel_button.place(x=300,y=180)
    
  
    screen.mainloop()

def close():
	screen.destroy()

def selectexcel():
    
    global filename
    global emailrow
    global listofmail
    global filecheck
    try:
        filename = filedialog.askopenfilename(initialdir = "/",
                                              filetypes = (("Excel files","*.xlsx*"),
                                                           ("all files","*.*")))
        # reading email from xl
        data = hi.read_excel(str(filename))
        filecheck=type(data)
        text="<class 'pandas.core.frame.DataFrame'>"
     
        if str(filecheck)==text :
            emailrow = data.get("EMAIL")
            listofmail = list(emailrow)
            file_selected_lbl=Label(text="\n\nFile Selected:"+filename+"\n\npls click next",
                                fg="green")
            file_selected_lbl.place(x=180,y=220)
            Label(screen,text=" ").pack()
            next_button=Button(text = "NEXT",
                               command = next,
                               font=("Comic sans ms",10),
                               relief=GROOVE)
            next_button.place(x=325,y=390)
            messagebox.showinfo("Success",str(emailrow))
            
            
    except Exception as e:
            print(e)
            messagebox.showerror("Failed","pls upload excel file with column header 'EMAIL'")
            file_selected.config(text="pls upload excel file with column header 'EMAIL'",fg="red").pack()
            
        
def next():
    
    global screen1
    global email
    global password
    global email_entry

    
    screen1=Toplevel(screen)
    screen1.title("LOGIN PAGE for PY-Auto MAIL")
    screen1.geometry("850x850")
    
    to_ = listofmail
    email=StringVar()
    password=StringVar()

    Label(screen1,text="Pls ENTER your EMAIL & PASSWORD ",fg="black",height="2").pack()
    Label(screen1,text=" ").pack()
    Label(screen1,text=" EMAIL",height="2").pack(padx=10,pady=10)
    
    email_entry=Entry(screen1,textvariable=email).pack(padx=10,pady=10)
    Label(screen1,text=" ").pack()
    Label(screen1,text=" PASSWORD",height="2").pack()
    
    password_entry=Entry(screen1,textvariable=password,show='*').pack(padx=10,pady=10)
    Label(screen1,text=" ").pack()
    Button(screen1,text = "Login",command = connect_Server1,relief=GROOVE).pack()
     
def connect_Server1():
    
    global subject_entry
    global subject
    global body_entry
    global body
    global server
    
    email_info=email.get()
    password_info=password.get()
    
    
    try:
        # object of SMTP
        server = mail.SMTP("smtp.gmail.com", 587)
        server.starttls()
        # login
        server.login(str(email_info), str(password_info))
        login_success=str(server)
        if login_success:
            subject=StringVar()
            body=StringVar()
            Label(screen1,text="Login success!",fg="green").pack()
            Label(screen1,text=" ").pack()
            sub_lbl=Label(screen1,text="SUBJECT OF EMAIL")
            sub_lbl.place(x=60,y=393)
            Label(screen1,text=" ").pack()
            subject_entry=Entry(screen1,textvariable=subject,width=50)
            subject_entry.place(x=205,y=394)
            Label(screen1,text=" ").pack()
            Label(screen1,text=" ").pack()
            body_lbl=Label(screen1,text="BODY OF EMAIL")
            body_lbl.place(x=60,y=464)
            body_entry=Entry(screen1,textvariable=body)
            body_entry.place(x=205,y=470,
                            width=400,height=150)
            
            send_msg_btn=Button(screen1,text = "SEND MSG",command = send_msg,relief=GROOVE)
            send_msg_btn.place(x=360,y=650)
            exit_button = Button(screen1, text="Exit", command=screen.destroy,relief=GROOVE)
            exit_button.place(x=375,y=720)
            
    except Exception as e:
            print(e)
            messagebox.showerror(" Failed","FAILED!\n pls check USERNAME and PASSWORD!!\n make sure your GMAIL ID's LESS SECURE APP ACCESS is ON")
                        
    
def send_msg():
    
    global from_ ,to_
    global message
    global email_entry
    global password_entry
    global server
    
    email_info=email.get()
    password_info=password.get()
    subject_entry=subject.get()
    body_entry=body.get()
   
    from_ = str(email_info)
    to_ = listofmail

    # object of SMTP
    server = mail.SMTP("smtp.gmail.com", 587)
    server.starttls()
    # login
    server.login(str(email_info), str(password_info))
    login_success=str(server)
    message = MIMEMultipart("alternative")
    message['subject'] = str(subject_entry)
    email_info=email.get()
    message['from'] = str(email_info)
    message['to'] = str(email_info)

      
    html=str(body_entry)
    email_body=MIMEText(html,"html")
    #attaching the msg
    message.attach(email_body)
    
    # sending the msg
    server.sendmail(from_, to_, message.as_string())
    messagebox.showinfo("Success","mail sented sucessfully")




tkinter()


