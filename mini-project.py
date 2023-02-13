import tkinter as tk
import mysql.connector
from tkinter import filedialog 
import xlrd
import os
from os import listdir
import smtplib


cnx = mysql.connector.connect(user='root',password='Kanna12@',auth_plugin='mysql_native_password',host='127.0.0.1',database="kanna")
mycursor=cnx.cursor()
#back function for page2
def back_function():
    frame2.pack_forget()
    entry.delete(0, tk.END)
    entry1.delete(0, tk.END)
    frame.pack()

#back funtion for page3 of add button
def aback_function():
    aframe3.pack_forget()
    frame2.pack()

#back funtion for remove employee button
def rback_function():
    rframe3.pack_forget()
    frame2.pack()

#back funtion for view emplyee button
def vback_funtion():
    vframe3.pack_forget()
    frame2.pack()
# funtion for update back button
def Uback_funtion():
    Uframe.pack_forget()
    frame2.pack()

#back funtion for generate back
def gbBack():
    Gframe.pack_forget()
    Uframe.pack()

#funtion for adding employee
def AddEmp():
    
    global counta
    if(counta==0):
        frame2.pack_forget()
        label=tk.Label(master=aframe3,text="you can add know",fg="black")
        label.pack()
        button1 = tk.Button(master=aframe3,text="back",fg="black",command=aback_function)
        button1.pack()
        aframe3.pack()
        counta+=1
    else:
        frame2.pack_forget()
        aframe3.pack()

        
def mailData():
    mycursor.execute("select * from employee_personnel")
    resultset=mycursor.fetchall()
    for i in resultset:
        with smtplib.SMTP('smtp.gmail.com',587) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login("saiscommercial@gmail.com","Saiscommercial@955")
            sub="monthly salary slip"
            mycursor.execute("select mail from employee where id='{}'".format(i[0]))
            result=mycursor.fetchone()
            msg=str(i)
            smtp.sendmail("saiscommercial@gmail.com",str(result[0]),msg)
    



        

#funtion for removing employee

def RemEmp():
    global countr
    if(countr==0):
        frame2.pack_forget()
        label=tk.Label(master=rframe3,text="you can remove know",fg="black")
        label.pack()
        buttonr = tk.Button(master=rframe3,text="back",fg="black",command=rback_function)
        buttonr.pack()
        rframe3.pack()
        countr+=1
    else:
        frame2.pack_forget()
        rframe3.pack()    
        
###this is upload method that traverse all xlfiles and puts data to employee_personnel
def uploadData():
    prolabel=tk.Label(master=Uframe,text="Results processig",fg="black")
    prolabel.grid(row=3,column=3)    
    l1=[]
    for i in range(0,15):
        l1.append("")
    file_list=os.listdir(path)
    m=0
    st=""
    for j in file_list:
        wb = xlrd.open_workbook(str(path)+"\\"+str(j))
        sheet = wb.sheet_by_index(0)
        global k
        k=sheet.nrows
        global inputtxt 
        inputtxt= tk.Text(master=Uframe,height=30,width=120)
        inputtxt.grid(row=5,column=2)
        for i in range(9,k):
            if(len(str(sheet.cell_value(i,1))[:-2])==8):
                #mycursor.execute("select attendance from employee where id='{}'".format(sheet.cell_value(i,1)))
                if("PR" in str(sheet.cell_value(i,12)) or "PR" in str(sheet.cell_value(i,13))):
                    #print(i)
                    l1[i-9]+='1'
                    #up=str(mycursor.fetchone()[0])+'1'
                    #mycursor.execute("update employee set attendance='{}' where id='{}'".format(up,sheet.cell_value(i,1)))
                else:
                    #print(i)
                    l1[i-9]+='0'
                    #up=str(mycursor.fetchone()[0])+'0'
                    #mycursor.execute("update employee set attendance='{}' where id='{}'".format(up,sheet.cell_value(i,1)))

        m+=1
        st+=str(m)+"\t wait ur file"+str(j)+"just now uploaded\n"
        inputtxt.insert(tk.END,str(m)+st)
    prolabel.config(text="finished uploading")
    for i in range(9,k):
        mycursor.execute("update employee set attendance='{}' where id='{}'".format("".join(l1[i-9]),sheet.cell_value(i,1)))
        cnx.commit()


        
###function that generates slip
def generateData():
    global sheet
    for i in range(9,k):
        if(len(str(sheet.cell_value(i,1))[:-2])==8):
            mycursor.execute("select attendance from employee where id='{}'".format(int(sheet.cell_value(i,1))))
            attendance=str(mycursor.fetchone()[0])
            days=len(attendance)
            no_of_atd=str(attendance)
            no_of_atd=no_of_atd.count('1')
            no_of_abd=str(attendance).count('0')
            mycursor.execute("select basic from employee_personnel where id='{}'".format(int(sheet.cell_value(i,1))))
            basic=float(mycursor.fetchone()[0])
            hra=2000
            oa=2000
            mycursor.execute("select holidas from holidays2021 where month='{}'".format('jun'))
            holidas=int(mycursor.fetchone()[0])
            workingdays=days-holidas
            mycursor.execute("select leaves_left from employee_personnel where id='{}'".format(int(sheet.cell_value(i,1))))
            leaves_left=int(mycursor.fetchone()[0])
            if(workingdays>no_of_atd and leaves_left-(workingdays-no_of_atd)>0):
                leaves_left-=workingdays-no_of_atd
                mycursor.execute("update employee_personnel set leaves_left='{}' where id='{}'".format(leaves_left,int(sheet.cell_value(i,1)))) 
                cnx.commit()
                lop=0.00
            elif(workingdays>no_of_atd and leaves_left==0):
                lop=basic*((workingdays-no_of_atd)/workingdays)
                basic=basic*(no_of_atd/workingdays)
            else:
                lop=0.00 
            gross=basic+hra+oa
            it=basic*0.12
            pf=basic*0.10
            pt=basic*0.05
            net=gross-it-pf-pt
            mycursor.execute("update employee set no_of_atd='{}',gross='{}',net='{}',hra='{}',oa='{}',it='{}',pf='{}',pt='{}',lop='{}' where id='{}'".format(no_of_atd,gross,net,hra,oa,it,pf,pt,lop,sheet.cell_value(i,1)))
            cnx.commit()
    inputtxt.insert(tk.END,"the salary details have been caluculated\n")
            

    
#function that lets user to select the directory
def getDir():
    global path
    path= filedialog.askdirectory(initialdir="/", title="Select file")
    plabel=tk.Label(master=Uframe,text="("+path+")",fg="black")
    plabel.grid(row=2,column=3)
    

#funtion for viewing employee details
def ViewEmp():
    global countv
    if(countv==0):
        frame2.pack_forget()
        label=tk.Label(master=vframe3,text="you can View know",fg="black")
        label.pack()
        buttonv = tk.Button(master=vframe3,text="back",fg="black",command=vback_funtion)
        buttonv.pack()
        vframe3.pack()
        countv+=1
    else:
        frame2.pack_forget()
        vframe3.pack()
    
# funtion for update the data base with new xl file
def Update():
    global Ucount
    if(Ucount==0):
        frame2.pack_forget()
        ubutton=tk.Button(master=Uframe,text="ChooseDir",fg="black",command=getDir)
        ubutton.grid(row = 2, column = 2,padx = 10,pady=10,ipadx=100,ipady=5)
        gbutton=tk.Button(master=Uframe,text="Upload",fg="black",command=uploadData)
        gbutton.grid(row = 3, column = 2,padx = 10,pady=10,ipadx=100,ipady=5)
        buttonu = tk.Button(master=Uframe,text="back",fg="black",command=Uback_funtion)
        buttonu.grid(row = 8, column = 2,padx = 10,pady=10,ipadx=100,ipady=5)
        buttonu = tk.Button(master=Uframe,text="generate",fg="black",command=generateData)
        buttonu.grid(row = 6, column = 2,padx = 10,pady=10,ipadx=100,ipady=5)
        label2=tk.Label(master=Uframe,text="click here to generate sal details",fg="black")
        label2.grid(row = 6, column = 3,padx = 10,pady=10)
        buttonu = tk.Button(master=Uframe,text="send_to_mail",fg="black",command=mailData)
        buttonu.grid(row = 7, column = 2,padx = 10,pady=10,ipadx=100,ipady=5)
        label2=tk.Label(master=Uframe,text="click here to send mail",fg="black")
        label2.grid(row = 7, column = 3,padx = 10,pady=10)

        Uframe.pack()
        
        #cbutton=tk.Button(master=Uframe,text="ChooseDir",fg="black",command=getDir)
        #cbutton.grid(row = 3, column = 2,padx = 10,pady=10,ipadx=100,ipady=5)
        Ucount+=1
    else:
        frame2.pack_forget()
        Uframe.pack()
#funtion for generate page
def Genarate_funtion():
    global gcount
    if(gcount==0):
        Uframe.pack_forget()
        dbutton=tk.Button(master=Gframe,text="ViewDB",fg="black")
        dbutton.grid(row = 1, column = 0,padx = 10,pady=10,ipadx=100,ipady=5)
        gbbutton=tk.Button(master=Gframe,text="back",fg="black",command=gbBack)
        gbbutton.grid(row = 2, column = 0,padx = 10,pady=10,ipadx=100,ipady=5)
        Gframe.pack()
        gcount+=1
    else:
        Uframe.pack_forget()
        Gframe.pack()
    





#funtion for displaying page2   
def function():
    global count
    if(count==0):
        uname=entry.get()
        password=entry1.get()
        if(uname=="" or password==""):
            flabel=tk.Label(master=frame,text="details cant be empty",fg="black")
            flabel.grid(row=3,column = 1,padx = 10,pady=10)
            
        
        else:
            mycursor.execute("select id,password from admin where id='{}' and password='{}'".format(uname,password))
            details=mycursor.fetchone()
            if(not details):
                flabel=tk.Label(master=frame,text="invalid authentication",fg="black")
                flabel.grid(row=3,column = 1,padx = 10,pady=10)

            elif(str(details[0])==uname and str(details[1])==password):
                frame.pack_forget()
                label2=tk.Label(master=frame2,text="Hai "+uname,fg="black")
                label2.grid(row = 0, column = 1,padx = 10,pady=10)
                add_emp = tk.Button(master=frame2,text="Add_Emp", fg="black",command=AddEmp)
                add_emp.grid(row = 1, column = 0,padx = 10,pady=10,ipadx=100,ipady=5)
                rem_emp = tk.Button(master=frame2,text="Rem_Emp", fg="black",command=RemEmp)
                rem_emp.grid(row = 2, column = 0,padx = 10,pady=10,ipadx=100,ipady=5)
                view_emp = tk.Button(master=frame2,text="View_Emp", fg="black",command=ViewEmp)
                view_emp.grid(row = 1, column = 2,padx = 10,pady=10,ipadx=100,ipady=5)
                buttonU=tk.Button(master=frame2,text="update File",fg="black",command=Update)
                buttonU.grid(row = 2, column = 2,padx = 10,pady=10,ipadx=100,ipady=5)
                frame2.pack()
                button1 = tk.Button(master=frame2,text="back",fg="black",command=back_function)
                button1.grid(row = 3, column = 1,padx = 10,pady=10,ipadx=100,ipady=5)
                count+=1

    else:
        frame.pack_forget()
        frame2.pack()
#creating main window and frames for all required sub pages

window = tk.Tk()

frame = tk.Frame()
frame2 = tk.Frame()
aframe3=tk.Frame()
rframe3=tk.Frame()
vframe3=tk.Frame()
Uframe=tk.Frame()
Gframe=tk.Frame()

label = tk.Label(master=frame,text="User Name", fg="black")
label.grid(row = 0, column = 0,padx = 10,pady=10)

entry = tk.Entry(master=frame)
entry.grid(row = 0, column = 1,ipadx=50,ipady=5)


label1 = tk.Label(master=frame,text="Password", fg="black")
label1.grid(row = 1, column = 0,padx = 10,pady=10)


entry1 = tk.Entry(master=frame)
entry1.grid(row = 1, column = 1,ipadx=50,ipady=5)



count=0
counta=0
countr=0
countv=0
Ucount=0
gcount=0


    



button = tk.Button(master=frame,text="login",fg="black",command=function) ###this one login button goes to function
button.grid(row=2,column = 1,padx = 10,pady=10)
frame.pack()
window.mainloop()
cnx.close()



