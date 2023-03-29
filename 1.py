from tkinter import *
import mysql.connector
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import qrcode
from PIL import ImageTk,Image
import pandas as pd
from pyqrcode import create
import pyqrcode
import smtplib, ssl

root = Tk()
root.geometry("900x600")
root.configure(bg="lightyellow")

def insert():
    Id=e1.get()
    name=e2.get()
    salary=e3.get()
    address=e4.get()
    
    mydb=mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="company"
        )
    mycursor = mydb.cursor()
    sql="insert into employee(ID,NAME,SALARY,ADDRESS) values (%s,%s,%s,%s)"
    val=(Id,name,salary,address)
    mycursor.execute(sql,val)
    mydb.commit()
    messagebox.showinfo("Record","Insert Successfully..!!")

def csv():

    mydb = mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="company")

    mycursor = mydb.cursor()
    sql = "SELECT * FROM employee"
    mycursor.execute(sql)
    myresult = mycursor.fetchall()
    all_id=[]
    all_name=[]
    all_salary=[]
    all_address=[]
    for Id,Name,Salary,Address in myresult:
        all_id.append(Id)
        all_name.append(Name)
        all_salary.append(Salary)
        all_address.append(Address)

    dic={'Emp_id':all_id,'Emp_name':all_name,'Emp_salary':all_salary,'Emp_address':all_address}
    df=pd.DataFrame (dic)
    df_csv=df.to_csv('D:/PYTHON/raj1.csv',index=False)
    messagebox.showinfo("Record","CSV File Generate Successfully....!!")

    


def update():

    root1=Tk()
    root1.geometry("600x600")
    root1.configure(bg="lightyellow")

    def update1():

        Id=e1.get()
        name=e2.get()
        salary=e3.get()
        address=e4.get()
        
        
        connection = mysql.connector.connect(host='localhost',
                                             database='company',
                                             user='root',
                                             password='')

        cur = connection.cursor()
        
        sql="UPDATE employee SET NAME=%s,SALARY=%s,ADDRESS=%s  WHERE ID=%s"
        val=(name,salary,address,Id)
            
        cur.execute(sql,val)
        connection.commit()
        messagebox.showinfo("Record","Updated Successfully....!!")
    l1=Label(root1,text="Update Employee Management",font=("arial,bold,24"))
    l1.place(x=170,y=20)

    l2=Label(root1,text="Employee Id :",font=("arial,2"))
    l2.place(x=80,y=90)

    e1=Entry(root1,bd=6,width=30)
    e1.place(x=280,y=90)

    l3=Label(root1,text="Employee Name :",font=("arial,2"))
    l3.place(x=80,y=160)

    e2=Entry(root1,bd=6,width=30)
    e2.place(x=280,y=160)

    l4=Label(root1,text="Employee Salary :",font=("arial,2"))
    l4.place(x=80,y=230)

    e3=Entry(root1,bd=6,width=30)
    e3.place(x=280,y=230)

    l5=Label(root1,text="Employee Address :",font=("arial,2"))
    l5.place(x=80,y=300)

    e4=Entry(root1,bd=6,width=30)
    e4.place(x=280,y=300)

    b1=Button(root1,text="UPDATE",width=15,relief = GROOVE,command=update1)
    b1.place(x=200,y=380)

    root1.mainloop()
    
def view():

    root3=Tk()
    root3.geometry("600x600")
    root3.configure(bg="lightyellow")

    table_frame=Frame(root3,bd=5,relief=RIDGE,bg="lightblue")
    table_frame.place(x=0,y=0,width=500,height=500)

    scroll_x=Scrollbar(table_frame,orient=HORIZONTAL)
    scroll_y=Scrollbar(table_frame,orient=VERTICAL)

    item_table = ttk.Treeview(table_frame,column=("ID","NAME","SALARY","ADDRESS"),xscrollcommand=scroll_x.set,yscrollcommand=scroll_y.set)
    

    scroll_x.pack(side=BOTTOM,fill=X)
    scroll_y.pack(side=RIGHT,fill=Y)

    scroll_x.config(command=item_table.xview)
    scroll_y.config(command=item_table.yview)

    item_table.heading("ID",text="Id")
    item_table.heading("NAME",text="Name")
    item_table.heading("SALARY",text="Salary")
    item_table.heading("ADDRESS",text="Address")
    

    item_table['show'] = 'headings'
    item_table.column("ID",width=100)
    item_table.column("NAME",width=100)
    item_table.column("SALARY",width=100)
    item_table.column("ADDRESS",width=100)
   
    item_table.pack(fill=BOTH,expand=1)


    


    mydb = mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="company")

    mycursor = mydb.cursor()
    sql = "SELECT * FROM employee"
    mycursor.execute(sql)
    myresult = mycursor.fetchall()
    if len(myresult)!=0:
        #stu_table.delete(stu_table.get_children())
        for row in myresult:
            item_table.insert('',END,values=row)
        mydb.commit()
    mydb.close()
    
    root3.mainloop()
def email():

    sender_email = "sender@xyz.com"
    receiver_email = "receiver@xyz.com"
    message = """\
        Subject: It Worked!

        Simple Text email from your Python Script."""

    port = 465  
    app_password = input("Enter Password: ")

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login("sender@xyz.com", app_password)
        server.sendmail(sender_email, receiver_email, message)

    


def pdf():
      

    stmt = "Select * from employee"
    conn = mysql.connector.connect(user="root",password="",database="company")
    cursor = conn.cursor()
    cursor.execute(stmt)
    row = cursor.fetchall()
    for i in row:
        try:
            with open(i[0], 'wb') as outfile:
                outfile.write(i[1])
                outfile.close()
                print("Filename Saved as: " + i[0])
        except:
            pass
    
        
def delete():

    root2=Tk()
    root2.geometry("600x600")
    root2.configure(bg="lightyellow")

    def delete1():

        Id=e1.get()

        mydb=mysql.connector.connect(
            host="localhost",
            user="root",
            password="",
            database="company"
            )
        mycursor=mydb.cursor()
        sql="DELETE FROM employee WHERE ID=%s"
        val=(Id)
        mycursor.execute(sql,(val,))
        mydb.commit()
        messagebox.showinfo("Record","delete Successfully...!!")
        mydb.close()


    l1=Label(root2,text="Delete Employee Management",font=("arial,bold,24"))
    l1.place(x=170,y=20)

    l2=Label(root2,text="Employee Id :",font=("arial,2"))
    l2.place(x=80,y=90)

    e1=Entry(root2,bd=6,width=30)
    e1.place(x=280,y=90)

    b1=Button(root2,text="DELETE",width=15,relief = GROOVE,command=delete1)
    b1.place(x=200,y=200)


    
    root2.mainloop()

   

        

    
    
def gene():
    
    
    global my_image
    #root4.geometry("500x500")
    #root4=Tk()
    #root4.configure(bg="lightyellow")
    l1=Label(Frame1)
    l1.place(x=100,y=70)
    
    Id=e1.get()
    name=e2.get()
    salary=e3.get()
    address=e4.get()
    qrdata=pyqrcode.create(f"Employee Id:{Id}\n Employee Name:{name}\n Employee Salary:{salary}\n Employee Address:{address}")
    my1_qr=qrdata.xbm(scale=2)
    my_image=tk.BitmapImage(data=my1_qr)
    l1.config(image=my_image)
    #qr_code=qrcode.make(qrdata)
    #qr_code.save('image2.py')
    #print(qr_code)
    #root4.mainloop()


    
Frame1=Frame(root,height=250,width=350,bg="deep sky blue",bd=1,relief=FLAT)    
Frame1.place(x=500,y=90)



ll1=Label(Frame1,text="QR Generate Here...!!",font=("arial,6,bold"),bg="lightblue",fg="black")
ll1.place(x=70,y=30)

l1=Label(root,text="Employee Management",font=("arial,bold,24"),bg="lightyellow")
l1.place(x=170,y=20)

l2=Label(root,text="Employee Id :",font=("arial,2,bold"),bg="lightyellow")
l2.place(x=80,y=90)

e1=Entry(root,bd=6,width=30)
e1.place(x=280,y=90)

l3=Label(root,text="Employee Name :",font=("arial,2,bold"),bg="lightyellow")
l3.place(x=80,y=160)

e2=Entry(root,bd=6,width=30)
e2.place(x=280,y=160)

l4=Label(root,text="Employee Salary :",font=("arial,2,bold"),bg="lightyellow")
l4.place(x=80,y=230)

e3=Entry(root,bd=6,width=30)
e3.place(x=280,y=230)

l5=Label(root,text="Employee Address :",font=("arial,2,bold"),bg="lightyellow")
l5.place(x=80,y=300)

e4=Entry(root,bd=6,width=30)
e4.place(x=280,y=300)

b1=Button(root,text="INSERT",width=15,relief = GROOVE,bg="#16a085", fg="white",bd=2,font=("Calibri", 10, "bold"),command=insert)
b1.place(x=30,y=380)

b2=Button(root,text="UPDATE",width=15,relief = GROOVE,bg="#f39c12", fg="white",bd=2,font=("Calibri", 10, "bold"),command=update)
b2.place(x=170,y=380)

b3=Button(root,text="DELETE",width=15,relief = GROOVE,bg="brown4", fg="white",bd=2,font=("Calibri", 10, "bold"),command=delete)
b3.place(x=310,y=380)

b4=Button(root,text="VIEW",width=15,relief = GROOVE,bg="purple4", fg="white",bd=2,font=("Calibri", 10, "bold"),command=view)
b4.place(x=450,y=380)

b5=Button(root,text= "QR Generator",width=20,relief= RAISED,bg="dark green", fg="white",bd=2,font=("Calibri", 10, "bold"),command=gene)
b5.place(x=30,y=450)

b6=Button(root,text= "Export to CSV",width=20,relief= RAISED,bg="dark green", fg="white",bd=2,font=("Calibri", 10, "bold"),command=csv)
b6.place(x=200,y=450)

b7=Button(root,text= "Export to PDF",width=20,relief= RAISED,bg="dark green", fg="white",bd=2,font=("Calibri", 10, "bold"),command=pdf)
b7.place(x=370,y=450)

b8=Button(root,text="EMAIL",width=15,relief = GROOVE,bg="green3", fg="white",bd=2,font=("Calibri", 10, "bold"),command=email)
b8.place(x=590,y=380)






root.mainloop()
