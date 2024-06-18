from tkinter import *
from customtkinter import * 
import customtkinter as ttk
from tkinter import font, StringVar, messagebox
import pymysql

# Database connection parameters
hostname1 = 'localhost'
porta = 10011
username1 ='root'
passwd1 = 'root'
database1='local'

# GUI colors and fonts
mainFont = 'Heebo'
primary = '#386fcf'
secondary = '#adb5bd'
success = '#02b875'
info = '#17a2b8'
warning = "#f0ad4e"
danger = '#d9534f'
light = '#f8f9fa'
dark = '#343a40'

class LogIn:
    def __init__(self):
        self.root = CTk()
        self.root.geometry('800x600+550+200')
        self.root.title('ניהול עובדים, פותח על ידי °Anis Zakaria Muhammad°')
        self.root.configure(bd=0, fg_color=light)
        self.root.iconbitmap('safety.ico')
        self.root.minsize(True, True)
        self.fname_var = StringVar()
        self.lname_var = StringVar()
        self.email_var = StringVar()
        self.password_var = StringVar()

        self.signUpWedget()

    def signUpWedget(self):
        
        container_frame = CTkFrame(master=self.root, width=700, height=600)
        container_frame.place(relx=0.5, rely=0.5, anchor=CENTER)
        container_frame.pack_propagate(False)

        h1_label = CTkLabel(master=container_frame, text='דף הרשמה', font=(mainFont, 40,'bold'), bg_color=primary, fg_color=dark)
        h1_label.grid(row=0, column=0, columnspan=2, pady=10)

        def inputs(label_text,rows,textvariable):
            label = CTkLabel(master=container_frame, text=label_text, font=(mainFont, 17,'bold'), fg_color=primary, bg_color=light)
            label.grid(row=rows, column=1, pady=5, padx=10, sticky=E)
            entry = CTkEntry(master=container_frame, textvariable=textvariable, fg_color=dark, justify='center', font=(mainFont, 12))
            entry.grid(row=rows, column=0, pady=5, padx=10, sticky=W+E)
        inputs(label_text='שם',rows=1,textvariable=self.fname_var)
        inputs(label_text='שם משפחה',rows=2,textvariable=self.lname_var)
        inputs(label_text='דוא"ל',rows=3,textvariable=self.email_var)
        inputs(label_text='סיסמה',rows=4,textvariable=self.password_var)
        signup_button = CTkButton(master=container_frame, text='הרשמה', font=(mainFont, 17,'bold'), fg_color=success, bg_color=dark, width=150)
        signup_button.grid(row=5, column=0, columnspan=2, pady=10)

    def destroy(self):
        self.root.destroy()

    def login(self):
        email = self.email_var.get()
        user = self.check_existing_email(email)
        if user:
            stored_password = user[1]
            password = self.password_var.get()
            if password == stored_password:
                messagebox.showinfo("تسجيل الدخول", f"{email}  مرحبًا بك مرة أخرى\n")
                self.destroy()
            else:
                messagebox.showerror("خطأ", "كلمة المرور غير صحيحة.")
        else:
            messagebox.showerror("خطأ", "البريد الإلكتروني غير مسجل.")

    def check_existing_email(self, email):
        try:
            connection = pymysql.connect(host=hostname1, port=porta, user=username1, passwd=passwd1,
                                        database=database1)
            with connection.cursor() as cursor:
                query = "SELECT UserName, password FROM users WHERE UserName = %s"
                cursor.execute(query, (email,))
                user = cursor.fetchone()
            connection.close()
            return user
        except pymysql.Error as e:
            print("Error:", e)
            return None

    def sign_up(self):
        new_fname = self.fname_var.get()
        new_lname = self.lname_var.get()
        new_email = self.email_var.get()
        new_password = self.password_var.get()
        if new_fname == '' or new_lname == '' or new_email == '' or new_password == '':
            messagebox.showerror('Error', 'الرجاء إدخال كافة التفاصيل')
            return

        try:
            connection = pymysql.connect(host=hostname1, port=porta, user=username1, passwd=passwd1,
                                        database=database1)
            with connection.cursor() as cursor:
                query = "INSERT INTO users (FirstName, LastName, UserName, password) VALUES (%s, %s, %s, %s)"
                cursor.execute(query, (new_fname, new_lname, new_email, new_password))
            connection.commit()
            connection.close()
            messagebox.showinfo("تم بنجاح", "تم التسجيل بنجاح!")
        except pymysql.Error as e:
            print("Error:", e)
            messagebox.showerror("خطأ", "فشل التسجيل!")

if __name__ == '__main__':
    LogIn().root.mainloop()
