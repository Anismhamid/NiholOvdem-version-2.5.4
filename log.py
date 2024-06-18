from tkinter import StringVar, messagebox
import pymysql
import customtkinter as ctk
from tkinter import *

mainFont = font='heebo'
blueColor = '#3d3dff'
mainColor = '#fc8653'
secondColor = '#182929'
SecondMainColor = '#2F4F4F'
buttonColor = '#1eac99'
redColor = '#ff4336'
whiteColor = "#ffffff"
greenColor = '#64ff4f'
underButtonColor = '#09795f'
blackColor = "black"

hostname1 = 'localhost'
porta = 10011
username1 ='root'
passwd1 = 'root'
database1='local'

class Login:
    def __init__(self):
        super().__init__()
        root = ctk.CTk()
        root.geometry('400x600+600+100')
        root.title('ניהול עוהדים, Develobed by °Anis Zkaria Mhamid°')
        root.config(bd=0,bg=SecondMainColor)
        root.iconbitmap('C:\\NiholOvdem-version-2.5.4\\safety.ico')
        root.minsize(True,True)
        root.config(relief=SUNKEN)

        img = PhotoImage(file='C:\\NiholOvdem-version-2.5.4\\gallery\\sme22.png')
        im = Label(master=root,image=img)
        im.place(x=0,y=0,relheight=1,relwidth=1)
        # تعريف المتغيرات والعناصر البصرية هنا
        Email_Var = StringVar()
        passowrd_Var = StringVar()



        def login(self):
            email = Email_Var.get()
            password = passowrd_Var.get()
            user = self.cooection_(email, password)
            if user:
                messagebox.showinfo("تسجيل الدخول", "تم تسجيل الدخول بنجاح!")
                # هنا يمكنك فتح النافذة الرئيسية للتطبيق أو أي عمليات أخرى بعد تسجيل الدخول بنجاح
            else:
                messagebox.showerror("خطأ", "البريد الإلكتروني أو كلمة المرور غير صحيحة.")

            # استدعاء دالة login() عند النقر على زر تسجيل الدخول
            signInButton = ctk.CTkButton(master=root, text='Sign In', command=self.login, fg_color=buttonColor, bg_color=mainColor, text_color=blackColor, border_width=1, border_color=mainColor)
            
            root.mainloop()


if __name__=='__main__':
    Login()