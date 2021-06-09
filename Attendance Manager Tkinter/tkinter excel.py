import tkinter as tk
from tkinter import ttk
import openpyxl
from openpyxl import Workbook
from tkinter import messagebox

LARGEFONT = ("Verdana", 35)

class tkinterApp(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (StartPage, Page1):
            frame = F(container, self)

            self.frames[F] = frame

            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(StartPage)

    # to display the current frame passed as
    # parameter
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        def logfun():
            p = y2.get()
            if (p == "Passkey"):
                print('p')
                messagebox.showinfo("showinfo", "Logged in successfully")
                controller.show_frame(Page1)
            else:
                print("FAILED")
                messagebox.showerror("showerror", "Invalid Password")

        x1 = tk.Label(self, text="Username")
        x2 = tk.Label(self, text="Password")

        y1 = tk.Entry(self)
        y2 = tk.Entry(self)

        x1.grid(row=0, column=0)
        x2.grid(row=1, column=0)

        y1.grid(row=0, column=1)
        y2.grid(row=1, column=1)

        login = tk.Button(self, text="Login", command=logfun)
        login.grid(row=4, column=1)


# second window frame page1
class Page1(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        def addfun():
            ex1 = openpyxl.load_workbook('students.xlsx')
            sheet = ex1.active

            a = e1.get()
            b = e2.get()
            c = e3.get()
            d = e4.get()
            print(a, b, c, d)
            sheet.append((a, b, c, d))
            ex1.save('students.xlsx')
            print("Data Added")
            messagebox.showinfo("showinfo", "Data Added")

        def delfun():
            ex1 = openpyxl.load_workbook('students.xlsx')
            sheet = ex1.active
            a = e1.get()

            find = sheet['A']
            count = 0
            for i in find:
                count = count + 1
                if (i.value == a):
                    print(i.value)
                    print(i)
                    sheet.delete_rows(count)
                    ex1.save('students.xlsx')
            print("Data Deleted")
            messagebox.showinfo("showinfo", "Data Deleted")
            count = 0

        def exitfun():
            reply = messagebox.askquestion("Exit?", "Do you want to exit", icon='warning')
            if reply == "yes":
                self.destroy()



        l1 = tk.Label(self, text="Roll No ", font=('calibri', 14, 'bold'))
        l2 = tk.Label(self, text="Name ", font=('calibri', 14, 'bold'))
        l3 = tk.Label(self, text="Days Attended", font=('calibri', 14, 'bold'))
        l4 = tk.Label(self, text="Days Bunked", font=('calibri', 14, 'bold'))
        l5 = tk.Label(self, text="Attendance Management for FYDS", font=('calibri', 18, 'bold'), bd=30,
                      bg='light slate grey')

        e1 = tk.Entry(self)
        e2 = tk.Entry(self)
        e3 = tk.Entry(self)
        e4 = tk.Entry(self)

        l1.grid(row=1, column=0)
        l2.grid(row=2, column=0)
        l3.grid(row=3, column=0)
        l4.grid(row=4, column=0)
        l5.grid(row=0, column=1)

        e1.grid(row=1, column=1)
        e2.grid(row=2, column=1)
        e3.grid(row=3, column=1)
        e4.grid(row=4, column=1)
        addb = tk.Button(self, text="Add", font=('calibri', 10, 'bold'), bd=10, bg='grey', command=addfun)
        delb = tk.Button(self, text="Delete", font=('calibri', 10, 'bold'), bd=10, bg='grey', command=delfun)
        exitb = tk.Button(self, text="Exit", font=('calibri', 10, 'bold'), bd=10, bg='grey', command=exitfun)

        addb.grid(row=10, column=0)
        delb.grid(row=10, column=1)
        exitb.grid(row=10, column=4)

app = tkinterApp()
app.mainloop()
