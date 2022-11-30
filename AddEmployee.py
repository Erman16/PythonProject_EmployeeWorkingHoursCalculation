import json
from tkinter import *
import tkinter.font as tkFont
from unicode_tr import unicode_tr
import shutil
import os


def forDailyAttendanceChecker():
    f = open('Daily Attendance Checker/Employee_List/Employees.json', encoding="UTF-8")
    data = json.load(f)
    f.close()
    return data


def forMonthlyReportCreator():
    f = open('Monthly Report Creator/Employee_List/Employees.json', encoding="UTF-8")
    data = json.load(f)
    f.close()
    return data


ForDailyAttendanceDict = forDailyAttendanceChecker()
ForMonthlyReportDict = forMonthlyReportCreator()


def PrintToExcel(MyName, MyID):
    ForDailyAttendanceDict[MyID] = MyName
    with open("NewEmployee.json", 'w', encoding='UTF-8') as json_file:
        json.dump(ForDailyAttendanceDict, json_file, ensure_ascii=False)

    src_path = r"NewEmployee.json"
    dst_path1 = r"Daily Attendance Checker/Employee_List/Employees.json"
    shutil.copy(src_path, dst_path1)

    dst_path2 = r"Monthly Report Creator/Employee_List/Employees.json"
    shutil.copy(src_path, dst_path2)

    os.remove("NewEmployee.json")

    root = Tk()
    Label(root, text="Başarılı", font=('Arial 12')).pack(pady=20, padx=20)
    root.mainloop()


def DeleteFromJson(MyID):
    try:
        ForDailyAttendanceDict.pop(MyID)
        with open('NewEmployee.json', 'w', encoding='UTF-8') as json_file:
            json.dump(ForDailyAttendanceDict, json_file, ensure_ascii=False)

        src_path = r"NewEmployee.json"
        dst_path1 = r"Daily Attendance Checker/Employee_List/Employees.json"
        shutil.copy(src_path, dst_path1)

        dst_path2 = r"Monthly Report Creator/Employee_List/Employees.json"
        shutil.copy(src_path, dst_path2)

        os.remove("NewEmployee.json")
        return True
    except:
        return False


class control2:
    control = 0


def AddEmployeeToJson():
    a = control2()

    root = Tk()
    root.geometry("600x420")

    myLabel = Label(root, text="Add Employee", padx=50, pady=10, font=('Arial 12'))
    myLabel.pack()

    myCanvas = Canvas(root, borderwidth=0, highlightthickness=0)

    myLabel = Label(myCanvas, text="Name & Surname: ", padx=50, pady=10, font=('Arial 12'))
    myLabel.grid(row=0, column=0)

    entry_text = StringVar()
    entry_text.set("")

    def to_uppercase(*args):
        varstr = entry_text.get()
        varstr = unicode_tr(varstr).upper()
        entry_text.set(varstr)

    e = Entry(myCanvas, textvariable=entry_text, width=20, font=('Arial 12'))
    e.grid(row=0, column=1)
    entry_text.trace_add('write', to_uppercase)

    myLabel2 = Label(myCanvas, text="ID: ", padx=50, pady=10, font=('Arial 12'))
    myLabel2.grid(row=1, column=0)

    entry_text2 = StringVar()
    entry_text2.set("")

    def to_uppercase(*args):
        varstr = entry_text2.get()
        varstr = unicode_tr(varstr).upper()
        entry_text2.set(varstr)

    e2 = Entry(myCanvas, textvariable=entry_text2, width=20, font=('Arial 12'))
    e2.grid(row=1, column=1)
    entry_text2.trace_add('write', to_uppercase)

    Label(myCanvas, padx=50, pady=10, font=('Arial 12')).grid(row=2, column=0)

    def goBack():
        root.destroy()
        ChooseProgram()

    def Approve():
        root.destroy()
        a.control = 1

    myButton = Button(myCanvas, text="Onayla", padx=60, pady=20, command=Approve, font=('Arial 12'))
    myButton.grid(row=3, column=1, pady=20)
    myButton = Button(myCanvas, text="Geri Dön", padx=60, pady=20, command=lambda: goBack(), font=('Arial 12'))
    myButton.grid(row=3, column=0, pady=20)

    myCanvas.pack(pady=60)

    root.mainloop()

    MyName = entry_text.get()
    MyID = entry_text2.get()

    if a.control == 1:
        PrintToExcel(MyName, MyID)


class control3:
    control = 0


def DeleteEmployee():
    a = control3()

    root = Tk()
    root.geometry("600x420")

    myLabel = Label(root, text="Remove Employee", padx=50, pady=10, font=('Arial 12'))
    myLabel.pack()

    myCanvas = Canvas(root, borderwidth=0, highlightthickness=0)

    myLabel2 = Label(myCanvas, text="ID: ", padx=50, pady=10, font=('Arial 12'))
    myLabel2.grid(row=0, column=0)

    entry_text2 = StringVar()
    entry_text2.set("")

    def to_uppercase(*args):
        varstr = entry_text2.get()
        varstr = unicode_tr(varstr).upper()
        entry_text2.set(varstr)

    e2 = Entry(myCanvas, textvariable=entry_text2, width=20, font=('Arial 12'))
    e2.grid(row=0, column=1)
    entry_text2.trace_add('write', to_uppercase)

    Label(myCanvas, padx=50, pady=10, font=('Arial 12')).grid(row=1, column=0)

    def goBack():
        root.destroy()
        ChooseProgram()

    def Approve():
        root.destroy()
        a.control = 1

    myButton = Button(myCanvas, text="Onayla", padx=60, pady=20, command=Approve, font=('Arial 12'))
    myButton.grid(row=2, column=1, pady=20)
    myButton = Button(myCanvas, text="Geri Dön", padx=60, pady=20, command=lambda: goBack(), font=('Arial 12'))
    myButton.grid(row=2, column=0, pady=20)

    myCanvas.pack(pady=60)

    root.mainloop()

    MyID = entry_text2.get()

    returnVal = DeleteFromJson(MyID)

    if a.control == 1:
        if returnVal:
            root = Tk()
            Label(root, text="Deleted Successfully", font=('Arial 12')).pack(pady=20, padx=20)
            root.mainloop()

        else:
            root = Tk()
            Label(root, text="Some Error Occurred!!!", font=('Arial 12')).pack(pady=20, padx=20)
            root.mainloop()


class control:
    control = 0


def ChooseProgram():
    a = control()

    root = Tk()
    helv12 = tkFont.Font(family='Helvetica', size=12)

    myLabel2 = Label(root, text="Please select", padx=50, pady=10, font=('Helvetica 12'))
    myLabel2.pack(pady=(40, 5))

    items2 = ["Add Employee", "Remove Employee"]
    entry_text2 = StringVar()
    entry_text2.set(items2[0])
    menu2 = OptionMenu(root, entry_text2, *items2)
    menu2.config(font=helv12)
    menu2.pack(ipadx=30)
    menuString2 = root.nametowidget(menu2.menuname)  # Get menu widget.
    menuString2.config(font=helv12)

    myLabel = Label(root, padx=500, pady=10)
    myLabel.pack()

    def Approve():
        a.control = 1
        root.destroy()

    myButton = Button(root, text="Onayla", padx=60, pady=20, command=Approve, font=('Helvetica 12'))
    myButton.pack(pady=20)

    root.mainloop()

    if a.control == 1:
        if entry_text2.get() == "Add Employee":
            AddEmployeeToJson()
        else:
            DeleteEmployee()


ChooseProgram()
