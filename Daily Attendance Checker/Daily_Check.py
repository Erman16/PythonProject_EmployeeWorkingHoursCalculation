"""
Designed and Codded by: Erman Yalçın
Linkedln: https://www.linkedin.com/in/ermanyalcin/
"""
import datetime
import json
import numpy as np
import openpyxl
from tkinter import *
import tkinter as tk
from tkcalendar import Calendar
from tkinter import ttk, font
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import os.path
from babel.dates import format_datetime
import babel.numbers
import time
import os


# -------- Get Employees ----------

def getEmployees():
    f = open('Employee_List/Employees.json', encoding="utf8")
    data = json.load(f)
    f.close()
    return data


EmplooyeesDict = getEmployees()


# ---------------------------------
# ------ Get Date From User -------

class getDays:
    day = ""
    month = ""
    year = ""


def getDateFromUser(day, month, year):
    root = Tk()
    root.title("Select Date")
    root.geometry("400x400")

    cal = Calendar(root, selectmode='day', year=int(year), month=int(month), day=int(day), locale="en")

    cal.pack(pady=20)

    def grad_date():
        getDate = cal.get_date()
        getDate = getDate.split("/")
        getDay = getDate[0]
        if int(getDay) < 10:
            getDay = "0" + str(int(getDay))
        getMounth = getDate[1]
        if int(getMounth) < 10:
            getMounth = "0" + str(int(getMounth))
        getYear = getDate[2]
        root.destroy()
        getDays.day = getDay
        getDays.month = getMounth
        getDays.year = getYear

    myFont = font.Font(family='Helvetica', size=14)
    button = Button(root, text="Select Date", command=grad_date, font=myFont)
    button.pack(pady=20)
    root.mainloop()

    getDay = getDays.day
    getMounth = getDays.month
    getYear = getDays.year
    if int(getDay) < 10:
        getDay = "0" + str(int(getDay))
    if int(getMounth) < 10:
        getMounth = "0" + str(int(getMounth))

    return getDay, getMounth, getYear


myDate = datetime.datetime.now()

yearstart = str(myDate.year)
mymonth = str(myDate.month)
myday = str(myDate.day)

if int(myday) < 10:
    myday = "0" + myday
if int(mymonth) < 10:
    mymonth = "0" + mymonth

tarih = myday + "." + mymonth + "." + yearstart

selectedDate = getDateFromUser(myday, mymonth, yearstart)


# ---------------------------------
# -------- Take Attandence --------
class closeClass:
    closeValue = False


p1 = closeClass()
p1.closeValue = False


def main():
    notSelectedOnes = []
    root = tk.Tk()
    root.title('Select File')
    root.geometry('800x450')

    mydate = format_datetime(datetime.datetime.now(), locale='en').split(" ")
    for i in range(len(mydate)):
        mydate[i] = mydate[i].replace(",", "")
    NameDate = mydate[0] + " " + mydate[2]

    try:
        os.mkdir('Data/' + NameDate)
    except FileExistsError:
        pass

    try:
        os.mkdir('Result/' + NameDate)
    except FileExistsError:
        pass

    for i in EmplooyeesDict.keys():
        check = os.path.isfile('Data/' + NameDate + "/" + EmplooyeesDict[i] + ".txt")
        if not check:
            notSelectedOnes.append(EmplooyeesDict[i])

    label = tk.Label(root, text="Persons whose attendance has not been taken", font=("Arial", 14))
    label.pack(pady=20)

    myText = ""
    for i in notSelectedOnes:
        myText = myText + i + "\n"

    T = tk.Text(root, height=15, width=70)

    T.pack()

    T.insert(tk.END, myText)

    def take_attandence(NameDate):
        closeForNow()
        addedOnes = []
        for i in notSelectedOnes:
            root = tk.Tk()
            root.title("Attandence For: " + i)
            root.geometry('800x450')

            t = time.localtime()
            attandenceTime = time.strftime("%H:%M", t)

            myDate = getDays.day + "." + getDays.month + "." + getDays.year
            name = i

            def here():
                try:
                    saveF = open("Data/" + NameDate + "/" + name + ".txt", "r")
                    oldInfo = saveF.read()
                    saveF.close()
                    saveF = open("Data/" + NameDate + "/" + name + ".txt", "w")
                    addedOnes.append(name)
                    saveF.write(oldInfo + "\n" + myDate + ": " + attandenceTime)
                    saveF.close()
                except:
                    saveF = open("Data/" + NameDate + "/" + name + ".txt", "w")
                    addedOnes.append(name)
                    saveF.write(myDate + "= " + attandenceTime)
                    saveF.close()

                try:
                    for i in addedOnes:
                        notSelectedOnes.remove(i)
                except:
                    pass

                root.destroy()

            def notHere():
                try:
                    saveF = open("Data/" + NameDate + "/" + name + ".txt", "r")
                    oldInfo = saveF.read()
                    saveF.close()
                    saveF = open("Data/" + NameDate + "/" + name + ".txt", "w")
                    addedOnes.append(name)
                    saveF.write(oldInfo + "\n" + myDate + ": " + "-")
                    saveF.close()
                except:
                    saveF = open("Data/" + NameDate + "/" + name + ".txt", "w")
                    addedOnes.append(name)
                    saveF.write(myDate + "= " + "-")
                    saveF.close()

                root.destroy()

            label = tk.Label(root, text="Name:" + i + "\n" + "Attandence Time: " + attandenceTime, font=("Arial", 14))
            label.pack(pady=20)
            save_button = ttk.Button(root, text='Here', command=lambda: here())
            save_button.pack()
            save_button = ttk.Button(root, text='Not Here', command=lambda: notHere())
            save_button.pack()

            root.mainloop()
            break

    def close():
        root.destroy()
        p1.closeValue = True

    def closeForNow():
        root.destroy()

    open_button = ttk.Button(root, text='Take Attendance', command=lambda: take_attandence(NameDate))
    open_button.pack(expand=True)

    close_button = ttk.Button(root, text='Close', command=close)
    close_button.pack(expand=True)

    root.mainloop()


while True:
    main()
    if p1.closeValue:
        break

