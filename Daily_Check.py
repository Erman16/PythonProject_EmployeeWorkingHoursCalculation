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
from tktimepicker import SpinTimePickerModern
from tktimepicker import constants
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

myJobStart = ""
myJobEnd = ""


def main():
    for i in EmplooyeesDict:
        def getTimes():
            JobStart = str(time_picker.hours()) + "." + str(time_picker.minutes())
            JobEnd = str(time_picker2.hours()) + "." + str(time_picker2.minutes())
            try:
                os.mkdir('Daily_Attendance_Checker/Data/' + selectedDate[2])
            except FileExistsError:
                pass

            check = os.path.isfile('Daily_Attendance_Checker/Data/' + selectedDate[2] + "/" +
                                   EmplooyeesDict[i] + ".txt")
            if not check:
                file = open('Daily_Attendance_Checker/Data/' + selectedDate[2] + "/" +
                            EmplooyeesDict[i] + ".txt", "w")
                file.write(selectedDate[1] + "_" + selectedDate[0] + ": " + str(JobStart) + "-" + str(JobEnd) + "\n")
            else:
                file = open('Daily_Attendance_Checker/Data/' + selectedDate[2] + "/" +
                            EmplooyeesDict[i] + ".txt", "a")
                file.write(selectedDate[1] + "_" + selectedDate[0] + ": " + str(JobStart) + "-" + str(JobEnd) + "\n")
            root.quit()

            file.close()
        root = Tk()
        myLabel = Label(root, text=EmplooyeesDict[i], padx=50, pady=50, font='Arial 12')
        myLabel.pack()

        myLabel = Label(root, text="Job Start Time: ", padx=20, pady=20, font='Arial 12')
        myLabel.pack()

        time_picker = SpinTimePickerModern(root)
        time_picker.addAll(constants.HOURS24)  # adds hours clock, minutes and period
        time_picker.configureAll(bg="#404040", height=1, fg="#ffffff", font=("Times", 16), hoverbg="#404040",
                                 hovercolor="#d73333", clickedbg="#2e2d2d", clickedcolor="#d73333")
        time_picker.configure_separator(bg="#404040", fg="#ffffff")
        time_picker.set24Hrs(7)
        time_picker.setMins(0)
        time_picker.pack()

        myLabel = Label(root, text="Job End Time: ", padx=20, pady=20, font='Arial 12')
        myLabel.pack()

        time_picker2 = SpinTimePickerModern(root)
        time_picker2.addAll(constants.HOURS24)  # adds hours clock, minutes and period
        time_picker2.configureAll(bg="#404040", height=1, fg="#ffffff", font=("Times", 16), hoverbg="#404040",
                                  hovercolor="#d73333", clickedbg="#2e2d2d", clickedcolor="#d73333")
        time_picker2.configure_separator(bg="#404040", fg="#ffffff")
        time_picker2.set24Hrs(17)
        time_picker2.setMins(0)
        time_picker2.pack()

        myLabel = Label(root, "", padx=500, pady=10)
        myLabel.pack()
        myButton = Button(root, text="Onayla", padx=60, pady=20, command=lambda: getTimes(), font='Arial 12')
        myButton.pack()
        root.mainloop()
        root.destroy()


main()
