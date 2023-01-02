import tkinter as tk
from tkinter import ttk
import os.path

"""
Welcome, this project include 2 different programs that can work separately but uses each other's data.
Program 1 Daily Attendance Checker: In this Program one employee takes attendance and these attendances saved to the
data folder.
Program 2 Monthly Report Creator: In this Program Data that collected in first program used for creating an Excel file
that shows work hours for every employee.

Note: You can run every program individually or Run This main program and select the program that you want to run
"""


def runDailyAttendanceChecker():
    os.system('python Daily_Check.py')


def runMonthlyReportCreator():
    os.system('python Monthly_Report.py')


def runAddEmploye():
    os.system('python AddEmployee.py')


def main():
    root = tk.Tk()
    root.title('Main Page')
    root.geometry('800x450')

    s = ttk.Style()
    s.configure('.', font=('Helvetica', 13), background='blue', width=20, borderwidth=1,
                focusthickness=3)

    label = tk.Label(root, text="Please Select One Program to Run", font=("Arial", 18))
    label.pack(pady=40)

    DAC_button = ttk.Button(root, text='Run Daily_Attendance_Checker', command=lambda: runDailyAttendanceChecker())
    DAC_button.pack(pady=10)

    MRC_button = ttk.Button(root, text='Run Monthly_Report_Creator', command=lambda: runMonthlyReportCreator())
    MRC_button.pack(pady=10)

    AddEmployee_button = ttk.Button(root, text='Add Employee', command=lambda: runAddEmploye())
    AddEmployee_button.pack(pady=10)

    Close_button = ttk.Button(root, text='Close', command=lambda: root.destroy())
    Close_button.pack(pady=80)

    root.mainloop()


def myInfo():
    root = tk.Tk()
    root.title('Info Page')
    root.geometry('800x450')

    label = tk.Label(root, text="Welcome to Info Page", font=("Arial", 14))
    label.pack(pady=20)
    label = tk.Label(root, text="This Program includes 2 different subprograms.\nOne program needs the others output." +
                                "\n\nHow to Execute: First You need to run Daily_Attendance_Checker for daily outputs\n"
                                + "(note: normally it will be run by certain employee on each working day.)\n" +
                                "Then at the end of the month, the second program(Monthly_Report_Creator)\n" +
                                "will be run. There will be outputs as the first program runs throughout the month." +
                                "\nAnd these outputs will be used by the second program to output Excel reports."
                                + " ", font=("Arial", 14))
    label.pack(pady=20)

    open_button = ttk.Button(root, text='Close Info Page', command=lambda: root.destroy())
    open_button.pack(expand=True)

    root.mainloop()
    main()


myInfo()
