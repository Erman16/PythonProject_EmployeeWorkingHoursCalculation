"""
Designed and codded by Erman Yalçın
Linkedln: https://www.linkedin.com/in/ermanyalcin/
"""
import datetime
import json
import shutil
import openpyxl
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import os.path
from babel.dates import format_datetime
from datetime import datetime

""" 
In This Program you need to provide output data from Daily Attendance Checker and with this data this program will
produce you result as Excel. Employee list can be changed From AddEployee program
"""
# ----- Take Employee List -----

MyCompanyName = "Company Name"  # You can Enter Your Company Name


def getEmployees():
    f = open('Employee_List/Employees.json', encoding="utf8")
    data = json.load(f)
    f.close()
    return data


EmployeesDict = getEmployees()


# ---------------------------------
# --------- Save as Excel ---------


def SaveAsExcel(data, excelname, name, company):
    xfile = openpyxl.load_workbook(excelname)
    sheet = xfile['Sheet1']

    sheet["D2"] = name
    sheet["D3"] = company

    counter = 5
    boxNameDate = "A" + str(counter)
    boxNameDay = "B" + str(counter)
    boxNameWorkingHours = "C" + str(counter)
    boxNameEntrance = "D" + str(counter)
    boxNameExit = "E" + str(counter)
    boxNameExplanation = "F" + str(counter)
    boxNameTotalWorkHours = "G" + str(counter)
    boxNameMissing = "H" + str(counter)
    boxNameOvertime = "I" + str(counter)

    for i in data:
        try:
            sheet[boxNameDate] = i[0].split("_")[0]
            sheet[boxNameDay] = i[0].split("_")[1]
            sheet[boxNameWorkingHours] = "7:00-17:00"

            myI1 = i[1].split(".")
            myI2 = i[2].split(".")
            if int(myI1[0]) < 10 or int(myI1[1]) < 10:
                if int(myI1[0]) < 10:
                    myI1[0] = "0" + myI1[0]
                if int(myI1[1]) < 10:
                    myI1[1] = "0" + myI1[1]

            if int(myI2[0]) < 10 or int(myI2[1]) < 10:
                if int(myI2[0]) < 10:
                    myI2[0] = "0" + myI2[0]
                if int(myI2[1]) < 10:
                    myI2[1] = "0" + myI2[1]
            myI1 = myI1[0] + "." + myI1[1].replace("\n", "")
            myI2 = myI2[0] + "." + myI2[1].replace("\n", "")

            tdelta = (datetime.strptime(myI2, '%H.%M') - datetime.strptime(myI1, '%H.%M')).seconds/3600

            sheet[boxNameEntrance] = i[1]
            sheet[boxNameExit] = i[2]
            sheet[boxNameExplanation] = ""
            sheet[boxNameTotalWorkHours] = "{0:.2f}".format(tdelta)
            print("0")
            if tdelta > 10:
                sheet[boxNameOvertime] = "{0:.2f}".format(tdelta - 10)
                sheet[boxNameMissing] = "-"
            else:
                sheet[boxNameOvertime] = "-"
                sheet[boxNameMissing] = "{0:.2f}".format(10 - tdelta)
            print("1")
        except:
            sheet["A5"] = "Something Went Wrong!!!"
        print("2")
        counter += 1
        boxNameDate = "A" + str(counter)
        boxNameDay = "B" + str(counter)
        boxNameWorkingHours = "C" + str(counter)
        boxNameEntrance = "D" + str(counter)
        boxNameExit = "E" + str(counter)
        boxNameExplanation = "F" + str(counter)
        boxNameTotalWorkHours = "G" + str(counter)
        boxNameMissing = "H" + str(counter)
        boxNameOvertime = "I" + str(counter)
        xfile.save(excelname)
        print("3")


def CopyTheExcel(person, company, getInfo, filePath):
    getNecessaryInfo = []
    print(person)
    for i in getInfo:
        date = i.split(":")[0]
        enterHour = (i.split(" ")[1]).split("-")[0]
        endHour = (i.split(" ")[1]).split("-")[1]
        getNecessaryInfo.append([date,enterHour,endHour])

    shutil.copy("Monthly_Report_Creator/sample.xlsx", "Monthly_Report_Creator/Result/" + filePath + "/" + person + ".xlsx")
    SaveAsExcel(getNecessaryInfo, "Monthly_Report_Creator/Result/" + filePath + "/" + person + ".xlsx", person, company)


# -------------------------------

def GetDataFromDatabase(person, path):
    filename = "Monthly_Report_Creator/Add_Employee_Data/" + path + "/" + person + ".txt"
    getInfo = []
    with open(filename) as f:
        for line in f:
            getInfo.append(line)
    return getInfo


def startProg(filePath):
    missingEmployeesData = []

    for i in EmployeesDict.keys():
        try:
            getInfo = GetDataFromDatabase(EmployeesDict[i], filePath)
            CopyTheExcel(EmployeesDict[i], MyCompanyName, getInfo, filePath)
        except:
            missingEmployeesData.append(EmployeesDict[i])


class closeClass:
    closeValue = False


p1 = closeClass()
p1.closeValue = False


def main():
    notChosenOnes = []
    root = tk.Tk()
    root.title('File Selection')
    root.geometry('800x450')

    mydate = format_datetime(datetime.now(), locale='en').split(" ")
    for i in range(len(mydate)):
        mydate[i] = mydate[i].replace(",", "")
    NameDate = mydate[0] + "_" + mydate[2]

    try:
        os.mkdir('Monthly_Report_Creator/Add_Employee_Data/' + NameDate)
    except FileExistsError:
        pass

    try:
        os.mkdir('Monthly_Report_Creator/Result/' + NameDate)
    except FileExistsError:
        pass

    for i in EmployeesDict.keys():
        check = os.path.isfile('Monthly_Report_Creator/Add_Employee_Data/' + NameDate + "/" + EmployeesDict[i] + ".txt")
        if not check:
            notChosenOnes.append(EmployeesDict[i])

    label = tk.Label(root, text="Missing Data", font=("Arial", 14))
    label.pack(pady=20)

    myText = ""
    for i in notChosenOnes:
        myText = myText + i + "\n"

    T = tk.Text(root, height=15, width=70)

    T.pack()

    T.insert(tk.END, myText)

    def select_files(NameDate):
        filetypes = (('text files', '*.txt'), ('All files', '*.*'))

        filenames = fd.askopenfilenames(title='Open files', initialdir='/', filetypes=filetypes)

        for files in filenames:
            f = open(files)
            writedata = ""
            for i in f.readlines():
                writedata = writedata + i

            f.close()
            isim = files.split("/")[-1]
            print(isim)
            saveF = open("Monthly_Report_Creator/Add_Employee_Data/" + NameDate + "/" + isim, "w")
            saveF.write(writedata)
            saveF.close()
        root.destroy()

    def close():
        root.destroy()
        p1.closeValue = True

    open_button = ttk.Button(root, text='Chose File', command=lambda: select_files(NameDate))
    open_button.pack(expand=True)

    close_button = ttk.Button(root, text='Close', command=close)
    close_button.pack(expand=True)

    root.mainloop()
    startProg(NameDate)


while True:
    main()
    print(p1.closeValue)
    if p1.closeValue:
        break
