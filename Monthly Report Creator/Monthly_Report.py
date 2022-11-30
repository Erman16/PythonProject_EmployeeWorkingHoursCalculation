"""
Designed and codded by Erman Yalçın
Linkedln: https://www.linkedin.com/in/ermanyalcin/
"""
import datetime
import json
import shutil
import openpyxl
from openpyxl import Workbook
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import os.path
from babel.dates import format_datetime
import babel.numbers

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

    for key in data.keys():
        try:
            if len(data[key]) == 3:
                sheet[boxNameDate] = key
                sheet[boxNameDay] = data[key][0]
                sheet[boxNameWorkingHours] = "8:00-17:00"

                sheet[boxNameEntrance] = data[key][1]
                sheet[boxNameExit] = data[key][2]
                sheet[boxNameExplanation] = "-"
                totalhours = round((abs(sum(x * int(t) for x, t in zip([3600, 60], data[key][1].split(":"))) - sum(
                    x * int(t) for x, t in zip([3600, 60], data[key][2].split(":")))) - 5400) / 3600, 2)
                sheet[boxNameTotalWorkHours] = totalhours
                if totalhours < 9:
                    sheet[boxNameMissing] = 9 - totalhours
                else:
                    sheet[boxNameMissing] = "-"

                if int(data[key][2].split(":")[0]) > 18:
                    if int(data[key][2].split(":")[0]) == 18:
                        if int(data[key][2].split(":")[1]) > 30:
                            sheet[boxNameOvertime] = round(abs(64800 - sum(
                                x * int(t) for x, t in zip([3600, 60], data[key][2].split(":")))) / 3600, 2)
                    else:
                        sheet[boxNameOvertime] = round(
                            abs(64800 - sum(x * int(t) for x, t in zip([3600, 60], data[key][2].split(":")))) / 3600, 2)
                else:
                    sheet[boxNameOvertime] = "-"

            elif len(data[key]) > 3:
                sheet[boxNameDate] = key
                sheet[boxNameDay] = data[key][0]
                sheet[boxNameWorkingHours] = "8:00-17:00"

                sheet[boxNameExplanation] = data[key][3]
                sheet[boxNameEntrance] = data[key][1]
                sheet[boxNameExit] = data[key][2]
                totalhours = round((abs(sum(x * int(t) for x, t in zip([3600, 60], data[key][1].split(":"))) - sum(
                    x * int(t) for x, t in zip([3600, 60], data[key][2].split(":")))) - 5400) / 3600, 2)
                sheet[boxNameTotalWorkHours] = totalhours
                if totalhours < 9:
                    sheet[boxNameMissing] = 9 - totalhours
                else:
                    sheet[boxNameMissing] = "-"

                if int(data[key][2].split(":")[0]) > 18:
                    if int(data[key][2].split(":")[0]) == 18:
                        if int(data[key][2].split(":")[1]) > 30:
                            sheet[boxNameOvertime] = round(abs(64800 - sum(
                                x * int(t) for x, t in zip([3600, 60], data[key][2].split(":")))) / 3600, 2)
                    else:
                        sheet[boxNameOvertime] = round(
                            abs(64800 - sum(x * int(t) for x, t in zip([3600, 60], data[key][2].split(":")))) / 3600, 2)
                else:
                    sheet[boxNameOvertime] = "-"
            else:
                sheet[boxNameDate] = key
                sheet[boxNameDay] = data[key][0]
                sheet[boxNameWorkingHours] = "8:00-17:00"

                sheet[boxNameEntrance] = "-"
                sheet[boxNameExit] = "-"
                sheet[boxNameExplanation] = data[key][1]
                sheet[boxNameTotalWorkHours] = "-"
                sheet[boxNameMissing] = "-"
                sheet[boxNameOvertime] = "-"
        except:
            sheet["A5"] = "Something Went Wrong!!!"

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


def CopyTheExcel(person, company, getInfo, filePath):
    getNecessaryInfo = []
    for i in getInfo:
        i = i.replace("    ", "", 1)
        if "." in i:
            i = i.split(" ")
            j = []
            for ele in i:
                if ele != "":
                    j.append(ele)
            getNecessaryInfo.append(j)
    daysDict = {}
    for i in getNecessaryInfo:
        try:
            try:
                if i[7] != "\n":
                    value = ""
                    for x in range(7, len(i)):
                        try:
                            value = value + str(i[x].split("\n")[0]) + " "
                        except:
                            value = value + str(i[x]) + " "
                    daysDict[i[0]] = [i[1], i[4], i[6], value]
                else:
                    daysDict[i[0]] = [i[1], i[4], i[6]]
            except:
                if "." in i[0]:
                    value = ""
                    for x in range(3, len(i)):
                        value = value + str(i[x])
                    daysDict[i[0]] = [i[1], value]
                else:
                    pass
        except:
            pass
    shutil.copy("sample.xlsx", "Result/" + filePath + "/" + person + ".xlsx")
    SaveAsExcel(daysDict, "Result/" + filePath + "/" + person + ".xlsx", person, company)


# -------------------------------

def GetDataFromDatabase(person, path):
    filename = "Add_Employee_Data/" + path + "/" + person + ".txt"
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

    mydate = format_datetime(datetime.datetime.now(), locale='en').split(" ")
    for i in range(len(mydate)):
        mydate[i] = mydate[i].replace(",", "")
    NameDate = mydate[0] + " " + mydate[2]

    try:
        os.mkdir('Add_Employee_Data/' + NameDate)
    except FileExistsError:
        pass

    try:
        os.mkdir('Result/' + NameDate)
    except FileExistsError:
        pass

    for i in EmployeesDict.keys():
        check = os.path.isfile('Add_Employee_Data/' + NameDate + "/" + EmployeesDict[i] + ".txt")
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

        chosenOnes = []
        for files in filenames:
            f = open(files)
            writedata = ""
            for i in f.readlines():
                writedata = writedata + i
                if "Name & Surname" in i:
                    neededData = []
                    i = i.split(" ")
                    for x in range(len(i)):
                        if i[x] != "":
                            try:
                                i[x] = i[x].split("\n")[0]
                            except:
                                pass
                            neededData.append((i[x]))
            isim = ""
            for i in range(2, len(neededData)):
                if i == len(neededData) - 1:
                    isim = isim + neededData[i]
                else:
                    isim = isim + neededData[i] + " "
            f.close()
            saveF = open("Add_Employee_Data/" + NameDate + "/" + isim + ".txt", "w")
            chosenOnes.append(isim)
            saveF.write(writedata)
            saveF.close()
        root.destroy()
        try:
            for i in chosenOnes:
                notChosenOnes.remove(i)
        except:
            pass

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
