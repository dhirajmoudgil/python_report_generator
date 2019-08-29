from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import numpy as np
import pandas as pd
import os.path as path
import matplotlib.pyplot as plt
import pyperclip as clip
import os

#Prerequisites:
# Install xlrd module
# File should have headers


isCsv = True
allFilesPath = ""
fileExtension = ""
sortedData = pd.DataFrame()
allHeaders = []
dropList = []
sortedFilePath = ""
options = ['Select Chart Name', 'Closure', 'Work Areas']
caseData = {}
totalCasesDataFrame = pd.DataFrame()

window = Tk()
window.title("Alps Support")
mod = StringVar()
mod.set('teamconnect')

def captureFileData():
    filePathField.config(state=NORMAL)
    window.filename = filedialog.askopenfilename(initialdir="/", title='Select File with all Weekly Pages', filetypes=(('CSV Files(.csv)', '*.csv'), ('Excel 2007(.xls)', '*.xls'), ('Excel 2003(.xlsx)', '*.xlsx')))
    filePathField.delete(0, 'end')
    filePathField.insert(INSERT, window.filename)
    filePathField.config(state=DISABLED)

def captureClosedFile():
    closedfilePathField.config(state=NORMAL)
    window.filename = filedialog.askopenfilename(initialdir="/", title='Select File with Closed Cases', filetypes=(('CSV Files(.csv)', '*.csv'), ('Excel 2007(.xls)', '*.xls'), ('Excel 2003(.xlsx)', '*.xlsx')))
    closedfilePathField.delete(0, 'end')
    closedfilePathField.insert(INSERT, window.filename)
    closedfilePathField.config(state=DISABLED)

def selectSortedFile():
    window.filename = filedialog.askopenfilename(initialdir="/", title='Select Sorted File', filetypes=(('CSV Files(.csv)', '*.csv'), ('Excel 2007(.xls)', '*.xls'), ('Excel 2003(.xlsx)', '*.xlsx')))
    chartFilePath.delete(0, 'end')
    chartFilePath.insert(INSERT, window.filename)

def sortData():
    global allFilesPath
    global isCsv
    global sortedData
    global fileExtension
    global allHeaders

    allFilesPath = path.dirname(filePathField.get())
    closedFilePath = path.dirname(closedfilePathField.get())
    fileExtension = path.splitext(filePathField.get())[1]
    closedFileExt = path.splitext(closedfilePathField.get())[1]

    cPath = closedfilePathField.get()
    aPath = filePathField.get()
    if len(cPath.strip()) > 0 and len(aPath.strip()) > 0:
        try:
            #fetching closed cases data
            if closedFileExt == '.csv':
                closedData = pd.read_csv(closedfilePathField.get())
            else:
                closedData = pd.read_excel(closedfilePathField.get())

            #fetching data depending upon the file type
            if fileExtension == '.csv':
                unsortedData = pd.read_csv(filePathField.get())
            else:
                unsortedData = pd.read_excel(filePathField.get())
                isCsv = False

            #getting column headers
            allHeaders = unsortedData.columns.values
            clsHeaders = closedData.columns.values

            #checking if any cell is blank/null
            if unsortedData.isnull().values.any():
                messagebox.showinfo('Empty Fields', 'Some of the fields are empty')
                sys.exit()
            else:
                #removing duplicates
                sortedData = unsortedData.drop_duplicates(allHeaders[0], keep='first')

                #merging closed cases in sorted data
                for closedIndex in range(closedData.__len__()):
                    for index in range(sortedData.__len__()):
                        if sortedData.iloc[index][allHeaders[0]] == closedData.iloc[closedIndex][clsHeaders[0]]:
                            sortedData.iloc[index][allHeaders[2]] = 'Closed'
                            dropList.append(closedIndex)
                            break

                #removing already added cases from closed data
                for x in dropList:
                    closedData = closedData.drop(x,axis=0)

                # changing headers of closed data(similar to that of all data) so thar cases in closed data
                # be appended under first column of all data dataframe
                clsHeaders = []
                for x in range(len(closedData.columns)):
                    clsHeaders.append(allHeaders[x])
                closedData.columns = clsHeaders

                #adding new closed cases
                sortedData = sortedData.append(closedData, ignore_index=True)[sortedData.columns.tolist()]

                messagebox.showinfo('Success', 'Report Sorted!')
                generatedSortedFile['state'] = 'normal'
        except FileNotFoundError:
            messagebox.showerror('Error', 'File not found at entered path\nPlease enter correct path or select file from button')
    else:
        messagebox.showerror('Error', 'Enter File Path')

def createSeparateFiles(fileExt, data):
    global caseData

    caseData[allHeaders[2]] = {}  # Status
    caseData[allHeaders[1]] = {}  # Work Areas

    # grouping by work areas
    groupByArea = data.groupby(allHeaders[1])
    # grouping by status
    groupByStatus = data.groupby(allHeaders[2])


    # adding data for closure
    for value, subset in groupByStatus:
        caseData[allHeaders[2]][value] = len(subset)

        if fileExt == '.csv':
            pd.DataFrame(groupByStatus.get_group(value)).to_csv(allFilesPath + '/' +value + fileExt)
        else:
            pd.DataFrame(groupByStatus.get_group(value)).to_excel(allFilesPath + '/' + value + fileExt)
    caseData[allHeaders[2]]['Total'] = len(data[allHeaders[2]])

    # adding data for work areas
    for value, subset in groupByArea:
        caseData[allHeaders[1]][value] = len(subset)

    print(caseData)

def generateFile():
    global sortedFilePath
    global allHeaders

    if isCsv:
        sortedFilePath = allFilesPath+'/sortedReport.csv'
        sortedData.to_csv(sortedFilePath)
        chartFilePath.delete(0, 'end')
        chartFilePath.insert(INSERT,sortedFilePath)
        messagebox.showinfo('Success', 'File Created!\nPlease fill empty fields')
        printBtn['state'] = 'normal'
    else:
        sortedFilePath = allFilesPath+'/sortedReport'+fileExtension
        sortedData.to_excel(sortedFilePath)
        chartFilePath.delete(0, 'end')
        chartFilePath.insert(INSERT,sortedFilePath)
        messagebox.showinfo('Success', 'File Created!\nPlease fill empty fields')
        printBtn['state'] = 'normal'


def printBarChart(status, caseCount, name):
    plt.figure(figsize=(6,3.5))
    plt.rcParams.update({'font.size': 9})
    plt.barh(status, caseCount)
    print(mod.get().capitalize())
    plt.title(mod.get().capitalize() + ' Weekly '+ name + ' Bar Representation', fontsize=10)
    plt.ylabel('Status')
    plt.xlabel('Number of Cases')
    for value, index in enumerate(caseCount):
        plt.text(index, value, str(index))
    plt.tight_layout()
    plt.show()

def printPieChart(caseCount, status, name):
    plt.figure(figsize=(6,3.5))
    plt.rcParams.update({'font.size': 9})
    plt.pie(caseCount, labels=status)
    plt.legend(status)
    plt.legend(fontsize='small')
    print(mod.get().capitalize())
    plt.title(mod.get().capitalize() + ' Weekly '+ name + ' Pie Representation', fontsize=10)
    plt.tight_layout()
    plt.show()

def getDataForClosure(dataFromFile):
    # grouping by status
    groupByStatus = dataFromFile.groupby(allHeaders[2])

    status = []
    caseCount = []
    for value, subset in groupByStatus:
        status.append(value)
        caseCount.append(len(subset))

    return (status, caseCount)

def getDataForWorkAreas(dataFromFile):
    # grouping by work areas
    groupByArea = dataFromFile.groupby(allHeaders[1])

    workArea = []
    caseCount = []
    for value, subset in groupByArea:
        workArea.append(value)
        caseCount.append(len(subset))

    return (workArea, caseCount)

def printChart():
    global allFilesPath
    global allHeaders

    fileExtension = path.splitext(chartFilePath.get())[1]
    allFilesPath = path.dirname(chartFilePath.get())

    try:
        if fileExtension == '.csv':
            chartData = pd.read_csv(chartFilePath.get())
        else:
            chartData = pd.read_excel(chartFilePath.get())

        allHeaders = chartData.columns.values
        allHeaders = np.delete(allHeaders, 0)

        if chartData.isnull().values.any():
            messagebox.showinfo('Empty Fields', 'Please enter values in empty fields')
        else:
            createSeparateFiles(fileExtension, chartData)

            # closure chart
            if chName.get() == options[1]:
                cookedData = getDataForClosure(chartData)

                if rd.get() == 1:
                    printBarChart(cookedData[0], cookedData[1], options[1])

                elif rd.get() == 2:
                    printPieChart(cookedData[1], cookedData[0], options[1])

                else:
                    messagebox.showinfo('Success', 'Nothing!')

            # work areas chart
            elif chName.get() == options[2]:
                cookedData = getDataForWorkAreas(chartData)

                if rd.get() == 1:
                    printBarChart(cookedData[0], cookedData[1], options[2])

                elif rd.get() == 2:
                    printPieChart(cookedData[1], cookedData[0], options[2])

                else:
                    messagebox.showinfo('Success', 'Nothing!')

            else:
                messagebox.showerror('Error', 'Please Specify Chart Name')
    except FileNotFoundError:
        messagebox.showerror('Error', 'File not found at entered path\nPlease enter correct path or select file from button')

# header frame
headFrame = Frame(window, padx = 5, pady = 5)
headFrame.pack()
heading = Label(headFrame, text = 'Weekly Report Generator', font = 'Times 18')
heading.grid(row = 0, columnspan = 2)

# module selection frame
modFrame = LabelFrame(window, text = 'Create Report For', padx = 5, pady = 5)
modFrame.pack(padx = 20, pady = 5)

teamConnect = Radiobutton(modFrame, text = 'TeamConnect', variable = mod, value = 'teamconnect', width=23)
teamConnect.grid(row = 0, column = 0, sticky = W)
collaborati = Radiobutton(modFrame, text = 'Collaborati', variable = mod, value = 'collaborati', width=23)
collaborati.grid(row = 0, column = 1, sticky = W)

# file selection frame
fileFrame = LabelFrame(window, text = 'File Selection', padx = 5, pady = 5)
fileFrame.pack(padx = 50, pady = 20)

closedFileBtn = Button(fileFrame, text = 'Select Closure File', width = 15, command = captureClosedFile)
closedFileBtn.grid(row = 0, column = 0)
closedfilePathField = Entry(fileFrame, width = 40, borderwidth = 5)
closedfilePathField.grid(row = 0, column = 1, padx = 5)

selectFileBtn = Button(fileFrame, text = 'Select All Cases File', command = captureFileData, width = 15)
selectFileBtn.grid(row = 1, column = 0, pady = 5)
filePathField = Entry(fileFrame, width = 40, borderwidth = 5)
filePathField.grid(row = 1, column = 1, padx = 5)

#sort button
sortBtn = Button(fileFrame, text = 'Sort', width = 50, bg = 'grey', command = sortData)
sortBtn.grid(row = 2, columnspan = 2, pady = 5)
generatedSortedFile = Button(fileFrame, text = 'Generate Sorted File', state = DISABLED,  width = 50, bg = 'grey', command = generateFile)
generatedSortedFile.grid(row = 3, columnspan = 2)

#sorted file frame
chartFieldFrame = LabelFrame(window, padx = 5, pady = 5)
chartFieldFrame.pack(padx = 10, pady = 10)
chartFileBtn = Button(chartFieldFrame, text = 'Select Sorted File', width = 15, command = selectSortedFile)
chartFileBtn.grid(row = 0, column = 0, pady = 5)
chartFilePath = Entry(chartFieldFrame, width = 40, borderwidth = 5)
chartFilePath.grid(row = 0, column = 1, padx = 5)

# charts selection frame
chartFrame = LabelFrame(window, text = 'Chart Selection', padx = 5, pady = 5)
chartFrame.pack(padx = 10, pady = 10)

chartName = Label(chartFrame, text = 'Chart Name:', width = 18)
chartName.grid(row = 0, column = 0)
#chart name drop down
chName = StringVar()
chName.set(options[0])
chartDropDown = OptionMenu(chartFrame, chName, *options)
chartDropDown.config(width=20)
chartDropDown.grid(row = 0, column = 1)

chartType = Label(chartFrame, text = 'Chart Type:')
chartType.grid(row = 2, column = 0, pady = 5)
#chart type  radio buttons
rd = IntVar()
barChrt = Radiobutton(chartFrame, text = 'Bar Chart', variable = rd, value = 1)
barChrt.grid(row = 3, column = 0, sticky = W)
pieChrt = Radiobutton(chartFrame, text = 'Pie Chart', variable = rd, value = 2)
pieChrt.grid(row = 3, column = 1, sticky = W)
lineChrt = Radiobutton(chartFrame, text = 'Line Chart', variable = rd, value = 3)
lineChrt.grid(row = 3, column = 2, sticky = W)

printBtn = Button(chartFrame, text = 'Print Chart', width = 50, bg = 'grey', command = printChart)
printBtn.grid(row = 4, columnspan = 3, pady = 5)

window.mainloop()