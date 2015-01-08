#!/usr/bin/python

import os
import sys
import re
import codecs
import datetime
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from Things import MenuItemThings
from tkinter import filedialog
from tempfile import TemporaryFile
from xlwt import Workbook, easyxf
from xlrd import open_workbook

priceArrayMatch = re.compile(r'(?<=\{)[^(\{|\})].+?(?=\})')
IG_EXPORT = 1
SIMPLE_EXPORT = 3
UNKNOWN_EXPORT = 10
CSV_EXPORT = 2
itemList = []
itemMap = {}

def ezPrint(string):
    print(str(string))
    
def openFile(options=None):
    if options == None:
        options = {}
        options['defaultextension'] = '.txt' 
        options['filetypes'] = [('Text Files', '.txt'), ('CSV Files', '*.csv*'), ('All Files', '.*')]
        options['title'] = 'Open...'
    file_opt = options
    global file_path
    file_path = filedialog.askopenfilename(**file_opt)
    options = None
    if file_path == None or file_path == "":
        print("No file selected")
    openFileString.set(str(os.path.basename(file_path)))
    if determineExportType(file_path) == IG_EXPORT:
        conversionButtonText.set("Simplify")
    else:
        conversionButtonText.set("Generate IG Update")
        
    showButton()
    
def saveFile(options):
        
    file_opt = options
    global csv_file
    global save_file
    csv_file = str(file_path)[:-4] + ".csv"
    save_file = filedialog.asksaveasfilename(**file_opt)
    if save_file == None or save_file == "":
        print("No file selected")
        
def fixArray(match):
    match = str(match.group(0))
    return match.replace(",",";")
    
def preParse(export, output):
    for x in export:
        itemDetails = re.sub(priceArrayMatch, fixArray, x)
        item = itemDetails.split(",")
        i = MenuItemThings.MenuItem(
                                item[1], item[2], item[3], item[4], item[5],
                                item[6], item[7], item[8], item[9], item[10],
                                item[11], item[12], item[13], item[14], item[15],
                                item[16], item[18], item[19], item[20], item[21],
                                item[22], item[23], item[24], item[25], item[26],
                                item[28], item[29], item[30], item[31]
                                )
        itemList.append(i)
        itemMap[i.id] = i
        try:
            output.write(itemDetails)
        except UnicodeEncodeError:
            errorText = "\n\n!!!!!!!!!!!!!!!!!!!!!!!\nerror encoding string for print/output\n!!!!!!!!!!!!!!!!!!!!!!!!!\n\n"
            print(errorText)
            output.write("error processing item " + str(i.id) + "\n")
    print("completed")
    
def readFromExcel(file=None):
    if file == None:
        file = 'simple_export.xls'
    
    book = open_workbook(file)
    print('openning ' + str(file))
    sheet = book.sheet_by_index(0)
    
    print('name: ' + sheet.name)
    print('number of rows: ' + str(sheet.nrows))
    print('number of columns: ' + str(sheet.ncols))
    print(str(sheet.row(0)))
    
    
def enumeratePriceLevels():
    numberOfPriceLevels = 0
    for item in itemList:
        if len(item.separatePriceLevels()) > numberOfPriceLevels:
            numberOfPriceLevels = len(item.separatePriceLevels())
    return numberOfPriceLevels

def generateSimpleExport(items=itemList, altered=True):
    simpleOutput = codecs.open(save_file, 'w+', 'utf8')
    
    book = Workbook()
    heading = easyxf(
        'font: bold True;'
        'alignment: horizontal center;'
        )
    sheet = book.add_sheet('Sheet 1')
    sheet.panes_frozen = True
    sheet.remove_splits = True
    sheet.horz_split_pos = 1
    row1 = sheet.row(0)
    row1.write(0, 'ID', heading)
    row1.write(1, 'Name', heading)
    
    numberOfPriceLevels = enumeratePriceLevels()
    for x in range(numberOfPriceLevels):
        col = x + 2
        row1.write(col, 'Price Level ' + str(x + 1), heading)
        sheet.col(col).width = 4260
    
    for i,item in zip(range(1, len(items) + 1),items):
        if altered:
            if item.priceLevels != "{}":
                simpleOutput.write(str(item.id) + "," + str(item.name) + "," + str(item.priceLevels) + "\r\n")
        else:
            simpleOutput.write(str(item.id) + "," + str(item.name) + "," + str(item.priceLevels) + "\r\n")
        row = sheet.row(i)
        row.write(0, str(item.id))
        row.write(1, str(item.name))
        for p in range(1, (numberOfPriceLevels + 1)):
            if p in item.separatePriceLevels():
                price = item.separatePriceLevels()[p]
            else:
                price = ''
            row.write((p + 1), str(price))
    
    sheet.col(1).width = 12780
    book.save('simple_export.xls')
    print('excel workbook saved')

    messagebox.showinfo(title='Success', message='Simplified item export created successfully.')
        
def convertToExcel(items=itemList, altered=True):
    simpleOutput = codecs.open(save_file, 'w+', 'utf8')
    
    book = Workbook()
    heading = easyxf(
        'font: bold True;'
        'alignment: horizontal center;'
        )
    sheet = book.add_sheet('Sheet 1')
    sheet.panes_frozen = True
    sheet.remove_splits = True
    sheet.horz_split_pos = 1
    row1 = sheet.row(0)
    row1.write(0, 'ID', heading)
    row1.write(1, 'Name', heading)
    
    for i,item in zip(range(1, len(items) + 1),items):
        row = sheet.row(i)
        row.write(0, str(item.id))
        row.write(1, str(item.name))
        row.write(2, str(item.priceLevels))
    
    sheet.col(1).width = 12780
    book.save('complete_export.xls')
    print('excel workbook saved')

    messagebox.showinfo(title='Success', message='Excel export created successfully.')
        
def generateIGPriceUpdate(inputFile, updateFile):
    print('preparing to generate IG Update file')
    if inputFile[-3:] == 'xls' or inputFile[-4:] == 'xlsx':
        print('generating IG Update from xls')
        book = open_workbook(inputFile)
        sheet = book.sheet_by_index(0)
        if sheet.ncols > 3:
            print('Extra Price Levels found.')
            for row in range(1, sheet.nrows):
                prices = []
                for col in range(2, sheet.ncols):
                    if sheet.cell_value(row,col) != '':
                        priceLevel = str(col - 1) + ',' + str(sheet.cell_value(row,col))
                        prices.append(priceLevel)
                priceImport = '{' + ','.join(prices) + '}'
                line = '"U",' + str(sheet.cell_value(row, 0)) + ',,,,,' + str(priceImport) + ',,,,,,,,,,,,,,,,,\r\n'
                updateFile.write(line)
        else:
            print('Processing with a single price level')
            for row in range(1, sheet.nrows):
                line = '"U",' + str(sheet.cell_value(row, 0)) + ',,,,,' + str(cell_value(row,2)) + ',,,,,,,,,,,,,,,,,\r\n'
    else:
        print('generating IG Update from txt or csv')
        with codecs.open(file_path, 'r', 'utf8') as file:
            for x in file:
                details = x.split(",")
                details[2] = details[2].replace(";", ",").strip("\r\n")
                line = '"U",' + str(details[0]) + ',,,,,' + str(details[2]) + ',,,,,,,,,,,,,,,,,\r\n'
                updateFile.write(line)

    messagebox.showinfo(title='Success', message='IG Item Import created successfully.')
    
def determineExportType(f):
    if f[-3:] == 'xls':
        print('Input file is xls, processing as SIMPLE_EXPORT')
        return SIMPLE_EXPORT
    else:
        file = codecs.open(f, 'r', 'utf8')
        if len(file.readline().split(",")) > 20:
            return IG_EXPORT
        else:
            return SIMPLE_EXPORT
        
def runConversion():

	export = file_path
		
	if conversionButtonText.get() == "Simplify":
		options = {}
		options['title'] = 'Save As'
		options['initialfile'] = str(file_path)[:-4] + "_simplified" + str(file_path)[-4:]
		saveFile(options)
		output = codecs.open(csv_file, 'w+', 'utf8')
		try:
			preParse(export, output)
		except UnicodeDecodeError:
			export = codecs.open(file_path, 'r', 'latin-1')
			preParse(export, output)
		generateSimpleExport(altered=truncate.get())
	else:
		options = {}
		options['title'] = 'Save As'
		options['initialfile'] = 'MI_IMP.txt'
		saveFile(options)
		output = codecs.open(save_file, 'w+', 'latin-1')
		generateIGPriceUpdate(export, output)

def hideButton():
    thatButton.grid_remove()
    
def showButton():
    thatButton.grid()

root = Tk()
root.option_add('*tearOff', FALSE)
root.title("Agilysy File Tools")

save_syserr = sys.stderr
fsock = open('error.log', 'a+')
sys.stderr = fsock

openFileString = StringVar()
conversionButtonText = StringVar()
truncate = StringVar()

menubar = Menu(root)
menu_file = Menu(menubar)
menu_options = Menu(menubar)
menubar.add_cascade(menu=menu_file, label='File')
menubar.add_cascade(menu=menu_options, label='Options')

menu_file.add_command(label='Open...', command=openFile)
menu_file.add_command(label='Close', command=root.quit)

menu_options.add_checkbutton(label='Condense Simplified Output', variable=truncate, onvalue=1, offvalue=0)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=1, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(1, weight=1)

ttk.Label(mainframe, text="Input File:").grid(column=1, row=1, sticky=(N, W, E))
openFile_entry = ttk.Entry(mainframe, width=40, textvariable=openFileString)
openFile_entry.grid(column=1, row=2, sticky=(W, E))

global thatButton
thatButton = ttk.Button(mainframe, textvariable=conversionButtonText, command=runConversion)
thatButton.grid(column=1, row=3)

for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)

root.config(menu=menubar)
hideButton()
root.mainloop()