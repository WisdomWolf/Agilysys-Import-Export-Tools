#!python3

import os
import sys
import re
import codecs
import pdb
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from Things.MenuItemThings import MenuItem
from xlwt import Workbook, easyxf
from xlrd import open_workbook

priceArrayMatch = re.compile(r'(?<=\{)[^(\{|\})].+?(?=\})')
commaQuoteMatch = re.compile(r'((?<=")[^",\{\}]+),([^"\{\}]*(?="))')
fileTypeFilters = [('Supported Files', '.xls .xlsx .txt'), ('Text Files', '.txt'), ('Excel Files', '.xls .xlsx .csv'), ('All Files', '.*')]
IG_EXPORT = 1
SIMPLE_EXPORT = 3
UNKNOWN_EXPORT = 10
CSV_EXPORT = 2
itemList = []
itemMap = {}

def ezPrint(string):
    print(str(string))
    
def openFile(options=None):
    if itemList:
        itemList.clear()
    if options == None:
        options = {}
        options['defaultextension'] = '.xls*, .txt' 
        options['filetypes'] = fileTypeFilters
        options['title'] = 'Open...'
    file_opt = options
    global file_path
    file_path = filedialog.askopenfilename(**file_opt)
    options = None
    if file_path == None or file_path == "":
        print("No file selected")
        return
    try:
        if determineExportType(file_path) == IG_EXPORT:
            for button in simplifyButtons:
                showButton(button)
        else:
            showButton(thatButton)
        openFileString.set(str(os.path.basename(file_path)))
    except IOError:
        messagebox.showinfo(title='Oops', message='This file is not supported.')
        return
        
    
def saveFile(options):
        
    file_opt = options
    save_file = filedialog.asksaveasfilename(**file_opt)
    if save_file == None or save_file == "":
        print("No file selected")
    return save_file
        
def fixArray(match,):
    match = str(match.group(0))
    return match.replace(",", ";")
    
def preParse(file_name):
    with codecs.open(file_name, 'r', 'latin-1') as export:
        print('pre-parse initiated')
        for line in export:
            itemDetails = re.sub(priceArrayMatch, fixArray, line)
            itemDetails = re.sub(commaQuoteMatch, fixArray, itemDetails)
            item = itemDetails.split(",")
            i = MenuItem(
                                    item[1], item[2], item[3], item[4], item[5],
                                    item[6], item[7], item[8], item[9], item[10],
                                    item[11], item[12], item[13], item[14], item[15],
                                    item[16], item[18], item[19], item[20], item[21],
                                    item[22], item[23], item[24], item[25], item[26],
                                    item[28], item[29], item[30], item[31]
                                    )
#             print('Item:\n' + i.toString())
            itemList.append(i)
    print("completed")
        
def enumeratePriceLevels():
    numberOfPriceLevels = 0
    for item in itemList:
        keys = item.separatePriceLevels().keys()
        for k in keys:
            if int(k) > numberOfPriceLevels:
                numberOfPriceLevels = int(k)
    return numberOfPriceLevels

def generateSimpleExport(save_file, items=None, excludeUnpriced=True):
    items = items or itemList
    print('Generating Simple Export')
    simpleOutput = codecs.open(save_file, 'w+', 'utf8')
    
    for item in items:
        if excludeUnpriced:
            if item.priceLevels != "{}":
                simpleOutput.write(str(item.id) + "," + str(item.name) + "," + str(item.priceLevels) + "\r\n")
        else:
            simpleOutput.write(str(item.id) + "," + str(item.name) + "," + str(item.priceLevels) + "\r\n")
        
    messagebox.showinfo(title='Success', message='Simplified item export created successfully.')

def generateSimpleExcel(save_file, items=None, excludeUnpriced=True):
    items = items or itemList
    print('Generating Simple Excel')
    
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
    
    #Write Column Headings
    for x in range(1, (numberOfPriceLevels + 1)):
        col = x + 1
        row1.write(col, 'Price Level ' + str(x), heading)
        sheet.col(col).width = 4260
    
    #Write items
    for i,item in zip(range(1, len(items) + 1),items):
        if excludeUnpriced:
            if item.priceLevels == "{}":
                continue
        else:
            row = sheet.row(i)
            row.write(0, str(item.id))
            row.write(1, str(item.name).strip('"'))
            for p in range(1, (numberOfPriceLevels + 1)):
                if p in item.separatePriceLevels():
                    price = item.separatePriceLevels()[p]
                else:
                    price = ''
                row.write((p + 1), str(price))
    
    sheet.col(1).width = 12780
    book.save(save_file)
        
    messagebox.showinfo(title='Success', message='Simplified excel sheet created successfully.')
            
def generateFullExcel(save_file, items=None, excludeUnpriced=True, expandPriceLevels=False):
    items = items or itemList
    print('preparing to convert to Excel')
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
    row1.write(0, '0', heading)
    sheet.col(0).hidden = True
    headers = ['"U"', 'ID', 'Name', 'Abbr1', 'Abbr2',
               'Kitchen Printer Label', 'Price', 'Product Class',
               'Revenue Category', 'Tax Group', 'Security Level',
               'Report Category', 'Use Weight', 'Tare', 'SKU',
               'Bar Gun Code', 'Cost', 'Reserved', 'Price Prompt',
               'Print on Check', 'Discountable', 'Voidable', 'Not Active',
               'Tax Included', 'Menu Item Group', 'Receipt',
               'Price Override', 'Reserved', 'Choice Groups', 'KPs',
               'Covers', 'Store ID']
    
    for h,i in zip(headers, range(1,32)):
        if i < 3:
            sheet.row(1).set_cell_boolean(i, True)
        else:
            sheet.row(1).set_cell_boolean(i, False)
        row1.write(i, h, heading)
    
    for i,item in zip(range(2, len(items) + 2),items):
        global isMisaligned
        isMisaligned = False
        columnMap = MenuItem.attributeMap
        
        row = sheet.row(i)
        #row.write(2, int(item.id))
        row.write(columnMap['id'], int(item.id))
        row.write(columnMap['name'], str(item.name))
        row.write(columnMap['abbr1'], str(item.abbr1))
        row.write(columnMap['abbr2'], str(item.abbr2))
        row.write(columnMap['printerLabel'], str(item.printerLabel))
        
        if expandPriceLevels:
            pass
        else:
            row.write(columnMap['priceLevels'], str(item.priceLevels))
            row.write(columnMap['classID'], safeIntCast(item.classID))
            row.write(columnMap['revCategoryID'], safeIntCast(item.revCategoryID))
            row.write(columnMap['taxGroup'], safeIntCast(item.taxGroup))
            row.write(columnMap['securityLevel'], safeIntCast(item.securityLevel))
            row.write(columnMap['reportCategory'], safeIntCast(item.reportCategory))
            row.write(columnMap['useWeightFlag'], safeIntCast(item.useWeightFlag))
            row.write(columnMap['weightTareAmount'], str(item.weightTareAmount))
            row.write(columnMap['sku'], str(item.sku))
            row.write(columnMap['gunCode'], str(item.gunCode))
            row.write(columnMap['costAmount'], str(item.costAmount))
            row.write(columnMap['costAmount'] + 1, 'N/A')
            row.write(columnMap['pricePrompt'], safeIntCast(item.pricePrompt))
            row.write(columnMap['checkPrintFlag'], safeIntCast(item.checkPrintFlag))
            row.write(columnMap['discountableFlag'], safeIntCast(item.discountableFlag))
            row.write(columnMap['voidableFlag'], safeIntCast(item.voidableFlag))
            row.write(columnMap['inactiveFlag'], safeIntCast(item.inactiveFlag))
            row.write(columnMap['taxIncludeFlag'], safeIntCast(item.taxIncludeFlag))
            row.write(columnMap['itemGroupID'], safeIntCast(item.itemGroupID))
            row.write(columnMap['receiptText'], str(item.receiptText))
            row.write(columnMap['priceOverrideFlag'], safeIntCast(item.priceOverrideFlag))
            row.write(columnMap['priceOverrideFlag'] + 1, 'N/A')
            row.write(columnMap['choiceGroups'], str(item.choiceGroups))
            row.write(columnMap['kitchenPrinters'], str(item.kitchenPrinters))
            row.write(columnMap['covers'], str(item.covers))
            row.write(columnMap['storeID'], str(item.storeID))
        
        if isMisaligned:
            oopsStyle = (easyxf('pattern: pattern solid, fore_color rose'))
            row.write(1, 'X', oopsStyle)
    
    try:
        book.save(save_file)
    except PermissionError:
        messagebox.showerror(title= 'Error', message='Unable to save file')

    messagebox.showinfo(title='Success', message='Excel export created successfully.')
        
def generateCustomExcel(save_file, items=None, excludeUnpriced=True):
    pass
    
def generateIGPriceUpdate(inputFile, updateFile):
    print('preparing to generate IG Price Update file')
    if inputFile[-3:] == 'xls' or inputFile[-4:] == 'xlsx':
        print('generating IG Update from xls')
        book = open_workbook(inputFile)
        sheet = book.sheet_by_index(0)
        if sheet.ncols > 3:
            if sheet.cell_value(1,1) == True or sheet.cell_value(1,1) == False:
                generateIGUpdate(book, updateFile)
                return
            print('Extra Price Levels found.')
            for row in range(1, sheet.nrows):
                prices = []
                isNegative = False
                for col in range(2, sheet.ncols):
                    if sheet.cell_value(row,col) != '':
                        priceLevelNumber = str(col - 1) + ','
                        if '(' in str(sheet.cell_value(row,col)) or '(' in str(sheet.cell_value(row,col)):
                            isNegative = True
                        price = '{0:.2f}'.format(float(str(sheet.cell_value(row,col)).strip('$()')))
                        if isNegative == True:
                            priceLevel = priceLevelNumber + '($' + price + ')'
                        else:
                            priceLevel = priceLevelNumber + '$' + price
                        prices.append(priceLevel)
                priceImport = '{' + ','.join(prices) + '}'
                line = '"U",' + str(sheet.cell_value(row, 0)) + ',,,,,' + str(priceImport) + ',,,,,,,,,,,,,,,,,\r\n'
                updateFile.write(line)
        else:
            print('Processing with a single price level')
            for row in range(1, sheet.nrows):
                line = '"U",' + str(sheet.cell_value(row, 0)) + ',,,,,' + str(sheet.cell_value(row,2)) + ',,,,,,,,,,,,,,,,,\r\n'
    else:
        print('generating IG Update from txt or csv')
        with codecs.open(file_path, 'r', 'latin-1') as file:
            for x in file:
                details = x.split(",")
                details[2] = details[2].replace(";", ",").strip("\r\n")
                line = '"U",' + str(details[0]) + ',,,,,' + str(details[2]) + ',,,,,,,,,,,,,,,,,\r\n'
                updateFile.write(line)

    messagebox.showinfo(title='Success', message='IG Item Import created successfully.')

def generateIGUpdate(book, updateFile):
    print('preparing to generate IG Update file')
    sheet = book.sheet_by_index(0)
    includeColumns = set()
    quotedFields = (3, 4, 5, 26)
    
    for col in range(3, sheet.ncols):
        if sheet.cell_value(1, col) == True:
            includeColumns.add(col)
            
    includeColumns = sorted(includeColumns)
            
    for row in range(2, sheet.nrows):
        itemProperties = []
        updateType = sheet.cell_value(row,1)
        if updateType != 'A' and updateType != 'U' and updateType != 'D' and updateType != 'X':
            continue
        elif updateType == 'X':
            messagebox.showwarning(title='File Error', 
                message='One or more lines are not aligned properly.\nPlease correct and retry.')
            return
        else:
            itemProperties.append('"' + str(sheet.cell_value(row,1)) + '"')
        itemProperties.append(safeIntCast((sheet.cell_value(row,2))))
        previousIndex = 2
        for col in includeColumns:
            emptySpaces = col - previousIndex - 1
            for _ in range(emptySpaces):
                itemProperties.append('')
            if col in quotedFields:
                itemProperties.append('"' + str(sheet.cell_value(row,col)) + '"')
            else:
                itemProperties.append(safeIntCast(sheet.cell_value(row,col)))
            previousIndex = col
        if len(itemProperties) < 32:
            for _ in range(32 - len(itemProperties)):
                itemProperties.append('')
        line = ",".join(itemProperties).replace(";",",")
        updateFile.write(line + "\r\n")

    messagebox.showinfo(title='Success', message='IG Item Import created successfully.')
    
def determineExportType(f):
    if f[-3:] == 'xls':
        print('Input file is xls, processing as SIMPLE_EXPORT')
        return SIMPLE_EXPORT
    elif f[-3:] == 'txt':
        file = codecs.open(f, 'r', 'utf8')
        if len(file.readline().split(",")) > 20:
            return IG_EXPORT
        else:
            return SIMPLE_EXPORT
    else:
        raise IOError('UnsupportedFileExtensionError')
        
def runConversion():

    export = file_path
        
    options = {}
    options['title'] = 'Save As'
    options['initialfile'] = str(os.path.dirname(file_path)) + '/' + 'MI_IMP.txt'
    save_file = saveFile(options)
    if save_file:
        with codecs.open(save_file, 'w+', 'latin-1') as output:
            generateIGPriceUpdate(export, output)

def convertToText():
    print('simplifying to txt')
    export = file_path
    
    try:
        preParse(export)
    except UnicodeDecodeError:
        with codecs.open(file_path, 'r', 'latin-1') as export:
            preParse(export)
    fileParts = str(os.path.basename(file_path)).rsplit('.', maxsplit=1)
    options = {}
    options['title'] = 'Save As'
    options['initialfile'] = str(os.path.dirname(file_path)) + '/' + fileParts[0] + '_simplified.txt'
    options['filetypes'] = fileTypeFilters
    save_file = saveFile(options)
    if save_file:
        generateSimpleExport(save_file, excludeUnpriced=noUnpriced.get())

def convertToExcelSimple():
    print('simplifying to excel')
    export = file_path
    
    try:
        preParse(export)
    except UnicodeDecodeError:
        with codecs.open(file_path, 'r', 'latin-1') as export:
            preParse(export)
    fileParts = str(os.path.basename(file_path)).rsplit('.', maxsplit=1)
    options = {}
    options['title'] = 'Save As'
    options['initialfile'] = str(os.path.dirname(file_path)) + '/' + fileParts[0] + '_simplified.xls'
    options['filetypes'] = fileTypeFilters
    save_file = saveFile(options)
    if save_file:
        generateSimpleExcel(save_file, excludeUnpriced=noUnpriced.get())

def convertToExcelFull():
    print('converting to excel')
    export = file_path
    
    try:
        preParse(export)
    except UnicodeDecodeError:
        with codecs.open(file_path, 'r', 'latin-1') as export:
            preParse(export)
    fileParts = str(os.path.basename(file_path)).rsplit('.', maxsplit=1)
    options = {}
    options['title'] = 'Save As'
    options['initialfile'] = str(os.path.dirname(file_path)) + '/' + fileParts[0] + '_complete.xls'
    options['filetypes'] = fileTypeFilters
    save_file = saveFile(options)
    if save_file:
        generateFullExcel(save_file, excludeUnpriced=noUnpriced.get(), expandPriceLevels=expandPriceLevels.get())

def separatePriceLevels():
    pass
    
def safeIntCast(value):
    try:
        return str(int(value))
    except ValueError:
        global isMisaligned
        isMisaligned = True
        return str(value)
    
def hideButton(button):
    button.grid_remove()

def hideAllButtons():
    thatButton.grid_remove()
    simpleTxtButton.grid_remove()
    simpleXlsButton.grid_remove()
    fullXlsButton.grid_remove()
    
def showButton(button):
    button.grid()

def displayAbout():
    messagebox.showinfo(title='About', message='v0.3.16')

root = Tk()
root.option_add('*tearOff', FALSE)
root.title("Agilysys File Tools")
try:
    root.iconbitmap(default='Format_Gears.ico')
except TclError:
    print('Unable to locate icon')
    
openFileString = StringVar()
noUnpriced = BooleanVar()
expandPriceLevels = BooleanVar()

menubar = Menu(root)
menu_file = Menu(menubar)
menu_options = Menu(menubar)
menu_help = Menu(menubar)
menubar.add_cascade(menu=menu_file, label='File')
menubar.add_cascade(menu=menu_options, label='Options')
menubar.add_cascade(menu=menu_help, label='Help')

menu_file.add_command(label='Open...', command=openFile)
menu_file.add_command(label='Close', command=root.quit)

menu_options.add_checkbutton(label='Remove Unpriced Items', variable=noUnpriced, onvalue=1, offvalue=0)
menu_options.add_checkbutton(label='Separate Price Level', variable=expandPriceLevels, onvalue=1, offvalue=0)

menu_help.add_command(label='About', command=displayAbout)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=1, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(1, weight=1)

ttk.Label(mainframe, text="Input File:").grid(column=1, row=1, sticky=(N, W, E))
openFile_entry = ttk.Entry(mainframe, width=40, textvariable=openFileString)
openFile_entry.grid(column=1, row=2, sticky=(W, E))

global thatButton, simpleTxtButton, simpleXlsButton, fullXlsButton
thatButton = ttk.Button(mainframe, text='Generate IG Update', command=runConversion)
thatButton.grid(column=1, row=3)

simpleTxtButton = ttk.Button(mainframe, text='Create txt', command=convertToText)
simpleTxtButton.grid(column=1, row=4)

simpleXlsButton = ttk.Button(mainframe, text='Create Prices xls', command=convertToExcelSimple)
simpleXlsButton.grid(column=1, row=5)

fullXlsButton = ttk.Button(mainframe, text='Create Full xls', command=convertToExcelFull)
fullXlsButton.grid(column=1, row=6)

simplifyButtons = [simpleTxtButton, simpleXlsButton, fullXlsButton]

for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)

root.config(menu=menubar)
hideAllButtons()
root.mainloop()