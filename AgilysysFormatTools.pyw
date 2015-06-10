#!python3

#todo fix empty barcode description generation

import os
import sys
import re
import codecs
import pdb
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from MenuItem import MenuItem
from xlwt import Workbook, easyxf
from xlrd import open_workbook
from configparser import ConfigParser

priceArrayMatch = re.compile(r'(?<=\{)[^(\{|\})].+?(?=\})')
commaQuoteMatch = re.compile(r'((?<=")[^",\{\}]+),([^"\{\}]*(?="))')
fileTypeFilters = [('Supported Files', '.xls .xlsx .txt'), ('Text Files', '.txt'), ('Excel Files', '.xls .xlsx .csv'), ('All Files', '.*')]
app_directory = os.path.join(os.getenv('APPDATA'), 'Agilysys Format Tools')
config_file = os.path.join(app_directory, 'config.ini')
config = ConfigParser()
config.read(config_file)

IG_EXPORT = 1
SIMPLE_EXPORT = 3
UNKNOWN_EXPORT = 10
CSV_EXPORT = 2
itemList = []
itemMap = {}

def ezPrint(string):
    print(str(string))
    
def openFile(options=None):
    hideAllButtons()
    init_dir = ''
    if itemList:
        itemList.clear()
        
    try:
        init_dir = config['Paths']['last dir']
    except KeyError:
        init_dir = os.path.expanduser('~')
        
    if options == None:
        options = {}
        options['defaultextension'] = '.xls*, .txt' 
        options['filetypes'] = fileTypeFilters
        options['title'] = 'Open...'
        options['initialdir'] = init_dir
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
            
        config['Paths'] = {'last dir': file_path}
        with open(config_file, 'w') as f:
            config.write(f)
            
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
            if str(i.storeID) == '0':
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
            if item.priceLvls != "{}":
                simpleOutput.write(str(item.id) + "," + str(item.name) + "," + str(item.priceLvls) + "\r\n")
        else:
            simpleOutput.write(str(item.id) + "," + str(item.name) + "," + str(item.priceLvls) + "\r\n")
        
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
            if item.priceLvls == "{}":
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
    
    if expandPriceLevels:
        startHeaders = headers[:6]
        endHeaders = headers[7:]
        headers.clear()
        priceHeaders = []
        numberOfPriceLevels = enumeratePriceLevels()
        
        for x in range(0, (numberOfPriceLevels)):
            priceHeaders.append('Price Level ' + str(x + 1))
            #sheet.col(x + 1).width = 4260
            
        headers = startHeaders + priceHeaders + endHeaders
        
    for h,i in zip(headers, range(1,len(headers))):
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
            for p in range(1, (numberOfPriceLevels + 1)):
                if p in item.separatePriceLevels():
                    price = item.separatePriceLevels()[p]
                else:
                    price = ''
                r = p - 1
                row.write(columnMap['priceLvls'] + r, str(price))
                
            row.write(columnMap['classID'] + (numberOfPriceLevels - 1), safeIntCast(item.classID))
            row.write(columnMap['revCat'] + (numberOfPriceLevels - 1), safeIntCast(item.revCat))
            row.write(columnMap['taxGrp'] + (numberOfPriceLevels - 1), safeIntCast(item.taxGrp))
            row.write(columnMap['securityLvl'] + (numberOfPriceLevels - 1), safeIntCast(item.securityLvl))
            row.write(columnMap['reportCat'] + (numberOfPriceLevels - 1), safeIntCast(item.reportCat))
            row.write(columnMap['byWeight'] + (numberOfPriceLevels - 1), safeIntCast(item.byWeight))
            row.write(columnMap['tare'] + (numberOfPriceLevels - 1), str(item.tare))
            row.write(columnMap['sku'] + (numberOfPriceLevels - 1), str(item.sku))
            row.write(columnMap['gunCode'] + (numberOfPriceLevels - 1), str(item.gunCode))
            row.write(columnMap['cost'] + (numberOfPriceLevels - 1), str(item.cost))
            row.write(columnMap['cost'] + (numberOfPriceLevels), 'N/A')
            row.write(columnMap['pricePrompt'] + (numberOfPriceLevels - 1), safeIntCast(item.pricePrompt))
            row.write(columnMap['prntOnChk'] + (numberOfPriceLevels - 1), safeIntCast(item.prntOnChk))
            row.write(columnMap['disc'] + (numberOfPriceLevels - 1), safeIntCast(item.disc))
            row.write(columnMap['voidable'] + (numberOfPriceLevels - 1), safeIntCast(item.voidable))
            row.write(columnMap['inactive'] + (numberOfPriceLevels - 1), safeIntCast(item.inactive))
            row.write(columnMap['taxIncluded'] + (numberOfPriceLevels - 1), safeIntCast(item.taxIncluded))
            row.write(columnMap['itemGrp'] + (numberOfPriceLevels - 1), safeIntCast(item.itemGrp))
            row.write(columnMap['receipt'] + (numberOfPriceLevels - 1), str(item.receipt))
            row.write(columnMap['priceOver'] + (numberOfPriceLevels - 1), safeIntCast(item.priceOver))
            row.write(columnMap['priceOver'] + (numberOfPriceLevels), 'N/A')
            row.write(columnMap['choiceGrps'] + (numberOfPriceLevels - 1), str(item.choiceGrps))
            row.write(columnMap['ktchnPrint'] + (numberOfPriceLevels - 1), str(item.ktchnPrint))
            row.write(columnMap['covers'] + (numberOfPriceLevels - 1), str(item.covers))
            row.write(columnMap['storeID'] + (numberOfPriceLevels - 1), str(item.storeID))
        else:
            row.write(columnMap['priceLvls'], str(item.priceLvls))
            row.write(columnMap['classID'], safeIntCast(item.classID))
            row.write(columnMap['revCat'], safeIntCast(item.revCat))
            row.write(columnMap['taxGrp'], safeIntCast(item.taxGrp))
            row.write(columnMap['securityLvl'], safeIntCast(item.securityLvl))
            row.write(columnMap['reportCat'], safeIntCast(item.reportCat))
            row.write(columnMap['byWeight'], safeIntCast(item.byWeight))
            row.write(columnMap['tare'], str(item.tare))
            row.write(columnMap['sku'], str(item.sku))
            row.write(columnMap['gunCode'], str(item.gunCode))
            row.write(columnMap['cost'], str(item.cost))
            row.write(columnMap['cost'] + 1, 'N/A')
            row.write(columnMap['pricePrompt'], safeIntCast(item.pricePrompt))
            row.write(columnMap['prntOnChk'], safeIntCast(item.prntOnChk))
            row.write(columnMap['disc'], safeIntCast(item.disc))
            row.write(columnMap['voidable'], safeIntCast(item.voidable))
            row.write(columnMap['inactive'], safeIntCast(item.inactive))
            row.write(columnMap['taxIncluded'], safeIntCast(item.taxIncluded))
            row.write(columnMap['itemGrp'], safeIntCast(item.itemGrp))
            row.write(columnMap['receipt'], str(item.receipt))
            row.write(columnMap['priceOver'], safeIntCast(item.priceOver))
            row.write(columnMap['priceOver'] + 1, 'N/A')
            row.write(columnMap['choiceGrps'], str(item.choiceGrps))
            row.write(columnMap['ktchnPrint'], str(item.ktchnPrint))
            row.write(columnMap['covers'], str(item.covers))
            row.write(columnMap['storeID'], str(item.storeID))
        
        if isMisaligned:
            oopsStyle = (easyxf('pattern: pattern solid, fore_color rose'))
            row.write(1, 'X', oopsStyle)
    
    try:
        book.save(save_file)
        messagebox.showinfo(title='Success', message='Excel export created successfully.')
    except PermissionError:
        messagebox.showerror(title= 'Error', message='Unable to save file')
        
def generateCustomExcel(save_file, items=None, excludeUnpriced=True):
    items = items or itemList
    print('preparing to convert to custom Excel')
    book = Workbook()
    heading = easyxf(
        'font: bold True;'
        'alignment: horizontal center;'
        )
    oopsStyle = (easyxf('pattern: pattern solid, fore_color rose'))
    sheet = book.add_sheet('Sheet 1')
    sheet.panes_frozen = True
    sheet.remove_splits = True
    sheet.horz_split_pos = 1
    row1 = sheet.row(0)
    row2 = sheet.row(1)
    headers = ['"A"', 'ID']
    colKeys = ['modType', 'id']
    pricePos = 100
    
    for k,v in sorted(checkVarMap.items(), key=lambda x: MenuItem.attributeMap.get(x[0])):
        if str(v.get()) == '1':
            headers.append(MenuItem.textMap[k])
            colKeys.append(k)
            
    print('Maps created...')
                        
    if 'Prices' in headers:
        print('priceLvls found in headers')
        pricePos = headers.index('Prices')
        del headers[pricePos]
        del colKeys[pricePos]
        numberOfPriceLevels = enumeratePriceLevels()
        priceHeaders = []
        print(str(numberOfPriceLevels) + ' price levels found.')
        
        for x in range(numberOfPriceLevels):
            priceHeaders.append('Price Level ' + str(x + 1))
            
        priceHeaders.reverse()
        for i,x in zip(reversed(range(1, len(priceHeaders) + 1)), priceHeaders):
            headers.insert(pricePos, x)
            colKeys.insert(pricePos, ('priceLvl' + str(i)))
    
    #Write Headers
    for h,x,i in zip(headers, colKeys, range(len(headers))):
        print('writing header: ' + str(h))
        row1.write(i, h, heading)
        row2.write(i, x, heading)
        
    sheet.row(1).hidden = True
    
    #Write Rows
    for r,item in zip(range(2, len(items) + 2), items):
        global isMisaligned
        isMisaligned = False
        
        row = sheet.row(r)
        
        #Write item values to columns
        for col, key in zip(range(len(colKeys)), colKeys):
            if 'priceLvl' in key:
                #Strip number from priceLvl key and pass to index of separatePriceLevels
                p = int(key[key.find('l') + 1:])
                if p in item.separatePriceLevels():
                    price = item.separatePriceLevels()[p]
                else:
                    price = ''
                    
                row.write(col, str(price))
            else:
                if key in MenuItem.integerItems:
                    row.write(col, safeIntCast(item.__dict__[key]))
                elif key == 'modType':
                    continue
                else:
                    row.write(col, str(item.__dict__[key]))
                    
        if isMisaligned:
            row.write(0, 'X', oopsStyle)
    
    try:
        book.save(save_file)
        messagebox.showinfo(title='Success', message='Custom Excel export created successfully')
    except PermissionError:
        messagebox.showerror(title= 'Error', message='Unable to save file')
        
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
            elif sheet.cell_value(1,0) == 'modType':
                generateCustomIGUpdate(book, updateFile)
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
    
def generateCustomIGUpdate(book, updateFile):
    print('preparing to generate IG Update file from custom xls')
    sheet = book.sheet_by_index(0)
    quotedFields = (3, 4, 5, 26)
            
    for row in range(2, sheet.nrows):
        itemProperties = []
        itemPropertyMap = {}
        priceLevelMap = {}
        updateType = sheet.cell_value(row,0)
        if updateType != 'A' and updateType != 'U' and updateType != 'D' and updateType != 'X':
            continue
        elif updateType == 'X':
            messagebox.showwarning(title='File Error', 
                message='One or more lines are not aligned properly.\nPlease correct and retry.')
            return
        else:
            itemProperties.append('"' + updateType + '"')
            
        for col in range(1, sheet.ncols):
            key = sheet.cell_value(1, col)
            if 'priceLvl' in key:
                priceLevelMap[key] = sheet.cell_value(row, col)
            itemPropertyMap[key] = (sheet.cell_value(row, col))
            
        if priceLevelMap:
            itemPropertyMap['priceLvls'] = rebuildPriceRecord(priceLevelMap)
            
        for k,v in sorted(MenuItem.attributeMap.items(), key=lambda x: x[1]):
            if k in itemPropertyMap.keys():
                if v in quotedFields:
                    itemProperties.append('"' + str(itemPropertyMap[k]) + '"')
                else:
                    itemProperties.append(safeIntCast(itemPropertyMap[k]))
            else:
                itemProperties.append('')
                
        itemProperties.append('')
        itemProperties.append('') #appending two additional comma for parity
        line = ','.join(itemProperties).replace(';',',')
        updateFile.write(line + '\r\n')

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

def convertToExcelCustom():
    print('customizing for excel')
    export = file_path
    
    try:
        preParse(export)
    except UnicodeDecodeError:
        with codecs.open(file_path, 'r', 'latin-1') as export:
            preParse(export)
            
    displayColumnSelection()
    root.wait_window(csWin)
    fileParts = str(os.path.basename(file_path)).rsplit('.', maxsplit=1)
    options = {}
    options['title'] = 'Save As'
    options['initialfile'] = str(os.path.dirname(file_path)) + '/' + fileParts[0] + '_custom.xls'
    options['filetypes'] = fileTypeFilters
    save_file = saveFile(options)
    if save_file:
        generateCustomExcel(save_file, excludeUnpriced=noUnpriced.get())
        
def rebuildPriceRecord(priceMap):
    prices = []
    isNegative = False
    for k,v in sorted(priceMap.items()):
        if v != '':
            #Strip number from priceLvl key and pass to index of separatePriceLevels
            p = str(k[k.find('l') + 1:])
            level = ''
            if '(' in str(v) or '(' in str(v):
                isNegative = True
            price = '{0:.2f}'.format(float(str(v).strip('$()')))
            if isNegative == True:
                level = p + ',($' + price + ')'
            else:
                level = p + ',$' + price
            prices.append(level)
    
    record = '{' + ','.join(prices) + '}'
    return record

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
    customXlsButton.grid_remove()
    
def showButton(button):
    button.grid()

def displayAbout():
    messagebox.showinfo(title='About', message='v0.6.5')
    
def displayColumnSelection():
    global csWin
    csWin = Toplevel(root)
    colSelectFrame = ttk.Frame(csWin)
    colSelectFrame.grid(column=0, row=1, sticky=(N,S,E,W))
    colSelectFrame.columnconfigure(0, weight=1)
    colSelectFrame.rowconfigure(1, weight=1)
    
    global checkVarMap
    row_count = 0
    counter = 0
    
    for k,v in sorted(MenuItem.attributeMap.items(), key=lambda x: x[1]):
        col = 0
        if k not in checkVarMap:
            checkVarMap[k] = Variable()
        if k != 'id':
            if counter % 2 == 0:
                col = 0
                row_count += 1
            else:
                col = 3
            l = ttk.Checkbutton(colSelectFrame, text=MenuItem.textMap[k], variable=checkVarMap[k]).grid(column=col, row=row_count, sticky=(N,W))
            counter += 1
            
    ttk.Button(colSelectFrame, text='OK', command=csWin.destroy).grid(column=1, row=100)
    return

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
colPrice = BooleanVar()
checkVarMap = {}

if not os.path.exists(app_directory):
    os.mkdir(app_directory)

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

customXlsButton = ttk.Button(mainframe, text='Create Custom xls', command=convertToExcelCustom)
customXlsButton.grid(column=1, row=7)

simplifyButtons = [simpleTxtButton, simpleXlsButton, fullXlsButton, customXlsButton]

for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)

root.config(menu=menubar)
hideAllButtons()
root.mainloop()