#!python3

# todo fix empty barcode description generation
# todo add sentinel item generation for import verification

import os
import codecs
import datetime
import pdb
import time
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from configparser import ConfigParser

from xlwt import Workbook, easyxf
from xlrd import open_workbook

from MenuItem import MenuItem

__version__ = 'v0.10.7'

priceArrayMatch = re.compile(r'(?<=\{)[^(\{|\})].+?(?=\})')
commaQuoteMatch = re.compile(r'((?<=")[^",\{\}]+),([^"\{\}]*(?="))')
file_type_filters = [('Supported Files', '.xls .xlsx .txt'),
                     ('Text Files', '.txt'),
                     ('Excel Files', '.xls .xlsx .csv'), ('All Files', '.*')]
app_directory = os.path.join(os.getenv('APPDATA'), 'Agilysys Format Tools')
config_file = os.path.join(app_directory, 'config.ini')
log_file = os.path.join(app_directory, 'errors.log')
config = ConfigParser()
config.read(config_file)

IG_EXPORT = 1
SIMPLE_EXPORT = 3
UNKNOWN_EXPORT = 10
CSV_EXPORT = 2
itemList = []
itemMap = {}


def ez_print(string):
    print(str(string))


def open_file(options=None):
    hideAllButtons()
    init_dir = ''
    if itemList:
        itemList.clear()

    try:
        init_dir = config['Paths']['last dir']
    except KeyError:
        init_dir = os.path.expanduser('~')

    if not options:
        options = dict()
        options['defaultextension'] = '.xls*, .txt'
        options['filetypes'] = file_type_filters
        options['title'] = 'Open...'
        options['initialdir'] = init_dir
    file_opt = options
    global file_path
    file_path = filedialog.askopenfilename(**file_opt)
    options = None
    if not file_path or file_path == "":
        print("No file selected")
        return

    try:
        if determineExportType(file_path) == IG_EXPORT:
            for button in simplifyButtons:
                showButton(button)
        else:
            showButton(ig_button)

        config['Paths'] = {'last dir': file_path}
        with open(config_file, 'w') as f:
            config.write(f)

    except IOError:
        messagebox.showinfo(title='Oops',
                            message='This file is not supported.')
        print('{0}\n{1}'.format(sys.exc_info()[0], sys.exc_info()[1]))
        return


def saveFile(options):
    file_opt = options
    save_file = filedialog.asksaveasfilename(**file_opt)
    if save_file is None or save_file == "":
        print("No file selected")
    return save_file


def fixArray(match, ):
    match = str(match.group(0))
    return match.replace(",", ";")


def pre_parse_ig_file(file_name):
    with codecs.open(file_name, 'r', 'latin-1') as export:
        print('pre-parse initiated')
        for line in export:
            itemDetails = re.sub(priceArrayMatch, fixArray, line)
            itemDetails = re.sub(commaQuoteMatch, fixArray, itemDetails)
            item = itemDetails.split(",")
            try:
                i = MenuItem(
                    item[1], item[2], item[3], item[4], item[5],
                    item[6], item[7], item[8], item[9], item[10],
                    item[11], item[12], item[13], item[14], item[15],
                    item[16], item[18], item[19], item[20], item[21],
                    item[22], item[23], item[24], item[25], item[26],
                    item[28], item[29], item[30], item[31]
                )
            except IndexError:
                response = None
                if item[1]:
                    response = messagebox.askokcancel(
                        title='Error reading file',
                        message='Unable to parse line {0}'.format(item[1]))
                else:
                    response = messagebox.askokcancel(
                        title='Error reading file',
                        message='Unable to parse info:\n' + line)
                if not response:
                    os._exit(1)
            # Skip Store Items
            if str(i.storeID) == '0':
                itemList.append(i)
            else:
                continue
    print("parse completed")


def enumeratePriceLevels():
    numberOfPriceLevels = 0
    # Seems to still miss price levels above range and has potential for issues
    #  if no item contains all prices
    for item in itemList:
        levels = item.separatePriceLevels()
        if max(k for k, _ in levels.items()) > numberOfPriceLevels:
            numberOfPriceLevels = max(k for k, _ in levels.items())
    return numberOfPriceLevels


# noinspection PyShadowingNames
def generateFullExcel(save_file, items=None,
                      excludeUnpriced=True, expandPriceLevels=False):
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

        for x in range(0, numberOfPriceLevels):
            priceHeaders.append('Price Level ' + str(x + 1))
            # sheet.col(x + 1).width = 4260

        headers = startHeaders + priceHeaders + endHeaders

    for h, i in zip(headers, range(1, len(headers))):
        if i < 3:
            sheet.row(1).set_cell_boolean(i, True)
        else:
            sheet.row(1).set_cell_boolean(i, False)
        row1.write(i, h, heading)

    for i, item in zip(range(2, len(items) + 2), items):
        global isMisaligned
        isMisaligned = False
        columnMap = MenuItem.attributeMap

        row = sheet.row(i)
        # row.write(2, int(item.id))
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

            row.write(columnMap['classID'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.classID))
            row.write(columnMap['revCat'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.revCat))
            row.write(columnMap['taxGrp'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.taxGrp))
            row.write(columnMap['securityLvl'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.securityLvl))
            row.write(columnMap['reportCat'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.reportCat))
            row.write(columnMap['byWeight'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.byWeight))
            row.write(columnMap['tare'] + (numberOfPriceLevels - 1),
                      str(item.tare))
            row.write(columnMap['sku'] + (numberOfPriceLevels - 1),
                      str(item.sku))
            row.write(columnMap['gunCode'] + (numberOfPriceLevels - 1),
                      str(item.gunCode))
            row.write(columnMap['cost'] + (numberOfPriceLevels - 1),
                      str(item.cost))
            row.write(columnMap['cost'] + numberOfPriceLevels, 'N/A')
            row.write(columnMap['pricePrompt'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.pricePrompt))
            row.write(columnMap['prntOnChk'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.prntOnChk))
            row.write(columnMap['disc'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.disc))
            row.write(columnMap['voidable'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.voidable))
            row.write(columnMap['inactive'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.inactive))
            row.write(columnMap['taxIncluded'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.taxIncluded))
            row.write(columnMap['itemGrp'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.itemGrp))
            row.write(columnMap['receipt'] + (numberOfPriceLevels - 1),
                      str(item.receipt))
            row.write(columnMap['priceOver'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.priceOver))
            row.write(columnMap['priceOver'] + numberOfPriceLevels, 'N/A')
            row.write(columnMap['choiceGrps'] + (numberOfPriceLevels - 1),
                      str(item.choiceGrps))
            row.write(columnMap['ktchnPrint'] + (numberOfPriceLevels - 1),
                      str(item.ktchnPrint))
            row.write(columnMap['covers'] + (numberOfPriceLevels - 1),
                      str(item.covers))
            row.write(columnMap['storeID'] + (numberOfPriceLevels - 1),
                      str(item.storeID))
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
        messagebox.showinfo(title='Success',
                            message='Excel export created successfully.')
    except PermissionError:
        messagebox.showerror(title='Error', message='Unable to save file')


# noinspection PyShadowingNames,PyShadowingNames
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

    for k, v in sorted(checkVarMap.items(),
                       key=lambda x: MenuItem.attributeMap.get(x[0])):
        if str(v.get()) == '1':
            headers.append(MenuItem.textMap[k])
            colKeys.append(k)

    if 'Prices' in headers:
        pricePos = headers.index('Prices')
        del headers[pricePos]
        del colKeys[pricePos]
        numberOfPriceLevels = enumeratePriceLevels()
        priceHeaders = []

        for x in range(numberOfPriceLevels):
            priceHeaders.append('Price Level ' + str(x + 1))

        priceHeaders.reverse()
        for i, x in zip(reversed(range(1, len(priceHeaders) + 1)),
                        priceHeaders):
            headers.insert(pricePos, x)
            colKeys.insert(pricePos, ('priceLvl' + str(i)))

    # Write Headers
    for h, x, i in zip(headers, colKeys, range(len(headers))):
        row1.write(i, h, heading)
        row2.write(i, x, heading)

    sheet.row(1).hidden = True

    # Write Rows
    for r, item in zip(range(2, len(items) + 2), items):
        global isMisaligned
        isMisaligned = False

        row = sheet.row(r)

        # Write item values to columns
        for col, key in enumerate(colKeys):
            if 'priceLvl' in key:
                # Strip number from priceLvl key and pass to index of separatePriceLevels
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
        messagebox.showinfo(title='Success',
                            message='Custom Excel export created successfully')
    except PermissionError:
        messagebox.showerror(title='Error', message='Unable to save file')


def generateIGUpdate(book, updateFile):
    print('preparing to generate IG Update file')
    sheet = book.sheet_by_index(0)
    includeColumns = set()
    quotedFields = (3, 4, 5, 26)

    for col in range(3, sheet.ncols):
        if sheet.cell_value(1, col):
            includeColumns.add(col)

    includeColumns = sorted(includeColumns)

    for row in range(2, sheet.nrows):
        itemProperties = []
        updateType = sheet.cell_value(row, 1)
        if updateType != 'A' and updateType != 'U' and\
                        updateType != 'D' and updateType != 'X':
            continue
        elif updateType == 'X':
            messagebox.showwarning(
                title='File Error',
                message='One or more lines are not aligned properly.'
                        '\nPlease correct and retry.')
            return
        else:
            itemProperties.append('"' + str(sheet.cell_value(row, 1)) + '"')
        itemProperties.append(safeIntCast((sheet.cell_value(row, 2))))
        previousIndex = 2
        for col in includeColumns:
            emptySpaces = col - previousIndex - 1
            for _ in range(emptySpaces):
                itemProperties.append('')
            if col in quotedFields:
                itemProperties.append(
                    '"' + str(sheet.cell_value(row, col)) + '"')
            else:
                itemProperties.append(safeIntCast(sheet.cell_value(row, col)))
            previousIndex = col
        if len(itemProperties) < 32:
            for _ in range(32 - len(itemProperties)):
                itemProperties.append('')
        line = ",".join(itemProperties).replace(";", ",")
        updateFile.write(line + "\r\n")

    messagebox.showinfo(title='Success',
                        message='IG Item Import created successfully.')


def generateCustomIGUpdate(book, updateFile):
    sheet = book.sheet_by_index(0)
    quotedFields = (3, 4, 5, 26)
    updated_items = 0

    for row in range(2, sheet.nrows):
        itemProperties = []
        itemPropertyMap = {}
        priceLevelMap = {}
        update_type = sheet.cell_value(row, 0)
        if update_type != 'A' and update_type != 'U' and \
                        update_type != 'D' and update_type != 'X':
            continue
        elif update_type == 'X':
            messagebox.showwarning(
                title='File Error',
                message='One or more lines are not aligned properly.'
                        '\nPlease correct and retry.')
            return
        else:
            itemProperties.append('"{0}"'.format(update_type))
            updated_items += 1

        for col in range(1, sheet.ncols):
            key = sheet.cell_value(1, col)
            if 'priceLvl' in key:
                priceLevelMap[key] = sheet.cell_value(row, col)
            itemPropertyMap[key] = (sheet.cell_value(row, col))

        if priceLevelMap:
            itemPropertyMap['priceLvls'] = rebuildPriceRecord(priceLevelMap)

        for key, position in sorted(MenuItem.attributeMap.items(),
                                    key=lambda x: x[1]):
            if key in itemPropertyMap.keys():
                if position in quotedFields:
                    itemProperties.append('"{0}"'.format(itemPropertyMap[key]))
                else:
                    itemProperties.append(safeIntCast(itemPropertyMap[key]))
            else:
                itemProperties.append('')

        line = ','.join(itemProperties).replace(';', ',')
        updateFile.write(line + '\r\n')

    # adding sentinel item
    updateFile.write(
        '"A",7110001,"{0}",,,,{{1,$0.00}},,,,,,,,,,,,,,,,,,,,,,,,,'.format(
            time.strftime('%c', time.localtime())))
    if updated_items:
        messagebox.showinfo(title='Success',
                            message='IG Item Import created successfully.')
    else:
        messagebox.showinfo(
            title='Oops',
            message="No items processed.  "
                    "Did you remember to put a 'U' or 'A'"
                    " in the first column?")


def determineExportType(filename):
    if filename[-3:] == 'xls':
        print('Input file is xls, processing as SIMPLE_EXPORT')
        return SIMPLE_EXPORT
    elif filename[-3:] == 'txt':
        file = codecs.open(filename, 'r', 'utf8')
        if len(file.readline().split(",")) > 20:
            return IG_EXPORT
        else:
            return SIMPLE_EXPORT
    else:
        raise IOError('UnsupportedFileExtensionError')


# Use lambda: runConversion(x) to pass values with button command call
def convert_to_ig_format():
    export = file_path

    options = {
        'title': 'Save As',
        'initialfile': os.path.join(os.path.dirname(file_path), 'MI_IMP.txt')
        }

    save_file = saveFile(options)
    if save_file:
        with codecs.open(save_file, 'w+', 'latin-1') as output:
            generateIGPriceUpdate(export, output)


def convert_to_excel(type='custom'):
    """
    Initiates conversion from IG Format to Excel spreadsheet

    Keyword arguments:
    type -- full or custom spreadsheet

    :return:

    """
    export = file_path

    try:
        pre_parse_ig_file(export)
    except UnicodeDecodeError:
        with codecs.open(file_path, 'r', 'latin-1') as export:
            pre_parse_ig_file(export)

    file_parts = str(os.path.basename(file_path)).rsplit('.', maxsplit=1)

    if type == 'complete':
        default_filename = file_parts[0] + '_complete.xls'
    elif type == 'custom':
        default_filename = file_parts[0] + '_custom.xls'
        displayColumnSelection()
        root.wait_window(csWin)
    else:
        raise TypeError("Can't convert to excel using {0} type".format(type))

    options = {'title': 'Save As',
               'initialfile': os.path.join(os.path.dirname(file_path),
                                           default_filename),
               'filetypes': file_type_filters}
    save_file = saveFile(options)

    if save_file and type == 'complete':
        generateFullExcel(save_file, excludeUnpriced=noUnpriced.get(),
                          expandPriceLevels=expandPriceLevels.get())
    elif save_file and type == 'custom':
        generateCustomExcel(save_file, excludeUnpriced=noUnpriced.get())


def rebuildPriceRecord(priceMap):
    prices = []
    isNegative = False
    for k, v in sorted(priceMap.items()):
        if v != '':
            # Strip number from priceLvl and pass to index of separatePriceLevels
            p = str(k[k.find('l') + 1:])
            level = ''
            if '(' in str(v) or '(' in str(v):
                isNegative = True
            price = '{0:.2f}'.format(float(str(v).strip('$(){}')))
            if isNegative:
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
    ig_button.grid_remove()
    fullXlsButton.grid_remove()
    customXlsButton.grid_remove()


def showButton(button):
    button.grid()


def displayAbout():
    messagebox.showinfo(title='About', message=__version__)


def displayColumnSelection():
    """
    Display sub window for selecting properties to include in Excel spreadsheet
    :return:
    """

    global csWin
    csWin = Toplevel(root)
    colSelectFrame = ttk.Frame(csWin)
    colSelectFrame.grid(column=0, row=1, sticky=(N, S, E, W))
    colSelectFrame.columnconfigure(0, weight=1)
    colSelectFrame.rowconfigure(1, weight=1)

    global checkVarMap
    row_count = 0
    counter = 0

    for k, v in sorted(MenuItem.attributeMap.items(), key=lambda x: x[1]):
        col = 0
        if k not in checkVarMap:
            checkVarMap[k] = IntVar()
        if k != 'id' and k[:-3] != 'reserved':
            if counter % 2 == 0:
                col = 0
                row_count += 1
            else:
                col = 3
            l = ttk.Checkbutton(colSelectFrame, text=MenuItem.textMap[k],
                                variable=checkVarMap[k]).grid(column=col,
                                                              row=row_count,
                                                              sticky=(N, W))
            counter += 1

    ttk.Button(colSelectFrame, text='OK',
               command=csWin.destroy).grid(column=1, row=100)
    ttk.Button(colSelectFrame, text='Select All',
               command=select_all_properties)
    return


def select_all_properties():
    """
    Selects all properties in checkVarMap
    :return:
    """
    global checkVarMap

    for k,v in MenuItem.attributeMap.items():
        if k not in checkVarMap:
            checkVarMap[k] = IntVar()
        checkVarMap[k].set(1)


def show_var_states(ttk_var):
    if type(ttk_var) is dict:
        for k,v in ttk_var.items():
            print('{0}: {1}'.format(k, v.get()))
    elif type(ttk_var) is list:
        for item in ttk_var:
            print('{0}: {1}'.format(item, item.get()))
    else:
        print('{0}: {1}'.format(ttk_var, ttk_var.get()))

root = Tk()
root.option_add('*tearOff', FALSE)
root.title("Agilysys Import Export Tool")
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

FSOCK = open(log_file, 'a+')
sys.stderr = FSOCK

menubar = Menu(root)
menu_file = Menu(menubar)
menu_options = Menu(menubar)
menu_help = Menu(menubar)
menubar.add_cascade(menu=menu_file, label='File')
menubar.add_cascade(menu=menu_options, label='Options')
menubar.add_cascade(menu=menu_help, label='Help')

menu_file.add_command(label='Open...', command=open_file)
menu_file.add_command(label='Close', command=root.quit)

menu_options.add_checkbutton(label='Remove Unpriced Items',
                             variable=noUnpriced, onvalue=1, offvalue=0)
menu_options.add_checkbutton(label='Separate Price Level',
                             variable=expandPriceLevels, onvalue=1, offvalue=0)
menu_options.add_command(label='Select All', command=select_all_properties)
menu_options.add_command(label='Display Vars',
                         command=lambda: show_var_states(checkVarMap))

menu_help.add_command(label='About', command=displayAbout)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=1, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(1, weight=1)

ttk.Label(mainframe, text="Input File:").grid(
    column=1, row=1, sticky=(N, W, E))
openFile_entry = ttk.Entry(mainframe, width=40, textvariable=openFileString)
openFile_entry.grid(column=1, row=2, sticky=(W, E))

# Use lambda: runConversion(x) to pass values with button command call
ig_button = ttk.Button(mainframe, text='Generate IG Update',
                       command=convert_to_ig_format)
ig_button.grid(column=1, row=3)

fullXlsButton = ttk.Button(mainframe, text='Create Full xls',
                           command=lambda: convert_to_excel('complete'))
fullXlsButton.grid(column=1, row=6)

customXlsButton = ttk.Button(mainframe, text='Create Custom xls',
                             command=lambda: convert_to_excel('custom'))
customXlsButton.grid(column=1, row=7)

simplifyButtons = [fullXlsButton, customXlsButton]

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)

root.config(menu=menubar)
hideAllButtons()
root.mainloop()
