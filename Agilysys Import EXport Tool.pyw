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
pretty_print_text_map = MenuItem.pretty_print_text_map
ig_field_sequence = MenuItem.ig_field_sequence
integer_fields = MenuItem.integer_fields
string_fields = MenuItem.string_fields

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
EXCEL_FILE = 3
itemList = []
itemMap = {}


def ez_print(string):
    print(str(string))


def open_file(options=None):
    hide_all_buttons()
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
    global in_file
    in_file = filedialog.askopenfilename(**file_opt)
    options = None
    if not in_file or in_file == "":
        print("No file selected")
        return

    try:
        if get_file_type(in_file) == IG_EXPORT:
            for button in simplifyButtons:
                show_button(button)
        else:
            show_button(button_ig)

        config['Paths'] = {'last dir': in_file}
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
            if str(i.store_id) == '0':
                itemList.append(i)
            else:
                continue
    print("parse completed")


# Might be worth moving this to MenuItem class
def enumeratePriceLevels():
    """Returns total number of price levels"""
    for item in itemList:
        levels = item.separate_price_levels()
        if max(k for k, _ in levels.items()) > numberOfPriceLevels:
            numberOfPriceLevels = max(k for k, _ in levels.items())
    return numberOfPriceLevels


# noinspection PyShadowingNames
def generateFullExcel(save_file, items=None,
                      excludeUnpriced=True, expandPriceLevels=False):
    """Legacy function to convert IG Export to complete Excel spreadsheet."""
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
        global row_is_misaligned
        row_is_misaligned = False

        row = sheet.row(i)
        # row.write(2, int(item.id))
        row.write(ig_field_sequence['id'], int(item.id))
        row.write(ig_field_sequence['name'], str(item.name))
        row.write(ig_field_sequence['abbr1'], str(item.abbr1))
        row.write(ig_field_sequence['abbr2'], str(item.abbr2))
        row.write(ig_field_sequence['print_label'], str(item.printerLabel))

        if expandPriceLevels:
            for p in range(1, (numberOfPriceLevels + 1)):
                if p in item.separate_price_levels():
                    price = item.separate_price_levels()[p]
                else:
                    price = ''
                r = p - 1
                row.write(ig_field_sequence['price_levels'] + r, str(price))

            row.write(ig_field_sequence['product_class'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.classID))
            row.write(ig_field_sequence['revenue_category'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.revCat))
            row.write(ig_field_sequence['tax_group'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.taxGrp))
            row.write(ig_field_sequence['security_level'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.securityLvl))
            row.write(ig_field_sequence['report_category'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.reportCat))
            row.write(ig_field_sequence['sell_by_weight'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.byWeight))
            row.write(ig_field_sequence['tare'] + (numberOfPriceLevels - 1),
                      str(item.tare))
            row.write(ig_field_sequence['sku'] + (numberOfPriceLevels - 1),
                      str(item.sku))
            row.write(ig_field_sequence['gun_code'] + (numberOfPriceLevels - 1),
                      str(item.gunCode))
            row.write(ig_field_sequence['cost'] + (numberOfPriceLevels - 1),
                      str(item.cost))
            row.write(ig_field_sequence['cost'] + numberOfPriceLevels, 'N/A')
            row.write(ig_field_sequence['prompt_for_price'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.pricePrompt))
            row.write(ig_field_sequence['print_on_check'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.prntOnChk))
            row.write(ig_field_sequence['is_discountable'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.disc))
            row.write(ig_field_sequence['voidable'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.voidable))
            row.write(ig_field_sequence['inactive'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.inactive))
            row.write(ig_field_sequence['tax_included'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.taxIncluded))
            row.write(ig_field_sequence['item_group'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.itemGrp))
            row.write(ig_field_sequence['receipt_text'] + (numberOfPriceLevels - 1),
                      str(item.receipt))
            row.write(ig_field_sequence['allow_price_override'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.priceOver))
            row.write(ig_field_sequence['allow_price_override'] + numberOfPriceLevels, 'N/A')
            row.write(ig_field_sequence['choice_groups'] + (numberOfPriceLevels - 1),
                      str(item.choiceGrps))
            row.write(ig_field_sequence['kitchen_printers'] + (numberOfPriceLevels - 1),
                      str(item.ktchnPrint))
            row.write(ig_field_sequence['covers'] + (numberOfPriceLevels - 1),
                      str(item.covers))
            row.write(ig_field_sequence['store_id'] + (numberOfPriceLevels - 1),
                      str(item.storeID))
        else:
            row.write(ig_field_sequence['price_levels'], str(item.priceLvls))
            row.write(ig_field_sequence['product_class'], safeIntCast(item.classID))
            row.write(ig_field_sequence['revenue_category'], safeIntCast(item.revCat))
            row.write(ig_field_sequence['tax_group'], safeIntCast(item.taxGrp))
            row.write(ig_field_sequence['security_level'], safeIntCast(item.securityLvl))
            row.write(ig_field_sequence['report_category'], safeIntCast(item.reportCat))
            row.write(ig_field_sequence['sell_by_weight'], safeIntCast(item.byWeight))
            row.write(ig_field_sequence['tare'], str(item.tare))
            row.write(ig_field_sequence['sku'], str(item.sku))
            row.write(ig_field_sequence['gun_code'], str(item.gunCode))
            row.write(ig_field_sequence['cost'], str(item.cost))
            row.write(ig_field_sequence['cost'] + 1, 'N/A')
            row.write(ig_field_sequence['prompt_for_price'], safeIntCast(item.pricePrompt))
            row.write(ig_field_sequence['print_on_check'], safeIntCast(item.prntOnChk))
            row.write(ig_field_sequence['is_discountable'], safeIntCast(item.disc))
            row.write(ig_field_sequence['voidable'], safeIntCast(item.voidable))
            row.write(ig_field_sequence['inactive'], safeIntCast(item.inactive))
            row.write(ig_field_sequence['tax_included'], safeIntCast(item.taxIncluded))
            row.write(ig_field_sequence['item_group'], safeIntCast(item.itemGrp))
            row.write(ig_field_sequence['receipt_text'], str(item.receipt))
            row.write(ig_field_sequence['allow_price_override'], safeIntCast(item.priceOver))
            row.write(ig_field_sequence['allow_price_override'] + 1, 'N/A')
            row.write(ig_field_sequence['choice_groups'], str(item.choiceGrps))
            row.write(ig_field_sequence['kitchen_printers'], str(item.ktchnPrint))
            row.write(ig_field_sequence['covers'], str(item.covers))
            row.write(ig_field_sequence['store_id'], str(item.storeID))

        if row_is_misaligned:
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

    for k, v in sorted(checkbox_variable_map.items(),
                       key=lambda x: ig_field_sequence.get(x[0])):
        if str(v.get()) == '1':
            headers.append(pretty_print_text_map[k])
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
        global row_is_misaligned
        row_is_misaligned = False

        row = sheet.row(r)

        # Write item values to columns
        for col, key in enumerate(colKeys):
            if 'priceLvl' in key:
                # Strip number from priceLvl key and pass to index of separatePriceLevels
                p = int(key[key.find('l') + 1:])
                if p in item.separate_price_levels():
                    price = item.separate_price_levels()[p]
                else:
                    price = ''

                row.write(col, str(price))
            else:
                if key in integer_fields:
                    row.write(col, safeIntCast(item.__dict__[key]))
                elif key == 'modType':
                    continue
                else:
                    row.write(col, str(item.__dict__[key]))

        if row_is_misaligned:
            row.write(0, 'X', oopsStyle)

    try:
        book.save(save_file)
        messagebox.showinfo(title='Success',
                            message='Custom Excel export created successfully')
    except PermissionError:
        messagebox.showerror(title='Error', message='Unable to save file')


def generateIGUpdate(excel_file, ig_text_file):
    """Legacy function for generating IG import from full Excel workbook."""
    print('preparing to generate IG Update file using legacy function')

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
        ig_text_file.write(line + "\r\n")

    messagebox.showinfo(title='Success',
                        message='IG Item Import created successfully.')


def generate_ig_import(book, ig_text_file):
    """Generates IG Import File from custom Excel workbook.

    Keyword arguments:
    book -- Excel workbook (custom)
    ig_text_file -- text file to be generated for Agilysys
    """
    sheet = book.sheet_by_index(0)
    quotedFields = (3, 4, 5, 26)
    updated_items = 0

    for row in range(2, sheet.nrows):
        item_properties = []
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
            item_properties.append('"{0}"'.format(update_type))
            updated_items += 1

        for col in range(1, sheet.ncols):
            key = sheet.cell_value(1, col)
            if 'priceLvl' in key:
                priceLevelMap = {key: sheet.cell_value(row, col)}
            itemPropertyMap= {key: (sheet.cell_value(row, col))}

        if priceLevelMap:
            itemPropertyMap['price_levels'] = build_ig_price_array(priceLevelMap)

        for key, position in sorted(ig_field_sequence.items(),
                                    key=lambda x: x[1]):
            if key in itemPropertyMap.keys():
                if position in quotedFields:
                    item_properties.append('"{0}"'.format(itemPropertyMap[key]))
                else:
                    item_properties.append(safeIntCast(itemPropertyMap[key]))
            else:
                item_properties.append('')

        line = ','.join(item_properties).replace(';', ',')
        ig_text_file.write(line + '\r\n')

    # adding sentinel item
    ig_text_file.write(
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


def get_file_type(filename):
    """Return the file type for filename"""
    if filename.rsplit('.', maxsplit=1)[1] == 'xls':
        print('Input file is xls, processing as EXCEL_FILE')
        return EXCEL_FILE
    elif filename.rsplit('.', maxsplit=1)[1] == 'txt':
        file = codecs.open(filename, 'r', 'utf8')
        if len(file.readline().split(",")) > 20:
            return IG_EXPORT
        else:
            raise IOError('File is corrupt or contains incomplete data.')
    else:
        raise IOError('Unsupported file extension')


def convert_to_ig_format():
    """Initiates conversion from Excel spreadsheet to IG text file"""
    text_file = in_file
    options = {
        'title': 'Save As',
        'initialfile': os.path.join(os.path.dirname(in_file), 'MI_IMP.txt')
        }
    save_file = saveFile(options)
    if save_file:
        with codecs.open(save_file, 'w+', 'latin-1') as text_file:
            file_extension = file.rsplit('.', maxsplit=1)[1]

        if file_extension == 'xls' or file_extension == 'xlsx':
            book = open_workbook(file)
            sheet = book.sheet_by_index(0)
            if book.nsheets > 1:
                generate_standardized_ig_imports(book, text_file)
            elif sheet.cell_value(1, 0) == 'modType':
                generate_ig_import(book, text_file)
            else:
                generateIGUpdate(book, text_file)


def convert_to_excel(type='custom'):
    """Initiates conversion from IG Format to Excel spreadsheet

    Keyword arguments:
    type -- full or custom spreadsheet
    """
    export = in_file

    try:
        pre_parse_ig_file(export)
    except UnicodeDecodeError:
        with codecs.open(in_file, 'r', 'latin-1') as export:
            pre_parse_ig_file(export)

    file_parts = str(os.path.basename(in_file)).rsplit('.', maxsplit=1)

    if type == 'complete':
        default_filename = file_parts[0] + '_complete.xls'
    elif type == 'custom':
        default_filename = file_parts[0] + '_custom.xls'
        display_item_property_selections()
        root.wait_window(csWin)
    else:
        raise TypeError("Can't convert to excel using {0} type".format(type))

    options = {'title': 'Save As',
               'initialfile': os.path.join(os.path.dirname(in_file),
                                           default_filename),
               'filetypes': file_type_filters}
    save_file = saveFile(options)

    if save_file and type == 'complete':
        generateFullExcel(save_file, excludeUnpriced=noUnpriced.get(),
                          expandPriceLevels=expandPriceLevels.get())
    elif save_file and type == 'custom':
        generateCustomExcel(save_file, excludeUnpriced=noUnpriced.get())


def build_ig_price_array(price_map):
    """Returns IG price array from dictionary of price levels."""

    prices = []
    for price_level, price in sorted(price_map.items()):
        if price != '':
            # Extract number from priceLvl
            level = str(price_level[price_level.find('l') + 1:])
            price_sequence = ''
            if '(' in str(price) or '(' in str(price):
                price_is_negative = True
            price = '{0:.2f}'.format(float(str(price).strip('$(){}')))
            if price_is_negative:
                price_sequence = level + ',($' + price + ')'
            else:
                price_sequence = level + ',$' + price
            prices.append(price_sequence)

    record = '{' + ','.join(prices) + '}'
    return record


def safeIntCast(value):
    """
    Attempts to cast value to an integer, falls back to a string if it fails.
    Will also set row_is_misaligned to True if int cast fails.
    """

    try:
        return str(int(value))
    except ValueError:
        global row_is_misaligned
        row_is_misaligned = True
        return str(value)


def hide_button(button):
    button.grid_remove()


def hide_all_buttons():
    for b in hideable_buttons:
        hide_button(b)


def show_button(button):
    button.grid()


def display_about():
    messagebox.showinfo(title='About', message=__version__)


def display_item_property_selections():
    """
    Display sub window for selecting properties to include in Excel spreadsheet
    """
    global csWin
    csWin = Toplevel(root)
    colSelectFrame = ttk.Frame(csWin)
    colSelectFrame.grid(column=0, row=1, sticky=(N, S, E, W))
    colSelectFrame.columnconfigure(0, weight=1)
    colSelectFrame.rowconfigure(1, weight=1)

    global checkbox_variable_map
    row_count = 0
    counter = 0

    for k, v in sorted(ig_field_sequence.items(), key=lambda x: x[1]):
        col = 0
        if k not in checkbox_variable_map:
            checkbox_variable_map[k] = IntVar()
        if k != 'id' and k[:-3] != 'reserved':
            if counter % 2 == 0:
                col = 0
                row_count += 1
            else:
                col = 3
            l = ttk.Checkbutton(colSelectFrame, text=pretty_print_text_map[k],
                                variable=checkbox_variable_map[k]).grid(
                                    column=col, row=row_count, sticky=(N, W))
            counter += 1

    ttk.Button(colSelectFrame, text='OK',
               command=csWin.destroy).grid(column=1, row=100)
    ttk.Button(colSelectFrame, text='Select All',
               command=select_all_properties)
    return


def select_all_properties():
    """Selects all properties in checkVarMap"""
    global checkbox_variable_map

    for k,v in MenuItem.ig_field_sequence.items():
        if k not in checkbox_variable_map:
            checkbox_variable_map[k] = IntVar()
        checkbox_variable_map[k].set(1)


def show_var_states(ttk_var):
    """prints state of ttk variables."""
    if type(ttk_var) is dict:
        for k,v in ttk_var.items():
            print('{0}: {1}'.format(k, v.get()))
    elif type(ttk_var) is list:
        for item in ttk_var:
            print('{0}: {1}'.format(item, item.get()))
    else:
        print('{0}: {1}'.format(ttk_var, ttk_var.get()))


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)

    return os.path.join(os.path.abspath("."), relative_path)


root = Tk()
root.option_add('*tearOff', FALSE)
root.title("Agilysys Import Export Tool")
ICON = resource_path('resources/Format_Gears.ico')
try:
    root.iconbitmap(default=ICON)
except TclError:
    print('Unable to locate icon at {0}'.format(ICON))

openFileString = StringVar()
noUnpriced = BooleanVar()
expandPriceLevels = BooleanVar()
colPrice = BooleanVar()
checkbox_variable_map = {}

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
                         command=lambda: show_var_states(checkbox_variable_map))

menu_help.add_command(label='About', command=display_about)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=1, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(1, weight=1)

ttk.Label(mainframe, text="Input File:").grid(
    column=1, row=1, sticky=(N, W, E))
openFile_entry = ttk.Entry(mainframe, width=40, textvariable=openFileString)
openFile_entry.grid(column=1, row=2, sticky=(W, E))

button_ig = ttk.Button(mainframe, text='Generate IG Update',
                       command=convert_to_ig_format)
button_ig.grid(column=1, row=3)

button_excel_complete = ttk.Button(mainframe, text='Create Full xls',
                           command=lambda: convert_to_excel('complete'))
button_excel_complete.grid(column=1, row=6)

button_excel_custom = ttk.Button(mainframe, text='Create Custom xls',
                             command=lambda: convert_to_excel('custom'))
button_excel_custom.grid(column=1, row=7)

simplifyButtons = [button_excel_complete, button_excel_custom]
hideable_buttons = [button_excel_complete, button_excel_custom, button_ig]

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=5)

root.config(menu=menubar)
hide_all_buttons()
root.mainloop()
