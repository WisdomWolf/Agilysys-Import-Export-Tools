#!python3

# TODO fix empty barcode description generation
# TODO Move to openpyxl module for *.xls(x/m) support
# TODO Streamline UI
# TODO Create Unit tests (may be easier if migrated to OO design

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

__version__ = 'v0.10.30'

TEXT_HEADERS = MenuItem.TEXT_HEADERS
IG_FIELD_SEQUENCE = MenuItem.IG_FIELD_SEQUENCE
INTEGER_FIELDS = MenuItem.INTEGER_FIELDS
STRING_FIELDS = MenuItem.STRING_FIELDS
PRICE_ARRAY_REGEX = re.compile(r'(?<=\{)[^(\{|\})].+?(?=\})')
QUOTED_COMMAS_REGEX = re.compile(r'((?<=")[^",\{\}]+),([^"\{\}]*(?="))')
PRICE_COLUMN = 8
MAX_STRING_LENGTH = 38
file_type_filters = [('Supported Files', '.xls .xlsx .txt'),
                     ('Text Files', '.txt'),
                     ('Excel Files', '.xls .xlsx .csv'), ('All Files', '.*')]
APP_DIR = os.path.join(os.getenv('APPDATA'), 'Agilysys Format Tools')
CONFIG_FILE = os.path.join(APP_DIR, 'config.ini')
LOG_FILE = os.path.join(APP_DIR, 'errors.log')
config = ConfigParser()
config.read(CONFIG_FILE)

IG_EXPORT = 1
EXCEL_FILE = 3
itemList = []
itemMap = {}


def open_file(options=None):
    """Generates File Open dialog and alters UI based on file selected."""
    hide_all_buttons()
    init_dir = ''
    global checkbox_variable_map, all_boxes_selected
    all_boxes_selected = False
    checkbox_variable_map = {}
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
    file_display_string.set(directory_display(in_file))
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
        with open(CONFIG_FILE, 'w') as f:
            config.write(f)

    except IOError:
        messagebox.showinfo(title='Oops',
                            message='This file is not supported.')
        print('{0}\n{1}'.format(sys.exc_info()[0], sys.exc_info()[1]))
        return


def directory_display(directory):
    """Returns truncated directory string when over length over 38 chars."""
    if len(directory) > MAX_STRING_LENGTH:
        split_dirs = directory.split('/')
        directory = '/'.join([split_dirs[0], '...',
                              split_dirs[-2], split_dirs[-1]])
        if len(directory) > MAX_STRING_LENGTH:
            return split_dirs[-1]
    return directory


# Not working in current incarnation, requires modification
def write_to_text_file(file, *args):
    output_list = []
    with open(file, 'w+') as f:
        for arg in args:
            f.write('\n'.join(arg))


def save_file_as(options):
    """Generates Save As dialog and returns chosen file path."""
    file_opt = options
    file_save_path = filedialog.asksaveasfilename(**file_opt)
    if file_save_path is None or file_save_path == "":
        print("No file selected")
    return file_save_path


def replace_commas(match, ):
    """returns string with semi-colon substitutes for matched commas."""
    match = str(match.group(0))
    return match.replace(",", ";")


def pre_parse_ig_file(file_name):
    """Parses lines from IG export into workable objects."""
    with codecs.open(file_name, 'r', 'latin-1') as export:
        print('pre-parse initiated')
        for line in export:
            itemDetails = re.sub(PRICE_ARRAY_REGEX, replace_commas, line)
            itemDetails = re.sub(QUOTED_COMMAS_REGEX, replace_commas, itemDetails)
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
# Leads to lots of empty columns with high price level ids
def count_price_levels():
    """Returns total number of price levels"""
    num_price_levels = 0
    price_level_list = []
    for item in itemList:
        levels = item.get_prices_dict()
        for level in levels.keys():
            if level not in price_level_list:
                price_level_list.append(level)
        if max(k for k in levels.keys()) > num_price_levels:
            num_price_levels = max(k for k in levels.keys())
    return num_price_levels, price_level_list


# noinspection PyShadowingNames
def generateFullExcel(excel_file, items=None,
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
        numberOfPriceLevels, _ = count_price_levels()

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
        row.write(IG_FIELD_SEQUENCE['id'], int(item.id))
        row.write(IG_FIELD_SEQUENCE['name'], str(item.name))
        row.write(IG_FIELD_SEQUENCE['abbr1'], str(item.abbr1))
        row.write(IG_FIELD_SEQUENCE['abbr2'], str(item.abbr2))
        row.write(IG_FIELD_SEQUENCE['print_label'], str(item.print_label))

        if expandPriceLevels:
            for p in range(1, (numberOfPriceLevels + 1)):
                if p in item.get_prices_dict():
                    price = item.get_prices_dict()[p]
                else:
                    price = ''
                r = p - 1
                row.write(IG_FIELD_SEQUENCE['price_levels'] + r, str(price))

            row.write(IG_FIELD_SEQUENCE['product_class'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.product_class))
            row.write(IG_FIELD_SEQUENCE['revenue_category'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.revenue_category))
            row.write(IG_FIELD_SEQUENCE['tax_group'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.tax_group))
            row.write(IG_FIELD_SEQUENCE['security_level'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.security_level))
            row.write(IG_FIELD_SEQUENCE['report_category'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.report_category))
            row.write(IG_FIELD_SEQUENCE['sell_by_weight'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.report_category))
            row.write(IG_FIELD_SEQUENCE['tare'] + (numberOfPriceLevels - 1),
                      str(item.tare))
            row.write(IG_FIELD_SEQUENCE['sku'] + (numberOfPriceLevels - 1),
                      str(item.sku))
            row.write(IG_FIELD_SEQUENCE['gun_code'] + (numberOfPriceLevels - 1),
                      str(item.gun_code))
            row.write(IG_FIELD_SEQUENCE['cost'] + (numberOfPriceLevels - 1),
                      str(item.cost))
            row.write(IG_FIELD_SEQUENCE['cost'] + numberOfPriceLevels, 'N/A')
            row.write(IG_FIELD_SEQUENCE['prompt_for_price'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.prompt_for_price))
            row.write(IG_FIELD_SEQUENCE['print_on_check'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.print_on_check))
            row.write(IG_FIELD_SEQUENCE['is_discountable'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.is_discountable))
            row.write(IG_FIELD_SEQUENCE['voidable'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.voidable))
            row.write(IG_FIELD_SEQUENCE['inactive'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.inactive))
            row.write(IG_FIELD_SEQUENCE['tax_included'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.tax_included))
            row.write(IG_FIELD_SEQUENCE['item_group'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.item_group))
            row.write(IG_FIELD_SEQUENCE['receipt_text'] + (numberOfPriceLevels - 1),
                      str(item.receipt_text))
            row.write(IG_FIELD_SEQUENCE['allow_price_override'] + (numberOfPriceLevels - 1),
                      safeIntCast(item.allow_price_override))
            row.write(IG_FIELD_SEQUENCE['allow_price_override'] + numberOfPriceLevels, 'N/A')
            row.write(IG_FIELD_SEQUENCE['choice_groups'] + (numberOfPriceLevels - 1),
                      str(item.choice_groups))
            row.write(IG_FIELD_SEQUENCE['kitchen_printers'] + (numberOfPriceLevels - 1),
                      str(item.kitchen_printers))
            row.write(IG_FIELD_SEQUENCE['covers'] + (numberOfPriceLevels - 1),
                      str(item.covers))
            row.write(IG_FIELD_SEQUENCE['store_id'] + (numberOfPriceLevels - 1),
                      str(item.store_id))
        else:
            row.write(IG_FIELD_SEQUENCE['price_levels'], str(item.price_levels))
            row.write(IG_FIELD_SEQUENCE['product_class'], safeIntCast(item.product_class))
            row.write(IG_FIELD_SEQUENCE['revenue_category'], safeIntCast(item.revenue_category))
            row.write(IG_FIELD_SEQUENCE['tax_group'], safeIntCast(item.tax_group))
            row.write(IG_FIELD_SEQUENCE['security_level'], safeIntCast(item.security_level))
            row.write(IG_FIELD_SEQUENCE['report_category'], safeIntCast(item.report_category))
            row.write(IG_FIELD_SEQUENCE['sell_by_weight'], safeIntCast(item.sell_by_weight))
            row.write(IG_FIELD_SEQUENCE['tare'], str(item.tare))
            row.write(IG_FIELD_SEQUENCE['sku'], str(item.sku))
            row.write(IG_FIELD_SEQUENCE['gun_code'], str(item.gun_code))
            row.write(IG_FIELD_SEQUENCE['cost'], str(item.cost))
            row.write(IG_FIELD_SEQUENCE['cost'] + 1, 'N/A')
            row.write(IG_FIELD_SEQUENCE['prompt_for_price'], safeIntCast(item.prompt_for_price))
            row.write(IG_FIELD_SEQUENCE['print_on_check'], safeIntCast(item.print_on_check))
            row.write(IG_FIELD_SEQUENCE['is_discountable'], safeIntCast(item.is_discountable))
            row.write(IG_FIELD_SEQUENCE['voidable'], safeIntCast(item.voidable))
            row.write(IG_FIELD_SEQUENCE['inactive'], safeIntCast(item.inactive))
            row.write(IG_FIELD_SEQUENCE['tax_included'], safeIntCast(item.tax_included))
            row.write(IG_FIELD_SEQUENCE['item_group'], safeIntCast(item.item_group))
            row.write(IG_FIELD_SEQUENCE['receipt_text'], str(item.receipt_text))
            row.write(IG_FIELD_SEQUENCE['allow_price_override'], safeIntCast(item.allow_price_override))
            row.write(IG_FIELD_SEQUENCE['allow_price_override'] + 1, 'N/A')
            row.write(IG_FIELD_SEQUENCE['choice_groups'], str(item.choice_groups))
            row.write(IG_FIELD_SEQUENCE['kitchen_printers'], str(item.kitchen_printers))
            row.write(IG_FIELD_SEQUENCE['covers'], str(item.covers))
            row.write(IG_FIELD_SEQUENCE['store_id'], str(item.store_id))

        if row_is_misaligned:
            oopsStyle = (easyxf('pattern: pattern solid, fore_color rose'))
            row.write(1, 'X', oopsStyle)

    try:
        book.save(excel_file)
        messagebox.showinfo(title='Success',
                            message='Excel export created successfully.')
    except PermissionError:
        messagebox.showerror(title='Error', message='Unable to save file')


# noinspection PyShadowingNames,PyShadowingNames
def generate_custom_excel_spreadsheet(
        excel_file, items=None, excludeUnpriced=True
):
    """Generates Excel spreadsheet from IG Export file

    Keyword arguements:
    excel_file -- file path to save resulting workbook
    items -- list of items to parse
    excludeUnpriced -- ignore items lacking a price when generating export
    """
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
    heading_row = sheet.row(0)
    keyname_row = sheet.row(1)
    headers = ['"A"', 'ID']
    keynames = ['modType', 'id']
    pricePos = 100

    for k, v in sorted(checkbox_variable_map.items(),
                       key=lambda x: IG_FIELD_SEQUENCE.get(x[0])):
        if str(v.get()) == '1':
            headers.append(TEXT_HEADERS[k])
            keynames.append(k)

    if 'Prices' in headers:
        pricePos = headers.index('Prices')
        del headers[pricePos]
        del keynames[pricePos]
        num_price_levels, price_level_list = count_price_levels()
        price_headers = []

        # Need additional logic to filter empty price levels
        if price_level_list:
            for price_level in price_level_list:
                price_headers.append('Price Level ' + str(price_level))
        else:
            for price_level in range(num_price_levels):
                price_headers.append('Price Level ' + str(price_level + 1))

        price_headers.reverse()
        for level, price in zip(reversed(range(1, len(price_headers) + 1)),
                                price_headers):
            headers.insert(pricePos, price)
            keynames.insert(pricePos, ('priceLvl' + str(level)))

    # Write Headers
    for header, key, row in zip(headers, keynames, range(len(headers))):
        heading_row.write(row, header, heading)
        keyname_row.write(row, key, heading)

    # Hiding keyname row
    sheet.row(1).hidden = True

    # Write Rows
    for r, item in zip(range(2, len(items) + 2), items):
        global row_is_misaligned
        row_is_misaligned = False

        row = sheet.row(r)

        # Write item values to columns
        for col, key in enumerate(keynames):
            if 'priceLvl' in key:
                # Extract number from priceLvl key
                p = int(key[key.find('l') + 1:])
                if p in item.get_prices_dict():
                    price = item.get_prices_dict()[p]
                else:
                    price = ''

                row.write(col, str(price))
            else:
                if key in INTEGER_FIELDS:
                    row.write(col, safeIntCast(item.__dict__[key]))
                elif key == 'modType':
                    continue
                else:
                    row.write(col, str(item.__dict__[key]))

        if row_is_misaligned:
            row.write(0, 'X', oopsStyle)

    try:
        book.save(excel_file)
        messagebox.showinfo(title='Success',
                            message='Custom Excel export created successfully')
    except PermissionError:
        messagebox.showerror(title='Error', message='Unable to save file')


def generateIGUpdate(excel_file, ig_text_file):
    """Legacy function for generating IG import from full Excel workbook."""
    print('preparing to generate IG Update file using legacy function')

    sheet = book.sheet_by_index(0)
    includeColumns = set()

    for col in range(3, sheet.ncols):
        if sheet.cell_value(1, col):
            includeColumns.add(col)

    includeColumns = sorted(includeColumns)

    for row in range(2, sheet.nrows):
        itemProperties = []
        updateType = sheet.cell_value(row, 1)
        if updateType != 'A' and updateType != 'U' and \
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
            if col in STRING_FIELDS:
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


# TODO rewrite to leverage MenuItem class
def generate_ig_import(book, ig_text_file):
    """Generates IG Import File from custom Excel workbook.

    Keyword arguments:
    book -- Excel workbook (custom)
    ig_text_file -- text file to be generated for Agilysys
    """
    print('Generating IG Import from custom Excel')
    sheet = book.sheet_by_index(0)
    updated_items = 0

    for row in range(2, sheet.nrows):
        item_properties = []
        item_property_map = dict()
        price_level_map = dict()
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
                price_level_map[key] = sheet.cell_value(row, col)
            item_property_map[key] = sheet.cell_value(row, col)

        if price_level_map:
            item_property_map['price_levels'] = \
                build_ig_price_array(price_level_map)

        for key, field in sorted(IG_FIELD_SEQUENCE.items(),
                                 key=lambda x: x[1]):
            if key in item_property_map.keys():
                if field in STRING_FIELDS:
                    item_properties.append('"{0}"'.format(
                        item_property_map[key].strip('\r\n')))
                else:
                    item_properties.append(safeIntCast(item_property_map[key]))
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


# TODO Correct/Replace save path selection
def generate_standardized_ig_imports(book, save_path):
    """Generates IG import files from POS Configuration Worksheet.

    Keyword arguments:
    book -- Excel workbook
    base_filename -- filename used as base for priced and unpriced output files
    """
    print('Generating standardized IG Import files')
    sheet = book.sheet_by_name('Menu Items & Pricing')
    ig_sheet = book.sheet_by_name('InfoGenesis')
    product_sheet = book.sheet_by_name('Product Classes')
    revenue_sheet = book.sheet_by_name('Revenue Categories')
    category_sheet = book.sheet_by_name('Categories')
    product_classes = get_product_classes(product_sheet)
    revenue_categories = get_revenue_categories(revenue_sheet)
    ig_priced_file = '{0} - MI_Imp_priced.txt'.format(save_path)
    ig_unpriced_file = '{0} - MI_Imp_unpriced.txt'.format(save_path)
    product_class_error_file = os.path.join(
        os.path.dirname(save_path), 'missing_product_classes.txt')
    revenue_category_error_file = os.path.join(
        os.path.dirname(save_path), 'missing_revenue_categories.txt')

    max_item_id = int(ig_sheet.cell_value(ig_sheet.nrows - 1, 1))
    last_item_id = 1
    item_ids = set()

    revenue_category_errors = []
    product_class_errors = []
    priced_items = []
    unpriced_items = []
    item_list = []

    for row in range(5, sheet.nrows):
        fields = []
        price_map = dict()
        update_type = '"A"'
        prices = '{1,$0.00}'

        for col in range(1, 8):
            property = sheet.cell_value(row, col)
            fields.append(property)
        for price_level in range(8, sheet.ncols):
            level = price_level - 7
            if sheet.cell_value(row, price_level) == '':
                continue
            else:
                price = sheet.cell_value(row, price_level)
            price_map[level] = price

        if price_map:
            prices = build_ig_price_array(price_map)

        item_id = fields[0]
        if item_id in item_ids:
            raise IndexError('duplicate id used')
        if item_id > max_item_id or last_item_id < item_id < 1000000:
            item_id = last_item_id
            last_item_id += 1
        rev_cat = fields[5]
        name = fields[3]
        prod_class = fields[6]
        sku = fields[2]

        try:
            revenue_category = revenue_categories[rev_cat]
        except KeyError:
            if rev_cat not in revenue_category_errors:
                revenue_category_errors.append(rev_cat)
                print('Unable to get value for {0} category'.format(rev_cat))
            revenue_category = None

        try:
            product_class = product_classes[prod_class]
        except KeyError:
            if prod_class not in product_class_errors:
                product_class_errors.append(prod_class)
                print('Unable to get value for {0} product class'.format(
                    prod_class))
            product_class = None

        item = MenuItem(id=item_id, name=name,
                        revenue_category=revenue_category,
                        product_class=product_class, sku=sku,
                        priceLvls=prices)
        item_list.append(item)
        line = '{0},{1}'.format(update_type, item)
        if price_map:
            priced_items.append(line)
        else:
            unpriced_items.append(line)

    print('iteration complete, writing files')
    if priced_items:
        write_to_text_file(ig_priced_file, priced_items)
    if unpriced_items:
        write_to_text_file(ig_unpriced_file, unpriced_items)
    if product_class_errors:
        write_to_text_file(product_class_error_file, product_class_errors)
    if revenue_category_errors:
        write_to_text_file(revenue_category_error_file, revenue_category_errors)

    print('IG import file creations complete.')
    messagebox.showinfo(title='Success',
                        message='IG import files created successfully.')


def get_revenue_categories(sheet):
    """Returns revenue categories dictionary from internal spreadsheet"""
    revenue_categories = dict()
    for row in range(1, sheet.nrows):
        try:
            revenue_categories[sheet.cell_value(row, 1)] = \
                sheet.cell_value(row, 0)
        except IndexError:
            print("oops, couldn't read row {0} from Revenue Categories".format(
                row))

    return revenue_categories


def get_product_categories(sheet):
    """Returns product classes dict from custom spreadsheet"""
    product_classes = dict()
    for row in range(1, sheet.nrows):
        try:
            product_classes[sheet.cell_value(row, 0)] = \
                sheet.cell_value(row, 1)
        except IndexError:
            print("oops, couldn't read row {0} from Product Classes".format(
                row))

    return product_classes


def get_product_classes(sheet):
    """Returns product classes dict from internal spreadsheet"""
    product_classes = dict()
    for row in range(1, sheet.nrows):
        pc = sheet.cell_value(row, 1)
        pid = sheet.cell_value(row, 0)
        try:
            product_classes[pc] = pid
        except IndexError:
            print("oops, couldn't read row {0} from Product Classes".format(
                row))
    return product_classes


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
    print('starting conversion to IG Format')
    options = {
        'title': 'Save As',
        'initialfile': os.path.join(os.path.dirname(in_file), 'MI_IMP.txt')
    }
    file_extension = in_file.rsplit('.', maxsplit=1)[1]
    if file_extension == 'xls' or file_extension == 'xlsx':
        book = open_workbook(in_file)
        sheet = book.sheet_by_index(0)
        if book.nsheets > 1:
            generate_standardized_ig_imports(book, in_file)
        else:
            file_save_path = save_file_as(options)
            if file_save_path:
                with codecs.open(file_save_path, 'w+', 'latin-1') as text_file:
                    if sheet.cell_value(1, 0) == 'modType':
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
    file_save_path = save_file_as(options)

    if file_save_path and type == 'complete':
        generateFullExcel(file_save_path, excludeUnpriced=noUnpriced.get(),
                          expandPriceLevels=expandPriceLevels.get())
    elif file_save_path and type == 'custom':
        generate_custom_excel_spreadsheet(file_save_path, excludeUnpriced=noUnpriced.get())


def build_ig_price_array(price_map):
    """Returns IG price array from dictionary of price levels."""
    prices = []
    price_is_negative = False
    for price_level, price in sorted(price_map.items()):
        if price != '':
            # Extract number from priceLvl
            if type(price_level) == str:
                level = str(price_level[price_level.find('l') + 1:])
            else:
                level = price_level
            price_sequence = ''
            if '(' in str(price) or ')' in str(price):
                price_is_negative = True
            price = '{0:.2f}'.format(float(str(price).strip('$(){}')))
            if price_is_negative:
                price_sequence = '{0},(${1})'.format(level, price)
            else:
                price_sequence = '{0},${1}'.format(level, price)
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
    row = 5
    counter = 0

    for k, v in sorted(IG_FIELD_SEQUENCE.items(), key=lambda x: x[1]):
        col = 0
        if k not in checkbox_variable_map:
            checkbox_variable_map[k] = IntVar()
        if k != 'id' and k[:-3] != 'reserved':
            if counter % 2 == 0:
                col = 0
                row += 1
            else:
                col = 3
            l = ttk.Checkbutton(colSelectFrame, text=TEXT_HEADERS[k],
                                variable=checkbox_variable_map[k]).grid(
                column=col, row=row, sticky=(N, W))
            counter += 1

    ttk.Button(colSelectFrame, text='OK',
               command=csWin.destroy).grid(column=1, row=100)
    ttk.Button(colSelectFrame, text='Select All',
               command=select_all_properties).grid(column=1, row=0)
    return


def select_all_properties():
    """Selects all properties in checkVarMap"""
    global checkbox_variable_map
    global all_boxes_selected

    if all_boxes_selected:
        check_mark = 0
        all_boxes_selected = False
    else:
        check_mark = 1
        all_boxes_selected = True

    for k, v in MenuItem.IG_FIELD_SEQUENCE.items():
        if k not in checkbox_variable_map:
            checkbox_variable_map[k] = IntVar()
        # id is mandatory and shouldn't be included in variable map
        if k != 'id':
            checkbox_variable_map[k].set(check_mark)


def show_var_states(ttk_var):
    """prints state of ttk variables. For debugging purposes."""
    if type(ttk_var) is dict:
        for k, v in ttk_var.items():
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
    else:
        return os.path.join(os.path.abspath("."), relative_path)


root = Tk()
root.option_add('*tearOff', FALSE)
root.title("Agilysys Import Export Tool")
ICON = resource_path('resources/Format_Gears.ico')
try:
    root.iconbitmap(default=ICON)
except TclError:
    print('Unable to locate icon at {0}'.format(ICON))

file_display_string = StringVar()
noUnpriced = BooleanVar()
expandPriceLevels = BooleanVar()
colPrice = BooleanVar()

if not os.path.exists(APP_DIR):
    os.mkdir(APP_DIR)

if hasattr(sys, '_MEIPASS'):
    FSOCK = open(LOG_FILE, 'a+')
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
menu_options.add_command(label='Display Vars',
                         command=lambda: show_var_states(checkbox_variable_map))

menu_help.add_command(label='About', command=display_about)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=1, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(1, weight=1)

ttk.Label(mainframe, text="Input File:").grid(
    column=1, row=1, sticky=(N, W, E))
openFile_entry = ttk.Entry(mainframe, width=40, textvariable=file_display_string)
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
