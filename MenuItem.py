#!/usr/bin/python

import re
import logging

quoteMatch = re.compile(r'(^"+|"+$)')

class MenuItem(object):
    """An object to simplify item property assignment"""

    IG_FIELD_SEQUENCE = {
        'id': 2, 'name': 3, 'abbr1': 4, 'abbr2': 5,
        'print_label': 6, 'price_levels': 7, 'product_class': 8,
        'revenue_category': 9, 'tax_group': 10, 'security_level': 11,
        'report_category': 12, 'sell_by_weight': 13,
        'tare': 14, 'sku': 15, 'gun_code': 16,
        'cost': 17, 'reserved_18': 18, 'prompt_for_price': 19,
        'print_on_check': 20, 'is_discountable': 21, 'voidable': 22,
        'inactive': 23, 'tax_included': 24,
        'item_group': 25, 'receipt_text': 26,
        'allow_price_override': 27, 'reserved_28': 28, 'choice_groups': 29,
        'kitchen_printers': 30, 'covers': 31, 'store_id': 32
    }
    
    TEXT_HEADERS = {
        'id':'ID', 'name':'Name', 'abbr1':'Abbr1', 'abbr2':'Abbr2',
        'print_label':'Printer Label', 'price_levels':'Prices',
        'revenue_category':'Revenue Category', 'tax_group':'Tax Group',
        'security_level':'Security Level', 'report_category':'Report Category',
        'sell_by_weight':'By Weight', 'tare':'Tare Weight', 'sku':'SKU',
        'gun_code':'Gun Code', 'cost':'Cost', 'product_class':'Product Class',
        'prompt_for_price':'Prompt For Price',
        'print_on_check':'Print on Check', 'is_discountable':'Discountable',
        'voidable':'Voidable', 'inactive':'Inactive',
        'tax_included':'Tax Included', 'item_group':'Item Group',
        'receipt_text':'Receipt Text',
        'allow_price_override':'Allow Price Override',
        'choice_groups':'Choice Groups', 'kitchen_printers':'Kitchen Printers',
        'covers':'Covers', 'store_id':'Store ID', 'reserved_18': '',
        'reserved_28': ''
    }
                
    INTEGER_FIELDS = [
        'product_class', 'revenue_category', 'tax_group', 'security_level',
        'report_category','sell_by_weight', 'prompt_for_price',
        'print_on_check', 'is_discountable','voidable', 'inactive',
        'tax_included', 'item_group', 'allow_price_override'
    ]

    STRING_FIELDS = (
        IG_FIELD_SEQUENCE['name'],
        IG_FIELD_SEQUENCE['abbr1'],
        IG_FIELD_SEQUENCE['abbr2'],
        IG_FIELD_SEQUENCE['receipt_text']
    )
    
    def __init__(
                self, id, name, abbr1='', abbr2='', print_label=None,
                priceLvls=None, product_class=None, revenue_category=None,
                taxGrp=None, securityLvl=0, reportCat=None, byWeight=None,
                tare=None, sku=None, gunCode=None, cost=None, pricePrompt=0,
                prntOnChk=1, disc=1, voidable=1, inactive=0, taxIncluded=0,
                itemGrp=None, receipt='', priceOver=1, choiceGrps=None,
                ktchnPrint=None, covers=0, storeID=None, reserved_18=0,
                reserved_28=0
                ):

        self.id = int(id) #seq 2
        self.name = re.sub(quoteMatch, remove_quotes, name).strip('\r\n') #seq 3
        self.abbr1 = re.sub(quoteMatch, remove_quotes, abbr1) #seq 4
        self.abbr2 = re.sub(quoteMatch, remove_quotes, abbr2) #seq 5
        self.print_label = print_label #seq 6
        self.price_levels = priceLvls #array in seq 7
        self.product_class = int_cast(product_class) or 0 #seq 8
        self.revenue_category = int_cast(revenue_category) or 0 #seq 9
        self.tax_group = int_cast(taxGrp) #seq 10
        self.security_level = int_cast(securityLvl) #seq 11
        self.report_category = int_cast(reportCat) #seq 12
        self.sell_by_weight = byWeight #seq 13
        self.tare = tare #seq 14
        self.sku = str(sku).strip().split(sep='.', maxsplit=1)[0] #seq 15
        self.gun_code = gunCode #seq 16
        self.cost = cost #seq 17
        self.prompt_for_price = pricePrompt #seq 19
        self.print_on_check = prntOnChk #seq 20
        self.is_discountable = disc #seq 21
        self.voidable = voidable #seq 22
        self.inactive = inactive #seq 23
        self.tax_included = taxIncluded #seq 24
        self.item_group = itemGrp #seq 25
        self.receipt_text = re.sub(quoteMatch, remove_quotes, receipt) #seq 26
        self.allow_price_override = priceOver #seq 27
        self.choice_groups = choiceGrps #array in seq 29
        self.kitchen_printers = ktchnPrint #array in seq 30
        self.covers = covers #seq 31
        self.store_id = int_cast(storeID) #seq 32
        self.reserved_18 = reserved_18
        self.reserved_28 = reserved_28

    def print_item_details(self):
        for k,v in self.__dict__.items():
            print(k + ": " + v)

    def print_item_details_sorted(self):
        for k,v in sorted(self.__dict__.items()):
            print("{0}: {1}".format(k,v))

    def __str__(self):
        item_properties = []
        for key, position in sorted(
                self.IG_FIELD_SEQUENCE.items(), key=lambda x: x[1]):
            if position in self.STRING_FIELDS:
                attribute = '"{0}"'.format(getattr(self, key))
            elif key == 'sku':
                attribute = self.get_barcode_string()
            else:
                attribute = str(getattr(self, key)).replace(';', ',')

            if not attribute or attribute == 'None' or attribute == '""':
                attribute = ''
            item_properties.append(attribute)
        return ",".join(item_properties)
        
    def get_prices_dict(self):
        prices = self.price_levels.strip("{}").split(";")
        price_map = dict()
        level = None
        i = 1
        for x in prices:
            if int(i) % 2 != 0:
                level = x
            else:
                price_map[int(level)] = x
                level = None
            i += 1
        return price_map

    def get_barcode_dict(self):
        barcodes = self.sku.strip("{}").split(";")
        barcode_map = dict()
        for i, x in enumerate(barcodes, start=1):
            if i % 2:
                sku = x.strip('"')
            else:
                barcode_map[sku] = x.strip('"')
                sku = None
        return barcode_map or {sku: ''}

    def get_barcode_string(self):
        barcodes = []
        for sku, description in self.get_barcode_dict().items():
            if sku:
                barcodes.append('"{0}","{1}"'.format(sku, description))
        if barcodes:
            return '{{{0}}}'.format(','.join(barcodes))
        else:
            return ''


def count_price_levels(itemList):
    """Returns total number of price levels"""
    logging.debug('counting price levels')
    num_price_levels = 0
    price_level_list = []
    for item in itemList:
        levels = item.get_prices_dict()
        for level in levels.keys():
            if level not in price_level_list:
                price_level_list.append(level)

    logging.debug('returning results of price level count')
    return price_level_list


@staticmethod
def get_flag_as_text(number):
    if number == 0:
        return 'False'
    else:
        return 'True'


def remove_quotes(match):
    match = str(match.group(0))
    return match.replace('"', '')


def int_cast(value):
    """Cast to int if possible, return unmodified otherwise"""
    try:
        return int(value)
    except (TypeError, ValueError) as e:
        return value
