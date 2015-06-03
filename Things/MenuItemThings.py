#!/usr/bin/python

import re

quoteMatch = re.compile(r'(^"+|"+$)')

class MenuItem:
    """An object to simplify item property assignment"""
    
    attributeMap = {'id':2, 'name':3, 'abbr1':4, 'abbr2':5,
                     'printerLabel':6, 'priceLvls':7, 'classID':8,
                     'revCat':9, 'taxGrp':10, 'securityLvl':11,
                     'reportCat':12, 'byWeight':13,
                     'tare':14, 'sku':15, 'gunCode':16,
                     'cost':17, 'pricePrompt':19, 'prntOnChk':20,
                     'disc':21, 'voidable':22,
                     'inactive':23, 'taxIncluded':24,
                     'itemGrp':25, 'receipt':26,
                     'priceOver':27, 'choiceGrps':29,
                     'ktchnPrint':30, 'covers':31, 'storeID':32}
    
    textMap = {'id':'ID', 'name':'Name', 'abbr1':'Abbr1', 'abbr2':'Abbr2',
                'printerLabel':'Printer Label', 'priceLvls':'Prices',
                'revCat':'Revenue Category', 'taxGrp':'Tax Group',
                'securityLvl':'Security Level', 'reportCat':'Report Category',
                'byWeight':'By Weight', 'tare':'Tare Weight', 'sku':'SKU',
                'gunCode':'Gun Code', 'cost':'Cost', 'classID':'Class ID',
                'pricePrompt':'Prompt For Price',
                'prntOnChk':'Print on Check', 'disc':'Discountable',
                'voidable':'Voidable', 'inactive':'Inactive',
                'taxIncluded':'Tax Included', 'itemGrp':'Item Group',
                'receipt':'Receipt Text', 'priceOver':'Allow Price Override',
                'choiceGrps':'Choice Groups', 'ktchnPrint':'Kitchen Printers',
                'covers':'Covers', 'storeID':'Store ID'}
    
    def __init__(
                self, itemID, name, abbr1=None, abbr2=None, printLabel=None,
                priceLvls=None, classID=None, revCat=None, taxGrp=None,
                securityLvl=0, reportCat=None, byWeight=None, tare=None,
                sku=None, gunCode=None, cost=None, pricePrompt=0,
                prntOnChk=1, disc=1, voidable=1, inactive=0, taxIncluded=0,
                itemGrp=None, receipt=None, priceOver=1, choiceGrps=None,
                ktchnPrint=None, covers=0, storeID=0
                ):
        
        self.id = itemID #seq 2
        self.name = re.sub(quoteMatch, removeQuotes, name) #seq 3
        self.abbr1 = abbr1 #"""seq 4"""
        self.abbr2 = abbr2 #"""seq 5""""
        self.printerLabel = printLabel #"""seq 6"""
        self.priceLvls = priceLvls #"""array in seq 7"""
        self.classID = classID #"""seq 8"""
        self.revCat = revCat #"""seq 9"""
        self.taxGrp = taxGrp #"""seq 10"""
        self.securityLvl = securityLvl #"""seq 11"""
        self.reportCat = reportCat #"""seq 12"""
        self.byWeight = byWeight #"""seq 13"""
        self.tare = tare #"""seq 14"""
        self.sku = sku #"""seq 15"""
        self.gunCode = gunCode #"""seq 16"""
        self.cost = cost #"""seq 17"""
        self.pricePrompt = pricePrompt #"""seq 19"""
        self.prntOnChk = prntOnChk #"""seq 20"""
        self.disc = disc #"""seq 21"""
        self.voidable = voidable #"""seq 22"""
        self.inactive = inactive #"""seq 23"""
        self.taxIncluded = taxIncluded #"""seq 24"""
        self.itemGrp = itemGrp #"""seq 25"""
        self.receipt = receipt #"""seq 26"""
        self.priceOver = priceOver #"""seq 27"""
        self.choiceGrps = choiceGrps #"""array in seq 29"""
        self.ktchnPrint = ktchnPrint #"""array in seq 30"""
        self.covers = covers #"""seq 31"""
        self.storeID = storeID #"""seq 32"""
        

    def printItemDetails(self):
        for k,v in self.__dict__.items():
            print(k + ": " + v)

    def printItemDetailsSorted(self):
        for k,v in sorted(self.__dict__.items()):
            print(k + ": " + v)
        
    def toString(self):
        result = []
        for _,v in self.__dict__.items():
            result.append(v)
        return ",".join(result)
        
        
    def showOriginal(self):
        print("Orginal Line: " + str(self.original))
        print("Parsed Data: " + str(self.parsed))
        
    def separatePriceLevels(self):
        prices = self.priceLvls.strip("{}")
        details = prices.split(";")
        priceMap = {}
        level = None
        i = 1
        for x in details:
            if int(i) % 2 != 0:
                level = x
            else:
                priceMap[int(level)] = x
                level = None
            i += 1
        return priceMap

    def separatePriceLevelsSorted(self):
        return sorted(self.separatePriceLevels().items())

    def printPrices(self):
        prices = self.separatepriceLevelsSorted()
        for k,v in prices.items():
            print("Price Level " + str(k) + ": " + str(v))
            
    def getFlagText(self, number):
        if number == 0:
            return 'False'
        else:
            return 'True'
        
def removeQuotes(match):
    match = str(match.group(0))
    return match.replace('"', '')
