#!/usr/bin/python

class MenuItem:
    """An object to simplify item property assignment"""
    def __init__(
                self, itemID, name, abbr1=None, abbr2=None, printLabel=None, priceLevels=None,
                classID=None, revCategoryID=None, taxGroup=None, securityLevel=0,
                reportCategory=None, useWeightFlag=None, weightTareAmount=None, sku=None,
                gunCode=None, costAmount=None, pricePrompt=0, checkPrintFlag=1,
                discountableFlag=1, voidableFlag=1, inactiveFlag=0,
                taxIncludeFlag=0, itemGroupID=None, receiptText=None,
                priceOverrideFlag=1, choiceGroups=None, kitchenPrinters=None, covers=0,
                storeID=0
                ):
        self.id = itemID #seq 2
        self.name = name.strip('"') #seq 3
        self.abbr1 = abbr1 #"""seq 4"""
        self.abbr2 = abbr2 #"""seq 5""""
        self.printerLabel = printLabel #"""seq 6"""
        self.priceLevels = priceLevels #"""array in seq 7"""
        self.classID = classID #"""seq 8"""
        self.revCategoryID = revCategoryID #"""seq 9"""
        self.taxGroup = taxGroup #"""seq 10"""
        self.securityLevel = securityLevel #"""seq 11"""
        self.reportCategory = reportCategory #"""seq 12"""
        self.useWeightFlag = useWeightFlag #"""seq 13"""
        self.weightTareAmount = weightTareAmount #"""seq 14"""
        self.sku = sku #"""seq 15"""
        self.gunCode = gunCode #"""seq 16"""
        self.costAmount = costAmount #"""seq 17"""
        self.pricePrompt = pricePrompt #"""seq 19"""
        self.checkPrintFlag = checkPrintFlag #"""seq 20"""
        self.discountableFlag = discountableFlag #"""seq 21"""
        self.voidableFlag = voidableFlag #"""seq 22"""
        self.inactiveFlag = inactiveFlag #"""seq 23"""
        self.taxIncludeFlag = taxIncludeFlag #"""seq 24"""
        self.itemGroupID = itemGroupID #"""seq 25"""
        self.receiptText = receiptText #"""seq 26"""
        self.priceOverrideFlag = priceOverrideFlag #"""seq 27"""
        self.choiceGroups = choiceGroups #"""array in seq 29"""
        self.kitchenPrinters = kitchenPrinters #"""array in seq 30"""
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
        prices = self.priceLevels.strip("{}")
        details = prices.split(";")
        priceList = {}
        level = None
        i = 1
        for x in details:
            if int(i) % 2 != 0:
                level = x
            else:
                priceList[int(level)] = x
                level = None
            i += 1
        return priceList

    def separatePriceLevelsSorted(self):
        return sorted(self.separatePriceLevels().items())

    def printPrices(self):
        prices = self.separatePriceLevelsSorted()
        for k,v in prices.items():
            print("Price Level " + str(k) + ": " + str(v))
            
    def getFlagText(self, number):
        if number == 0:
            return 'False'
        else:
            return 'True'