#!/usr/bin/python

class MenuItem:
	"""An object to simplify item property assignment"""
	def __init__(
				self, itemID, name, abbr1, abbr2, printLabel, priceLevels,
				classID, revCategoryID, taxGroup, securityLevel,
				reportCategory, useWeightFlag, weightTareAmount, sku,
				gunCode, costAmount, pricePrompt, checkPrintFlag,
				discountableFlag, voidableFlag, inactiveFlag,
				taxIncludeFlag, itemGroupID, receiptText,
				priceOverrideFlag, choiceGroups, kitchenPrinters, covers,
				storeID, original="unknown", parsed="unknown"
				):
		self.id = itemID #seq 2
		self.name = name #seq 3
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
		self.original = original
		self.parsed = parsed

	def printItemDetails(self):
		print("ID: " + str(self.id))
		try:
			print("Name: " + str(self.name))
		except UnicodeEncodeError:
			print("Name: *Error reading name*")
		print("Abbreviation 1: " + str(self.abbr1))
		print("Abbreviation 2: " + str(self.abbr2))
		print("Printer Label: " + str(self.printerLabel))
		print("Price Levels: " + str(self.priceLevels))
		print("Class ID: " + str(self.classID))
		print("Revenue Category ID: " + str(self.revCategoryID))
		print("Tax Group: " + str(self.taxGroup))
		print("Security Level: " + str(self.securityLevel))
		print("Report Category: " + str(self.reportCategory))
		print("Use Weight Flag: " + str(self.useWeightFlag))
		print("SKU: " + str(self.sku))
		print("Gun Code: " + str(self.gunCode))
		print("Cost Amount: " + str(self.costAmount))
		print("Price Prompt: " + str(self.pricePrompt))
		print("Print on Check Flag: " + str(self.checkPrintFlag))
		print("Discountable: " + str(self.discountableFlag))
		print("Voidable: " + str(self.voidableFlag))
		print("Inactive: " + str(self.inactiveFlag))
		print("Tax Include Flag: " + str(self.taxIncludeFlag))
		print("Item Group ID: " + str(self.itemGroupID))
		print("Receipt Text: " + str(self.receiptText))
		print("Price Override Flag: " + str(self.priceOverrideFlag))
		print("Choice Groups: " + str(self.choiceGroups))
		print("Kitchen Printers: " + str(self.kitchenPrinters))
		print("Covers: " + str(self.covers))
		print("Store ID: " + str(self.storeID))
		
	def toString(self):
		self.printItemDetails()
		
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
				priceList[level] = x
				level = None
			i += 1
		return priceList
	
	def printPrices(self):
		prices = self.separatePriceLevels()
		for k,v in prices.items():
			print("Price Level " + str(k) + ": " + str(v))