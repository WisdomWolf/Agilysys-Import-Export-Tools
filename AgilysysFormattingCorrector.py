import re
import codecs
import tkinter
import os
from tkinter import filedialog

class MenuItem:
	"""An object to simplify item property assignment"""
	def __init__(self, itemID, name, abbr1, abbr2, printLabel, priceLevels, classID, revCategoryID, taxGroup, securityLevel, reportCategory, useWeightFlag, weightTareAmount, sku, gunCode, costAmount, pricePrompt, checkPrintFlag, discountableFlag, voidableFlag, inactiveFlag, taxIncludeFlag, itemGroupID, receiptText, priceOverrideFlag, choiceGroups, kitchenPrinters, covers, storeID, original, parsed):
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


root = tkinter.Tk()
root.withdraw()

file_opt = options = {}
options['defaultextension'] = '.txt'
options['filetypes'] = [('text files', '.txt'), ('all files', '.*')]
options['title'] = 'Open Agilysys Export'

file_path = filedialog.askopenfilename(**file_opt)
save_file = filedialog.asksaveasfilename()
if (file_path or save_file == None) or (file_path or save_file == ""):
		print("No file selected")
		os._exit(1)
export = codecs.open(file_path, 'r', 'utf8')
output = codecs.open(save_file, 'w+', 'utf8')
priceArrayMatch = re.compile(r'(?<=\{)[^(\{|\})].+?(?=\})')
itemList = []

def fixArray(match):
	match = str(match.group(0))
	return match.replace(",",";")
	

def preParse():
	for x in export:
		itemDetails = re.sub(priceArrayMatch, fixArray, x)
		item = itemDetails.split(",")
		i = MenuItem(item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9], item[10], item[11], item[12], item[13], item[14], item[15], item[16], item[18], item[19], item[20], item[21], item[22], item[23], item[24], item[25], item[26], item[28], item[29], item[30], item[31], x, item)
		itemList.append(i)
		try:
			output.write(itemDetails)
		except UnicodeEncodeError:
			errorText = "\n\n!!!!!!!!!!!!!!!!!!!!!!!\nerror encoding string for print/output\n!!!!!!!!!!!!!!!!!!!!!!!!!\n\n"
			print(errorText)
			output.write("error processing item " + str(i.id) + "\n")
	print("completed")

def generateSimpleExport(items=itemList, altered=True):
	simple_file = str(save_file)[:-4] + "_simplified" + str(save_file)[-4:]
	simpleOutput = codecs.open(simple_file, 'w+', 'utf8')
	for item in items:
		if altered:
			if item.priceLevels != "{}":
				simpleOutput.write(str(item.id) + "," + str(item.name) + "," + str(item.priceLevels) + "\r\n")
		else:
			simpleOutput.write(str(item.id) + "," + str(item.name) + "," + str(item.priceLevels) + "\r\n")

def generateIGPriceUpdate(items=None):
	file_opt = options = {}
	options['defaultextension'] = '.txt'
	options['filetypes'] = [('text files', '.txt'), ('all files', '.*')]
	options['title'] = 'Open Simplified Export'

	file_path = filedialog.askopenfilename(**file_opt)
	if file_path == None or file_path == "":
		print("No file selected")
		return
	inputFile = codecs.open(file_path, 'r', 'utf8')
	while(True):
		save_path = filedialog.askdirectory()
		if save_path != None and save_path != "":
			save_file = str(save_path) + "/MI_IMP.txt"
		else:
			print("No file selected")
			return
		try:
			updateFile = codecs.open(save_file, 'x', 'utf8')
			break
		except FileExistsError:
			print("There is already an Agilysys import file in this directory.  Please try again.")
	
	for x in inputFile:
		details = x.split(",")
		details[2] = details[2].replace(";", ",").strip("\r\n")
		line = '"U",' + str(details[0]) + ',,,,,' + str(details[2]) + ',,,,,,,,,,,,,,,,,\r\n'
		updateFile.write(line)
	
	print("File output written successfully")

preParse()
generateSimpleExport()
print("fin")
	
