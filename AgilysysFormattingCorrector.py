#!/usr/bin/python

import re
import codecs
import tkinter
import os
from Things import MenuItemThings

from tkinter import filedialog


root = tkinter.Tk()
root.withdraw()

file_opt = options = {}
options['defaultextension'] = '.txt'
options['filetypes'] = [('text files', '.txt'), ('all files', '.*')]
options['title'] = 'Open Agilysys Export'

file_path = filedialog.askopenfilename(**file_opt)
save_file = filedialog.asksaveasfilename()
if file_path == None or save_file == None or file_path == "" or save_file == "":
		print("No file selected")
		x = input('--> ')
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
		i = MenuItemThings.MenuItem(item[1], item[2], item[3], item[4], item[5], item[6], item[7], item[8], item[9], item[10], item[11], item[12], item[13], item[14], item[15], item[16], item[18], item[19], item[20], item[21], item[22], item[23], item[24], item[25], item[26], item[28], item[29], item[30], item[31], x, item)
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
	
