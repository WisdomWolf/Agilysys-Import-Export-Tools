#!/usr/bin/python

import re
import codecs
import tkinter
import os
from Things import MenuItemThings
from tkinter import filedialog


def openExport(**options):
    if options == None:
        options = {}
        options['defaultextension'] = '.txt' 
        options['filetypes'] = [('Text Files', '.txt'), ('CSV Files', '*.csv*'), ('All Files', '.*')]
        options['title'] = 'Open Agilysys Export'
    file_opt = options
    file_path = filedialog.askopenfilename(**file_opt)
    if file_path == None or file_path == "":
        print("No file selected")
        os._exit(1)
    return options, file_path, file_opt

root = tkinter.Tk()
root.withdraw()

options, file_path, file_opt = openExport()

options['title'] = 'Save As'
options['initialfile'] = str(file_path)[:-4] + "_simplified" + str(file_path)[-4:]
save_file = str(file_path)[:-4] + ".csv"
simple_file = filedialog.asksaveasfilename(**file_opt)
if save_file == None or save_file == "":
        print("No file selected")
        os._exit(1)
        
export = codecs.open(file_path, 'r', 'utf8')
output = codecs.open(save_file, 'w+', 'utf8')
priceArrayMatch = re.compile(r'(?<=\{)[^(\{|\})].+?(?=\})')
itemList = []
itemMap = {}

def fixArray(match):
    match = str(match.group(0))
    return match.replace(",",";")
    
def preParse():
    for x in export:
        itemDetails = re.sub(priceArrayMatch, fixArray, x)
        item = itemDetails.split(",")
        i = MenuItemThings.MenuItem(
                                item[1], item[2], item[3], item[4], item[5],
                                item[6], item[7], item[8], item[9], item[10],
                                item[11], item[12], item[13], item[14], item[15],
                                item[16], item[18], item[19], item[20], item[21],
                                item[22], item[23], item[24], item[25], item[26],
                                item[28], item[29], item[30], item[31]
                                )
        itemList.append(i)
        itemMap[i.id] = i
        try:
            output.write(itemDetails)
        except UnicodeEncodeError:
            errorText = "\n\n!!!!!!!!!!!!!!!!!!!!!!!\nerror encoding string for print/output\n!!!!!!!!!!!!!!!!!!!!!!!!!\n\n"
            print(errorText)
            output.write("error processing item " + str(i.id) + "\n")
    print("completed")

def generateSimpleExport(items=itemList, altered=True):
    simpleOutput = codecs.open(simple_file, 'w+', 'utf8')
    for item in items:
        if altered:
            if item.priceLevels != "{}":
                simpleOutput.write(str(item.id) + "," + str(item.name) + "," + str(item.priceLevels) + "\r\n")
        else:
            simpleOutput.write(str(item.id) + "," + str(item.name) + "," + str(item.priceLevels) + "\r\n")
			
def generateCSV(items=itemList):
	numberOfLevels = 0
	itemLine = []
	for i in items:
		if len(i.separatePriceLevels()) > numberOfLevels:
			numberOfLevels = len(i.separatePriceLevels())
		for k,v in i.separatePriceLevels().items():
			pass
	priceLevelHeadings = []
	for x in range(numberOfLevels):
		priceLevelHeadings.append('"Price Level ' + str(x + 1) + '"')
	headings = ",".join(priceLevelHeadings)
	headings = '"ID","Name",' + headings
	output_file = str(file_path)[:-4] + "_simplified" + ".csv"
	output = codecs.open(output_file, 'w+', 'utf8')
	output.write(headings + "\r\n")
	tempString = '1,"Things and Stuff",$4.95,,,$5.95,,,,,,$3.99,,\r\n'
	output.write(tempString)

preParse()
generateSimpleExport()
