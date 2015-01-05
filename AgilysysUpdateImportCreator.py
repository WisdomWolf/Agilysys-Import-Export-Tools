#!/usr/bin/python

import re
import codecs
import tkinter
import os
from Things import MenuItemThings

from tkinter import filedialog


root = tkinter.Tk()
root.withdraw()

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

generateIGPriceUpdate()
os._exit(0)
