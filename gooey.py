#!/usr/bin/python

import os
import re
import codecs
from tkinter import *
from tkinter import ttk
from Things import MenuItemThings
from tkinter import filedialog

priceArrayMatch = re.compile(r'(?<=\{)[^(\{|\})].+?(?=\})')
itemList = []
itemMap = {}

def calculate(*args):
    try:
        value = float(feet.get())
        meters.set((0.3048 * value * 10000.0 + 0.5)/10000.0)
    except ValueError:
        pass
    
def openFile(**options):
    if options == None:
        options = {}
        options['defaultextension'] = '.txt' 
        options['filetypes'] = [('Text Files', '.txt'), ('CSV Files', '*.csv*'), ('All Files', '.*')]
        options['title'] = 'Open Agilysys Export'
    file_opt = options
    global file_path
    file_path = filedialog.askopenfilename(**file_opt)
    if file_path == None or file_path == "":
        print("No file selected")
    openFileString.set(str(file_path))
    
def saveFile(**options):
    options = {}
    options['title'] = 'Save As'
    options['initialfile'] = str(file_path)[:-4] + "_simplified" + str(file_path)[-4:]
    file_opt = options
    global save_file
    global simple_file
    save_file = str(file_path)[:-4] + ".csv"
    simple_file = filedialog.asksaveasfilename(**file_opt)
    if save_file == None or save_file == "":
        print("No file selected")
    saveFileString.set(str(save_file))
        
def fixArray(match):
    match = str(match.group(0))
    return match.replace(",",";")
    
def preParse(export, output):
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
        
def simplifyExport(export=None, output=None):
    if export == None:
        export = codecs.open(file_path, 'r', 'utf8')
    if output == None:
        output = codecs.open(save_file, 'w+', 'utf8')
    
    preParse(export, output)
    generateSimpleExport()
    
def newFile():
    pass
        
root = Tk()
root.option_add('*tearOff', FALSE)
root.title("Agilysy File Tools")

menubar = Menu(root)
menu_file = Menu(menubar)
menu_edit = Menu(menubar)
menubar.add_cascade(menu=menu_file, label='File')
menubar.add_cascade(menu=menu_edit, label='Edit')

menu_file.add_command(label='New', command=newFile)
menu_file.add_command(label='Open...', command=openFile)
menu_file.add_command(label='Close', command=root.quit)

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=1, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(1, weight=1)

openFileString = StringVar()
saveFileString = StringVar()

ttk.Label(mainframe, text="Input File:").grid(column=1, row=1, sticky=(N, W, E))
openFile_entry = ttk.Entry(mainframe, width=50, textvariable=openFileString)
openFile_entry.grid(column=1, row=2, sticky=(W, E))
ttk.Button(mainframe, text="Browse", command=openFile).grid(column=2, row=2, sticky=W)

ttk.Label(mainframe, text="Output File:").grid(column=1, row=3, sticky=(W,E))
saveFile_entry = ttk.Entry(mainframe, width=50, textvariable=saveFileString)
saveFile_entry.grid(column=1, row=4, sticky=(W,E))
ttk.Button(mainframe, text="Browse", command=saveFile).grid(column=2, row=4, sticky=W)

ttk.Button(mainframe, text="Simplify", command=simplifyExport).grid(column=1, row=5)


for child in mainframe.winfo_children(): child.grid_configure(padx=5, pady=5)

root.config(menu=menubar)

root.mainloop()