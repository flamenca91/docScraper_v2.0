#from __future__ import annotations
import docx
#from docx import Document
#from docx.shared import RGBColor
import re
from docSearch import *
from docExtract import *
#import xlwings

from tkinter import *
from docSearch import *
from tkinter import filedialog

#***************************************** CREATE GUI ***************************************
root = Tk()                # Initiates form with size and background color
root.title("TARGEST_App")
root.geometry("800x400")
root['background'] = '#afeae6'

# ******************************** Test File Search *********************************

'''
def GetDocDict(filename):
    global docFile
    dodFile = {}
    f = open(docFile)
    #f = open("DocTagsList.txt")
    for line in f:
        line = line.split()
        docFile[line[0]] = line[1]
    return docFile
'''

def openFile():
    global filename
    filename = filedialog.askopenfilename(parent=root, title='Choose a file')
    filePath.set(filename)
    print(filename)

def openDirectory():
    foldername = filedialog.askdirectory(parent=root,title='Choose a file')
    dirPath.set(foldername)
    print(foldername)

filePath = StringVar(None)
filepath = Entry(root, borderwidth=5, textvariable=filePath)
filepath.grid(column=0,row=0,sticky='EW')
filepath.update()
filepath.focus_set()
filepath.pack(padx = 20, pady = 20,anchor='n')
filepath.place(y = 25, x = 100, width = 525, height = 25)

button = Button(root, text='Select tags/docs file',command = openFile)
button.pack(padx = 1, pady = 1,anchor='ne')
button.place( x = 650, y = 25)

dirPath = StringVar(None)
dirpath = Entry(root, borderwidth=5, textvariable=dirPath)
dirpath.grid(column=0,row=0,sticky='EW')
dirpath.update()
#dirpath.focus_set()
dirpath.pack(padx = 20, pady = 20,anchor='n')
dirpath.place(y = 100, x = 100, width = 525, height = 25)

button = Button(root, text='Select docs directory',command = openDirectory)
button.pack(padx = 1, pady = 1,anchor='ne')
button.place( x = 650, y = 105)

#command=lambda: [fun1(), fun2()])
tagFileBtn = Button(root, text='Extract Data', padx=40, command = GetRelationalList ,pady=10, bg = 'white')
tagFileBtn.place(x=100,y=160)

root.mainloop()

'''
leadTags = GetLeadingTags()[0]
print("There are", len(leadTags), "leading tags")
print(leadTags[106])
reducedLeadingTags = GetRelLeadTags()
print("There are", len(reducedLeadingTags[0]), "leading tags excluding RISK and URS")
trailingTagsList = GetTrailingTags()
#print(trailingTagsList[66])
print("There are", len(trailingTagsList), "trailing tags")
DescriptionTagsList = GetLeadingTags()[1]
print("There are", len(DescriptionTagsList), "tags with descriptions")
uniqueTagsList = GetUniqueTrailingTags()
print("There are", len(uniqueTagsList), "unique trailing tags")
verifTags = GetVerifiedUniqueTrailingTags()
print("There are", len(verifTags[0]), "tags without PASS/FAIL data or Orphans")
print("There are", len(verifTags[1]), "orphan tags")
'''
