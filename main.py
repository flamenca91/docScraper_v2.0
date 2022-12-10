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
filepath = Entry(root, borderwidth=5, textvariable=filePath)    # Creates a textbox for the tags/Documents file path
filepath.grid(column=0,row=0,sticky='EW')
filepath.update()
filepath.focus_set()
filepath.pack(padx = 20, pady = 20,anchor='n')
filepath.place(y = 25, x = 100, width = 525, height = 25)

button = Button(root, text='Select tags/docs file',command = openFile)
button.pack(padx = 1, pady = 1,anchor='ne')
button.place( x = 650, y = 25)

dirPath = StringVar(None)
dirpath = Entry(root, borderwidth=5, textvariable=dirPath)   # Creates a textbox for the documents directory
dirpath.grid(column=0,row=0,sticky='EW')
dirpath.update()
#dirpath.focus_set()
dirpath.pack(padx = 20, pady = 20,anchor='n')
dirpath.place(y = 100, x = 100, width = 525, height = 25)

button = Button(root, text='Select docs directory',command = openDirectory) # Creates a button to run all the code
button.pack(padx = 1, pady = 1,anchor='ne')
button.place( x = 650, y = 105)

# When button is clicked, it runs the code in the docExtract.py file which contains the GetRelationaList function
tagFileBtn = Button(root, text='Extract Data', padx=40, command = GetRelationalList ,pady=10, bg = 'white')
tagFileBtn.place(x=100,y=160)

root.mainloop()

