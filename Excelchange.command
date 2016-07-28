#!/usr/bin/python
import pip
import subprocess

def install(package):
    subprocess.call(['sudo','pip', 'install', package])

if __name__ == '__main__':
    install('xlrd')

from xlrd import open_workbook
from Tkinter import *
import tkMessageBox
import tkFileDialog
import Tkinter
import os

class ExcelChange(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid()
        self.createGUI()

    def createGUI(self):
        self.excelBtn = Button(self,text ="Select File", command = self.readexcel)
        self.excelBtn.grid(row=0,column = 0)
        
        self.fileLabel = Label(self)
        self.fileLabel["text"] = ""
        self.fileLabel.grid(row=0,column=1)

        self.v = IntVar()
        btn = Radiobutton(self, text="ios", variable=self.v, value=1)
        btn.select()
        btn.grid(row=1,column=0)
        btn = Radiobutton(self, text="android", variable=self.v, value=2)
        btn.grid(row=1,column=1)

        self.list = Listbox(self)
        self.list["width"] = 50
        self.list.grid(row=2,column = 0,columnspan=50)
        self.list.activate(0)
        
        self.convertBtn = Button(self,text ="Convert File", command = self.convert)
        self.convertBtn.grid(row=3,column = 0)
        self.convertBtn.config(state="disabled")
    
        print(self.v.get())
    
    def convert(self):
        index = self.list.curselection()[0]
        name = self.wb.sheet_names()[index]
        sheet = self.wb.sheet_by_name(name)
        number_rows = sheet.nrows
        number_columns = sheet.ncols
        localStringArr = ["" for i in range(number_columns-1)]
        for row in range(1,number_rows):
            for col in range(1,number_columns):
                key = sheet.cell(row,0).value
                value = sheet.cell(row,col).value
                if self.v.get() == 1:
                    localStringArr[col-1] += ("\""+ key + "\" = " + "\""+ value + "\"" + ";\n")
                else:
                    localStringArr[col-1] += ("<string name=\""+ key + "\">"+ value + "</string>\n")

        title = "ILocalString_" if self.v.get()==1 else "ALocalString_"
        currentPath = os.path.dirname(os.path.realpath(__file__))
        for col in range(1,number_columns):
            stringType = sheet.cell(0,col).value
            print("PATH:" + currentPath)
            with open(os.path.join(currentPath,title + stringType + ".strings"),"wb") as f:
                f.write(localStringArr[col-1].encode('utf-8'))


        alertTitle = "IOS String File Completed" if self.v.get()==1 else "Android String File Completed"
        tkMessageBox.showinfo(alertTitle, alertTitle)
        
    def readexcel(self):
        Tk().withdraw()
        dirname = tkFileDialog.askopenfilename()
        self.wb = open_workbook(dirname)
        self.fileLabel["text"] = "File Name:" + dirname.split('/')[-1]
        self.list.delete(1,END)
        self.convertBtn.config(state="normal")
        for item in self.wb.sheet_names():
            self.list.insert(END, item)
        self.list.select_set(0)
def on_closing():
    root.destroy()

root = Tk()
root.wm_title("LocalString Helper")
root.protocol("WM_DELETE_WINDOW", on_closing)
app = ExcelChange(master=root)
app.mainloop()

#top = Tkinter.Tk()
#top.wm_title("Excel Helper")
#listbox = Listbox(top)
#listbox.pack()
#frame = Frame(top, width=300, height=300)
#frame.pack()
#path = ExcelChange()
#top.mainloop()