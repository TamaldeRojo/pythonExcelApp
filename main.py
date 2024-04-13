import tkinter as tk
from tkinter import ttk
import openpyxl
import datetime

class App(tk.Frame):
    def __init__(self,master):
        super().__init__(master)
        self.pack()
        
        self.createLeftWidgets()
        self.createRightWidgets()
        
        
    def setStyleInstance(self,style):
        self.style = style
        
    def toggleMode(self):
        self.style.theme_use("forest-dark") if self.modeSwitch.instate(["selected"]) else self.style.theme_use("forest-light")
        
    def createLeftWidgets(self):
        self.comboList = ["Aldama","Norte","Sur"]
        self.statusList = ["Cliente","Nuevo","Resagado"]

        self.widgetsFrame = ttk.LabelFrame(self, text="Insert Row")
        self.widgetsFrame.grid(row=0,column=0,padx=20,pady=10)
        
        self.nameEntry = ttk.Entry(self.widgetsFrame)
        self.nameEntry.insert(0,"Name")
        self.nameEntry.bind("<FocusIn>", lambda e: self.nameEntry.delete('0','end'))
        self.nameEntry.grid(row=0,column=0,padx=5,pady=(0,5),sticky="ew")
        
        self.ageSpinBox = ttk.Spinbox(self.widgetsFrame, from_=18,to=120)
        self.ageSpinBox.insert(0,"Age")
        self.ageSpinBox.bind("<FocusIn>", lambda e: self.ageSpinBox.delete('0','end'))
        self.ageSpinBox.grid(row=1,column=0,padx=5,pady=5,sticky="ew")
        
        self.phoneEntry = ttk.Entry(self.widgetsFrame)
        self.phoneEntry.insert(0,"Phone")
        self.phoneEntry.bind("<FocusIn>", lambda e: self.phoneEntry.delete('0','end'))
        self.phoneEntry.grid(row=2,column=0,padx=5,pady=(0,5),sticky="ew")
        
        self.locationCombo = ttk.Combobox(self.widgetsFrame,values=self.comboList)
        self.locationCombo.current(0)
        self.locationCombo.grid(row=3,column=0,padx=5,pady=5,sticky="ew")
        
        self.statusCombo = ttk.Combobox(self.widgetsFrame,values=self.statusList)
        self.statusCombo.current(0)
        self.statusCombo.grid(row=4,column=0,padx=5,pady=5,sticky="ew")
        
        self.btnInsert = ttk.Button(self.widgetsFrame,text="Insert", command=self.insertRow)
        self.btnInsert.grid(row=5,column=0,padx=5,pady=5,sticky="nsew")
        
        self.separator = ttk.Separator(self.widgetsFrame)
        self.separator.grid(row=6,column=0,padx=5,pady=(20,10),sticky="ew")

        self.b = tk.BooleanVar(value=True)
        self.modeSwitch = ttk.Checkbutton(self.widgetsFrame,text="Dark Mode",variable=self.b,command= self.toggleMode )
        self.modeSwitch.grid(row=7,column=0,padx=5,pady=10,sticky="nsew")
        
    def createRightWidgets(self):
        self.treeFrame = ttk.Frame(self)
        self.treeFrame.grid(row=0,column=1,pady=10)
        
        self.treeScrollBar = ttk.Scrollbar(self.treeFrame)
        self.treeScrollBar.pack(side="right",fill="y")
        
        self.cols = ("Name","Age","Phone","Status","Zona","Fecha")
        self.treeView = ttk.Treeview(self.treeFrame,show="headings",yscrollcommand=self.treeScrollBar.set,columns=self.cols,height=13)
        self.treeView.column('Name',width=100)
        self.treeView.column('Age',width=50)
        self.treeView.column('Phone',width=100)
        self.treeView.column('Status',width=100)
        self.treeView.column('Fecha',width=100)
        self.treeView.column('Zona',width=100)

        self.treeView.pack()
        self.treeScrollBar.config(command=self.treeView.yview)
        
        self.loadData()
        
    def loadData(self):
        self.path = r"C:\Users\josue\OneDrive\Escritorio\pythonExcelApp\terrenos.xlsx"
        self.workbook = openpyxl.load_workbook(self.path)
        
        self.sheet = self.workbook.active
        self.listValues = list(self.sheet.values)
        # print(self.listValues)
        for col in self.listValues[0]:
            self.treeView.heading(col,text=col)
        for colValue in self.listValues[1:]:
            self.treeView.insert('',tk.END,values=colValue)
        
    def insertRow(self):
        try:
            self.nameForm = self.nameEntry.get()
            self.ageForm = int(self.ageSpinBox.get())
            self.phoneForm = str(self.phoneEntry.get())
            self.locationForm = self.locationCombo.get()
            self.statusForm = self.statusCombo.get()
            self.date = datetime.datetime.now()
            
            # self.path = r"C:\Users\josue\OneDrive\Escritorio\pythonExcelApp\terrenos.xlsx"
            self.workbook = openpyxl.load_workbook(self.path)
            self.sheet = self.workbook.active
            self.rowVal = [self.nameForm,self.ageForm,self.phoneForm,self.locationForm,self.statusForm,self.date]
            self.sheet.append(self.rowVal)
            self.workbook.save(self.path)
            
            
            self.treeView.insert('',tk.END,values=self.rowVal)
            
            self.nameEntry.delete(0,'end')
            self.nameEntry.insert(0,'Name')
            self.ageSpinBox.delete(0,'end')
            self.ageSpinBox.insert(0,'Age')
            self.phoneEntry.delete(0,'end')
            self.phoneEntry.insert(0,'Phone')
            self.statusCombo.set(self.comboList[0])
            self.locationCombo.set(self.locationCombo[0])
            
        except Exception as e:
          print(f'[-] An exception occurred: {e}')
        
 
def main():
    root = tk.Tk()
    app = App(root)
    
    style = ttk.Style(app)
    app.setStyleInstance(style)
    app.tk.call("source","forest-light.tcl")
    app.tk.call("source","forest-dark.tcl")    
    style.theme_use("forest-dark")
    app.mainloop()

if __name__ == "__main__":
    main()