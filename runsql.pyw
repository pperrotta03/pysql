import tkinter as tk
from tkinter import messagebox
from tkinter.simpledialog import askstring
from tkinter import ttk
import subprocess
import os
import time
from pathlib import Path
import glob
import ctypes

class SelectionMenu(tk.Frame):
    def __init__(self, master, sqlDirectory):
        super(SelectionMenu, self).__init__()
        self.master = master
        self.sqlDirectory = sqlDirectory
        self.sqlSelection = tk.StringVar()
        self.outputSelection = tk.StringVar()
        self.createItems()
        
    def createItems(self):
        self.sqlDirContent = self.getFolderContent(self.sqlDirectory)
        ttk.Label(self, text="Output File").grid(column=1, row=1, padx=(10,0))
        self.optionMenu = ttk.OptionMenu(self, self.sqlSelection, 'SQL Files', *self.sqlDirContent, command=self.touchFile)
        self.optionMenu.grid(column=2, row=0)
        self.outputEntry = ttk.Entry(self, textvariable=self.outputSelection)
        self.outputEntry.grid(column=2, row=1)
        ttk.Button(self, text="Run", command=lambda: self.runProgram("runSQL.bat", [self.sqlSelection.get(), self.outputSelection.get()])).grid(column=0, row=0)
        
    def getFolderContent(self, directory):
        results = []
        for item in os.listdir(directory):
            if (os.path.isfile(os.path.join(directory, item))):
                results.append(item)
        return results
    
    def runProgram(self, fileName, args):
        if (len(args) < 1 or args[0] =='' or args[1] == ''):
            messagebox.showerror(title="Invalid Input", message="Please enter an output file name | Given: " + str(args[-1:]))
            return
        procOutput = subprocess.run([os.path.join(os.getcwd(), self.sqlDirectory, fileName), args[0], args[1]], capture_output=True, text=True).stdout
        self.showResult(procOutput)

    def getCurrentSqlSelection(self):
        return self.sqlSelection.get()
    
    def getCurrentSqlDirectory(self):
        return self.sqlDirectory

    def setSqlSelection(self, item):
        self.sqlSelection = item
    
    def touchFile(self, context):
        Path(os.path.join(self.getCurrentSqlDirectory(), self.getCurrentSqlSelection())).touch()

    def addItems(self):
        nextList = self.getFolderContent(self.sqlDirectory)
        print(nextList)
        for item in nextList:
            if item not in self.sqlDirContent:
                self.sqlDirContent.append(item)
        self.optionMenu = ttk.OptionMenu(self, self.sqlSelection, self.sqlDirContent[-1], *self.sqlDirContent, command=self.touchFile)
        self.optionMenu.grid(column=2, row=0)

    def showResult(self, value):
        window = tk.Toplevel(self)
        window.title("SQL Ouput")
        window.geometry("500x500")
        Label(window, text=value).pack()


class Notebook(tk.Frame):
    def __init__(self, master, sqlDirectory):
        super(Notebook, self).__init__()
        self.master = master
        self.sqlDirectory = sqlDirectory
        self.notebook = ttk.Notebook(self)
        self.notebook.pack()
        self.textAreas = []
        self.notebookIndex = 0

    def addTab(self, title, text=''):
        frame = tk.Frame(self.notebook, bg="grey25", width="100")
        self.verticalScrollbar = ttk.Scrollbar(frame)
        self.horizontalScrollbar = ttk.Scrollbar(frame, orient=tk.HORIZONTAL)
        self.textAreas.append(tk.Text(frame, bg='#33ffff', font='Avenir', spacing1=2, wrap=tk.NONE, width=70, height=255, undo=True, xscrollcommand=self.horizontalScrollbar.set, yscrollcommand=self.verticalScrollbar.set))
        self.textAreas[-1].configure(yscrollcommand=self.verticalScrollbar.set)
        self.horizontalScrollbar.config(command=self.textAreas[-1].xview)
        self.verticalScrollbar.config(command=self.textAreas[-1].yview)
        self.horizontalScrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.verticalScrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.textAreas[-1].pack(side=tk.LEFT, padx=5, pady=3)
        self.textAreas[-1].focus()

        if text != '':
            self.textAreas[-1].insert(tk.END, text)
        
        ttk.Button(frame, text="Refresh", command=self.refresh).pack(pady=20, padx=5)
        ttk.Button(frame, text="Save", command=self.update).pack(padx=5)
        ttk.Button(frame, text="Close Tab", command=lambda: self.notebook.forget(self.notebook.select())).pack(side=tk.BOTTOM, padx=5, pady=15)
        self.notebook.add(frame, text=title)

        self.notebookIndex += 1

    def refresh(self):
        currentTab = self.notebook.index(self.notebook.select())
        self.textAreas[currentTab].delete(1.0, tk.END)
        fileList = glob.glob(self.sqlDirectory + '/*')
        latest = max(fileList, key=os.path.getmtime)
        fileContent = ''
        with open(latest, 'r') as file:
            fileContent = file.read()
        self.textAreas[currentTab].insert(tk.END, fileContent)
        
    def update(self):
        fileList = glob.glob(self.sqlDirectory + '/*')
        latest = max(fileList, key=os.path.getmtime)
        latestPath = os.path.abspath(latest)
        payload = self.textArea.get("1.0", tk.END)
        with open(latest, 'w') as file:
            file.write(payload)

    def searchTextArea(self):
        i = 0
        searchTerm = askstring("Search", "Enter search term:")
        self.textArea.tag_remove('found', '1.0', tk.END)
        while 1:
            i = self.textArea.search(searchTerm, i, nocase=1, stopindex=tk.END)
            if not i:
                break
            lastI = '%s+%dc' % (i, len(searchTerm))
            self.textArea.tag_add('found', i, lastI)
            i = lastI
        self.textArea.tag_config('found', foreground='red')
        
    
class runSQL(tk.Tk):
    def __init__(self, sqlDirectory):
        super().__init__()
        self.sqlDirectory = sqlDirectory
        self.root = tk.Tk()
        self.root.withdraw()

        self.wm_title("runSQL")
        self.rootFrame = tk.Frame(self)
        
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.geometry("800x600")
        self.notebook = Notebook(self.rootFrame, 'sql')
        self.menu = SelectionMenu(self.rootFrame, 'sql')
        self.notebook.addTab(self.menu.getCurrentSqlSelection())
        self.menuBar = tk.Menu(self)
        self.fileMenu = tk.Menu(self.menuBar, tearoff=0)
        self.menu.pack()
        self.notebook.pack()
        self.fileMenu.add_command(label="New File", command=self.newFile)
        self.fileMenu.add_command(label="Sync List", command=self.menu.addItems)
        self.fileMenu.add_command(label="Search", command=self.notebook.searchTextArea)
        self.fileMenu.add_command(label="New Tab", command=self.addTabUtility)
        self.fileMenu.add_command(label="Exit", command=self.endProgram)
        self.fileMenu.add_separator()
        self.menuBar.add_cascade(label="File", menu=self.fileMenu)
        self.config(menu=self.menuBar)

    def mainloop(self):
        self.root.mainloop()

    def endProgram(self):
        self.destroy()
        self.root.destroy()

    def setIcon(self):
        self.iconbitmap(os.path.join(os.curdir, 'images/runsql.ico'))

    def newFile(self):
        fileName = askstring("New File", "Please enter a filename:")
        if not os.path.exists(os.path.join(os.getcwd(), self.sqlDirectory, fileName)):
            with open(os.path.join(os.getcwd(), self.directory, fileName), 'w') as file:
                pass
        else:
            messagebox.showerror(title="File Already Exists", message="Specify a new name as " + fileName + " already exists in " + self.directory)

    def addTabUtility(self):
        title = self.menu.getCurrentSqlSelection()
        content = ''
        with open(os.path.join(os.getcwd(), self.sqlDirectory, title)) as file:
            content = file.read()
        self.notebook.addTab(title, content)

if __name__ == "__main__":
        app = runSQL('sql')
        app.mainloop()
        #https://stackoverflow.com/questions/1551605/how-to-set-applications-taskbar-icon-in-windows-7/1552105
        appId = u'union.runSQL-1.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(appId)
