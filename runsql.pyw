import tkinter as tk
from tkinter import messagebox
from tkinter.simpledialog import askstring
from tkinter.filedialog import asksaveasfile
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
        self.serverName = tk.StringVar()
        self.serverList = []
        self.sqlDirectory = sqlDirectory
        self.sqlSelection = tk.StringVar()
        self.outputSelection = tk.StringVar()
        self.createItems()
        
    def createItems(self):
        self.sqlDirContent = self.getFolderContent(self.sqlDirectory)
        ttk.Label(self, text="Output File").grid(column=5, row=1, padx=(10,0))
        self.optionMenu = ttk.OptionMenu(self, self.sqlSelection, 'SQL Files', *self.sqlDirContent, command=self.touchFile)
        self.optionMenu.grid(column=6, row=0)
        self.outputEntry = ttk.Entry(self, textvariable=self.outputSelection)
        self.outputEntry.grid(column=6, row=1)
        self.serverLabel = tk.Label(self, textvariable=self.serverName)
        self.serverLabel.config(bg="grey25", fg="white")
        self.serverListMenu = ttk.OptionMenu(self, self.serverName, 'Located Servers', *self.serverList)
        ttk.Button(self, text="Run", command=lambda: self.runProgram('runSQL.bat', [self.serverName.get(), self.sqlSelection.get(), self.outputSelection.get()])).grid(column=7, row=5)
        
    def getFolderContent(self, directory):
        results = []
        for item in os.listdir(directory):
            if (os.path.isfile(os.path.join(directory, item))):
                results.append(item)
        return results

    def validateSqlArgs(self, args):
        if (len(args) < 1 or args[1] =='' or args[2] == ''):
            messagebox.showerror(title="Invalid Input", message="Enter output file name")
            return False
        elif (args[0] == '') or (args[0] == "Located Servers"):
            messagebox.showerror(title="Connect Server", message="Scan for nearby servers under the \"Server\" Menu")
            return False
        return True
    
    def runProgram(self, programName, args):
        if (self.validateSqlArgs(args)):
            procOutput = subprocess.run([os.path.join(os.getcwd(), programName), args[0],  args[1], args[2]], capture_output=True, text=True).stdout
            self.showResult(procOutput, os.path.join('txt\\', args[2]))

    def getCurrentSqlSelection(self):
        return self.sqlSelection.get()
    
    def getCurrentSqlDirectory(self):
        return self.sqlDirectory
    
    def touchFile(self, context):
        Path(os.path.join(self.getCurrentSqlDirectory(), self.sqlSelection.get())).touch()

    def addItems(self):
        nextList = self.getFolderContent(self.sqlDirectory)
        for item in nextList:
            if item not in self.sqlDirContent:
                self.sqlDirContent.append(item)
        self.optionMenu = ttk.OptionMenu(self, self.sqlSelection, self.sqlDirContent[-1], *self.sqlDirContent, command=self.touchFile)
        self.optionMenu.grid(column=2, row=0)

    def showResult(self, procValue, resultFileName):
        window = tk.Toplevel(self)
        window.title("SQL Ouput")
        window.geometry("500x500")
        tk.Label(window, text=procValue).pack()
        resultTextArea = tk.Text(window, width=60, height=20, bg='#33ffff', font='Avenir', spacing1=2, wrap=tk.NONE)
        
        resultValue = ''
        with open(resultFileName, 'r') as file:
            resultValue = file.read()

        resultTextArea.insert(tk.INSERT, resultValue)
        resultTextArea.config(state='disabled')

        verticalScrollbar = ttk.Scrollbar(window)
        horizontalScrollbar = ttk.Scrollbar(window, orient=tk.HORIZONTAL)
        resultTextArea.config(yscrollcommand=verticalScrollbar.set, xscrollcommand=horizontalScrollbar.set)
        horizontalScrollbar.config(command=resultTextArea.xview)
        verticalScrollbar.config(command=resultTextArea.yview)
        resultTextArea.pack()
        horizontalScrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        verticalScrollbar.pack(side=tk.LEFT, fill=tk.Y)

        
class Notebook(tk.Frame):
    def __init__(self, master, sqlDirectory):
        super(Notebook, self).__init__()
        self.master = master
        self.sqlDirectory = sqlDirectory
        self.notebook = ttk.Notebook(self)
        self.notebook.enable_traversal()
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
        if (self.notebookIndex != 0):
            ttk.Button(frame, text="Close Tab", command=self.closeTab).pack(side=tk.BOTTOM, padx=5, pady=15)
        self.notebook.add(frame, text=title, underline=0)
        self.notebookIndex += 1

    def closeTab(self):
        self.notebook.forget(self.notebook.select())
        self.notebookIndex -= 1
        

    def refresh(self):
        currentTab = self.notebook.index(self.notebook.select())
        if self.textAreas[currentTab].edit_modified():
            answer = messagebox.askquestion("Updating Modified Text", "Are you sure you want to discard your changes? They will be LOST! Forever...")
            if (answer.lower() != 'yes'):
                return
        self.textAreas[currentTab].delete(1.0, tk.END)
        fileList = glob.glob(self.sqlDirectory + '/*')
        latest = max(fileList, key=os.path.getmtime)
        self.notebook.tab(currentTab, text=latest)
        fileContent = ''
        with open(latest, 'r') as file:
            fileContent = file.read()
        self.textAreas[currentTab].insert(tk.END, fileContent)
        self.textAreas[currentTab].edit_modified(False)
            
        
    def update(self):
        currentTab = self.notebook.index(self.notebook.select())
        if self.textAreas[currentTab].edit_modified():
            fileList = glob.glob(self.sqlDirectory + '/*')
            latest = max(fileList, key=os.path.getmtime)
            with open(latest, 'w') as file:
                file.write(self.textAreas[currentTab].get("1.0", tk.END))

    def saveAs(self):
        currentTab = self.notebook.index(self.notebook.select())
        saveAsFileName = asksaveasfile(initialfile=currentTab, defaultextension=".sql", filetypes=[("All Files","*.*"),("SQL File","*.sql")])
        print(saveAsFileName.name)
        with open(saveAsFileName.name, "w") as file:
            file.write(self.textAreas[currentTab].get("1.0", tk.END))

    def searchTextArea(self):
        currentTab = self.notebook.index(self.notebook.select())
        i = 0
        searchTerm = askstring("Search", "Enter search term:")
        self.textAreas[currentTab].tag_remove('found', '1.0', tk.END)
        while 1:
            i = self.textAreas[currentTab].search(searchTerm, i, nocase=1, stopindex=tk.END)
            if not i:
                break
            lastI = '%s+%dc' % (i, len(searchTerm))
            self.textAreas[currentTab].tag_add('found', i, lastI)
            i = lastI
        self.textAreas[currentTab].tag_config('found', foreground='red')
        
    
class runSQL(tk.Tk):
    def __init__(self, sqlDirectory):
        super().__init__()
        self.sqlDirectory = sqlDirectory
        self.root = tk.Tk()
        self.root.withdraw()

        self.wm_title("runSQL")
        self.rootFrame = tk.Frame(self)
        
        self.geometry("800x600")
        
        self.notebook = Notebook(self.rootFrame, 'sql')
        self.menu = SelectionMenu(self.rootFrame, 'sql')
        self.notebook.addTab(self.menu.getCurrentSqlSelection())
        self.menuBar = tk.Menu(self)
        self.fileMenu = tk.Menu(self.menuBar, tearoff=0)
        
        self.fileMenu.add_command(label="New File", command=self.newFile)
        self.fileMenu.add_command(label="Sync List", command=self.menu.addItems)
        self.fileMenu.add_command(label="Search", command=self.notebook.searchTextArea)
        self.fileMenu.add_command(label="New Tab", command=self.addTabUtility)
        self.fileMenu.add_command(label="Save As", command=self.notebook.saveAs)
        self.fileMenu.add_command(label="Exit", command=self.endProgram)
        self.fileMenu.add_separator()
        self.menuBar.add_cascade(label="File", menu=self.fileMenu)
        
        self.serverMenu = tk.Menu(self.menuBar, tearoff=0)
        self.serverMenu.add_command(label="Find Local", command=self.findLocalServersSync)
        self.serverMenu.add_separator()
        self.menuBar.add_cascade(label="Server", menu=self.serverMenu)
        
        self.menu.pack(padx=20, pady=10)
        self.notebook.pack()
        self.config(menu=self.menuBar)
        self.setIcon()

    def findLocalServersSync(self):
        processObject = subprocess.run('sqlcmd -L', capture_output=True, text=True).stdout
        results = []
        for line in processObject.splitlines():
            line = line.strip()
            if not ((line == '') or (line == 'Servers:')):
                results.append(line)
        self.menu.serverList = results
        self.menu.serverListMenu = ttk.OptionMenu(self.menu, self.menu.serverName, 'Located Servers', *self.menu.serverList)
        self.menu.serverListMenu.grid(row=1, column=0)

    def mainloop(self):
        self.root.mainloop()

    def endProgram(self):
        self.destroy()
        self.root.destroy()

    def setIcon(self):
        self.iconbitmap(os.path.join(os.getcwd(), 'images\\runsql.ico'))

    def newFile(self):
        fileName = askstring("New File", "Please enter a filename:")
        if not os.path.exists(os.path.join(os.getcwd(), self.sqlDirectory, fileName)):
            with open(os.path.join(os.getcwd(), self.sqlDirectory, fileName), 'w') as file:
                pass
        else:
            messagebox.showerror(title="File Already Exists", message="Specify a new name as " + fileName + " already exists in " + self.sqlDirectory)

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
