import tkinter as tk
from tkinter import messagebox
from tkinter.simpledialog import askstring
from tkinter.filedialog import asksaveasfile
from tkinter.simpledialog import Dialog
from tkinter import ttk
import subprocess
import os
import time
from pathlib import Path
import glob
import ctypes
import json

class SqlConfigPopup(tk.Toplevel):
    def __init__(self, master, title='SQLCMD Arguments', width=400, height=400):
        tk.Toplevel.__init__(self, master)
        self.geometry(f'{width}x{height}')
        self.config(bg='grey50')
        self.master = master
        self.rowsPerHeader = tk.IntVar()
        self.printStatsStatus = tk.StringVar()
        self.trailingSpacesStatus = tk.StringVar()
        self.separator = tk.StringVar()
        self.printStatsStatus.set('Off')
        self.trailingSpacesStatus.set('Off')
        self.separator.set(',')
        self.configFile = os.path.join(os.getcwd(), 'config.data')
        
        self.sqlArguments = {'W':'Off','p':'Off','s':'","','h':'0'}
        
        self.createItems()
        
        if os.path.isfile(os.path.basename(self.configFile)):
            with open(self.configFile, 'r') as file:
                self.sqlArguments = json.loads(file.read())
                try:
                    self.rowsPerHeader.set(self.sqlArguments['h'])
                    self.incrementRPHValue(0)
                    self.trailingSpacesStatus.set(self.sqlArguments['W'])
                    if self.sqlArguments['W'].lower() == 'on':
                        self.trailingSpacesLabel.config(bg='green')
                    self.separator.set(self.sqlArguments['s'])
                    self.printStatsStatus.set(self.sqlArguments['p'])
                    if self.sqlArguments['p'].lower() == 'on':
                        self.printStatsLabel.config(bg='green')
                except(KeyError):
                    print(f'KeyError - key not stored in {self.configFile}')
        
    def createItems(self):
        self.rphFrame = tk.Frame(self, bg='grey50')
        self.printStatsFrame = tk.Frame(self, bg='grey50')
        self.trailingSpacesFrame = tk.Frame(self, bg='grey50')
        self.separatorFrame = tk.Frame(self, bg='grey50')
        
        self.incrementRPHButton = tk.Button(self.rphFrame, text='Up', font='system', command=lambda: self.incrementRPHValue(1))
        self.decrementRPHButton = tk.Button(self.rphFrame, text='Down', font='system', command=lambda: self.incrementRPHValue(-1))
        self.rphLabel = ttk.Label(self.rphFrame, text='-h | Rows Per Header', font='system')
        self.currentRPHLabel = tk.Label(self.rphFrame, textvariable=self.rowsPerHeader, bg='green', font='system 18')
        self.printStatsLabel = tk.Label(self.printStatsFrame, textvariable=self.printStatsStatus, bg='red', font='system 18')
        self.printStatsButton = tk.Button(self.printStatsFrame, text='-p | Print Statistics', relief='sunken', font='system', command=self.togglePrintStats)
        self.trailingSpacesLabel = tk.Label(self.trailingSpacesFrame, textvariable=self.trailingSpacesStatus, bg='red', font='system 18')
        self.trailingSpacesButton = tk.Button(self.trailingSpacesFrame, text='-W | Remove Trailing Whitespaces', relief='sunken', font='system', command=self.toggleTrailingSpaces)
        self.separatorLabel = tk.Label(self.separatorFrame, text='Separator', font='system')
        self.currentSeparatorLabel = tk.Label(self.separatorFrame, textvariable=self.separator, bg='red', font='system 18')
        self.separatorEntry = tk.Entry(self.separatorFrame, font='system 18', width=3)
        self.separatorButton = tk.Button(self.separatorFrame, text='Set', font='system', command=self.setSeparator)
        self.separatorEntry.insert(0, 'ie |,')
        self.separatorEntry.bind("<FocusIn>", self.resetSeparatorLabel)
        
        self.rphFrame.grid_rowconfigure(0, weight=1)
        
        self.rphFrame.pack(anchor="w", expand=True, padx=20, pady=(20,0))
        self.currentRPHLabel.grid(row=0, column=0, padx=5)
        self.rphLabel.grid(row=0, column=1)
        self.incrementRPHButton.grid(row=0, column=2, padx=10)
        self.decrementRPHButton.grid(row=0, column=3)
        
        self.printStatsFrame.pack(anchor="w", expand=True, padx=20, pady=(5,5))
        self.printStatsLabel.grid(row=0, column=0, padx=5)
        self.printStatsButton.grid(row=0, column=1)

        self.trailingSpacesFrame.pack(anchor="w", expand=True, padx=20, pady=(5,5))
        self.trailingSpacesLabel.grid(row=0, column=0, padx=5)
        self.trailingSpacesButton.grid(row=0, column=1)

        self.separatorFrame.pack(anchor="w", expand=True, padx=20, pady=(5,5))
        self.currentSeparatorLabel.grid(row=0, column=0, padx=5)
        self.separatorEntry.grid(row=0, column=1, padx=5)
        self.separatorLabel.grid(row=0, column=2, padx=5)
        self.separatorButton.grid(row=0, column=3, padx=5)

    def updateArg(self, key, value):
        try:
            self.sqlArguments[key] = value
        except(KeyError):
            print(f'KeyError: Updating {key} - does not exists in sqlArguments at this time')
        with open(self.configFile, 'w') as file:
            args = json.dumps(self.sqlArguments)
            file.write(args)

    def resetSeparatorLabel(self, context=None):
        self.separatorEntry.delete(0, tk.END)

    def setSeparator(self):
        self.separator.set(self.separatorEntry.get()[:3].strip())
        self.updateArg('s', '"' + self.separator.get() + '"')
        

    def toggleTrailingSpaces(self):
        if self.trailingSpacesStatus.get().lower() == 'off':
            self.sqlArguments['W'] = 'On'
            self.trailingSpacesButton.config(relief='raised')
            self.trailingSpacesStatus.set('On')
            self.trailingSpacesLabel.config(bg='green')
        else:
            self.trailingSpacesButton.config(relief='sunken')
            try:
                del self.sqlArguments['W']
            except(KeyError):
                print('Key Error: "W"')
            self.trailingSpacesStatus.set('Off')
            self.trailingSpacesLabel.config(bg='red')
        self.updateArg('W', self.trailingSpacesStatus.get())

    def togglePrintStats(self):
        if self.printStatsStatus.get().lower() == 'off':
            self.sqlArguments['p'] = 'On'
            self.printStatsButton.config(relief='raised')
            self.printStatsStatus.set('On')
            self.printStatsLabel.config(bg='green')
        else:
            self.printStatsButton.config(relief='sunken')
            try:
                del self.sqlArguments['p']
            except(KeyError):
                print('Key Error: "p"')
            self.printStatsStatus.set('Off')
            self.printStatsLabel.config(bg='red')
        self.updateArg('p', self.printStatsStatus.get())
        
    def __repr__(self):
        return str(type(self))

    def __str__(self):
        return str(self.sqlArguments)

    def incrementRPHValue(self, byValue):
        result = self.rowsPerHeader.get() + byValue
        if (result < 0):
            result = -1
            self.currentRPHLabel.config(bg='red')
        else:
            self.currentRPHLabel.config(bg='green')
        self.rowsPerHeader.set(result)
        self.updateArg('h', self.rowsPerHeader.get())


class SelectionMenu(tk.Frame):
    def __init__(self, master, sqlDirectory):
        super(SelectionMenu, self).__init__()
        self.master = master
        self.serverName = tk.StringVar()
        self.serverList = []
        self.sqlDirectory = sqlDirectory
        self.sqlSelection = tk.StringVar()
        self.outputSelection = tk.StringVar()
        self.sqlcmdArgs = {}
        self.createItems()
        
    def createItems(self):
        self.sqlDirContent = self.getFolderContent(self.sqlDirectory)
        ttk.Label(self, text='Output File').grid(column=5, row=1, padx=(10,0))
        self.optionMenu = ttk.OptionMenu(self, self.sqlSelection, 'SQL Files', *self.sqlDirContent, command=self.touchFile)
        self.optionMenu.grid(column=6, row=0)
        self.outputEntry = ttk.Entry(self, textvariable=self.outputSelection)
        self.outputEntry.grid(column=6, row=1)
        self.serverLabel = tk.Label(self, textvariable=self.serverName)
        self.serverLabel.config(bg="grey25", fg="white")
        self.serverListMenu = ttk.OptionMenu(self, self.serverName, 'Located Servers', *self.serverList)
        ttk.Button(self, text='Run', command=self.runProgram).grid(column=7, row=5, padx=50)
        
    def getFolderContent(self, directory):
        results = []
        for item in os.listdir(directory):
            if (os.path.isfile(os.path.join(directory, item))):
                results.append(item)
        return results

    def validateSqlArgs(self):
        if (self.outputSelection.get() == ''):
            messagebox.showerror(title='Invalid Input', message='Enter output file name')
            return False
        elif (self.serverName.get() == '') or (self.serverName.get() == 'Located Servers'):
            messagebox.showerror(title='Connect Server', message='Scan for nearby servers under the "Server" Menu')
            return False
        return True
    
    def runProgram(self):
        if (self.validateSqlArgs()):
            command = f'sqlcmd -S {self.serverName.get()} -i "{os.path.join(os.getcwd(), "sql", self.sqlSelection.get())}" -o "{os.path.join(os.getcwd(), "txt", self.outputSelection.get())}"'
            for key in self.sqlcmdArgs:
                value = self.sqlcmdArgs[key]
                if value == 'Off':
                    continue
                if (key == 'p') or (key == 'W'):
                    value = ''
                command = f'{command} -{key} {value}'
            procOutput = subprocess.run(command)
            self.showResult(procOutput, os.path.join("txt", self.outputSelection.get()))

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
        textFrame = tk.Frame(window, bg='grey50')
        resultTextArea = tk.Text(textFrame, width=60, height=20, bg='#33ffff', font='Avenir', spacing1=2, wrap=tk.NONE)
        
        resultValue = ''
        with open(resultFileName, 'r') as file:
            resultValue = file.read()

        resultTextArea.insert(tk.INSERT, resultValue)
        resultTextArea.config(state='disabled')

        verticalScrollbar = ttk.Scrollbar(textFrame)
        horizontalScrollbar = ttk.Scrollbar(textFrame, orient=tk.HORIZONTAL)

        textFrame.pack()
        resultTextArea.config(yscrollcommand=verticalScrollbar.set, xscrollcommand=horizontalScrollbar.set)
        horizontalScrollbar.config(command=resultTextArea.xview)
        verticalScrollbar.config(command=resultTextArea.yview)
        horizontalScrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        verticalScrollbar.pack(side=tk.LEFT, fill=tk.Y)
        resultTextArea.pack(side=tk.RIGHT)

        
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
        self.verticalScrollbar.pack(side=tk.RIGHT, fill=tk.Y)
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
        with open(saveAsFileName.name, "w") as file:
            file.write(self.textAreas[currentTab].get("1.0", tk.END))

    def searchTextArea(self):
        currentTab = self.notebook.index(self.notebook.select())
        i = '1.0'
        while True:
            searchTerm = askstring("Search", "Enter search term:")
            if searchTerm != '':
                continue
            break
            
        self.clearSearchTags()
        while 1:
            i = self.textAreas[currentTab].search(searchTerm, i, nocase=1, stopindex=tk.END)
            if not i:
                break
            lastI = '%s+%dc' % (i, len(searchTerm))
            self.textAreas[currentTab].tag_add('found', i, lastI)
            i = lastI
        self.textAreas[currentTab].tag_config('found', background='grey60')

    def clearSearchTags(self):
        currentTab = self.notebook.index(self.notebook.select())
        self.textAreas[currentTab].tag_remove('found', '1.0', tk.END)
        
    
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
        self.fileMenu.add_command(label="Reset Search", command=self.notebook.clearSearchTags)
        self.fileMenu.add_command(label="New Tab", command=self.addTabUtility)
        self.fileMenu.add_command(label="Save As", command=self.notebook.saveAs)
        self.fileMenu.add_separator()
        self.fileMenu.add_command(label="Exit", command=self.endProgram)
        self.menuBar.add_cascade(label="File", menu=self.fileMenu)
        
        self.serverMenu = tk.Menu(self.menuBar, tearoff=0)
        self.serverMenu.add_command(label="Find Instances", command=self.findLocalServersSync)
        self.serverMenu.add_separator()
        self.serverMenu.add_command(label="Configure SQLCMD", command=self.sqlConfigPopupWindow)
        self.menuBar.add_cascade(label="Server", menu=self.serverMenu)
        
        self.menu.pack(padx=20, pady=10)
        self.notebook.pack()
        self.config(menu=self.menuBar)
        self.setIcon()

    def sqlConfigPopupWindow(self):
        self.sqlConfigPopup = SqlConfigPopup(self.rootFrame)
        self.sqlConfigPopup.wait_window()
        self.menu.sqlcmdArgs = self.sqlConfigPopup.sqlArguments

    def findLocalServersSync(self):
        processOutput = subprocess.run('sqlcmd -Lc', capture_output=True, text=True).stdout
        results = []
        for line in processOutput.splitlines():
            line = line.strip()
            results.append(line)
        self.menu.serverList = results
        self.menu.serverListMenu = tk.OptionMenu(self.menu, self.menu.serverName, *self.menu.serverList)
        self.menu.serverListMenu.config(bg='grey75')
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
