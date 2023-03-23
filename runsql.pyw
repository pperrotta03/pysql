import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import subprocess
import os
import time
from pathlib import Path
import glob


class EditorSection(tk.Frame):
    def __init__(self, master, directory='sql'):
        super(EditorSection, self).__init__()
        self.master = master
        self.editorFrame = tk.Frame(self.master, bg="grey25", padx=5, pady=20)
        self.createItems()
        self.editorFrame.grid(row=0, column=2)
        self.directory = directory
    def createItems(self):
        self.editorLabel = ttk.Label(self.editorFrame, background='#33ccff' , text="SQL Editor").pack(padx=5, pady=3)
        self.scrollbar = ttk.Scrollbar(self.editorFrame)
        self.textArea = tk.Text(self.editorFrame, bg='#33ffff', font='Avenir', width=60, height=255)
        self.textArea.configure(yscrollcommand=self.scrollbar.set)
        self.textArea.pack(side=tk.LEFT, padx=5, pady=3)
        self.scrollbar.config(command=self.textArea.yview)
        self.scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.textArea.focus()
        self.refreshButton = ttk.Button(self.editorFrame, text="Refresh", command=self.refresh).pack(pady=20, padx=5)
        self.updateButton = ttk.Button(self.editorFrame, text="Save", command=self.update).pack(padx=5)
    def refresh(self):
        self.textArea.delete(1.0, tk.END)
        fileList = glob.glob(self.directory + '/*')
        latest = max(fileList, key=os.path.getmtime)
        fileContent = ''
        with open(latest, 'r') as file:
            fileContent = file.read()
        self.textArea.insert(tk.END, fileContent)
    def update(self):
        fileList = glob.glob(self.directory + '/*')
        latest = max(fileList, key=os.path.getmtime)
        latestPath = os.path.abspath(latest)
        payload = self.textArea.get("1.0", tk.END)
        with open(latest, 'w') as file:
            file.write(payload)
            

class SelectionSection(tk.Frame):
    def __init__(self, master):
        super(SelectionSection, self).__init__()
        self.master = master
        self.sqlSelection = tk.StringVar()
        self.outputSelection = tk.StringVar()
        self.selectionFrame = tk.Frame(self.master, padx=50, pady=50, bg='grey25')
        self.createItems()
        self.selectionFrame.grid(row=0, column=0, sticky='N')

    def createItems(self):
        self.sqlDirContent = self.getFolderContent('sql')
        self.label = ttk.Label(self.selectionFrame, text="Output File").grid(column=1, row=1, padx=(10,0))
        self.optionMenu = ttk.OptionMenu(self.selectionFrame, self.sqlSelection, 'SQL Files', *self.sqlDirContent, command=self.touchFile).grid(column=2, row=0)
        self.outputEntry = ttk.Entry(self.selectionFrame, textvariable=self.outputSelection).grid(column=2, row=1)
        self.runButton = ttk.Button(self.selectionFrame, text="Run", command=lambda: self.runProgram("runSQL.bat", [self.sqlSelection.get(), self.outputSelection.get()])).grid(column=0, row=0)
        self.quitButton = ttk.Button(self.selectionFrame, text="Quit", command=self.master.destroy).grid(column=0, row=5)
        
    def getFolderContent(self, dirName):
        results = []
        for item in os.listdir(dirName):
            if (os.path.isfile(os.path.join(dirName, item))):
                results.append(item)
        return results
    
    def runProgram(self, fileName, args):
        if (len(args) < 1 or args[0] =='' or args[1] == ''):
            messagebox.showerror(title="Invalid Input", message="Please enter an output file name | Given: " + str(args[-1:]))
            return
        procOutput = subprocess.run([os.path.join(os.getcwd(), fileName), args[0], args[1]], capture_output=True, text=True).stdout

    def getCurrentSqlSelection(self):
        return self.sqlSelection.get()
    def touchFile(self, master):
        Path(os.path.join('sql', self.getCurrentSqlSelection())).touch()


#api
class runSQL(object):
    def __init__(self, master):
        self.master = master
        self.editorSection = EditorSection(self.master)
        self.sectionOne = SelectionSection(self.master)

root = tk.Tk()
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(0, weight=1)
app = runSQL(root)
root.geometry("1100x350")

root.mainloop()
