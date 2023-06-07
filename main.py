#Libraries Setup
from pywinauto import Application
import win32process, win32gui, win32con, time, win32api, wmi

from tkinter import *
from tkinter import ttk

SEPARATOR = "------------------------------------"

root = Tk()
root.title("wlock0.0")
root.geometry("300x500+150+200")

s = ttk.Style()
s.configure("Header.TLabel", anchor=CENTER)

c = wmi.WMI()

class Main:
    def __init__(self):
        #Set up the information-storing lists.
        self.wlist = []
        self.wtargt = ["Discord", "Telegram"]
        self.wtargs = []

        self.getWinds()

        #Top Bar
        self.topLabel = ttk.Label(root, text="Test", style="Header.TLabel")
        self.topBar = Label(root, bg="#000000")
       
        #Body
        self.body = ttk.Frame(root)

        #Selection
        self.cBoxVar = StringVar()
        self.cBox = ttk.Combobox(self.body, textvariable=self.cBoxVar, state="readonly", values=self.wlist)
        self.cBox.set("Select a window.")
        self.cBox.bind("<<ComboboxSelected>>", self.CBoxSelected)

        self.UISetup()

    def CBoxSelected(self,n):
        print("Selected")

    def UISetup(self):
        #Place UI Objects
        self.topBar.place(relwidth=0.8, relheight=0.002, relx=0.5, rely=0.1, anchor="s")
        self.topLabel.place(relwidth=1, relheight=0.1, relx=0.5, rely=0, anchor="n")

        self.body.place(relwidth=0.9, relheight=0.8, relx=0.5, rely=0.11, anchor="n")

        self.cBox.place(relwidth=0.95, relheight=0.07, relx=0.5, rely=0, anchor="n")

    def getFullName(self,pid):
        try:
            for p in c.query('SELECT ExecutablePath FROM Win32_Process WHERE ProcessId = %s' % str(pid)): #Query for a PID match.
                return p.ExecutablePath #Return the full path.
        except:
            return None

    def getName(self,fn):
        try:
            language, codepage = win32api.GetFileVersionInfo(fn, '\\VarFileInfo\\Translation')[0]
            stringFileInfo = u'\\StringFileInfo\\%04X%04X\\%s' % (language, codepage, "FileDescription")
            return win32api.GetFileVersionInfo(fn, stringFileInfo)
        except:
            return None
    
    def EnumWindows(self,hwnd,ctx):
        if win32gui.IsWindowVisible(hwnd): #Check if window is visible on the screen.
            wtext = win32gui.GetWindowText(hwnd)
            if len(wtext) > 0:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)

                fullname = self.getFullName(pid)
                name = self.getName(fullname)

                if name and len(name) > 0:
                    self.wlist.append(name) #If so, append it into the list of all active windows.
                    #Loop through the targets specified for window locking.
                    for i in self.wtargt:
                        if i.lower() in wtext.lower(): #Does the given name match the current window?
                            #It does, add its info to the targets list.
                            self.wtargs.append({"KEYWORD": i, "NAME": name, "PID": pid})
                            break #Break the loop as we have gotten what we want already.

    def getWinds(self):
        #Clear the list of windows & lock targets.
        self.wlist.clear() 
        self.wtargs.clear()

        #Use EnumWindows to loop through every process.
        win32gui.EnumWindows(self.EnumWindows,None)

        return self.wlist, self.wtargs
        

m = Main()

root.mainloop()