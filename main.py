#Libraries Setup
from pywinauto import Application
import win32process, win32gui, win32con, time, win32api, wmi

from tkinter import *
from tkinter import ttk

SEPARATOR = "------------------------------------"
INITIAL_CBOX_TEXT = "Select a window."

root = Tk()
root.title("wlock0.0")
root.geometry("300x500+150+200")

s = ttk.Style()
s.configure("Header.TLabel", anchor=CENTER)
s.configure("SetTarg.TLabel", anchor=LEFT)

c = wmi.WMI()

class UnderlinedLabel():
    def __init__(self, parent, text, style="Header.TLabel", bg="#000000"):
        self.label = ttk.Label(parent, text=text, style=style)
        self.underline = Label(parent, bg = bg)
    
    def place(self, relwidth, relheight, relx, rely, anchor):
        self.label.place(relwidth=relwidth ,relheight=relheight ,relx=relx ,rely=rely ,anchor=anchor)
        self.underline.place(relwidth=relwidth*0.9, relheight=0.002, relx=relx, rely=rely+relheight, anchor="s")

class Main:
    def __init__(self):
        #Set up the information-storing lists.
        self.wlist = []
        self.wdict = {}
        self.wtargt = []
        self.target = None

        self.getWinds()

        #Top Bar
        self.topLabel = UnderlinedLabel(root, "wlock 0.0")
       
        #Body
        self.body = ttk.Frame(root)

        #Selection
        self.cBoxVar = StringVar()
        self.cBox = ttk.Combobox(self.body, textvariable=self.cBoxVar, state="readonly", values=self.wlist)
        self.cBox.set(INITIAL_CBOX_TEXT)
        self.cBox.bind("<<ComboboxSelected>>", self.CBoxSelected)
        
        self.targLabel = ttk.Label(self.body, text="Target: None", style="SetTarg.TLabel")
        self.targLabel.bind('<Configure>', lambda e: self.targLabel.config(wraplength=self.targLabel.winfo_width()))
        self.setButton = ttk.Button(self.body, command=self.SetTarget, text="Set Target")

        #Whitelist
        self.whitelist = ttk.Frame(self.body)
        self.wlLabel = UnderlinedLabel(root, "Whitelist")

        self.sTBox = Text(self.whitelist)
        self.sTBox.config(state="disabled")
        self.addButton = ttk.Button(self.whitelist, command=self.AddItem, text="Add")
        self.remButton = ttk.Button(self.whitelist, command=self.RemItem, text="Remove")
        self.cleButton = ttk.Button(self.whitelist, command=self.ClearItems, text="Clear")

        self.UISetup()

    def CBoxSelected(self,n):
        print("Selected: " + self.cBox.get())

    def SetTarget(self):
        procname = self.cBox.get()
        if procname != INITIAL_CBOX_TEXT:
            self.target = procname
            self.targLabel.configure(text="Target: " + procname)
        
    def AddItem(self):
        procname = self.cBox.get()
        if procname != INITIAL_CBOX_TEXT and not procname in self.wtargt:
            self.sTBox.config(state="normal")
            self.sTBox.insert(END, procname+"\n")
            self.sTBox.config(state="disabled")

            self.wtargt.append(procname)

    def RemItem(self):
        procname = self.cBox.get()
        self.sTBox.config(state="normal")
        self.sTBox.delete("1.0", END)
        i = 0
        for v in self.wtargt.copy():
            if procname == v:
                del self.wtargt[i]
            else:
                self.sTBox.insert(END, v+"\n")
            i += 1
        self.sTBox.config(state="disabled")

    def ClearItems(self):
        self.sTBox.config(state="normal")
        self.wtargt.clear()
        self.sTBox.delete("1.0", END)
        self.sTBox.config(state="disabled")

    def UISetup(self):
        #Place UI Objects
        self.topLabel.place(relwidth=1, relheight=0.08, relx=0.5, rely=0, anchor="n")

        self.body.place(relwidth=0.9, relheight=0.8, relx=0.5, rely=0.11, anchor="n")

        self.cBox.place(relwidth=0.95, relheight=0.07, relx=0.5, rely=0, anchor="n")
        self.targLabel.place(relwidth=0.95, relheight=0.08, relx=0.5, rely=0.09, anchor="n")
        self.setButton.place(relwidth=0.95, relheight=0.08, relx=0.5, rely=0.19, anchor="n")

        self.whitelist.place(relwidth=0.95, relheight=0.3, relx=0.5, rely=0.42, anchor="n")
        self.wlLabel.place(relwidth=1, relheight=0.06, relx=0.5, rely=0.35, anchor="n") 
    
        self.sTBox.place(relwidth=0.65, relheight=1, relx=0, rely=0, anchor="nw")

        self.addButton.place(relwidth=0.32, relheight=0.32, relx=0.975, rely=0, anchor="ne")
        self.remButton.place(relwidth=0.32, relheight=0.32, relx=0.975, rely=1/3, anchor="ne")
        self.cleButton.place(relwidth=0.32, relheight=0.32, relx=0.975, rely=2/3, anchor="ne")

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
                    self.wlist.append(name)
                    self.wdict[name] = {"PID": pid} #If so, append it into the list of all active windows.


    def getWinds(self):
        #Clear the list of windows & lock targets.
        self.wlist.clear() 
        self.wdict.clear()

        #Use EnumWindows to loop through every process.
        win32gui.EnumWindows(self.EnumWindows,None)

        return self.wlist
        

m = Main()

root.mainloop()