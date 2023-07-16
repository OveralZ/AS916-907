#Libraries Setup
from pywinauto import Application
import win32process, win32gui, win32con, time, win32api, wmi, keyboard

from tkinter import *
from tkinter import ttk, messagebox

SEPARATOR = "------------------------------------"
INITIAL_CBOX_TEXT = "Select a window."
TITLE = "wlock0.0"
SIZE = "300x500+150+200"

root = Tk()
root.title(TITLE)
root.geometry(SIZE)

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
        self.timer = 10
        self.target = None

        self.getWinds()

        keyboard.on_press_key("p", lambda e: print("p pressed"))

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

        #Timer
        self.timerLabel = UnderlinedLabel(root, "Timer")
        self.timerFrame = ttk.Frame(self.body)
        self.timerEntry = Entry(self.timerFrame)
        self.timerCVar = StringVar()
        self.timerOption = ttk.Combobox(self.timerFrame, textvariable=self.timerCVar, state="readonly", values=["Seconds", "Minutes", "Hours"])
        self.timerOption.set("Seconds")

        #Force
        self.forceButton = ttk.Button(self.body, command=self.Force, text="Force")

        self.UISetup()
    
    def Force(self):
        try:
            pid = self.wdict[self.target]["PID"]
        except: 
            messagebox.showerror(title="Error", message="You have not set a target!")
            return
        try:
            mode = self.timerOption.get()
            timer = float(self.timerEntry.get())
            if mode == "Minutes": timer *= 60 
            elif mode == "Hours": timer *= 360
        except:
            messagebox.showerror(title="Error", message="The timer entry is not a valid number!")
            return

        st = time.time()
        try: 
            app = Application().connect(process=pid) 
        except: 
            messagebox.showerror(title="Alert", message="Application not found, try refreshing the available windows under 'File'.")
            return
        app.top_window().set_focus()
        root.title("Focusing..")

        while True:
            try:
                hwnd = win32gui.GetForegroundWindow()
                _, npid = win32process.GetWindowThreadProcessId(hwnd)
                if npid != pid:
                    app.top_window().set_focus()
                time.sleep(0.1)
                if time.time() - st > timer:
                    break
            except Exception as e:
                print(e)
                pass
        messagebox.showinfo(title="Alert", message="Focus timer complete.")
        root.title(TITLE)

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
        self.topLabel.place(relwidth=1, relheight=0.06, relx=0.5, rely=0, anchor="n")

        self.body.place(relwidth=0.9, relheight=0.9, relx=0.5, rely=0.08, anchor="n")

        self.cBox.place(relwidth=0.95, relheight=0.07, relx=0.5, rely=0, anchor="n")
        self.targLabel.place(relwidth=0.95, relheight=0.06, relx=0.5, rely=0.09, anchor="n")
        self.setButton.place(relwidth=0.95, relheight=0.08, relx=0.5, rely=0.17, anchor="n")

        self.whitelist.place(relwidth=0.95, relheight=0.3, relx=0.5, rely=0.37, anchor="n")
        self.wlLabel.place(relwidth=1, relheight=0.06, relx=0.5, rely=0.32, anchor="n") 
    
        self.sTBox.place(relwidth=0.65, relheight=1, relx=0, rely=0, anchor="nw")

        self.addButton.place(relwidth=0.32, relheight=0.32, relx=0.975, rely=0, anchor="ne")
        self.remButton.place(relwidth=0.32, relheight=0.32, relx=0.975, rely=1/3, anchor="ne")
        self.cleButton.place(relwidth=0.32, relheight=0.32, relx=0.975, rely=2/3, anchor="ne")

        self.timerLabel.place(relwidth=1, relheight=0.06, relx=0.5, rely=0.7, anchor="n")
        self.timerFrame.place(relwidth=1, relheight=0.06, relx=0.5, rely=0.78, anchor="n")
        self.timerEntry.place(relwidth=0.7, relheight=1, relx=0, rely=0, anchor="nw")
        self.timerOption.place(relwidth=0.3, relheight=1, relx=1, rely=0, anchor="ne")

        self.forceButton.place(relwidth=0.32, relheight=0.08, relx=0.5, rely=0.98, anchor="s")

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
        try: self.cBox.configure(values=self.wlist) 
        except: pass

        return self.wlist
        

m = Main()

menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Refresh", command=m.getWinds)
menubar.add_cascade(label="File", menu=filemenu)

root.configure(menu=menubar)

root.mainloop()