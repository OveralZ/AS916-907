#Libraries Setup
from pywinauto import Application
import win32process, win32gui, win32con, time, win32api, wmi, keyboard, json, copy, sys, os, threading

from tkinter import *
from tkinter import ttk, messagebox, filedialog

#Constants
SEPARATOR = "------------------------------------"
INITIAL_CBOX_TEXT = "Select a window."
TITLE = "wlock0.1"
SIZE = "300x500+150+200"

def getpath(rp): #Used to get the real path of the file when in .EXE form.
    try:
        #PyInstaller creates a temp folder and stores path in _MEIPASS, so we use it instead.
        bp = sys._MEIPASS
    except Exception:
        bp = os.path.abspath(".") #For when we are in VSCode.

    return os.path.join(bp, rp)

ICON = getpath("./assets/icon.ico")

#Setup the root GUI.
root = Tk()
root.title(TITLE)
root.geometry(SIZE)
root.iconbitmap(ICON)

#Styles for our GUI objects, mainly for text centering.
s = ttk.Style()
s.configure("Header.TLabel", anchor=CENTER)
s.configure("SetTarg.TLabel", anchor=LEFT)

#Initialize WMI for later use.
c = wmi.WMI()

class Hotkey(): #Class for a hotkey and its assigned function.
    def __init__(self, key, function):
        #Set up variables.
        self.key = key
        self.function = function
        self.disconnected = False
        
        keyboard.on_press_key(key, self.func) #Create the callback for when the key is pressed.
    
    def func(self,n):
        if self.disconnected == False:
            self.function() #When the key is pressed and if the key is not disconnected, run the originally assigned function.
        else: del self #If not, delete this class object.
    
    def disconnect(self):
        self.disconnected = True #Disconnect this object.
        del self

class UnderlinedLabel(): #Class for convenience, acts as one GUI object but contains 2.
    def __init__(self, parent, text, style="Header.TLabel", bg="#000000"):
        #Initialize, create the 2 GUI objects.
        self.label = ttk.Label(parent, text=text, style=style, anchor=CENTER)
        self.underline = Label(parent, bg = bg)
    
    def place(self, relwidth, relheight, relx, rely, anchor):
        #Replace the .place() function with our own.
        self.label.place(relwidth=relwidth, relheight=relheight, relx=relx, rely=rely, anchor=anchor)
        self.underline.place(relwidth=relwidth*0.9, relheight=0.002, relx=relx, rely=rely+relheight, anchor="s")

class SettingsMenu():
    def __init__(self, menu):
        #Initialize the new GUI.
        self.menu = menu
        self.root = Tk()
        self.root.iconbitmap(ICON)
        self.root.title("Settings")
        self.root.geometry("300x300")
        self.settings = copy.deepcopy(menu.settings) #Create a clone of the main settings, since we will have the ability to cancel changes.

        #Callback to window closure, since we have to tell the main program that the settings have been closed.
        self.root.protocol("WM_DELETE_WINDOW", self.closed)

        #Top title
        self.top = UnderlinedLabel(self.root, "Settings")
        self.top.place(relwidth=1, relheight=0.1, relx=0.5, rely=0, anchor="n")

        #Settings
        self.whitelistToggle = Checkbutton(self.root, text="Add window to whitelist on select", onvalue=True, offvalue=False, command=lambda: self.toggleSetting("AddToWhitelistOnSelect"))
        if self.settings["AddToWhitelistOnSelect"] == True: self.whitelistToggle.select() #Select the checkbox if the value is true.
        self.whitelistToggle.place(relwidth=0.9, relheight=0.1, relx=0.5, rely=0.11, anchor="n")

        #Hotkeys
        i2 = 0
        self.hk = UnderlinedLabel(self.root, "Hotkeys")
        self.hk.place(relwidth=1, relheight=0.1, relx=0.5, rely=0.3, anchor="n")
        self.hotkeyFrame = Frame(self.root)
        self.hotkeyFrame.place(relwidth=0.9, relheight=0.4, relx=0.5, rely=0.41, anchor="n")

        self.hotkeyButtons = {} #Setup another dictionary for the specific button objects created so we can refer to them later.

        for i in self.settings["Hotkeys"]: #Loop through the settings dictionary.
            v = self.settings["Hotkeys"][i]
            #Create the GUI objects.
            label = ttk.Label(self.hotkeyFrame, text=i, style="Header.TLabel", anchor=CENTER)
            label.place(relwidth=0.4, relheight=0.25, relx=0.5, rely=0 + i2/4, anchor="ne")

            btn = ttk.Button(self.hotkeyFrame, text=v.upper(), command=lambda i = i: self.adjustHotkey(i))
            btn.place(relwidth=0.4, relheight=0.25, relx=0.5, rely=0 + i2/4, anchor="nw")
            self.hotkeyButtons[i] = btn

            i2 += 1

        #Bottom UI
        self.bFrame = Frame(self.root)
        self.bFrame.place(relwidth=0.7, relheight=0.1, relx=0.99, rely=0.99, anchor="se")
        self.applyButton = ttk.Button(self.bFrame, text="Apply", command=lambda: menu.changeSettings(self.settings))
        self.applyButton.place(relwidth=0.45, relheight=1, relx=0.5, rely=0, anchor="ne")
        self.cancelButton = ttk.Button(self.bFrame, text="Cancel", command=self.closed)
        self.cancelButton.place(relwidth=0.45, relheight=1, relx=1, rely=0, anchor="ne")

    def closed(self): #Callback for when the settings GUI is closed.
        self.root.destroy()
        self.menu.settingsClosed()

    def toggleSetting(self,val): #Function for toggling a setting.
        self.settings[val] = not self.settings[val]

    def adjustHotkey(self, i): #Hotkey change function.
        nroot = Tk() #Create the new GUI.
        nroot.iconbitmap(ICON)
        nroot.geometry("300x80") 
        label = ttk.Label(nroot, text="Press any key..", anchor=CENTER) #Create the label.
        label.place(relheight=1,relwidth=1,relx=0,rely=0,anchor="nw")
        threading.Thread(target=lambda: self.hkeyWait(i, nroot)).start() #Call a new thread so the window does not count as frozen.
        nroot.mainloop() #Call mainloop so the window remains.

    def hkeyWait(self, i, nroot):
        k = keyboard.read_key() #Wait for an input from the user.
        nroot.destroy() #Destroy the GUI.
        self.hotkeyButtons[i].configure(text=k.upper())
        self.settings["Hotkeys"][i] = k #Change our settings.

class Main:
    def __init__(self):
        #Set up the information-storing lists.
        self.wlist = []
        self.wdict = {}
        self.wtargt = []
        self.wpids = []
        self.timer = 10
        self.target = None
        self.paused = False

        #Settings and setting loading.
        self.settings = {
            "AddToWhitelistOnSelect": False,
            "Hotkeys": {
                "Pause": "f5", "Abort": "f6"
            }
        }

        try:
            setJson = open("settings.json","r+") #Attempt to open the data in read & write form.
            self.settings = json.load(setJson)
        except FileNotFoundError:
            setJson = open("settings.json","w+") #If it doesn't exist, use w+ to create it.
            json.dump(self.settings,setJson)
        setJson.close()

        #Setup windows
        self.getWinds()

        #Hotkey Setup
        self.hotkeyFuncs = {"Pause": self.pause, "Abort": self.abort}
        self.hotkeyObjects = {
            "Pause": Hotkey(self.settings["Hotkeys"]["Pause"], self.hotkeyFuncs["Pause"]),
            "Abort": Hotkey(self.settings["Hotkeys"]["Abort"], self.hotkeyFuncs["Abort"])
        }

        #Top Bar
        self.topLabel = UnderlinedLabel(root, TITLE)
       
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
        self.addButton = ttk.Button(self.whitelist, command=lambda: self.AddItem(None), text="Add")
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

        #Menus
        self.menubar = Menu(root)
        self.filemenu = Menu(self.menubar, tearoff=0)

        #Add the different commands alongside their functions.
        self.filemenu.add_command(label="Refresh", command=self.getWinds)
        self.filemenu.add_command(label="Save", command=self.savePreset)
        self.filemenu.add_command(label="Load", command=self.loadPreset)
        self.filemenu.add_command(label="Settings", command=self.settingsDisplay)
        self.settingsActive = False
       
        self.menubar.add_cascade(label="File", menu=self.filemenu)

        self.UISetup()

    def pause(self):
        self.paused = not self.paused #Pause the force function.
    
    def abort(self):
        self.aborted = True #Abort the force function.

    def settingsClosed(self): #The callback for when the settings menu is closed.
        self.settingsActive = False

    def changeSettings(self,n): #Function for toggling a setting.
        for i in n:
            if i == "Hotkeys": #We need to change the actual hotkeys, so we loop for it.
                for x in n[i]:
                    if n[i][x] != self.settings[i][x]: #Check if the hotkey has been changed.
                        self.hotkeyObjects[x].disconnect() #If so, disconnect the previous hotkey object.
                        self.hotkeyObjects[x] = Hotkey(n[i][x], self.hotkeyFuncs[x]) #And create a new one.

            self.settings[i] = n[i] #Change the settings.

        self.settingsMenu.closed() #Close the settings GUI.
        setJson = open("settings.json", "w+")
        json.dump(self.settings, setJson) #Update the settings JSON.
        setJson.close()

    def settingsDisplay(self): #Create the settings menu.
        if self.settingsActive == False:
            self.settingsMenu = SettingsMenu(self)
        self.settingsActive = not self.settingsActive
        self.settingsMenu.root.mainloop()

    def onClose(self): #Callback for when the main GUI is closed, so that we can close the other opened windows.
        try:
            self.settingsMenu.root.destroy()
        except:
            pass
        root.destroy()

    def savePreset(self): #Preset saving.
        fp = filedialog.asksaveasfilename(filetypes={("Preset files", "*.wlpreset")}) #Ask the user for the name of the preset.
        fp = fp.replace(".wlpreset", "") #Make sure there isn't a duplicate extension name in it.
        if fp: #See if the user has actually decided to save it, since cancelling will return fp as None.
            save = open(fp + ".wlpreset","w+") #Open the file, if it isn't there, create it.
            json.dump({
                "Target": self.target, "Timer": float(self.timerEntry.get() or 10), "Whitelist": self.wtargt, "TimerSetting": self.timerOption.get()
            },save) #Dump the information and close the file.
            save.close()
    
    def loadPreset(self): #Preset loading.
        fp = filedialog.askopenfilename(filetypes={("Preset files", "*.wlpreset")}) #Ask the user for the preset to load.
        save = open(fp,"r") #Open the file.
        tab = json.load(save) #Load the JSON information.
        
        self.getWinds() #Refresh the windows so that nothing errors when we load the data.

        self.timerEntry.delete(0, END) #Change the timer accordingly.
        self.timerEntry.insert(0, str(tab["Timer"]))
        self.timerOption.set(tab["TimerSetting"])

        if self.wdict[tab["Target"]]: #Check if the target in the preset is active.
            n = tab["Target"] #If so, set the target.
            self.targLabel.configure(text="Target: " + n)
            self.target = n
        
        for i in tab["Whitelist"]: #Loop through the whitelist.
            if self.wdict[i]: #If the window is active on the device, add it to the whitelist.
                self.AddItem(i)

        save.close() #Close the file.
    
    def Force(self): #Force function, the core of the program.
        self.paused = False
        self.aborted = False
        try:
            pid = self.wdict[self.target]["PID"] #Check if our target actually exists.
        except: 
            messagebox.showerror(title="Error", message="You have not set a target!") #If not, tell the user and return.
            return
        try: #Check if the timer exists.
            mode = self.timerOption.get()
            timer = float(self.timerEntry.get()) #If so, set the amount of time to wait for accordingly.
            if mode == "Minutes": timer *= 60 
            elif mode == "Hours": timer *= 360
        except:
            messagebox.showerror(title="Error", message="The timer entry is not a valid number!") #If not, tell the user and return.
            return

        st = time.time() #Get the current time for comparison.
        try: 
            app = Application().connect(process=pid) #See if our target is active on the user's device.
        except: 
            messagebox.showerror(title="Alert", message="Application not found, try refreshing the available windows under 'File'.") #If not, tell the user and return.
            return
        app.top_window().set_focus() #Set the target to be on top.
        root.title("Focusing..") #Change the GUI's name.

        while True: #Loop.
            if self.aborted == True: break
            if self.paused == True:
                time.sleep(0.1) #Wait
                st += 0.1 #Also increase the original time so the timer continues for the correct amount once resumed.
            else:
                try:
                    hwnd = win32gui.GetForegroundWindow() #Get the current active window.
                    hn = win32gui.GetWindowText(hwnd) #Get the name of it.
                    
                    _, npid = win32process.GetWindowThreadProcessId(hwnd)
                    fullname = self.getFullName(npid)
                    name = self.getName(fullname) #Get the proper name of the current active window.

                    if hn and (not name == "Windows Explorer"): #See if it exists & if it's not windows explorer.
                        if npid != pid: #If it doesn't match the target's pid, check if it's in the whitelist.
                            if not npid in self.wpids:
                                app.top_window().set_focus() #If it isn't in the whitelist, set the main target to be active.
                    time.sleep(0.1) #Wait.
                except Exception as e: #In case anything errors.
                    print(e)
                    break #If something errors, break the loop early.
                if time.time() - st > timer: #When the specified amount of time has passed, break the loop.
                    break
        #Display a message after the timer is done.
        if self.aborted == True:
            messagebox.showinfo(title="Alert", message="Timer aborted.")
        else: messagebox.showinfo(title="Alert", message="Focus timer complete.")
        root.title(TITLE)

    def CBoxSelected(self,n): #Function for when the main combobox is selected.
        if self.settings["AddToWhitelistOnSelect"] == True: self.AddItem(self.cBox.get())

    def SetTarget(self): #Function for setting the target.
        procname = self.cBox.get()
        if procname != INITIAL_CBOX_TEXT: #See if there is actual content in the combobox. If so, assign values.
            self.target = procname
            self.targLabel.configure(text="Target: " + procname)
        
    def AddItem(self, opt): #Function for adding an item to the whitelist.
        procname = opt or self.cBox.get()
        if procname != INITIAL_CBOX_TEXT and not procname in self.wtargt: #Check if the item is not the default value and is not in the whitelist.
            self.sTBox.config(state="normal")
            self.sTBox.insert(END, procname+"\n") #Add the item to the textbox.
            self.sTBox.config(state="disabled")

            self.wpids.append(self.wdict[procname]["PID"]) #Add the item to the whitelist lists.
            self.wtargt.append(procname)

    def RemItem(self): #Function for removing items from the whitelist.
        procname = self.cBox.get()
        self.sTBox.config(state="normal")
        self.sTBox.delete("1.0", END) #Clear the whitelist textbox.
        i = 0
        for v in self.wtargt.copy(): #Loop through a copy of the whitelist to avoid errors.
            if procname == v: #If it is the removed item, delete it from the whitelist.
                del self.wtargt[i]
            else: #If not, add it to the textbox.
                self.sTBox.insert(END, v+"\n")
            i += 1

        i2 = 0
        for i in self.wpids: #Update the whitelisted pids.
            if i == self.wdict[procname]["PID"]:
                del self.wpids[i2]
            i2 += 1

        self.sTBox.config(state="disabled")

    def ClearItems(self): #Function for clearing the whitelist.
        self.sTBox.config(state="normal")
        self.wtargt.clear() #Clear both tables.
        self.wpids.clear()
        self.sTBox.delete("1.0", END) #Remove all text in whitelist textbox.
        self.sTBox.config(state="disabled")

    def UISetup(self): #UI setup.
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

    def getFullName(self,pid): #Function for getting the full path of a process.
        try:
            for p in c.query('SELECT ExecutablePath FROM Win32_Process WHERE ProcessId = %s' % str(pid)): #Query for a PID match.
                return p.ExecutablePath #Return the full path.
        except: #If it errors, just return nothing.
            return None

    def getName(self,fn): #Function for getting the file description of a process, which is its display name.
        try:
            language, codepage = win32api.GetFileVersionInfo(fn, '\\VarFileInfo\\Translation')[0]
            stringFileInfo = u'\\StringFileInfo\\%04X%04X\\%s' % (language, codepage, "FileDescription")
            return win32api.GetFileVersionInfo(fn, stringFileInfo)
        except: #If it errors, just return nothing.
            return None
    
    def EnumWindows(self,hwnd,ctx):
        if win32gui.IsWindowVisible(hwnd): #Check if window is visible on the screen.
            wtext = win32gui.GetWindowText(hwnd)
            if len(wtext) > 0:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)

                fullname = self.getFullName(pid)
                name = self.getName(fullname) #Using the full name, get the display name of the window.

                if name and len(name) > 0: #Check if the name of the window is valid.
                    self.wlist.append(name)
                    self.wdict[name] = {"PID": pid} #If so, append it into the list of all active windows.


    def getWinds(self):
        #Clear the list of windows & lock targets.
        self.wlist.clear() 
        self.wdict.clear()
        self.wpids.clear()

        loadingL = ttk.Label(root, anchor=CENTER, text="Loading.. Please wait...") #Create the loading label.
        loadingL.place(relheight=1,relwidth=1)
        root.update() #Update the root so the label appears.

        #Use EnumWindows to loop through every process.
        win32gui.EnumWindows(self.EnumWindows,None)
        try: self.cBox.configure(values=self.wlist) 
        except: pass

        loadingL.destroy() #Destroy it after loading is complete.

        return self.wlist
        

m = Main() #Create the root window of the GUI.

root.configure(menu=m.menubar) #Assign the menu of the GUI to the root.
root.protocol("WM_DELETE_WINDOW", m.onClose) #Assign a callback for when the window is closed.
root.mainloop() #Call mainloop.
