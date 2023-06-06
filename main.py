#Libraries Setup
from pywinauto import Application
import win32process, win32gui, win32con, time

from tkinter import *
from tkinter import ttk

SEPARATOR = "------------------------------------"

root = Tk()
root.title("wlock0.0")
root.geometry("300x500+150+200")

class Main:
    def __init__(self):
        #Set up the information-storing lists.
        self.wlist = []
        self.wtargt = ["Discord","Telegram"]
        self.wtargs = []

        print("test")

    def EnumWindows(self,hwnd,ctx):
        if win32gui.IsWindowVisible(hwnd): #Check if window is visible on the screen.
            self.wlist.append(win32gui.GetWindowText(hwnd)) #If so, append it into the list of all active windows.
            #Loop through the targets specified for window locking.
            for i in self.wtargt:
                if i.lower() in win32gui.GetWindowText(hwnd).lower(): #Does the given name match the current window?
                    #It does, add its info to the targets list.
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    self.wtargs.append({"KEYWORD":i, "NAME": win32gui.GetWindowText(hwnd), "PID": pid})
                    break #Break the loop as we have gotten what we want already.

    def getWinds(self):
        #Clear the list of windows & lock targets.
        self.wlist.clear() 
        self.wtargs.clear()

        #Use EnumWindows to loop through every process.
        win32gui.EnumWindows(self.EnumWindows,None)
        

m = Main()
m.getWinds()
print(m.wlist)
print(SEPARATOR)
print(m.wtargs)

root.mainloop()