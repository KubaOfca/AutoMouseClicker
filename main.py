import pyautogui
import pyautogui as py
import tkinter as tk
import time
from tkinter import ttk, messagebox
import keyboard
import os
import ctypes
import tkinter.font as font
from tkinter import filedialog as fd
from string import *

# TODO: Load and save to file a sequence of doing click


class ActuallScreenManager(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.actuallFrame = None
        self.mousePosKey = 'f3'
        self.swichFrame(MainMenu)

    def swichFrame(self, _class):
        self.geometry("")
        newFrame = _class(self)
        if self.actuallFrame is not None:
            self.actuallFrame.destroy()
        self.actuallFrame = newFrame
        self.actuallFrame.pack()
        self.update()
        self.setWindowSize(self.winfo_width(), self.winfo_height())
        self.deiconify()

    def setWindowSize(self, width, height):
        self.state("normal")
        x_screen = int(self.winfo_screenwidth() / 2) - int(width/2)
        y_screen = int(self.winfo_screenheight() / 2) - int(height/2)
        self.geometry("{}x{}+{}+{}".format(width, height, x_screen, y_screen))

    def exit(self):
        self.destroy()


class MainMenu(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        # button
        self.createSeq_button = tk.Button(
            self, text="Create mouse click sequence",
            command=lambda: self.master.swichFrame(Creator))
        self.exit_button = tk.Button(self, text="Exit",
                                     command=lambda: self.master.exit())
        self.option_button = tk.Button(
            self, text="Configure",
            command=lambda: self.master.swichFrame(Configuration))
        # pack
        self.createSeq_button.pack(padx=20, pady=20, side='top', fill='x')
        self.option_button.pack(padx=20, pady=20, side='top', fill='x')
        self.exit_button.pack(padx=20, pady=20, side='top', fill='x')

        self.master.iconify()


class Configuration(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.master.setWindowSize(350, 200)
        # label
        self.mousePosKey_label = tk.Label(
            self,
            text="Key to save mouse position is: {}".format(self.master.mousePosKey))
        # button
        self.changeMousePosKey_button = tk.Button(self, text="Change key",
                                                  command=self.changeKey)
        self.back_button = tk.Button(
            self, text="Back to Main Menu",
            command=lambda: self.master.swichFrame(MainMenu))
        # pack
        self.mousePosKey_label.grid(row=0, column=0, padx=10, pady=10)
        self.changeMousePosKey_button.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        self.back_button.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

    def changeKey(self):
        self.label = tk.Label(self, text="Press key...")
        self.label.grid(row=1, column=1, padx=10, pady=10)
        self.update()
        while True:
            if keyboard.read_key():
                self.master.mousePosKey = keyboard.read_key()
                self.mousePosKey_label.config(text="Key to save mouse position is: {}".format(self.master.mousePosKey))
                self.label.destroy()
                break


class Creator(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        # global
        self.id = 0
        self.mousePos = {}
        self.isDoubleClick_var = tk.IntVar()
        self.clockParameters = [0, 0, 0]
        # frame
        self.entry_frame = tk.LabelFrame(self, text="Entry event name and entry Trigger time")
        self.tree_frame = tk.Frame(self)
        self.option_frame = tk.Frame(self)
        # label
        self.clock = tk.Label(self.option_frame, font=('ds-digital', 50), background='black', foreground='cyan',
                              text="00:00:00")
        self.trigerTime_label = tk.Label(self.entry_frame, text='Trigger time in [sek]', anchor='w')
        # fontsize
        myFont = font.Font(size=30)
        # button
        self.addMouseClick_button = tk.Button(self, text="AddMouseClick",
                                              command=self.caputreMousePosition, bg='green', fg='white')
        self.addMouseClick_button['font'] = myFont
        self.removeClick_button = tk.Button(self.option_frame, text="Remove Click",
                                            command=self.removeRecord)
        self.loadSeq_button = tk.Button(self.option_frame, text="Load Sequencce from File",
                                        command=self.loadSequenceFromFile)
        self.saveSeq_button = tk.Button(self.option_frame, text="Save Sequence to File",
                                        command=self.saveSeqToFileConfigure)
        self.setTimeToStart_button = tk.Button(self.option_frame, text="Set Time to Start",
                                               command=self.setClock)
        self.run_button = tk.Button(self.option_frame, text="RUN", bg='red', fg='white',
                                    command=self.run)
        # entry
        self.nameEvent_entry = tk.Entry(self.entry_frame, width=40)
        self.triggerTime_entry = tk.Entry(self.entry_frame, width=20)
        # checkbox
        self.isDoubleClick_checkbox = tk.Checkbutton(self.entry_frame, text="Double click",
                                                     variable=self.isDoubleClick_var)
        # treeview Scroll
        self.yscroll = ttk.Scrollbar(self.tree_frame)
        # treeview
        self.tree = ttk.Treeview(self.tree_frame, yscrollcommand=self.yscroll.set)
        self.tree['column'] = ('Id', 'EventName', 'TriggerTime', 'Status')
        self.tree.column('#0', width=0, stretch=False)
        self.tree.column('Id', width=20, anchor='center')
        self.tree.column('EventName', width=50, anchor='center')
        self.tree.column('TriggerTime', width=30, anchor='center')
        self.tree.column('Status', width=30, anchor='center')
        self.tree.heading('Id', text='Id', anchor='center')
        self.tree.heading('EventName', text='Event name', anchor='center')
        self.tree.heading('TriggerTime', text='Trigger time', anchor='center')
        self.tree.heading('Status', text='Status', anchor='center')
        # config scroll
        self.yscroll.config(command=self.tree.yview)
        # grid
        self.addMouseClick_button.grid(row=0, column=0, padx=20, pady=20, sticky='ewsn')
        self.entry_frame.grid(row=0, column=1, padx=20, pady=20)
        self.tree_frame.grid(row=1, column=0)
        self.option_frame.grid(row=1, column=1, sticky='n', pady=10)
        # pack
        self.tree.pack(padx=20, pady=20, ipadx=170, ipady=50)
        self.nameEvent_entry.grid(row=0, column=0, columnspan=2, padx=20, pady=20, sticky='w')
        self.triggerTime_entry.grid(row=1, column=0, padx=20, pady=20, sticky='e')
        self.trigerTime_label.grid(row=1, column=1, sticky='w', padx=5, pady=20)
        self.isDoubleClick_checkbox.grid(row=2, column=0, padx=20, pady=10, sticky='w')
        self.yscroll.pack(side="right", fill='y')
        self.clock.pack(pady=10, side='top', fill='x')
        self.setTimeToStart_button.pack(pady=10, side='top', fill='x')
        self.removeClick_button.pack(pady=10, side='top', fill='x')
        self.loadSeq_button.pack(pady=10, side='top', fill='x')
        self.saveSeq_button.pack(pady=10, side='top', fill='x')
        self.run_button.pack(pady=10, side='top', fill='x')

    # functions
    def changeCursour(self, cursourName):
        path = r"HKEY_CURRENT_USER\Control Panel\Cursors"
        cur_loc = r"C:\Windows\Cursors\{}.cur".format(cursourName)
        os.system(f"""REG ADD "{path}" /v Arrow /t REG_EXPAND_SZ /d "{cur_loc}" /f""")
        ctypes.windll.user32.SystemParametersInfoA(0x57)

    def saveSeqToFileConfigure(self):
        self.newWindow = tk.Toplevel(self.master)
        # frame
        self.enterFolder_frame = tk.LabelFrame(self.newWindow, text="Enter a folder location")
        self.enterFileName_frame = tk.LabelFrame(self.newWindow, text="Enter a file name")
        # entry
        self.folder_entry = tk.Entry(self.enterFolder_frame, width=40, borderwidth=3)
        self.fileName_entry = tk.Entry(self.enterFileName_frame, width=30, borderwidth=3)
        # button
        self.browseFolder_button = tk.Button(self.enterFolder_frame, text="Browse",
                                             command=self.setDict)
        self.save_button = tk.Button(self.newWindow, text="Save", width=10,
                                     command=self.saveToFile)
        # grid
        self.enterFolder_frame.pack(padx=20, pady=20)
        self.enterFileName_frame.pack(padx=20, pady=20)
        self.folder_entry.grid(row=0, column=0, padx=20, pady=20)
        self.fileName_entry.grid(row=0, column=0,  padx=20, pady=20)
        self.browseFolder_button.grid(row=0, column=1, padx=20, pady=20)
        self.save_button.pack(padx=20, pady=20)

        self.newWindow.update()
        ActuallScreenManager.setWindowSize(self.newWindow, self.newWindow.winfo_width(), self.newWindow.winfo_height())

    def saveToFile(self):
        file = open("{}/{}.txt".format(self.folder_entry.get(), self.fileName_entry.get()), 'w')
        for element in self.tree.get_children():
            line = " ".join(map(str, self.tree.item(element)['values']))
            x_pos = self.mousePos[self.tree.item(element)['values'][0]].x
            y_pos = self.mousePos[self.tree.item(element)['values'][0]].y
            line = line + " {} {}".format(x_pos, y_pos)
            file.write(line + '\n')
        self.newWindow.destroy()

    def setDict(self):
        self.folder_entry.delete(0, 'end')
        self.folder_entry.insert(0, fd.askdirectory())
        self.newWindow.deiconify()

    def removeRecord(self):
        recordToRemove = self.tree.selection()
        if recordToRemove == ():
            return
        for x in recordToRemove:
            self.mousePos.pop(int(self.tree.item(x)['values'][0]), None)
            self.tree.delete(x)

    def deleteAllRecords(self):
        for record in self.tree.get_children():
            self.mousePos.pop(int(self.tree.item(record)['values'][0]), None)
            self.tree.delete(record)

    def loadSequenceFromFile(self):
        self.deleteAllRecords()
        file = open("{}".format(fd.askopenfilename(), 'r'))
        for line in file:
            line = line.replace('\n', '').split(' ')
            self.tree.insert(parent='', index='end', values=(line[0], line[1],
                                                             int(line[2]), line[3]))
            self.mousePos[int(line[0])] = pyautogui.Point(line[4], line[5])

    def caputreMousePosition(self):
        if not self.nameEvent_entry.get():
            messagebox.showerror(title="ERROR!", message="Entry the step name")
            return
        triggerTime = self.triggerTime_entry.get()
        if not triggerTime.isnumeric():
            triggerTime = 0
        self.master.withdraw()
        time.sleep(1)
        self.changeCursour('aero_pen')
        while True:
            if keyboard.read_key() == self.master.mousePosKey:
                self.id += 1
                self.mousePos[self.id] = py.position()
                status = "doubleClick" if self.isDoubleClick_var.get() == 1 else "singleClick"
                self.tree.insert(parent='', index='end', values=(self.id, self.nameEvent_entry.get(),
                                                                 int(triggerTime), status))
                break
        self.changeCursour('aero_arrow')
        time.sleep(1)
        self.nameEvent_entry.delete(0, 'end')
        self.triggerTime_entry.delete(0, 'end')
        self.isDoubleClick_var.set(0)
        self.master.deiconify()

    def setClock(self):
        self.newWindow = tk.Toplevel(self.master)
        ActuallScreenManager.setWindowSize(self.newWindow, 450, 200)
        # frame
        self.frame = tk.Frame(self.newWindow)
        # label
        self.title_label = tk.Label(self.frame, text="SET TIME!")
        self.hours_label = tk.Label(self.frame, text='H')
        self.minutes_label = tk.Label(self.frame, text='M')
        self.seconds_label = tk.Label(self.frame, text='S')
        # entry
        self.hours_entry = tk.Entry(self.frame, width=10)
        self.minutes_entry = tk.Entry(self.frame, width=10)
        self.seconds_entry = tk.Entry(self.frame, width=10)
        # button
        self.button = tk.Button(self.frame, text='Ok', command=self.getClockSettings)
        # grid
        self.frame.pack(side='top', padx=20, pady=20)
        self.title_label.grid(row=0, column=0, columnspan=6, padx=10, pady=10)
        self.hours_entry.grid(row=1, column=0, padx=10, pady=10)
        self.hours_label.grid(row=1, column=1, padx=10, pady=10)
        self.minutes_entry.grid(row=1, column=2, padx=10, pady=10)
        self.minutes_label.grid(row=1, column=3, padx=10, pady=10)
        self.seconds_entry.grid(row=1, column=4, padx=10, pady=10)
        self.seconds_label.grid(row=1, column=5, padx=10, pady=10)
        self.button.grid(row=2, column=2, columnspan=2, padx=10, pady=10, sticky='ew')

    def getClockSettings(self):
        self.clockParameters = [self.hours_entry.get(), self.minutes_entry.get(), self.seconds_entry.get()]
        for x in range(0, len(self.clockParameters)):
            if not self.clockParameters[x].isnumeric():
                self.clockParameters[x] = 0

        self.clock.config(text="{:02d}:{:02d}:{:02d}".format(int(self.clockParameters[0]), int(self.clockParameters[1]),
                                                             int(self.clockParameters[2])))
        self.newWindow.destroy()

    def runClock(self):
        try:
            seconds = int(self.clockParameters[0]) * 3600 + int(self.clockParameters[1]) * 60 + int(self.clockParameters[2])
        except:
            seconds = 0
        while seconds > 0:
            time.sleep(1)
            seconds -= 1
            seconds_temp = seconds
            hours = int(seconds_temp / 3600)
            seconds_temp -= int(hours * 3600)
            minutes = int(seconds_temp / 60)
            seconds_temp -= int(minutes * 60)
            self.clock.config(text="{:02d}:{:02d}:{:02d}".format(hours, minutes, seconds_temp))
            self.update()
        self.clock.config(text="{:02d}:{:02d}:{:02d}".format(int(self.clockParameters[0]), int(self.clockParameters[1]),
                                                             int(self.clockParameters[2])))

    def run(self):
        self.runClock()
        self.master.withdraw()
        time.sleep(0.5)
        for x in self.tree.get_children():
            key = self.tree.item(x)['values'][0] # to get index of row in treeview
            time.sleep(self.tree.item(x)['values'][2])
            isDoubleClick = True if self.tree.item(x)['values'][3] == "doubleClick" else False
            if isDoubleClick:
                py.doubleClick(int(self.mousePos[key].x), int(self.mousePos[key].y))
            else:
                py.click(int(self.mousePos[key].x), int(self.mousePos[key].y))
            time.sleep(1)
        self.master.deiconify()

app = ActuallScreenManager()
app.mainloop()
