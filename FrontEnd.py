import glob
from tkinter import *
import tkinter as tk
import ctypes
import pandas as pd
from PIL import ImageTk, Image
from pandastable import Table, TableModel
import random
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from Stampa0inj import *
import json

myappid = 'mycompany.myproduct.subproduct.version'  # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


class AutoScrollbar(ttk.Scrollbar):
    ''' A scrollbar that hides itself if it's not needed.
        Works only if you use the grid geometry manager '''

    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            self.grid_remove()
        else:
            self.grid()
        ttk.Scrollbar.set(self, lo, hi)

    def pack(self, **kw):
        raise tk.TclError('Cannot use pack with this widget')

    def place(self, **kw):
        raise tk.TclError('Cannot use place with this widget')


class MyWindow:
    def __init__(self, win):
        self.lbl_a = None
        self.lbl_b = None
        self.lbl_c = None
        self.table_ser = None
        self.bservice = None
        self.optionmenu_c = None
        self.optionmenu_b = None
        self.optionmenu_a = None
        self.modello = None
        self.serie = None
        self.tipologia = None
        self.table = None
        self.opt_info_inst1 = None
        self.opt_info_inst2 = None
        self.dict_add_inst = None
        self.dict_add_info = None
        self.info_inst1 = None
        self.info_inst2 = None
        self.binfo = None
        self.info1 = None
        self.info2 = None
        self.opt_info1 = None
        self.opt_info2 = None
        self.newWindow = win

        self.dict = {'1PH': ph1,
                     '3PH': ph3,
                     'BATTERY': battery,
                     'ACCESSORI': accessori,
                     'WALLBOX': wallbox}

        self.dict_model = {
            "ZCS-1PH-1100_3300TL-V3": dict_all["1PH"]["ZCS-1PH-1100_3300TL-V3"]['MODELLO'],
            "ZCS-1PH-3000_6000TLM-V3": dict_all["1PH"]["ZCS-1PH-3000_6000TLM-V3"]['MODELLO'],
            "ZCS-1PH-1100_3300TL-V1": dict_all["1PH"]["ZCS-1PH-1100_3300TL-V1"]['MODELLO'],
            "ZCS-1PH-3000_6000TLM-V1": dict_all["1PH"]["ZCS-1PH-3000_6000TLM-V1"]['MODELLO'],
            "ZCS-1PH-3000_6000TLM-V2": dict_all["1PH"]["ZCS-1PH-3000_6000TLM-V2"]['MODELLO'],
            "ZCS-3000SP-V2": dict_all["1PH"]["ZCS-3000SP-V2"]['MODELLO'],
            "ZCS-1PH-HYD-3000_6000-ZSS": dict_all["1PH"]["ZCS-1PH-HYD-3000_6000-ZSS"]['MODELLO'],
            "ZCS-1PH-HYD-3000_6000-ZSS-HP": dict_all["1PH"]["ZCS-1PH-HYD-3000_6000-ZSS-HP"]['MODELLO'],
            "ZCS-3PH-HYD-5000_8000-ZSS": dict_all["3PH"]["ZCS-3PH-HYD-5000_8000-ZSS"]['MODELLO'],
            "ZCS-3PH-HYD-10000_20000-ZSS": dict_all["3PH"]["ZCS-3PH-HYD-10000_20000-ZSS"]['MODELLO'],
            "ZCS-3PH-50000_60000TL-V1": dict_all["3PH"]["ZCS-3PH-50000_60000TL-V1"]['MODELLO'],
            "ZCS-3PH-80_110KTL-LV": dict_all["3PH"]["ZCS-3PH-80_110KTL-LV"]['MODELLO'],
            "ZCS-3PH-100_136KTL-HV": dict_all["3PH"]["ZCS-3PH-100_136KTL-HV"]['MODELLO'],
            "ZCS-3PH-3.3_12KTL-V1": dict_all["3PH"]["ZCS-3PH-3.3_12KTL-V1"]['MODELLO'],
            "ZCS-3PH-3.3_12KTL-V3": dict_all["3PH"]["ZCS-3PH-3.3_12KTL-V3"]['MODELLO'],
            "ZCS-3PH-10_15KTL-V2": dict_all["3PH"]["ZCS-3PH-10_15KTL-V2"]['MODELLO'],
            "ZCS-3PH-25_50KTL-V3": dict_all["3PH"]["ZCS-3PH-25_50KTL-V3"]['MODELLO'],
            "ZCS-3PH-20000_33000TL-V2": dict_all["3PH"]["ZCS-3PH-20000_33000TL-V2"]['MODELLO'],
            "ZCS-3PH-15000_24000TL-V3": dict_all["3PH"]["ZCS-3PH-15000_24000TL-V3"]['MODELLO'],
            "ZST-BAT-2,4KWH-PL": dict_all["BATTERY"]["ZST-BAT-2,4KWH-PL"]['MODELLO'],
            "ZST-BAT-5KWH-PL": dict_all["BATTERY"]["ZST-BAT-5KWH-PL"]['MODELLO'],
            "ZZT-BAT-5KWH-W": dict_all["BATTERY"]["ZZT-BAT-5KWH-W"]['MODELLO'],
            "ZZT-BAT-5KWH-ZSX": dict_all["BATTERY"]["ZZT-BAT-5KWH-ZSX"]['MODELLO'],
            "ZST-BAT-2,4KWH-H": dict_all["BATTERY"]["ZST-BAT-2,4KWH-H"]['MODELLO'],
            "ZZT-BAT-6KWH-W": dict_all["BATTERY"]["ZZT-BAT-6KWH-W"]['MODELLO'],
            "Current sensors": dict_all["ACCESSORI"]["Current sensors"]['MODELLO'],
            "Meters": dict_all["ACCESSORI"]["Meters"]['MODELLO'],
            "Single unit loggers": dict_all["ACCESSORI"]["Single unit loggers"]['MODELLO'],
            "External loggers": dict_all["ACCESSORI"]["External loggers"]['MODELLO'],
            "Connext": dict_all["ACCESSORI"]["Connext"]['MODELLO']
        }

        self.b1 = Button(self.newWindow, text='Load Config file', command=self.load_config, width=width,
                         bg='lightgreen')
        self.b1.place(x=50, y=30)

    def load_config(self):
        #  reading config file
        df_manual = pd.read_excel("Config.xlsx", sheet_name='CONFIG')
        self.Window_mod = Toplevel(self.newWindow)
        if df_manual['FUNZIONE'][0] == '0-inj':
            self.Window_mod.title(" 0 injection")
        elif df_manual['FUNZIONE'][0] == 'DRMn':
            self.Window_mod.title(" DRMn")
        self.Window_mod.iconbitmap(r'img\azzurro.ico')
        self.Window_mod.geometry("620x250")

        self.modes = tk.StringVar(self.Window_mod)

        
        self.lbl_c = Label(self.Window_mod, text='Modello', bg=bg)
        self.lbl_c.place(x=400, y=10)
        self.optionmenu_c.place(x=400, y=30)

        self.binfo = Button(self.Window_mod, text='INFO', command=self.print_info, width=width, bg='lightgreen')
        self.binfo.place(x=400, y=140)


    def print_info(self):
        value = [self.tipologia.get(), self.serie.get(), self.modello.get(),
                 self.info_inst1.get(), self.info_inst2.get(),
                 # self.info1.get(), self.info2.get()
                 ]
        if value[3] == 'Zero Injection':
            if dict_all[value[0]][value[1]]['0-INJ']['ABLE'] == True:
                if value[4] in dict_all[value[0]][value[1]]['0-INJ']['TIPO 0-INJ']:
                    documento_0inj(dict_all, value[1], value[0], value[4])
                else:
                    Window_info2 = Toplevel(self.newWindow)
                    Window_info2.title("Error")
                    Window_info2.geometry("300x50")
                    Window_info2.iconbitmap('img/azzurro.ico')
                    self.lbl_info_err = Label(Window_info2, text='Questo metodo non è compreso in questo prodotto',
                                              bg=bg)
                    self.lbl_info_err.place(x=10, y=10)
            else:
                Window_info2 = Toplevel(self.newWindow)
                Window_info2.title("Error")
                Window_info2.geometry("200x50")
                Window_info2.iconbitmap('img/azzurro.ico')
                self.lbl_info_err = Label(Window_info2, text='Questo modello non ha lo 0-inj', bg=bg)
                self.lbl_info_err.place(x=10, y=10)


bg = "#f5f6f7"
width = 20

f = open('Database.json')
dict_all = json.load(f)

ph1 = list()
ph3 = list()
battery = list()
accessori = list()
wallbox = list()
for i in dict_all:
    if i == '1PH':
        for j in dict_all[i]:
            ph1.append(j)
    elif i == '3PH':
        for j in dict_all[i]:
            ph3.append(j)
    elif i == 'BATTERY':
        for j in dict_all[i]:
            battery.append(j)
    elif i == 'ACCESSORI':
        for j in dict_all[i]:
            accessori.append(j)
    elif i == 'WALLBOX':
        for j in dict_all[i]:
            wallbox.append(j)

window = Tk()
MyWindow(window)
window.title('Manual Printer')
window.geometry("250x80")
window.iconbitmap(r'img\azzurro.ico')
window.mainloop()
