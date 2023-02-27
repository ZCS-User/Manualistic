import ctypes
import json
import tkinter as tk
from tkinter import *
from tkinter import ttk
import pandas as pd
from Stampa0inj import *

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
        self.df_manual = pd.read_excel("Config.xlsx", sheet_name='CONFIG')
        self.Window_mod = Toplevel(self.newWindow)
        if self.df_manual['FUNZIONE'][0] == '0-inj':
            self.Window_mod.title(" 0 injection")
        elif self.df_manual['FUNZIONE'][0] == 'DRMn':
            self.Window_mod.title(" DRMn")
        self.Window_mod.iconbitmap(r'img\azzurro.ico')
        self.Window_mod.geometry("420x100")

        modes = tk.StringVar()
        self.modes_combobox = ttk.Combobox(self.Window_mod, textvariable=modes)
        if len(self.df_manual['MODELLO INVERTER']) != 0:
            if len(self.df_manual['MODELLO INVERTER']) == 1:
                self.modes_combobox['values'] = ['TA', 'METER', 'COMBOX', 'CCMASTER']
            else:
                a = 0
                for i in self.df_manual['MODELLO INVERTER']:
                    if ('V1' in i or 'V2' in i) and (i != 'ZCS-3PH-50000_60000TL-V1'):
                        a = 1
                if a == 0:
                    if len(self.df_manual['MODELLO INVERTER']) > 4 or float(self.df_manual['POTENZA TOTALE IMPIANTO [kW]'][0]) > 40:
                        self.modes_combobox['values'] = ['COMBOX']
                    else:
                        self.modes_combobox['values'] = ['COMBOX', 'CCMASTER']

        self.modes_combobox.pack()
        self.lbl_combobox = Label(self.Window_mod, text='Con cosa vuoi fare la 0-immissione?', bg=bg)
        self.lbl_combobox.place(x=50, y=10)
        self.modes_combobox.place(x=50, y=40)

        self.binfo = Button(self.Window_mod, text='INFO', command=self.print_info, width=width, bg='lightgreen')
        self.binfo.place(x=200, y=35)


    def print_info(self):
        zero_inj = 0
        for i in self.df_manual['MODELLO INVERTER']:
            try:
                if dict_all['1PH'][i]['0-INJ']['ABLE'] != True:
                    zero_inj += 1
                # debug
                # else:
                #     print(i)

            except:
                try:
                    if dict_all['3PH'][i]['0-INJ']['ABLE'] != True:
                        zero_inj += 1
                    # debug
                    # else:
                    #     print(i)
                except:
                    zero_inj += 1

            if zero_inj == 0 and self.modes_combobox.get() != '':
                probe = self.modes_combobox.get()
                if probe == 'COMBOX':
                    probe = 'ENERCLICK'
                try:
                    if probe in dict_all['1PH'][i]['0-INJ']['TIPO 0-INJ']:
                        documento_0inj(dict_all, i, '1PH', probe)
                except:
                    if probe in dict_all['3PH'][i]['0-INJ']['TIPO 0-INJ']:
                        documento_0inj(dict_all, i, '3PH', probe)

        if zero_inj != 0 or self.modes_combobox.get() == '':
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
