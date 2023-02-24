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


class Zoom(ttk.Frame):
    """ Simple zoom with mouse wheel """

    def __init__(self, mainframe, path):
        """ Initialize the main Frame """
        ttk.Frame.__init__(self, master=mainframe)
        # self.master.title('Simple zoom with mouse wheel')
        # Vertical and horizontal scrollbars for canvas
        vbar = AutoScrollbar(self.master, orient='vertical')
        hbar = AutoScrollbar(self.master, orient='horizontal')
        vbar.grid(row=0, column=1, sticky='ns')
        hbar.grid(row=1, column=0, sticky='we')
        # Open image
        self.image = Image.open(path)
        # Create canvas and put image on it
        self.canvas = tk.Canvas(self.master, highlightthickness=0,
                                xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        self.canvas.grid(row=0, column=0, sticky='nswe')
        vbar.configure(command=self.canvas.yview)  # bind scrollbars to the canvas
        hbar.configure(command=self.canvas.xview)
        # Make the canvas expandable
        self.master.rowconfigure(0, weight=1)
        self.master.columnconfigure(0, weight=1)
        # Bind events to the Canvas
        self.canvas.bind('<ButtonPress-1>', self.move_from)
        self.canvas.bind('<B1-Motion>', self.move_to)
        self.canvas.bind('<MouseWheel>', self.wheel)  # with Windows and MacOS, but not Linux
        self.canvas.bind('<Button-5>', self.wheel)  # only with Linux, wheel scroll down
        self.canvas.bind('<Button-4>', self.wheel)  # only with Linux, wheel scroll up
        # Show image and plot some random test rectangles on the canvas
        self.imscale = 1.0
        self.imageid = None
        self.delta = 0.75
        width_a, height = self.image.size
        minsize, maxsize = 5, 20
        for n in range(10):
            x0 = random.randint(0, width_a - maxsize)
            y0 = random.randint(0, height - maxsize)
            x1 = x0 + random.randint(minsize, maxsize)
            y1 = y0 + random.randint(minsize, maxsize)
            color = ('red', 'orange', 'yellow', 'green', 'blue')[random.randint(0, 4)]
            self.canvas.create_rectangle(x0, y0, x1, y1, outline='black', fill=color,
                                         activefill='black', tags=n)
        # Text is used to set proper coordinates to the image. You can make it invisible.
        self.text = self.canvas.create_text(0, 0, anchor='nw', text='')
        self.show_image()
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def move_from(self, event):
        """ Remember previous coordinates for scrolling with the mouse """
        self.canvas.scan_mark(event.x, event.y)

    def move_to(self, event):
        """ Drag (move) canvas to the new position """
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def wheel(self, event):
        """ Zoom with mouse wheel """
        scale = 1.0
        # Respond to Linux (event.num) or Windows (event.delta) wheel event
        if event.num == 5 or event.delta == -120:
            scale *= self.delta
            self.imscale *= self.delta
        if event.num == 4 or event.delta == 120:
            scale /= self.delta
            self.imscale /= self.delta
        # Rescale all canvas objects
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        self.canvas.scale('all', x, y, scale, scale)
        self.show_image()
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def show_image(self):
        """ Show image on the Canvas """
        if self.imageid:
            self.canvas.delete(self.imageid)
            self.imageid = None
            self.canvas.imagetk = None  # delete previous image from the canvas
        width, height = self.image.size
        new_size = int(self.imscale * width), int(self.imscale * height)
        imagetk = ImageTk.PhotoImage(self.image.resize(new_size))
        # Use self.text object to set proper coordinates
        self.imageid = self.canvas.create_image(self.canvas.coords(self.text),
                                                anchor='nw', image=imagetk)
        self.canvas.lower(self.imageid)  # set it into background
        self.canvas.imagetk = imagetk  # keep an extra reference to prevent garbage-collection


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
            "ZCS-SOFAR-1.1_3KTL-L-G3": dict_all["1PH"]["ZCS-SOFAR-1.1_3KTL-L-G3"]['MODELLO'],
            "ZCS-1PH-1100_3300TL-V3": dict_all["1PH"]["ZCS-1PH-1100_3300TL-V3"]['MODELLO'],
            "ZCS-1PH-3000_6000TLM-V3": dict_all["1PH"]["ZCS-1PH-3000_6000TLM-V3"]['MODELLO'],
            "ZCS-SOFAR-1100_3000TL": dict_all["1PH"]["ZCS-SOFAR-1100_3000TL"]['MODELLO'],
            "ZCS-1PH-1100_3300TL-V1": dict_all["1PH"]["ZCS-1PH-1100_3300TL-V1"]['MODELLO'],
            "ZCS-1PH-3000_6000TLM-V1": dict_all["1PH"]["ZCS-1PH-3000_6000TLM-V1"]['MODELLO'],
            "ZCS-SOFAR-3_6KTLM_L": dict_all["1PH"]["ZCS-SOFAR-3_6KTLM_L"]['MODELLO'],
            "ZCS-1PH-3000_6000TLM-V2": dict_all["1PH"]["ZCS-1PH-3000_6000TLM-V2"]['MODELLO'],
            "ZCS-SOFAR-3000SP": dict_all["1PH"]["ZCS-SOFAR-3000SP"]['MODELLO'],
            "ZCS-3000SP-V2": dict_all["1PH"]["ZCS-3000SP-V2"]['MODELLO'],
            "ZCS-SOFAR-HYD-3000_6000-ES": dict_all["1PH"]["ZCS-SOFAR-HYD-3000_6000-ES"]['MODELLO'],
            "ZCS-1PH-HYD-3000_6000-ZSS": dict_all["1PH"]["ZCS-1PH-HYD-3000_6000-ZSS"]['MODELLO'],
            "ZCS-1PH-HYD-3000_6000-ZSS-HP": dict_all["1PH"]["ZCS-1PH-HYD-3000_6000-ZSS-HP"]['MODELLO'],
            "ZCS-SOFAR-HYD-5000_8000-T": dict_all["3PH"]["ZCS-SOFAR-HYD-5000_8000-T"]['MODELLO'],
            "ZCS-3PH-HYD-5000_8000-ZSS": dict_all["3PH"]["ZCS-3PH-HYD-5000_8000-ZSS"]['MODELLO'],
            "ZCS-SOFAR-HYD-10000_20000-T": dict_all["3PH"]["ZCS-SOFAR-HYD-10000_20000-T"]['MODELLO'],
            "ZCS-3PH-HYD-10000_20000-ZSS": dict_all["3PH"]["ZCS-3PH-HYD-10000_20000-ZSS"]['MODELLO'],
            "ZCS-SOFAR-50000_70000TL-V1": dict_all["3PH"]["ZCS-SOFAR-50000_70000TL-V1"]['MODELLO'],
            "ZCS-3PH-50000_60000TL-V1": dict_all["3PH"]["ZCS-3PH-50000_60000TL-V1"]['MODELLO'],
            "ZCS-3PH-80_110KTL-LV": dict_all["3PH"]["ZCS-3PH-80_110KTL-LV"]['MODELLO'],
            "ZCS-SOFAR-80_110KTL": dict_all["3PH"]["ZCS-SOFAR-80_110KTL"]['MODELLO'],
            "ZCS-3PH-100_136KTL-HV": dict_all["3PH"]["ZCS-3PH-100_136KTL-HV"]['MODELLO'],
            "ZCS-SOFAR-100_136KTL-HV": dict_all["3PH"]["ZCS-SOFAR-100_136KTL-HV"]['MODELLO'],
            "ZCS-SOFAR-3.3_12KTL": dict_all["3PH"]["ZCS-SOFAR-3.3_12KTL"]['MODELLO'],
            "ZCS-3PH-3.3_12KTL-V1": dict_all["3PH"]["ZCS-3PH-3.3_12KTL-V1"]['MODELLO'],
            "ZCS-3PH-3.3_12KTL-V3": dict_all["3PH"]["ZCS-3PH-3.3_12KTL-V3"]['MODELLO'],
            "ZCS-3PH-10_15KTL-V2": dict_all["3PH"]["ZCS-3PH-10_15KTL-V2"]['MODELLO'],
            "ZCS-3PH-25_50KTL-V3": dict_all["3PH"]["ZCS-3PH-25_50KTL-V3"]['MODELLO'],
            "ZCS-SOFAR-10000_20000TL": dict_all["3PH"]["ZCS-SOFAR-10000_20000TL"]['MODELLO'],
            "ZCS-SOFAR-20000_33000TL-G2": dict_all["3PH"]["ZCS-SOFAR-20000_33000TL-G2"]['MODELLO'],
            "ZCS-3PH-20000_33000TL-V2": dict_all["3PH"]["ZCS-3PH-20000_33000TL-V2"]['MODELLO'],
            "ZCS-3PH-15000_24000TL-V3": dict_all["3PH"]["ZCS-3PH-15000_24000TL-V3"]['MODELLO'],
            "ZCS-SOFAR-30000_40000TL": dict_all["3PH"]["ZCS-SOFAR-30000_40000TL"]['MODELLO'],
            "ZCS-SOFAR-3000_6000TLM": dict_all["1PH"]["ZCS-SOFAR-3000_6000TLM"]['MODELLO'],
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

        self.b1 = Button(self.newWindow, text='Database Modelli', command=self.database_modelli, width=width,
                         bg='lightgreen')
        self.b1.place(x=50, y=30)

        self.b2 = Button(self.newWindow, text='Database Service', command=self.database_service, width=width,
                         bg='lightgreen')
        self.b2.place(x=200, y=30)

    def database_modelli(self):
        self.Window_mod = Toplevel(self.newWindow)
        self.Window_mod.title("R&D")
        self.Window_mod.iconbitmap(r'img\azzurro.ico')
        self.Window_mod.geometry("620x250")

        # self.lbl1 = Label(Window_mod, text='Modello')

        self.tipologia = tk.StringVar(self.Window_mod)
        self.serie = tk.StringVar(self.Window_mod)
        self.modello = tk.StringVar(self.Window_mod)

        self.tipologia.trace('w', self.update_options)
        self.serie.trace('w', self.update_options_b)

        self.optionmenu_a = tk.OptionMenu(self.Window_mod, self.tipologia, *self.dict.keys())
        self.optionmenu_b = tk.OptionMenu(self.Window_mod, self.serie, *self.dict_model.keys())
        self.optionmenu_c = tk.OptionMenu(self.Window_mod, self.modello, '')

        self.tipologia.set('1PH')
        self.optionmenu_a.pack()
        self.optionmenu_b.pack()
        self.optionmenu_c.pack()

        self.lbl_a = Label(self.Window_mod, text='Categoria', bg=bg)
        self.lbl_a.place(x=50, y=10)
        self.optionmenu_a.place(x=50, y=30)
        self.lbl_b = Label(self.Window_mod, text='Prodotto', bg=bg)
        self.lbl_b.place(x=150, y=10)
        self.optionmenu_b.place(x=150, y=30)
        self.lbl_c = Label(self.Window_mod, text='Modello', bg=bg)
        self.lbl_c.place(x=400, y=10)
        self.optionmenu_c.place(x=400, y=30)

        #self.add_info
        self.installazione()
        self.Window_info = self.Window_mod

    def database_service(self):
        Window_ser = Toplevel(self.newWindow)
        Window_ser.title("Service Database")
        Window_ser.iconbitmap(r'img\azzurro.ico')
        Window_ser.geometry("800x90")

        self.tipologia_ser = tk.StringVar(Window_ser)
        self.serie_ser = tk.StringVar(Window_ser)
        self.modello_ser = tk.StringVar(Window_ser)

        self.tipologia_ser.trace('w', self.update_options_ser)
        self.serie_ser.trace('w', self.update_options_b_ser)

        self.optionmenu_a_ser = tk.OptionMenu(Window_ser, self.tipologia_ser, *self.dict.keys())
        self.optionmenu_b_ser = tk.OptionMenu(Window_ser, self.serie_ser, *self.dict_model.keys())
        self.optionmenu_c_ser = tk.OptionMenu(Window_ser, self.modello_ser, '')

        self.tipologia_ser.set('1PH INV')
        self.optionmenu_a_ser.pack()
        self.optionmenu_b_ser.pack()
        self.optionmenu_c_ser.pack()

        self.optionmenu_a_ser.place(x=50, y=30)
        self.optionmenu_b_ser.place(x=150, y=30)
        self.optionmenu_c_ser.place(x=400, y=30)

        self.bservice = Button(Window_ser, text='ADD INFO', command=self.DB_service, width=width)
        self.bservice.place(x=600, y=30)

    def add_info(self):
        # self.Window_info.title("ADD INFOs")
        # self.Window_info.iconbitmap(r'img\azzurro.ico')
        # self.Window_info.geometry("800x100")

        self.dict_add_info = {
            "Ampliamento Impianto": ['Aumentare Potenza', 'Inserire Sistema di Accumulo',
                                     'Inserire Accumulo & Aumento Potenza'],
            "Nuovo Impianto": ['Puro PV', 'Impianto PV + Accumulo'],
            "Solo Ricarica veicoli elettrici": ['Monofase', 'Trifase']
        }

        self.info1 = tk.StringVar(self.Window_info)
        self.info2 = tk.StringVar(self.Window_info)
        self.info1.trace('w', self.add_info_opt)
        self.opt_info1 = tk.OptionMenu(self.Window_info, self.info1, *self.dict_add_info.keys())
        self.opt_info2 = tk.OptionMenu(self.Window_info, self.info2, '')
        self.info1.set("Ampliamento Impianto")
        self.opt_info1.pack()
        self.opt_info2.pack()
        self.lbl_info1 = Label(self.Window_info, text='Tipo di connessione', bg=bg)
        self.lbl_info1.place(x=50, y=70)
        self.opt_info1.place(x=50, y=100)
        self.lbl_info2 = Label(self.Window_info, text='Tipo di connessione', bg=bg)
        self.lbl_info2.place(x=250, y=70)
        self.opt_info2.place(x=250, y=100)

        # self.binfo = Button(Window_info, text='Tipologia Installazione', command=self.installazione, width=width)
        # self.binfo.place(x=600, y=50)

    def installazione(self):

        # if self.info2.get() == "Solo Ricarica veicoli elettrici":
        #     self.dict_add_inst = {
        #         'Accesso privato', 'Accesso limitato', 'Accesso con pagamento'
        #     }
        # else:
        self.dict_add_inst = {
            "Monitoraggio Remoto": ['WIFI', '4G', 'ETH'],
            "Zero Injection": ['TA', 'METER', 'ENERCLICK'],
            "Gestione Attiva Carichi Esterni": [''],
            "Monitoraggio Consumi": [''],
            "Ricarica Veicoli Elettrici": ['Accesso privato', 'Accesso limitato', 'Accesso con pagamento']
        }
        self.Window_info = self.Window_mod
        self.info_inst1 = tk.StringVar(self.Window_info)
        self.info_inst2 = tk.StringVar(self.Window_info)
        self.info_inst1.trace('w', self.add_inst_opt)
        self.opt_info_inst1 = tk.OptionMenu(self.Window_info, self.info_inst1, *self.dict_add_inst.keys())
        self.opt_info_inst2 = tk.OptionMenu(self.Window_info, self.info_inst2, '')
        self.info_inst1.set("Monitoraggio Remoto")
        self.opt_info_inst1.pack()
        self.opt_info_inst2.pack()
        self.lbl_info_inst1 = Label(self.Window_info, text='Connessione esterna', bg=bg)
        self.lbl_info_inst1.place(x=50, y=140)
        self.opt_info_inst1.place(x=50, y=170)
        self.lbl_info_inst2 = Label(self.Window_info, text='Tipo di connessione', bg=bg)
        self.lbl_info_inst2.place(x=250, y=140)
        self.opt_info_inst2.place(x=250, y=170)

        self.binfo = Button(self.Window_info, text='INFO', command=self.print_info, width=width, bg='lightgreen')
        self.binfo.place(x=400, y=140)

    def add_inst_opt(self, *args):
        info_instb = self.dict_add_inst[self.info_inst1.get()]
        self.info_inst2.set(info_instb[0])

        menu_info_inst2 = self.opt_info_inst2['menu']
        menu_info_inst2.delete(0, 'end')

        for name_product_b in info_instb:
            menu_info_inst2.add_command(label=name_product_b, command=lambda id=name_product_b: self.info_inst2.set(id))

    def add_info_opt(self, *args):
        infob = self.dict_add_info[self.info1.get()]
        self.info2.set(infob[0])

        menu_info2 = self.opt_info2['menu']
        menu_info2.delete(0, 'end')

        for name_product_b in infob:
            menu_info2.add_command(label=name_product_b, command=lambda id=name_product_b: self.info2.set(id))

    def update_options(self, *args):
        product_b = self.dict[self.tipologia.get()]
        self.serie.set(product_b[0])

        menu_b = self.optionmenu_b['menu']
        menu_b.delete(0, 'end')

        for name_product_b in product_b:
            menu_b.add_command(label=name_product_b, command=lambda id=name_product_b: self.serie.set(id))

    def update_options_b(self, *args):
        product_c = self.dict_model[self.serie.get()]
        self.modello.set(product_c[0])

        menu_c = self.optionmenu_c['menu']
        menu_c.delete(0, 'end')

        for name_product_c in product_c:
            menu_c.add_command(label=name_product_c, command=lambda id=name_product_c: self.modello.set(id))

    def update_options_ser(self, *args):
        product_b = self.dict[self.tipologia_ser.get()]
        self.serie_ser.set(product_b[0])

        menu_b = self.optionmenu_b_ser['menu']
        menu_b.delete(0, 'end')

        for name_product_b in product_b:
            menu_b.add_command(label=name_product_b, command=lambda id=name_product_b: self.serie_ser.set(id))

    def update_options_b_ser(self, *args):
        product_c = self.dict_model[self.serie_ser.get()]
        self.modello_ser.set(product_c[0])

        menu_c = self.optionmenu_c_ser['menu']
        menu_c.delete(0, 'end')

        for name_product_c in product_c:
            menu_c.add_command(label=name_product_c, command=lambda id=name_product_c: self.modello_ser.set(id))

    def print_info(self):
        value = [self.tipologia.get(), self.serie.get(), self.modello.get(),
                 self.info_inst1.get(), self.info_inst2.get(),
                 #self.info1.get(), self.info2.get()
                 ]
        if value[3] != 'Zero Injection':
            Window_info = Toplevel(self.newWindow)
            Window_info.title("INFOs")
            Window_info.geometry(
                "{0}x{1}+0+0".format(Window_info.winfo_screenwidth(), Window_info.winfo_screenheight()))
            Window_info.iconbitmap('img/azzurro.ico')

        df_all = pd.read_excel('./database/summa.xlsx')
        df = df_all[df_all['Modello'] == value[2]]
        df = df.dropna(axis=1, how='all')
        df = df.T
        df = df.reset_index()
        df = df.rename(columns={'index': "Inverter Datasheet", df.columns[1]: " "})

        if value[3] == 'Monitoraggio Remoto':
            # aggiunto i logger esterni
            df_concat = df_all[(df_all['Prodotto']=='Current sensors') | (df_all['Prodotto']=='Connext') |
                               (df_all['Prodotto']=='Meters') | (df_all['Prodotto']=='Single unit loggers') |
                               (df_all['Prodotto']=='External loggers')]
            df_concat = df_concat[df_concat['ConnAPP']==value[4]]
            df_concat = df_concat.dropna(axis=1, how='all')
            df_concat = df_concat.T.reset_index()
            df_concat = df_concat.rename(columns={'index': "Loggers Datasheet"})
            for i in df_concat.columns:
                if i != 'Loggers Datasheet':
                    df_concat=df_concat.rename(columns={i: " "})
            df = pd.concat([df, df_concat], axis=1)

        elif value[3] == 'Zero Injection':
            if dict_all[value[0]][value[1]]['0-INJ']['ABLE'] == True:
                if value[4] in dict_all[value[0]][value[1]]['0-INJ']['TIPO 0-INJ']:
                    documento_0inj(dict_all, value[1], value[0], value[4])
                else:
                    Window_info2 = Toplevel(self.newWindow)
                    Window_info2.title("Error")
                    Window_info2.geometry("300x50")
                    Window_info2.iconbitmap('img/azzurro.ico')
                    self.lbl_info_err = Label(Window_info2, text='Questo metodo non è compreso in questo prodotto', bg=bg)
                    self.lbl_info_err.place(x=10, y=10)
            else:
                Window_info2 = Toplevel(self.newWindow)
                Window_info2.title("Error")
                Window_info2.geometry("200x50")
                Window_info2.iconbitmap('img/azzurro.ico')
                self.lbl_info_err = Label(Window_info2, text='Questo modello non ha lo 0-inj', bg=bg)
                self.lbl_info_err.place(x=10, y=10)
        if value[3] != 'Zero Injection':
            f = Frame(Window_info)
            f.pack(fill=BOTH, expand=1)
            # df = TableModel.getSampleData()
            self.table = pt = Table(f,
                                    dataframe=df,
                                    showtoolbar=True, showstatusbar=True)
            pt.show()
            self.zoom_window(value[1])

    def DB_service(self):
        value = [self.tipologia_ser.get(), self.serie_ser.get(), self.modello_ser.get()]
        Window_info_ser = Toplevel(self.newWindow)
        Window_info_ser.title("INFOs")
        Window_info_ser.geometry(
            "{0}x{1}+0+0".format(Window_info_ser.winfo_screenwidth(), Window_info_ser.winfo_screenheight()))
        Window_info_ser.iconbitmap('img/azzurro.ico')
        df = pd.read_excel('./database/Service.xlsx')
        df = df[(df['Modello'] == value[2]) & (df['Prodotto'] == value[1])]
        f = Frame(Window_info_ser)
        f.pack(fill=BOTH, expand=1)
        # df = TableModel.getSampleData()
        self.table_ser = pt = Table(f,
                                    dataframe=df,
                                    showtoolbar=True, showstatusbar=True)
        pt.show()

        self.zoom_window(value[1])

    def zoom_window(self, value):
        window_graph = Toplevel(self.newWindow)
        window_graph.title("Datasheet")
        # window_graph.geometry('%dx%d+%d+%d' % (1100, 740, 500, 100))
        window_graph.geometry(
            "{0}x{1}+0+0".format(window_graph.winfo_screenwidth(), window_graph.winfo_screenheight()))
        window_graph.iconbitmap('img/azzurro.ico')

        path = self.serie.get()
        for file in glob.glob('./Datasheet/' + path + "/*.png"):
            img_print = file
        for file in glob.glob('./Datasheet/' + path + "/*.jpg"):
            img_print = file

        app = Zoom(window_graph, path=img_print)
        window_graph.mainloop()


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

# dict_datasheet = {
#     'Single unit loggers': 'scheda tecnica monitoring',
#     'External loggers': 'scheda tecnica monitoring',
#     'ZST-BAT-2,4KWH-PL': 'Datasheet batterie Pylontech US2000 (it 2021)',
#     'ZST-BAT-5KWH-PL': 'Datasheet batterie Pylontech US5000 (IT2022)0',
#     'ZZT-BAT-5KWH-W': 'Datasheet batterie Weco 4k4 LV (it 2021)',
#     'ZZT-BAT-5KWH-ZSX': 'Datasheet batterie LV ZSX5000 I ZSX5000 PRO (it 2022)',
#     'ZST-BAT-2,4KWH-H': 'Datasheet batterie Pylontech 48050 (it 2021)',
#     'ZZT-BAT-6KWH-W': 'Datasheet batterie Weco 5k3 HV (it 2021)',
#     'Current sensors': 'scheda tecnica meter e sensori-1',
#     'Connext': 'Datasheet connext (it 2020)-1',
#     'Meters': 'scheda tecnica meter e sensori-1',
#     'ZCS-SOFAR-30000_40000TL': 'Datasheet 30-40 TL (it 2020)',
#     'ZCS-3PH-15000_24000TL-V3': 'Datasheet trifase 15-24',
#     'ZCS-3PH-20000_33000TL-V2': 'scheda tecnica trifase 20-33 V2',
#     'ZCS-SOFAR-20000_33000TL-G2': 'Datasheet 20-33 KTL (it 2020)',
#     'ZCS-3PH-10_15KTL-V2': 'scheda tecnica trifase 10-15 V2',
#     'ZCS-SOFAR-10000_20000TL': 'Datasheet 10-20 TL (it 2020)',
#     'ZCS-3PH-3.3_12KTL-V3': 'scheda tecnica trifase 3.3-12 V3',
#     'ZCS-3PH-3.3_12KTL-V1': 'scheda tecnica trifase 3.3-12 V1',
#     'ZCS-SOFAR-3.3_12KTL': 'Datasheet 3.3-12 KTL (it 2020)',
#     'ZCS-SOFAR-100_136KTL-HV': 'scheda tecnica trifase 110-136 HV',
#     'ZCS-3PH-100_136KTL-HV': 'scheda tecnica trifase 110-136 HV',
#     'ZCS-SOFAR-80_110KTL': 'scheda tecnica trifase 80-110 LV',
#     'ZCS-3PH-80_110KTL-LV': 'scheda tecnica trifase 80-110 LV',
#     'ZCS-3PH-50000_60000TL-V1': '50000TL-60000TL',
#     'ZCS-SOFAR-50000_70000TL-V1': 'Datasheet 50-70 TL (it 2020)',
#     'ZCS-3PH-10000_20000-ZSS': 'scheda tecnica ibrido trifase 10-20',
#     'ZCS-SOFAR-HYD-10000_20000-T': 'Datasheet 10-20 TL (it 2020)',
#     'ZCS-3PH-5000_8000-ZSS': 'scheda tecnica ibrido trifase 5-8',
#     'ZCS-SOFAR-HYD-5000_8000-T': 'Datasheet ZCS HYD 5-8K T',
#     'ZCS-1PH-HYD-3000_6000-ZSS-HP': '1ph hyd3k-6k hp',
#     'ZCS-1PH-HYD-3000_6000-ZSS': '1ph hyd3k-6k',
#     'ZCS-SOFAR-HYD-3000_6000-ES': 'scheda tecnica ibrido monofase ES',
#     'ZCS-3000SP-V2': 'scheda tecnica ME3000SP',
#     'ZCS-SOFAR-3000SP': 'scheda tecnica ME3000SP',
#     'ZCS-1PH-3000_6000TLM-V2': 'scheda tecnica monofase 3-6 V2',
#     'ZCS-SOFAR-3_6KTLM_L': 'Datasheet 3-6 KTLM (it 2020)',
#     'ZCS-SOFAR-3000_6000TLM': 'Datasheet 3-6 TLM (it 2020)',
#     'ZCS-1PH-3000_6000TLM-V1': 'scheda tecnica monofase 3-6 V1',
#     'ZCS-1PH-1100_3300TL-V1': 'scheda tecnica monofase 1-3 V1',
#     'ZCS-SOFAR-1100_3000TL': 'Datasheet 1-3 TL (it 2020)',
#     'ZCS-1PH-3000_6000TLM-V3': 'Datasheet 3-6 TLM (it 2020)',
#     'ZCS-1PH-1100_3300TL-V3': 'scheda tecnica monofase 1-3 V3',
#     'ZCS-SOFAR-1.1_3KTL-L-G3': 'Datasheet 1-3 TL (it 2020)'
# }

window = Tk()
MyWindow(window)
window.title('R&D - Service')
window.geometry("400x90")
window.iconbitmap(r'img\azzurro.ico')
window.mainloop()
