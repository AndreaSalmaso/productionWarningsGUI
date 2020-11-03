from tkinter import *
from tkinter import ttk, font
import ctypes, center_tk_window, keyboard, math
import pandas as pd
import mytkinter as mytk
import excel_handler as eh
from popup import warning_msg


class InserimentoSegnalazioniGUI:

    def __init__(self, master, main_color, path_excel):
        self.master = master
        self.main_color = main_color
        self.path_excel = path_excel

        # ***************************************************
        #        SEZIONE SUPERIORE (INSERIMENTO DATI) 
        # ***************************************************
        self.root0 = Frame(self.master, bg=self.main_color)
        self.root0.pack(pady=20)

        # ID OPERATORE
        self.frm0 = Frame(self.root0, bg=self.main_color)
        self.frm0.pack(padx=10, pady=10, fill=BOTH)
        mytk.create_label(self.frm0, 'ID operatore', 12)
        self.opID = StringVar()
        self.opID_ent = mytk.create_entry(self.frm0, self.opID, 7)
        self.opID_ent.pack()
        self.opID_ent.focus()
        #---------------------------------------------

        self.root1 = Frame(self.master, bg=self.main_color)
        self.root1.pack()
        #----- LATO SX --------------------------------------

        self.container1 = Frame(self.root1, bg=self.main_color)
        self.container1.pack(side=LEFT, anchor=NW, padx=30, pady=15)

        # ORDINE DI PRODUZIONE
        self.frm1 = Frame(self.container1, bg=self.main_color)
        self.frm1.pack(padx=10, pady=20, fill=BOTH)
        mytk.create_label(self.frm1, 'Ordine di produzione')
        self.ordine = StringVar()
        self.ordine_ent = mytk.create_entry(self.frm1, self.ordine, 9)
        self.ordine_ent.bind('<FocusOut>', mytk.to_uppercase)
        self.ordine_ent.pack(side=LEFT)

        # CODICE ARTICOLO
        self.frm2 = Frame(self.container1, bg=self.main_color)
        self.frm2.pack(padx=10, pady=20, fill=BOTH)
        mytk.create_label(self.frm2, 'Codice articolo')
        self.codice = StringVar()
        self.codice_ent = mytk.create_entry(self.frm2, self.codice, 20)
        self.codice_ent.bind('<KeyRelease>', self.update_listbox)
        self.codice_ent.pack(side=LEFT)

        # LISTA CODICI ARTICOLO
        self.frm3 = Frame(self.container1, bg=self.main_color)
        self.frm3.pack(padx=10, pady=20, fill=BOTH)
        self.frm31 = Frame(self.frm3)
        self.frm31.pack(side=LEFT)
        mytk.create_label(self.frm31)
        self.frm32 = Frame(self.frm3, bg='white')
        self.frm32.pack(side=LEFT, fill=BOTH)
        self.lb = Listbox(self.frm32, bd=0, highlightthickness=0, highlightbackground='white', takefocus=0, selectmode=SINGLE)
        self.scrollb = Scrollbar(self.frm32, width=20, command=self.lb.yview)
        self.lb.config(yscrollcommand=self.scrollb.set)
        self.lb.bind('<Double-Button-1>', self.enter_item_code)
        self.scrollb.pack(side=RIGHT, fill=Y)
        self.lb.pack(padx=15, pady=15, fill=BOTH, expand=True)

        # add products code into the listbox
        self.item_list = eh.get_item_codes(self.path_excel)
        self.fill_listbox(self.item_list)
        # -------------------------------------------------------


        #----- LATO DX ---------------------------------------
        self.container2 = Frame(self.root1, bg=self.main_color)
        self.container2.pack(side=LEFT, anchor=NW, padx=30, pady=15, fill=X)

        # OGGETTO DELLA RILEVAZIONE
        self.frm5 = Frame(self.container2, bg=self.main_color)
        self.frm5.pack( padx=10, pady=20, anchor=NW, fill=BOTH)
        mytk.create_label(self.frm5, 'Oggetto della rilevazione')
        self.ogg_rilev = ttk.Combobox(self.frm5, width=15, values=['prodotto finito', 'semilavorato', 'altro'])
        self.ogg_rilev.pack(side=LEFT, padx=50)
        mytk.create_label(self.frm5, '', width=10)

        # NUMERO SERIALE
        self.frm6 = Frame(self.container2, bg=self.main_color)
        self.frm6.pack(padx=10, pady=20, fill=BOTH)
        mytk.create_label(self.frm6, 'Numero seriale', 15)
        self.seriale = StringVar()
        self.seriale_ent = mytk.create_entry(self.frm6, self.seriale)
        self.seriale_ent.bind('<FocusOut>', mytk.to_uppercase)
        self.seriale_ent.pack(side=LEFT)

        self.checkVar = IntVar()
        mytk.create_label(self.frm6, '', 5)
        Checkbutton(self.frm6, var=self.checkVar, bg=self.main_color, activebackground=self.main_color, takefocus=0, relief=FLAT).pack(side=LEFT, padx=5)
        self.chkbtn_lbl = Label(self.frm6, text='Seriale non disponibile', anchor='w', fg='white', bg=self.main_color, font=('ubuntu', 12))
        self.chkbtn_lbl.pack(side=LEFT)

        # DESCRIZIONE DEL PROBLEMA
        self.frm7 = Frame(self.container2, bg=self.main_color)
        self.frm7.pack(padx=10, pady=4, fill=BOTH)
        mytk.create_label(self.frm7, 'Descrizione')
        self.frm8 = Frame(self.container2, bg='white')
        self.frm8.pack(padx=10, pady=10, fill=BOTH)
        self.descr = Text(self.frm8, height=9, width=63, bd=0)
        self.descr.pack(side=LEFT, padx=15, pady=15)
        # -------------------------------------------------------

        # ***************************************************
        #               SEZIONE INFERIORE (TASTI)
        # ***************************************************
        self.root2 = Frame(self.master, bg=self.main_color)
        self.root2.pack(padx=250, fill=X)
        self.canc_btn = mytk.create_img_button(self.root2, 'media/button_cancella-tutto.png', 'media/button_cancella-tutto_hov.png', self.clear_all)
        self.canc_btn.pack(side=RIGHT, padx=10)
        self.agg_segn_btn = mytk.create_img_button(self.root2, 'media/button_aggiungi-segnalazione.png', 'media/button_aggiungi-segnalazione_hov.png', self.add_report)
        self.agg_segn_btn.pack(side=RIGHT)

        # ***************************************************
        #               SEZIONE INFERIORE (EXCEL RECORDS)
        # ***************************************************
        self.root3 = Frame(self.master, bg=self.main_color)
        self.root3.pack(padx=250, pady=30, fill=X)

        self.root4 = Frame(self.master, bg=self.main_color)
        self.root4.pack()
        self.df = pd.read_excel(self.path_excel) 
        self.records_frm = Frame(self.root4, bg=self.main_color)
        self.records_frm.pack(padx=20)
        self.tree_font = font.Font(family='ubuntu', size=12)

        self.last_records = ttk.Treeview(self.records_frm, height=7)
        ttk.Style().configure('Treeview', rowheight=30, font=('ubuntu', 12))
        ttk.Style().configure('Treeview.Heading',  font=('ubuntu', 12, 'bold'))

        self.column_list = list(self.df.columns)
        self.last_records['column'] = self.column_list

        # format columns
        for i, col in enumerate(self.column_list):
            if i == 1 or i == 3:
                width = self.tree_font.measure(col) + 70
            elif i == 4:
                width = 400
            else:
                width = self.tree_font.measure(col) + 30

            self.last_records.column(col, width=width, minwidth=width, anchor=CENTER)    

        # specify the labels to be displayed (thanks to the next 
        # definition of the headings, the #0 column won't be showned)
        self.last_records['show'] = 'headings' 

        # define the headings
        for column in self.last_records['column']:
            self.last_records.heading(column, text=column, anchor=CENTER)

        self.show_last_records()
        self.last_records.pack()

        center_tk_window.center_on_screen(self.master)

        self.update_treeview()


    # refresh the treeview to show possible new entries from other operators
    def update_treeview(self):
        self.master.after(20000, self.show_last_records)

    def show_last_records(self):
        # clear treeview
        self.last_records.delete(*self.last_records.get_children())

        self.df = pd.read_excel(self.path_excel)
        df_rows = self.df.to_numpy().tolist()
        # loop through rows
        for row in df_rows[-7:]:
            # troncate problem description if too long
            probl_maxlen = 50
            if len(row[4]) > probl_maxlen:
                i = 0
                while True:
                    if row[4][probl_maxlen - i] == ' ':
                        row[4] = row[4][:probl_maxlen - i] + ' [...]'
                        break
                    else:
                        i -= 1

            row[4] = row[4].capitalize() # capitalize the 1° letter in the problem description           
            self.last_records.insert('', 'end', values=row)

    def add_report(self):

        d = self.descr.get('1.0',END)[:-1]

        if self.seriale.get()=='' and self.checkVar.get()==0:
            win = Toplevel(self.master)
            msg = 'Il numero seriale non è stato inserito!'
            warning_msg(win, msg)
        elif self.ordine.get()=='' or self.codice.get()=='' or self.ogg_rilev.get()=='' or d=='' or self.opID.get()=='':
            win = Toplevel(self.master)
            msg = 'Tutti i campi devono essere riempiti!'
            warning_msg(win, msg)
        else:
            d = self.descr.get('1.0',END)[:-1]
            if self.checkVar.get() == 1:
                s = 'Non disponibile'
            else:
                s = self.seriale.get()

            report_data = [self.ordine.get(), self.codice.get(), self.ogg_rilev.get(), s, d, self.opID.get()]
            eh.add_data_to_sheet(self.path_excel, report_data)
            self.clear_all()
            self.show_last_records()

    def clear_all(self):
        self.ordine.set('')
        self.codice.set('')
        self.ogg_rilev.set('')
        self.seriale.set('')
        self.checkVar.set(0)
        self.descr.delete('1.0', END)
        self.fill_listbox(self.item_list)
        self.ordine_ent.focus()

    def fill_listbox(self, items):
        self.lb.delete(0, END)
        for item in items:
            self.lb.insert(END, item)

    def enter_item_code(self, event):
        self.codice_ent.delete(0, END)
        self.codice_ent.insert(0, self.lb.get(self.lb.curselection()[0]))

    def update_listbox(self, event):
        partial = self.codice.get().upper()
        new_list = []
        for item in eh.get_item_codes(self.path_excel):
            if partial in item:
                new_list.append(item)
        self.lb.delete(0, END)
        self.fill_listbox(new_list)


def main():
    fileExcel_PATH = 'Dati/SegnalazioniProduzione.xlsx'
    BG_COLOR = '#545451'

    root = Tk()
    myfont = font.Font(family='ubuntu', size=16)
    root.title('Segnalazioni Produzione')
    root.iconbitmap(root, 'media/sheet.ico')
    root.option_add('*font', myfont)
    root['bg'] = BG_COLOR

    app = InserimentoSegnalazioniGUI(root, BG_COLOR, fileExcel_PATH)

    # maximize window at opening
    keyboard.press('alt+space+n')
    keyboard.release('alt+space+n')

    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    root.mainloop()

if __name__ == "__main__":
    main()