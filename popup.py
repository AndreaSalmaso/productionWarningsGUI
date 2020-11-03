from tkinter import *
from PIL import ImageTk, Image
import ctypes, os, time, center_tk_window

def warning_msg(win, msg):
    
    win.geometry("350x100+0+0")
    win.wm_title('Dati mancanti')
    win.iconbitmap('media/alert.ico')
    bg_color = '#42423e'
    win['bg'] = bg_color

    pic = Image.open('media/alert.png')
    pic = pic.resize((64,64), Image.ANTIALIAS)
    img = ImageTk.PhotoImage(pic)
    frm = Frame(win, bg=bg_color)
    frm.pack(padx=20, pady=20)
    panel = Label(frm, image = img, bg=bg_color)
    panel.pack(side=LEFT, fill = "both", expand = "yes")
    Label(frm, text=msg, fg='white', bg=bg_color, font=('ubuntu', 14), wraplength=200).pack(side=LEFT, padx=10)

    center_tk_window.center_on_screen(win)
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
    win.after(5000, lambda: win.destroy())
    win.mainloop()

if __name__=='__main__':
    win = Tk()
    msg = 'Il numero seriale non è stato inserito!'
    warning_msg(win, msg)