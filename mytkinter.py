from tkinter import *
import ctypes


def create_frame(container):
    frm = Frame(container, bg='#545451')
    frm.pack(padx=10, pady=20, fill='both')
    return frm

def create_label(frame, lbl_text='', width=20):
    Label(frame, text=lbl_text, width=width, anchor='w', fg='white', bg='#545451').pack(side='left')

def create_entry(frame, var, width=10):
    return Entry(frame, textvariable=var, width=width, relief=FLAT, highlightthickness=1, highlightbackground='white')

def change_btn_img(event, img):
    event.widget['image'] = img

def create_img_button(frame, img, img_hov, action):
    btn_pic = PhotoImage(file=img)
    btn_pic_hov = PhotoImage(file=img_hov)
    btn = Button(frame, image=btn_pic, command=action, bd=0, bg='#545451', activebackground='#545451')
    btn.bind('<Enter>', lambda event, arg=btn_pic_hov: change_btn_img(event, arg))
    btn.bind('<Leave>', lambda event, arg=btn_pic: change_btn_img(event, arg))
    return btn

def to_uppercase(event):
    text = event.widget.get().upper()
    event.widget.delete(0, END)
    event.widget.insert(0, text)