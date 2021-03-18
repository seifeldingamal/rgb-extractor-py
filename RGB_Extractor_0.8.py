import os
from tkinter import *
import tkinter.filedialog
from pathlib import Path
from PIL import Image
import pandas as pd

def get_RGB(p, df):

    date = Image.open(p)._getexif()[36867]
    pic = Image.open(p)
    pic_pixel_map = pic.load()

    width, length = pic.size

    R = 0
    G = 0
    B = 0

    for i in range(width):
        for j in range(length):
            colors = pic_pixel_map[i, j]
            R += colors[0]
            G += colors[1]
            B += colors[2]

    Sum = R+G+B
    df = df.append({'Date': date, 'Total_R': R, 'Total_G': G, 'Total_B': B, 'Sum': Sum}, ignore_index=True)

    return df

def alg(name,path):
    df = pd.DataFrame(columns=['Date', 'Total_R', 'Total_G', 'Total_B', 'Sum'])
    filename = name + ".xlsx"

    pathlist = Path(path).glob('**/*.JPG')

    for path in pathlist:
        path_in_str = str(path)
        df = get_RGB(path_in_str, df)

    df.to_excel(filename)

window = Tk()
window.title("RGB Data into Excel")
window.geometry('350x200+200+150')

lbl = Label(window, text="Hello")
lbl.grid(column=0, row=0)

txt = Entry(window,width=10)
txt.grid(column=1, row=0)

def run():
    filename = txt.get()
    curr_directory = os.getcwd()
    path = tkinter.filedialog.askdirectory()
    alg(filename,path)

def clicked():
    res = "Welcome to " + txt.get()
    lbl.configure(text= res)

btn = Button(window, text="Enter", command=clicked)
btn.grid(column=2, row=0)

btn2 = Button(window, text="Choose Directory and Run", command=run)
btn2.grid(columnspan = 2)

window.mainloop()
