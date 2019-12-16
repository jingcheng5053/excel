import tkinter as ttk
from tkinter import *
from tkinter.filedialog import askopenfilename  # 文件打开对话框
from PIL import Image, ImageTk


path = StringVar()
path.set("eee")
print(path)

def choosepic():
    # global path
    path_ = askopenfilename()
    print(path_)
    path.set(path_)
    img_open = Image.open(path_)
    img = ImageTk.PhotoImage(img_open)
    labeltimu.config(image=img)
    labeltimu.image = img  # keep a reference


root = Tk()
root.title("Excel操作")
root.geometry('400x400+100+100')

butceshi = Button(root, text="读取", textvariable=path, command=choosepic)
butceshi.pack()
labeltimu = Label(root, image="")  # , bitmap='warning'
labeltimu.pack()

root.mainloop()
