import os
import sys
import numpy as np

from tkinter import *
from tkinter.filedialog import *
from tkinter import filedialog
from tkinter.messagebox import showinfo

class fromeDemo:
    def __init__(self):#主窗口设置参数

        def Select_Path():
            file_path = filedialog.askopenfilename()
        def Save_Path():
            file_path = filedialog.askdirectory()
            print(file_path)

        window = Tk()
        window_Width = window.winfo_screenwidth()
        window_Height = window.winfo_screenheight()
        window.title("PDF2Image")
        ww = 400
        wh = 150
        x = (window_Width-ww)/2
        y = (window_Height-wh)/2
        window.geometry("%dx%d+%d+%d" %(ww,wh,x,y))

        text_label_select = Label(window,width=20,text='Selected File')
        text_label_select.place(relx=-0.08,rely=0.2,anchor='w')
        textBox_select = Entry(window,width=30)
        textBox_select.place(relx=0.48,rely=0.2,anchor='center')

        text_label_select = Label(window,width=20,text='Save Path')
        text_label_select.place(relx=-0.08,rely=0.5,anchor='w')
        textBox_select = Entry(window,width=30)
        textBox_select.place(relx=0.48,rely=0.5,anchor='center')

        bt1=Button(window,width=10,height=1,text='Select File',command=Select_Path)        
        bt1.place(relx=0.87,rely=0.2,anchor='center')
        bt2=Button(window,width=10,height=1,text='Save Path',command=Save_Path)       
        bt2.place(relx=0.87,rely=0.5,anchor='center')
        bt2=Button(window,width=10,height=1,text='RUN')        
        bt2.place(relx=0.87,rely=0.8,anchor='center')

        window.resizable(0,0)
        window.mainloop()

fromeDemo()
    

