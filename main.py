#!/usr/bin/env python3
#encoding: utf-8

#開啟docx（聽說不能開doc檔，但實際開好像沒問題XD）
from docx import Document
doc1 = Document('../linux應用_v9.75.doc')
#print(f'Doc={Doc}')

#列出所有大綱
#for paragraph in doc1.paragraphs:
    #print(f'大綱階層為：{paragraph.style.name}')
    #print(f'內文為：{paragraph.text}')

#建立Tkinter GUI
import tkinter
from tkinter import ttk

root = tkinter.Tk()
root.title('sam的Word處理器')

#千萬別加show='headings'，否則會縮小到看不到展開的+號
#tree = ttk.Treeview(root, columns=['1'], show='headings')
tree = ttk.Treeview(root, columns=['欄位ID 1','欄位ID 2'])

#anchor='center'表示置中對齊
#tree.column('1', width=100, anchor='center')
#tree.heading('1', text='大綱')
#tree["columns"]=("one","two")
tree.column("欄位ID 1", width=100)
tree.column("欄位ID 2", width=100)
tree.heading("欄位ID 1", text="欄位文字1")
tree.heading("欄位ID 2", text="欄位文字2")

#第一階層-1
#第二欄位表順序（0表最優先）
tree.insert("" , 0,'階層1-1 ID', text="階層1-1文字", values=['a','b'])

#第一階層-2
apple = tree.insert("", 2, "階層1-2 ID", text="階層1-2文字")
#第二階層
tree.insert(apple, 0,"階層2-1 ID1", text="階層2-1文字a", values=['c','d'])
tree.insert(apple, 1,"階層2-1 ID2", text="階層2-1文字b", values=['e','f'])

from PIL import Image, ImageTk

linux_img2 = Image.open('/home/sam/PycharmProjects/linux.png')
linux_img2 = linux_img2.resize((30,30), Image.ANTIALIAS)
linux_img2 = ImageTk.PhotoImage(linux_img2)
tree.insert(apple, 3,"階層2-1 ID2 pic2", text='linux.png', open=True, image=linux_img2,
                 value=['0','1'])

#第一階層-3
tree.insert('',3,"階層1-3 ID", text="階層1-3文字", values=['g','h'])
#tree.grid()
tree.pack()
root.mainloop()
