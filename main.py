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
tree = ttk.Treeview(root, columns=['one','two'])

#anchor='center'表示置中對齊
#tree.column('1', width=100, anchor='center')
#tree.heading('1', text='大綱')
#tree["columns"]=("one","two")
tree.column("one", width=100)
tree.column("two", width=100)
tree.heading("one", text="coulmn A")
tree.heading("two", text="column B")

tree.insert("" , 0, text="Line 1", values=['山姆'])

apple = tree.insert("", 1, "iamD", text="DDDDD")
tree.insert(apple, "end","iamd", text="ddddd", values=['Tree'])

#第一階層
#hierarchy1 = tree.insert('', 'end', values=['山姆'])
#第二階層
#tree.insert(hierarchy1, 'end', values=['Tree'])

#第一階層
tree.insert('','end', values=['Canvas'])
#tree.grid()
tree.pack()
root.mainloop()
