#!/usr/bin/env python3
#encoding: utf-8

#開啟docx（聽說不能開doc檔）
from docx import Document
doc1 = Document('../linux應用_v9.75.doc')
#print(f'Doc={Doc}')

#列出所有大綱
for paragraph in doc1.paragraphs:
    print(f'大綱階層為：{paragraph.style.name}')
    print(f'內文為：{paragraph.text}')