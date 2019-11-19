#!/usr/bin/env python3
#encoding: utf-8

#開啟docx（聽說不能開doc檔）
from docx import Document
doc1 = Document('../linux應用_v9.75.doc')
#print(f'Doc={Doc}')

for paragraph in doc1.paragraphs:
    parStr = paragraph.text
    print(f'parStr={parStr}')