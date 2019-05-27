# -*- coding: utf-8 -*-
"""
Created on Tue Nov 27 03:38:25 2018

@author: Akash Singh
"""

#把所有扩展名为doc和docx的word文档全部转为pdf

import win32com.client
import glob #文件操作模块，用来查找符合自己目的的文件，类似于搜索
import os

input_direc = r"C:\Axis AI Challenge @ Akash_Abhishek\INPUT FILES"


wdFormatPDF = 17

word = win32com.client.Dispatch('Word.Application') #独立启动程序，启动word

docx_direc = r"C:\Axis AI Challenge @ Akash_Abhishek\INPUT FILES"
#total_docx_files = len(os.listdir(docx_direc))
#docx_lists=[]

#docx_files_name=os.listdir(input_direc)
   
results = [each for each in os.listdir(docx_direc) if each.endswith('.docx')] #路径下所有扩展名为.docx类型的word文档的名字
print(results)
docx = [word.Documents.Open(in_file) for in_file in glob.glob("C:\Axis AI Challenge @ Akash_Abhishek\INPUT FILES\*.docx")]

docx_no = len(docx)
for i in range (0, len(docx)):
    docx_no = i+1
    docx_orig = docx[i] 
    name = results[i]
    name_list = os.path.splitext(name)[0]#分离文件名和扩展名，得到的第一个是文件名
    out_file = r"C:\Axis AI Challenge @ Akash_Abhishek\INPUT FILES\{}.pdf".format(name_list)
    docx_orig.SaveAs(out_file, FileFormat=wdFormatPDF)
    docx_orig.Close()
#    os.rmdir( r"C:\Axis AI Challenge @ Akash_Abhishek\INPUT FILES\{}.docx".format(name_list) )
    
doc = [word.Documents.Open(in_file) for in_file in glob.glob("C:\Axis AI Challenge @ Akash_Abhishek\INPUT FILES\*.doc")]

doc_no = len(doc)
for i in range (0 , len(doc)):
    doc_no = i+1
    doc_orig = doc[i]    
    out_file = r"C:\Axis AI Challenge @ Akash_Abhishek\INPUT FILES\PDF of DOC_%d.pdf"%doc_no
    doc_orig.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc_orig.Close()

total_word_file = docx_no + doc_no    
#word.Quit()
