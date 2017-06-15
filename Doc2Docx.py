# -*- coding: utf-8 -*-
#import re
#import docx2txt
#import unicodedata
import os
import win32com.client
#import textract
#import sys
#import pythoncom
#import fnmatch
#from subprocess import Popen, PIPE
from shutil import copyfile

# doc2docx
files = os.listdir("./doc")

if not os.path.exists("./docx"):
	print("Creating \\docx directory")
	os.makedirs("./docx")

wrd = win32com.client.Dispatch("Word.Application")
wrd.Visible = False

for file in files:
	file = str(file)
	FILENAME = file.split('.')
	if FILENAME[1] == "docx":
		copyfile(os.getcwd() + "\\doc\\" + FILENAME + ".docx", os.getcwd() + "\\docx\\" + FILENAME + ".docx")
	elif FILENAME[1] != "doc":
		continue
	FILENAME = FILENAME[0]
	PATH = os.getcwd() + "\\doc\\" + FILENAME + ".doc"
	PATH_TO_DOCX = os.getcwd() + "\\docx\\" + FILENAME + ".docx"
	wb = wrd.Documents.Open(PATH)
	wb.SaveAs(PATH_TO_DOCX, FileFormat = 16)
	wb.Close()
	print("Transforming \"" + FILENAME + ".doc\"")

wrd.Quit()