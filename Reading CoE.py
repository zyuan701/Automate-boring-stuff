#! python 3
import PyPDF2
import os
from PyPDF2 import PdfFileReader, PdfFileWriter
from pathlib import Path, PureWindowsPath

filepath = input ('Please enter the file path')
os.chdir(PureWindowsPath(filepath).parents[0])
filename = Path(filepath).name
pdfFile = open (filename, 'rb')
reader = PdfFileReader(pdfFile)
#Test if this pdf is encrypted under higher Adobe version; extract text from only decrypted file
if reader.isEncrypted == True:
    print ('Decryption required: Please go to https://smallpdf.com/unlock-pdf')
else:
    for pageNum in range(reader.numPages):
        print(reader.getPage(pageNum).extractText())