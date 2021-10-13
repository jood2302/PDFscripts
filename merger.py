import os, os.path

import PyPDF2
import win32com.client
from PyPDF2 import PdfFileMerger, PdfFileReader

from docx2pdf import convert
from fpdf import FPDF
from PIL import Image
from tkinter import Tk, Label, Entry, Button
from tkinter import filedialog


Tk().withdraw()
choosen_dir = filedialog.askdirectory()
os.chdir(choosen_dir)

def mycom():
    e=edit.get()
    o=edit1.get()
    files = [f for f in os.listdir('.') ]
    for f in files:
        fbasename = os.path.splitext(os.path.basename(f))[0]
        if e in fbasename:
            # convert_doc to .docx
            word = win32com.client.Dispatch("Word.application")
            baseDir = choosen_dir
            for dir_path, dirs, files in os.walk(baseDir):
                for file_name in files:
                    file_path = os.path.join(dir_path, file_name)
                    file_name, file_extension = os.path.splitext(file_path)

                if "~$" not in file_name:
                    if file_extension.lower() == '.doc':  #
                        docx_file = '{0}{1}'.format(file_path, 'x')

                        if not os.path.isfile(docx_file):  # Skip conversion where docx file already exists

                            file_path = os.path.abspath(file_path)
                            docx_file = os.path.abspath(docx_file)
                            try:
                                wordDoc = word.Documents.Open(file_path)
                                wordDoc.SaveAs2(docx_file, FileFormat=16)
                                wordDoc.Close()
                            except Exception as e:
                                print('Failed to Convert: {0}'.format(file_path))
                                print(e)

            # convert .docx to .pdf
            if f.endswith('.docx'):
                convert(f, os.path.realpath('.') + '/' + fbasename + '.pdf')

            # convert .bmp to .pdf
            os.chdir(choosen_dir)
            for image in os.listdir('.'):
                if image.endswith('.bmp'):
                    canvas_image = Image.open(choosen_dir + '/' + image)
                    canvas_image.save(image + '.pdf', format='PDF', quality=200)
                    canvas_image.close()

            # convert .tiff to .pdf
            os.chdir(choosen_dir)
            for image in os.listdir('.'):
                if image.endswith('.tiff'):
                    canvas_image = Image.open(choosen_dir + '/' + image)
                    canvas_image.save(image + '.pdf', format='PDF', quality=200)
                    canvas_image.close()

            # convert .tif to .pdf
            os.chdir(choosen_dir)
            for image in os.listdir('.'):
                if image.endswith('.tif'):
                    canvas_image = Image.open(choosen_dir + '/' + image)
                    canvas_image.save(image + '.pdf', format='PDF', quality=200)
                    canvas_image.close()

            # convert .png to .pdf
            os.chdir(choosen_dir)
            pdf = FPDF()
            for image in os.listdir('.'):
                if image.endswith('.png'):
                    pdf = FPDF()
                    pdf.add_page()
                    pdf.image(image, x=50, y=100, w=pdf.w / 2.0, h=pdf.h / 4.0)
                    pdf.output(image + ".pdf", "F")

            # convert .jpg to .pdf
            os.chdir(choosen_dir)
            for image in os.listdir('.'):
                if image.endswith('.jpg'):
                    canvas_image = Image.open(choosen_dir + '/' + image)
                    canvas_image.save(image + '.pdf', format='PDF', quality=200)

            # convert .jpeg to .pdf
            os.chdir(choosen_dir)
            for image in os.listdir('.'):
                if image.endswith('.jpeg'):
                    canvas_image = Image.open(choosen_dir + '/' + image)
                    canvas_image.save(image + '.pdf', format='PDF', quality=200)

    # merger pdf
    merger = PdfFileMerger()
    for f in os.listdir('.'):
        if f.endswith('.pdf'):
            merger.append(PdfFileReader(f, 'rb'))
    merger.write(o + ".pdf")
    merger.close()

    win.destroy()

win=Tk()
win.geometry('600x300')
t1=Label(win, text = 'Введите наименование файлов, которые хотите выбрать')
t1.config(font=('Verdana', 10))
t1.pack()
edit = Entry(win, width = 20)
edit.pack()

t2=Label(win, text = 'Введите имя итогового файла')
t2.config(font=('Verdana', 10))
t2.pack()
edit1 = Entry(win, width = 20)
edit1.pack()

but = Button(win, text='Merger', command = mycom)
but.pack()

win.mainloop()


