import os, os.path

from PyPDF2 import PdfFileWriter, PdfFileReader
from tkinter import Tk, Label, Entry, Button
from tkinter import filedialog
from tkinter.filedialog import askopenfilename


Tk().withdraw()
choosen_dir = filedialog.askdirectory()
os.chdir(choosen_dir)
filename = askopenfilename()
name=os.path.basename(filename)
num_pages = PdfFileReader(open(name, "rb"))


def mycom():
    e=edit.get()
    res = sum(((list(range(*[int(b)-1 + c
                             for c, b in enumerate(a.split('-'))]))
                if '-' in a else [int(a) -1]) for a in e.split(', ')), [])
    pages_to_delete = res  # page numbering starts from 0
    infile = PdfFileReader(name, 'rb')
    output = PdfFileWriter()

    for i in range(infile.getNumPages()):
        if i not in pages_to_delete:
            p = infile.getPage(i)
            output.addPage(p)

    with open(name, 'wb') as f:
        output.write(f)
    win.destroy()

win=Tk()
win.geometry('600x300')
t1=Label(win, text = 'Введите количество страниц')
t1.config(font=('Verdana', 25))
t1.pack()


edit = Entry(win, width = 20)
edit.pack()

but = Button(win, text='Удалить', command = mycom)
but.pack()

win.mainloop()


