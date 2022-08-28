from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from PIL import ImageTk, Image
import tkinter.font as font
from tkinter import filedialog
import docx2pdf
import pdf2docx

win= Tk()
win.title('PDF Converter')
win.geometry('400x525')
win.configure(background='white')

def docx_to_pdf():
    global filename
    filename = filedialog.askopenfilename(initialdir="/",
                                          title='Select File',
                                          filetypes=(('Word files', '*.docx'),
                                                     ('All files', '*.*')))


    if messagebox.askyesno('Docx to PDF', 'Confirm Convert?') == True:
        docx2pdf.convert(filename)
        messagebox.showinfo('SUCCESS',
                            'Congrats! Your docx File is successfully converted into PDF File')

def pdf_to_docx():
    global filename
    filename = filedialog.askopenfilename(initialdir="/",
                                          title='Select File',
                                          filetypes=(('PDF files', '*.pdf'),
                                                     ('All files', '*.*')))

    if messagebox.askyesno('PDF to Docx', 'Confirm Convert?') == True:
        pdf2docx.parse(filename)
        messagebox.showinfo('SUCCESS',
                            'Congrats! Your PDF File is successfully converted into docx File')

# Converting Word to PDF
w2p = Image.open("E:\Python\Py IDLE Projects\PDF Converter\word-to-pdf.png")
resize = w2p.resize((204, 122))
new = ImageTk.PhotoImage(resize)

label = Label(image = new)
label.pack(side=TOP, padx=0, pady=15)

txtFont = font.Font(family='Century Gothic', size=16, weight='bold')
txt = Label(text = 'Convert Word to PDF', bg='white',
            fg = '#2C3333', font=txtFont)
txt.pack(side=TOP, padx=0, pady=5)

buttonFont = font.Font(family='Century Gothic', size=12, weight='bold')
button = Button(text ="  Select a File  ", bg = "#1473e6",
                fg = "white", font=buttonFont,
                border=0, cursor='hand2',
                command=docx_to_pdf)
button.pack(side=TOP, padx=0, pady=5)


# Split Line
txtFont = font.Font(size=16)
txt = Label(text = '', bg='white', font=txtFont)
txt.pack(side=TOP, padx=0, pady=0)


# Converting PDF to Word
p2w = Image.open("E:\Python\Py IDLE Projects\PDF Converter\pdf-to-word.png")
resize = p2w.resize((204, 122))
new2 = ImageTk.PhotoImage(resize)

label = Label(image = new2)
label.pack(side=TOP, padx=0, pady=15)

txtFont = font.Font(family='Century Gothic', size=16, weight='bold')
txt = Label(text = 'Convert PDF to Word', bg='white',
            fg = '#2C3333', font=txtFont)
txt.pack(side=TOP, padx=0, pady=5)

buttonFont = font.Font(family='Century Gothic', size=12, weight='bold')
button = Button(text ="  Select a File  ", bg = "#1473e6",
                fg = "white", font=buttonFont,
                border=0, cursor='hand2',
                command=pdf_to_docx)
button.pack(side=TOP, padx=0, pady=5)

win.mainloop()
