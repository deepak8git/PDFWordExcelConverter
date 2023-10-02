from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter.ttk import *
from pdf2docx import Converter, parse
import tabula
import subprocess
import os
import sys

iMaxStackSize = 1000
sys.setrecursionlimit(iMaxStackSize)

def convert_pdf_to_word():
    try:
        pdf_file = convert_source_folder_path.get()
        word_file=convert_des_folder_path.get()
        cv = Converter(pdf_file)
        cv.convert(word_file,start=0,end=None)
        cv.close()
        messagebox.showinfo("Info","PDF to Word Converted Successfully")
        subprocess.run(['explorer', os.path.realpath(convert_des_folder_path.get())])
    except:
        messagebox.showerror("Warning","Some Error Occured")
    finally:
        convert_source_folder_path.set("")
        convert_des_folder_path.set("")

def convert_pdf_to_excel():
    try:
        source = convert_source_folder_path.get()
        destintion = convert_des_folder_path.get()
        tabula.convert_into(source,destintion,pages="all",output_format="csv")
        messagebox.showinfo("Info","PDF to Excel Converted Successfully")
        subprocess.run(['explorer', os.path.realpath(convert_des_folder_path.get())])
    except:
        messagebox.showerror("Warning","Some Error Occured")
    finally:
        convert_source_folder_path.set("")
        convert_des_folder_path.set("")    

def convert_to_word_excel():
    if mvalue.get() == 1:
        convert_pdf_to_word()
    elif mvalue.get() == 2:
        convert_pdf_to_excel()


def browse_convert_source_button():
    filename = filedialog.askopenfilename(initialdir = last_opened_source_path.get(), title = "Select file",filetypes = \
                                        (("pdf files","*.pdf"),("all files","*.*")))
    last_opened_source_path.set(os.path.dirname(filename))
    convert_source_folder_path.set(filename)
     
def browse_convert_destination_button():
    if mvalue.get() == 1:
        filename=filedialog.asksaveasfilename(initialdir = last_opened_dest_path.get(),title = "Save file",defaultextension="*.docx",filetypes = \
                                        (("word files","*.docx"),("all files","*.*")))
    elif mvalue.get() == 2:
        filename=filedialog.asksaveasfilename(initialdir = last_opened_dest_path.get(),title = "Save file",defaultextension="*.csv",filetypes = \
                                        (("csv files","*.csv"),("all files","*.*")))
    
    last_opened_dest_path.set(os.path.dirname(filename))
    convert_des_folder_path.set(filename)
   
def radio_convert_pdf():
    if mvalue.get() == 1:
        browse_convert_source_button()
    elif mvalue.get() == 2:
        browse_convert_source_button()
   


#This section initiate Graphics and defines position and geometry of windows which opens
root =Tk()
root.title("Welcome to PDF To Word/Excel Converter           Developed By: Deepak Kumar Ram")
#root.iconbitmap(r"C:\Users\Deepak\Desktop\Installer\sbi.ico")
        #root.overrideredirect(1)
root.resizable(0,0)  
root.withdraw()
root.update_idletasks()

w,h=610,200
x=(root.winfo_screenwidth() - w)/2
y=(root.winfo_screenheight() - h)/2
root.geometry("{}x{}+{}+{}".format(w,h,int(x),int(y)))
#****************************************************************************************

tabcontrol = Notebook(root)
tab2 = Frame(tabcontrol)
tabcontrol.add(tab2,text=" PDF To Word/Excel ")

# This section for Tab2 PDF Merger ***************************************************
convert_source_folder_path = StringVar()
convert_des_folder_path=StringVar()
last_opened_source_path=StringVar()
last_opened_dest_path=StringVar()

mvalue = IntVar()
mvalue.set(1)
pdfconvert = LabelFrame(tab2, text=" Convert PDF to Word Or Excel File ", width=585, height=110)
pdfconvert.place(x=10,y=10)

convert_to_word_radio =Radiobutton(pdfconvert, text = "Convert to Word", variable=mvalue, value=1)
convert_to_word_radio.place(x=70, y=14, anchor=W)

convert_to_excel_radio = Radiobutton(pdfconvert,text="Convert to Excel", variable=mvalue, value=2)
convert_to_excel_radio.place(x=187, y=14, anchor=W)

convert_source_label = Label(pdfconvert, text="Source")
convert_source_label.place(x=5, y=40, anchor="w")
convert_source_entry = Entry(pdfconvert,width=70,textvariable=convert_source_folder_path)
convert_source_entry.place(x=70,y=40,anchor="w")
convert_source_button=Button(pdfconvert,text="Browse",command=radio_convert_pdf)
convert_source_button.place(x=500,y=40,anchor="w")

convert_destination_label = Label(pdfconvert, text="Destination")
convert_destination_label.place(x=5, y=70, anchor="w")
convert_destination_entry = Entry(pdfconvert,width=70,textvariable=convert_des_folder_path)
convert_destination_entry.place(x=70,y=70,anchor="w")
convert_destination_button=Button(pdfconvert,text="Browse",command=browse_convert_destination_button)
convert_destination_button.place(x=500,y=70,anchor="w")

convert_button = Button(tab2,text="  Convert Files  ",command=convert_to_word_excel)
convert_button.place(x=250,y=140,anchor="w")

tabcontrol.pack(expand="2", fill="both")

root.deiconify()
root.mainloop()
