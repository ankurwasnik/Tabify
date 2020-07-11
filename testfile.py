# -*- coding: utf-8 -*-
"""
@author:Aztec Acer Aspire E 15 Ankur Wasnik
"""

from tkinter import *
import  camelot
import os
from win32com.client import Dispatch

xl = Dispatch('Excel.Application')
from tkinter import messagebox
#function for opening the excel app
def gettable() :
    if filename.get()==None or pageentry.get()=="" or pdfpageentry.get()=="" :
        messagebox.showerror('Error' , "Invalid Responce ")
        return
    else:
        messagebox.showinfo('Aztec','Excel file is opening ')
    pdflist = list([i for i in range(1, int(pdfpageentry.get()) + 1)])
    pdfpages = "1"
    for i in range(1, len(pdflist)):
        pdfpages = pdfpages + ',' + str(pdflist[i])
    print(pdfpages)

    file = filename.get() + '.pdf'
    tables=camelot.read_pdf(file , password=passname.get(), pages=pdfpages)
    export = filename.get()+pageentry.get() + '.csv'  #here we export extracted table
    tables[int(pageentry.get())-1].to_csv(export)
    filename.delete(0, 'end')
    passname.delete(0, 'end')
    pageentry.delete(0, 'end')
    pdfpageentry.delete(0 , 'end' )

    abspath=os.path.abspath(export)
    print(abspath)
    wb=xl.Workbooks.Open(abspath)
    xl.Visible = True
#function to show table in terminal
def result () :
    if filename.get()==None or pageentry.get()=="" or pdfpageentry.get()=="" :
        messagebox.showerror('Error' , "Invalid Responce ")
        return
    #code to get string range from 1 to end
    pdfname = filename.get() + ".pdf"
    pdflist=list([ i for i in range(1,int(pdfpageentry.get())+1)])
    pdfpages= "1"
    for i in range( 1,len(pdflist) ) :
        pdfpages=pdfpages + ',' +str(pdflist[i])
    print(pdfpages +' pages to be read')

    tables = camelot.read_pdf(pdfname , pages= pdfpages )
    page=0
    try :
     if (int(pageentry.get()) == 0 ) :
         print('Sorry , we do not disclose table 0 .\n Try again \n Thank You ')
     if len(tables) < int( pageentry.get() ) :
        print('Sorry , we cannot display those table number. \n It is out of range ')
        info = 'There are only ' + str(len(tables)) + ' avialable tables in pdf  !'
        print(info)
        print(tables)

     elif int ( pageentry.get() ) <=len(tables) and int(pageentry.get()) >0  :
        page=int( pageentry.get() ) -1
        print(tables[page].df)
        print(tables)
        print(tables[page].parsing_report)
    except IndexError :
        info = 'There are only ' + str(len(tables)) + ' avialable tables in pdf  !'
        print(info)
        print(tables)
    finally:
        print(' \n Thank You ')
        messagebox.showinfo("Aztec" , 'Your response is submitted')


# start program
window = Tk()
window.title("Aztec")
window.config(background='#8ecaff')
window.geometry("890x650")

azteclabel =Label(window ,anchor=CENTER, text=" Tabify " , bg='#5bb2ff' ,width=20 ,height=2)
azteclabel.grid( row=1 , columnspan=4)
azteclabel.config(font=("Algerian", 56))

lbl = Label(window, text="Get table from pdf ",font = ( "Castellar" , 10), fg="blue", bg="yellow" ,  width=50 , height=2  )
lbl.grid(column=0, row=2 , columnspan=4)

filelbl=Label(window , text='  Enter file name  ', font = ( "Segoe UI" , 15),bg='#8ecaff').grid(row=4 , column=0 )
filename=Entry(window)
filename.grid(row=4, column=1 , pady=5 )
filename.focus_set() #set focus to label when window is open.

passlbl=Label(window , text='  Enter Password  ',bg='#8ecaff',font = (  "Segoe UI" , 15)).grid(row=5 , column=0 )
passname=Entry(window ,show='*' ) #show * when entered in passwordlabel
passname.grid(row=5, column=1 , pady=5 )

pdfpageslbl=Label(window , text='   Number of pages  ' ,bg='#8ecaff',font = (  "Segoe UI" , 15)).grid(row=6 , column=0 )
pdfpageentry=Entry(window)
pdfpageentry.grid(row=6, column=1,pady=5 )

pagelbl=Label(window , text='Enter table number ' , bg='#8ecaff' ,font = (  "Segoe UI" , 15) ).grid(row=7 , column=0)
pageentry=Entry(window)
pageentry.grid(row=7, column=1 ,pady=10 )

submitbtn=Button(window , text='Submit' , font = ( "fixedsys" , 8), command=result , activebackground='Green' , activeforeground='black' , height=1 , width=20)
submitbtn.grid(row=8,  columnspan = 5 ,pady=2)

tablebtn = Button(window , text="Get ExcelFile" , font= ("fixedsys", 8) , command=gettable , activebackground='Grey' , width=20 , height=1  , activeforeground='white')
tablebtn.grid(row=9, columnspan=5 ,pady=2)


#destroy function to exit label.
def destroy ( ) :
    messagebox.showinfo('Tabify' , 'press "ok" to Exit program  \n ')
    print("Tabify is closed .")
    window.destroy() #destroy main window

exitbtn= Button( window , text="Close" , font = ( "fixedsys" , 8), command=destroy , activebackground='#ff0f16'  , height=1 , width=20)
exitbtn.grid(row=10 , columnspan = 5 ,pady=2)

# start running
if __name__ == '__main__' :
    window.mainloop()
