import logging
import docx
from docx import Document
from docx.shared import RGBColor
from docx.shared import Inches
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from typing import Tuple

from tkinter import scrolledtext
from tkinter.scrolledtext import ScrolledText
import re
import copy
import time

# This libraries are for opening word document automatically
import os
import platform
import subprocess

# This library is for opening excel document automatically
import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt

import Targest2

import Targest
global scrolled_text_box

from tkinter import messagebox

import webbrowser


def GUI1():
   
    try:
        # Creates the gui
        window = Tk(className=' TARGEST v.1.4.1 ')
        # set window size #
        window.geometry("1000x750")
        window['background']='#009ce8'

        icon = PhotoImage(file='TARGEST.png')
        window.iconphoto(True, icon)

        # Create a style for the widgets
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10, "bold"), background="#b2d8ff", foreground="black")

         # button 1
        ttk.Button(window, text="Choose list of Documents", command=Targest2.generateReport, width = 22).place(x=146, y=10)

         # button 2
        global genRep
        genRep = ttk.Button(window, text="Generate Reports", state= DISABLED, command=Targest2.generateReport2, width = 22)
        genRep.place(x=340, y=10)

        # button 3
        global allTagsButton
        allTagsButton = ttk.Button(text="Open All Tags Table Report", state= DISABLED, command=Targest2.getDocumentTable, width = 30)

        # button 4
        global getDoc
        getDoc = ttk.Button(window, text="Open Child and Parent Tags Report", state= DISABLED, command=Targest2.getDocument, width = 30)

        # button 5
        global getOrphanDoc
        getOrphanDoc = ttk.Button(text="Open Orphan Tags Report", state= DISABLED, command=Targest2.getOrphanDocument, width = 30)

        # button 6
        global getChildlessDoc
        getChildlessDoc = ttk.Button(text="Open Childless Tags Report", state= DISABLED, command=Targest2.getChildlessDocument, width = 30)

        # button 7
        global getTBVdoc
        getTBVdoc = ttk.Button(text="Open TBV Word Report", state= DISABLED, command=Targest2.getTBV, width = 30)

        # button 8
        global getTBDdoc
        getTBDdoc = ttk.Button(text="Open TBD Word Report", state= DISABLED, command=Targest2.getTBD, width = 30)

        # button 9
        global getExcel
        getExcel = ttk.Button(text="Open Tags and Requirements Excel Report", state= DISABLED, command=Targest2.createExcel, width = 30)

        # button 10
        global getExcel2
        getExcel2 = ttk.Button(text="Open Relationship Trees Excel Report", state= DISABLED, command=Targest2.createExcel2, width = 30)
    #    getExcel2.place(x=620, y=185)
        
        # button 10
        global TreeDiagram
        TreeDiagram = ttk.Button(text="Create Family Trees", state= DISABLED, command =lambda: Targest.text3(window), width = 30)
    #     TreeDiagram.place(x=620, y=210)

        # button 11
        global Website
        Website = ttk.Button(text="Visit our Website", state= ACTIVE, command =lambda: open_website(), width = 30)
        Website.place(x=215, y=64)

        # button 11
        global button
        button = ttk.Button(text="End Program", command=lambda:[window.destroy(), Targest2.closeReports(), Targest2.closeExcelWorkbooks()], width = 30)
        button.place(x=215, y=40)

        # Create text widget and specify size.
        global Txt
        Txt = ScrolledText(window, wrap=tk.WORD, height = 45, width = 70)
        Txt.place(x=25, y=120)
        Txt.configure(bg='grey', fg='white')

        # Create a label for the developers
        labelDevs = Label(window, text="Developers:\nJan William Haug\nAdrian Bernardino\nStephania Rey", font=("Segoe UI", 10, "bold"), bg="#E5CCFF")
        labelDevs.place(x=690, y=670)
        labelDevs.config(borderwidth=2, relief="groove", padx=15, pady=5, fg="black")
        

        # Create ScrolledText widget
        scrolled_text_box = ScrolledText(window, wrap=tk.WORD, height=45, width=51)
        scrolled_text_box.place(x=566, y=70)
        scrolled_text_box.configure(bg='grey', fg='white') 

        # Load the image file
        global imageLogo
        imageLogo = PhotoImage(file="TARGEST3.png")

        # Create a label to display the image
        label2 = Label(window, image=imageLogo)
        label2.place(x=25, y=10)

        # Load the copy right image file
        global imageCopyRight
        imageCopyRight = PhotoImage(file="copyright.png")

        # Create a label to display the image
        label3 = Label(window, image=imageCopyRight)
        label3.place(x=692, y=674)
        label3.config(bg="#E5CCFF")

        msg3 = ('You need a text file with paths to your documents\n 1. Please choose your documents by clicking on \n    the "Choose list of Documents" button.\n 2. Once the documents are displayed, Click "Generate Reports"\n\n')
        Txt.insert(tk.END, msg3) #print in GUI

        # show a pop-up message
        #messagebox.showinfo("Welcome to TARGEST",  "Make sure you have closed all your previous Word Reports and Excel Reports, before running this application")
        messagebox.showinfo("Welcome to TARGEST",  "Make sure to save a text file with the paths to the documents you want to use, if you haven't already")
        
        def selection_changed(selection):
            if selection == "Open All Tags Table Report":
                allTagsButton.invoke()
            elif selection == "Open Child & Parent Tags Report":
                getDoc.invoke()
            elif selection == "Open Orphan Tags Report":
                getOrphanDoc.invoke()
            elif selection == "Open Childless Tags Report":
                getChildlessDoc.invoke()
            elif selection == "Open TBV Word Report":
                getTBVdoc.invoke()
            elif selection == "Open TBD Word Report":
                getTBDdoc.invoke()
            elif selection == "Open Tags & Requirements Excel Report":
                getExcel.invoke()
            elif selection == "Open Relationship Trees Excel Report":
                getExcel2.invoke()
            elif selection == "Create Family Trees":
                TreeDiagram.invoke()

        options = ["Open All Tags Table Report", "Open Child & Parent Tags Report", "Open Orphan Tags Report", "Open Childless Tags Report", "Open TBV Word Report", "Open TBD Word Report", "Open Tags & Requirements Excel Report", "Open Relationship Trees Excel Report", "Create Family Trees"]

        selected_option = StringVar()
        selected_option.set(options[0])

        dropdown = OptionMenu(window, selected_option, *options, command=selection_changed)
        dropdown.place(x=566, y=46)

    except Exception as e:
        # Log an error message
        logging.exception('main(): ERROR', exc_info=True)
    else:
        # Log a success message
        logging.info('main(): PASS')

        window.mainloop()


def open_website():
    webbrowser.open("https://targest-website.vercel.app/")