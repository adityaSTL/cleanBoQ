#Importing all the libraries
from tkinter import *
from tkinter import filedialog
from sys import exit
import pandas as pd
import numpy as np
import os
from extract import Extract
from create_table1 import Create_table
import openpyxl
import matplotlib.pyplot as plt
from openpyxl_image_loader import SheetImageLoader
import os
import os.path
import warnings
from array import *
import math
warnings.simplefilter(action='ignore', category=FutureWarning)
pd.options.mode.chained_assignment = None
from boq_checker import check
#from boq_calc import calculate
from boq_calc import filler

from summary_calc import summarycalc


file_format=""
file_mandal=""

global logs
logs=''
logs= logs+"Starting"+"\n"
popupRoot = Tk()

popupRoot.title('BoQ Generator A')    


def boq_format():
    global file_format
    t=filedialog.askopenfilename(initialdir="C:/")
    file_format=file_format+t
    global logs
    logs= logs+ "Captured BoQ Format Folder " + file_format + "\n"

def mandal_folder(): 
    global file_mandal   
    t=filedialog.askdirectory()
    file_mandal=file_mandal+t
    global logs
    logs= logs + "Captured BoQ Mandal Folder " + file_mandal + "\n"


def boq_checkup():
    global logs
    logs= logs + "Called BoQ Checkup Function "  + "\n"
    top= Toplevel(popupRoot)
    top.geometry("980x550")
    l = Label(top, text = "Here is a list of Errors!!")
    l.config(font =("Courier", 14))
    b2 = Button(top, text = "Exit",command = top.destroy)
    top.title("Error Check-up Window")
    text_box = Text(top,height=50,width=200)
    text_box.pack(expand=True)
    try:
        temp_logs=''
        superstring,temp_logs=check(file_mandal)
        logs= logs + temp_logs  + "\n" 
    except:
        logs= logs + "Red Flag: Could not complete BoQ Checker Function. Issues Found!!"  + "\n" 
        superstring="Not able to do a BoQ Checkup on given Mandal!!"   
    
    text_box.insert('end', superstring)
    text_box.config(state='disabled')
    l.pack()
    b2.pack()
    logs= logs + "Work in BoQ Checkup function done. Closed! "  + "\n"

def boq_calc(): 

    global logs
    logs= logs + "Called BoQ Calculate Function "  + "\n"   
    
    
    try:
        temp_logs=''
        temp_logs=filler(file_mandal,file_format)
        logs= logs + temp_logs  + "\n" 
    except:
        logs= logs + "Red Flag: Could not complete BoQ Calculate Function. Issues Found!!"  + "\n" 
      
    
    #print(logs)
    popupRoot.destroy()

def summary():
    global logs
    logs= logs + "Called BoQ Summary Function "  + "\n"
    temp_logs=''
    temp_logs=summarycalc(file_mandal)
    logs= logs + temp_logs  + "\n" 
    """
    try:
        temp_logs=''
        temp_logs=summarycalc(file_mandal)
        logs= logs + temp_logs  + "\n" 
    except:
        logs= logs + "Red Flag: Could not Fill-up Summary. Issues Found!!"  + "\n" 
    """    
    print(1)    


popuplabel = Label(popupRoot, text = 'BoQer Code V2.1',font = ("Times New Roman", 11)).grid(row=0,column=2)
popupButton = Label(popupRoot, text = 'STL',bg='white',fg='blue',font = ("Times New Roman", 13,'bold'), anchor="e",justify=RIGHT).grid(row=0,column=3)
popuplabel = Label(popupRoot, text = ' ',font = ("Times New Roman", 12)).grid(row=1,column=1)
popuplabel = Label(popupRoot, text = 'Select the BoQ Format',font = ("Times New Roman", 12)).grid(row=2,column=1)
popupButton = Button(popupRoot, text = 'BoQ Format', font = ("Times New Roman", 12), command = boq_format,width = 20).grid(row=2,column=2)
popuplabel = Label(popupRoot, font = ("Times New Roman", 12),text = 'Select Mandal Folder').grid(row=3,column=1)
popupButton = Button(popupRoot, text = 'Mandal Folder', font = ("Times New Roman", 12), command = mandal_folder,width = 20).grid(row=3,column=2)
popuplabel = Label(popupRoot, text = ' ',font = ("Times New Roman", 12)).grid(row=4,column=1)
popupButton = Button(popupRoot, text = 'Run BoQ Script', bg='light green',font = ("Times New Roman", 15), command = boq_calc,width = 12).grid(row=5,column=3)
popupButton = Button(popupRoot, text = 'Error Checkup', bg='orange',font = ("Times New Roman", 15), command = boq_checkup,width = 12).grid(row=5,column=1)
popupButton = Button(popupRoot, text = 'Summary fillup', bg='light blue',font = ("Times New Roman", 15), command = summary,width = 12).grid(row=5,column=2)
popupRoot.geometry('500x200')
mainloop()