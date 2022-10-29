from tkinter import *
from sys import exit
from tkinter import filedialog
import pandas as pd
import numpy as np
import os
from extract import Extract
from create_table1 import Create_table
import openpyxl
import matplotlib.pyplot as plt
from openpyxl_image_loader import SheetImageLoader
import os
import tkinter
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
#Parth Pandey12:12 PM
pd.options.mode.chained_assignment = None

def append_new_line(file_name, text_to_append):
    """Append given text as a new line at the end of file"""
    # Open the file in append & read mode ('a+')
    with open(file_name, "a+") as file_object:
        # Move read cursor to the start of file.
        file_object.seek(0)
        # If file is not empty then append '\n'
        data = file_object.read(100)
        if len(data) > 0:
            file_object.write("\n")
        # Append text at the end of file
        file_object.write(text_to_append)

#file2=r'C:\Users\Aditya.gupta\Downloads\BoQ_Format.xlsx'
file1=r'C:\Users\Aditya.gupta\Desktop\Mandal'

mandal_file=file1+'/'
print("Grabbed mandal location")
no=0
os.chdir(mandal_file)
for info in os.listdir():
    j=mandal_file+info
    print(j)
    blo=Extract.extract_blo(j)
    if len(blo)==0:
        e1=info
        append_new_line('sample3.txt', e1)
        #(blo,joint_closer)=Create_table.create_blo(blo)
        #joint_closer.to_csv('Joint_closure.csv')
    #blo.to_csv('blwoing.csv')
    #print("pas printing")
    drt=Extract.extract_drt(j)
    if len(drt)!=0:
        drt=Create_table.create_drt(drt)
        drt.reset_index(drop=True,inplace=True)
    #print("empty")
    #print(drt)
    #drt.to_csv('drt12.csv')
    ot=Extract.extract_ot(j)
    if len(ot)!=0:
        ot=Create_table.create_ot(ot)   
    hdd=Extract.extract_hdd(j)
    if len(hdd)!=0:
        hdd=Create_table.create_hdd(hdd)      
    #print(joint_closer)
    s=0
    fin=0
    temp=0
    no+=1