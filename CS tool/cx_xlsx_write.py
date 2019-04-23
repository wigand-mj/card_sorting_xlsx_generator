# -*- coding: utf-8 -*-
"""
Created on Tue Jun 12 17:48:02 2018

@author: Maik Wigand
"""

import sys,xlsxwriter,configparser,csv
import numpy as np
from time import time
from time import sleep
import random
from tkinter import *
from tkinter import messagebox
from pygame.locals import *
from pygame.compat import unichr_, unicode_

parlist = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY"]
parlist_small = ["a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z"]
newlist = list()
item_list = list()
horizontal_items = {}
vertical_items = {}

def main():
    STATE = "READ"
    while True:
        if STATE == "READ":
            try:
                item_list = read_items()
                STATE = "INIT"
            except Excpetion as e: 
                print("Couldn't Read Data:" + e)
                error_msg(e)
        if STATE == "INIT":
            try:
                horizontal_items = initialize(item_list, 0)
                vertical_items = initialize(item_list, 1)
                STATE = "WRITE"
            except Exception as e:
                print("Couldn't Initialize:" + e)
                error_msg(e)
        if STATE == "WRITE":       
            try:
                write_xlsx(item_list, read_data(), vertical_items, horizontal_items)
                STATE = "END"
            except Exception as e:
                print("Couldn't Write Data")
                print(e)
                error_msg(e)
        if STATE == "END":           
            sys.exit(0)
            
def initialize(item_list, x):
    new_dict = {}
    if x == 0:
        for i in range(len(item_list)):
            if item_list[i] in new_dict:
                pass
            else:
                new_dict[item_list[i]] = parlist[i+1]
    elif x == 1:
        for i in range(len(item_list)):
            if item_list[i] in new_dict:
                pass
            else:
                new_dict[item_list[i]] = i+2
    return new_dict
def read_data():
    with open ("proximity.txt", "r") as f:
        proximity_data = f.read().replace('\n','')
        print(proximity_data)

        
        return proximity_data
def read_items():
    with open("item_list.csv", "r") as csv_file:
        csv_reader = csv.reader(csv_file)
        #print(csv_reader[1])
        for line in csv_reader:
            for x in line:
                newlist.append(x)
    return newlist
          
def write_xlsx(item_list, proximity_data, vertical_list, horizontal_list):
   prox_list = list()
   number = False
   number_get=False
   c=""
   num=""
   workbook = xlsxwriter.Workbook("cs_data_pX.xlsx")
   counter = 0
   worksheet = workbook.add_worksheet()
   for x in parlist:
       command = "worksheet.set_column(" + "'" +x+":"+x+"'"+", 10)"
       eval(command)

   for i in range((len(item_list))):
       command2 = "worksheet.write(" + "'" + parlist[i+1] + "1', item_list[" + str(i) + "])"
       eval(command2)
   for i in range((len(item_list))):
      command2 = "worksheet.write(" + "'A" + str(i+2) + "', item_list[" + str(i) + "])"
      eval(command2)
   for i in range((len(item_list))):
       for j in range((len(item_list))):
           command2 = "worksheet.write(" + "'" + parlist[i+1] + str(j+2) + "', 0)"
           eval(command2)
           
           
   for char in proximity_data:
       #print(char)
       if number_get==True:
           if char != "-":
               num = num + str(char)   
               print(num)
           if char == "-":
               n = float(num)
               print(n)
               number=True
               number_get=False
               
       elif number == False:
           if char != ";":
               if char != ",":
                   c = c+char
                   print(c)
               elif char == ",":
                   prox_list.append(c)
                   print(prox_list)
                   c=""
           elif char ==";":
                prox_list.append(c)
                print(prox_list)
                c=""
                number_get=True
       if number == True:
           #print(prox_list)
           #print(n)
           for y in range(len(prox_list)):
               for x in range(len(prox_list)):
                  # if prox_list[y] == prox_list[x]:
                   #    print("2")
                    #   break
                   #(prox_list[y])
                   #print(prox_list[x])
                   if prox_list[y] != prox_list[x]:
                       #print("hallo")
                       s = horizontal_list[prox_list[y]] + str(vertical_list[prox_list[x]])
                       #print(s)
                       worksheet.write(s, n)
                       print("write")
           num=""
           n=0
           prox_list = list()
           number=False
           
   workbook.close()
def error_msg(e):
    file = open("error.txt", "w")
    file.write(e)
    file.close
    
main()