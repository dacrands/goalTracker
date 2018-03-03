#! python3.
# -*- coding: utf-8 -*-
"""
Created on Sat Apr 15 16:19:20 2017

@author: dacrands
"""
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import date


# 
try:
    wb = load_workbook(filename = "goal-tracker.xlsx")
except IOError:
    print("Creating workbook")
    wb = Workbook()
    
    # Create sheet names (meditate, walk, read)
    walkSheet = wb.active
    walkSheet.title = "walk"
    walkSheet.cell(row=1, column=1).value  = 0
                  
    medSheet = wb.create_sheet(title="meditate")
    medSheet.cell(row=1, column=1).value  = 0
                 
    readSheet = wb.create_sheet(title="read")
    readSheet.cell(row=1, column=1).value = 0
                  
    wb.save("goal-tracker.xlsx")
    

def chooseSheet():
    leaveSheet = False
    while not leaveSheet:
        print("Choose an activity: ")
        for i in range(len(wb.sheetnames)):
            print("{0}: {1}\r".format(i, wb.sheetnames[i]))
            
        sheetChoice = input(">>>")        
        try:
            sheetChoice = int(sheetChoice)
            currSheet = wb.sheetnames[sheetChoice]
            print(currSheet)
            leaveSheet = True
            return currSheet 
           
        except:
            print("That's not a valid option")


leave = False
while not leave:
    print("Press 'n' to update goal. Press 'q' to quit")
    choices = ['q', 'n']
    choice = input(">>>")
        
    try:
        if choice not in choices:
            raise ValueError()
            
        if choice == 'q':
            leave = True
            
        elif choice == 'n':
            sheetLeave = False
            while not sheetLeave:
                currSheet = chooseSheet()
                sheet = wb[currSheet]   
                print(sheet.max_row)
                time = input("How long were you doing said task? >>")
                
                try: 
                    int(time)
                    for rowNum in range(1, sheet.max_row+2):
                        if sheet.cell(row=rowNum, column=1).value is None:
                            sheet.cell(row=rowNum, column=1).value = str(date.today())
                            sheet.cell(row=rowNum, column=2).value = time    
                            sheetLeave = True
                            
                except ValueError:
                    print("Please enter an integer, representing minutes...REPRESENT, B!")
                
    except ValueError:
        print("That's not a valid option")


wb.save("goal-tracker.xlsx")


