#! python3.
# -*- coding: utf-8 -*-
"""
Created on Sat Apr 15 16:19:20 2017

@author: dacrands
"""
#from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import date


## Create the wb
#wb = Workbook()
## Create sheet names (meditate, walk, read)
#walkSheet = wb.active
#walkSheet.title = "walk"
#
#medSheet = wb.create_sheet(title="meditate")
#
#readSheet = wb.create_sheet(title="read")
#
#wb.save("goal-tracker.xlsx")

# Load the wb
wb = load_workbook(filename = "goal-tracker.xlsx")


currSheet = str()

# 1. Prompt user for activity (perhaps: "Are you sure?")
def chooseColor():
    while True:
        print("Choose an activity: ")
        for i in range(len(wb.sheetnames)):
            print("{0}: {1}\r".format(i, wb.sheetnames[i]))
        sheetChoice = input(">>>")
        try:
            sheetChoice = int(sheetChoice)
            currSheet = wb.sheetnames[sheetChoice]
            print(currSheet)
            return currSheet
            break
        except:
            print("That's not a valid option")

currSheet = chooseColor()
sheet = wb[currSheet]   
print(sheet.max_row)
# 2. Open corresponding sheet, prompt for data (perhaps: "Is the data accurate?")

while True:
    time = input("How long were you doing said task? >>")
    try: 
        int(time)
        break
    except ValueError:
        print("Please enter an integer, representing minutes...REPRESENT, B!")
        
for rowNum in range(1, sheet.max_row+2):
    if sheet.cell(row=rowNum, column=1).value is None:
        sheet.cell(row=rowNum, column=1).value = str(date.today())
        sheet.cell(row=rowNum, column=2).value = time

   

wb.save("goal-tracker.xlsx")
# 3. "Anything else" (if yes, go back to step one)

