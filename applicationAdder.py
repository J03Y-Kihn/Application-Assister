
from openpyxl import load_workbook

DEFAULT_FILE = "Applications Applied.xlsx"
#DEFAULT_PATH = "C:\Users\jokih\Documents\Programming\Personal Projects\Application Excel\Applications Applied.xlsx"
#test = "C:\Users\jokih\Documents\Programming\Personal Projects\Application Excel\ Applications Applied.xlsx"
TITLE_ROW = 1   #change if your titles are not in this row
STARTING_COLUMN = 2 #change if your format is different
STARTING_ROW = 2    #change if your format is different


def addRow(file):


    #load excel file
    workbook = load_workbook(file)

    #open workbook
    sheet = workbook.active

    currColumn = STARTING_COLUMN

    currRow = STARTING_ROW
    currCell = sheet.cell(currRow, currColumn)
    #find the lowest empty row, since the maximum row may not be correct
    while(currCell.value):
        currRow+=1
        currCell = sheet.cell(currRow, currColumn)
    

    updateRow = sheet.max_row+1
    if(updateRow > currRow):
        updateRow = currRow
        
    while(currColumn<sheet.max_column+1):
        print("Enter the", sheet.cell(row= TITLE_ROW, column= currColumn).value)
        currInput = input()
        currCell = sheet.cell(updateRow, currColumn)
        currCell.value = currInput

        currColumn+=1

    #save file
    workbook.save(file)

def updateRow(file):

    #load excel file
    workbook = load_workbook(file)
    #open workbook
    sheet = workbook.active

    currColumn = STARTING_COLUMN
    updateRow = sheet.max_row+1
    while(currColumn<sheet.max_column+1):
        print("Enter the", sheet.cell(row= TITLE_ROW, column= currColumn).value)
        currInput = input()
        currCell = sheet.cell(updateRow, currColumn)
        currCell.value = currInput

        currColumn+=1

    #save file
    workbook.save(file)

#def viewRow(file):
    

def numAppsComp(file):
    #load excel file
    workbook = load_workbook(file, data_only=True)
    #open workbook
    sheet = workbook.active
    #Print value in first cell, which is the number of rows written to (number of applications)
    print(sheet.max_row -1)


def main():
    default = False
    print("Welcome to the Application Adder!! \nIf adding to a file that is not", DEFAULT_FILE, "please type it now, otherwise hit enter.")
    x = input()
    filenames = ""
    if(x != ""):
        print(x)
        filenames = x + ".xlsx"
    else:
        print("Default File")
        filenames = DEFAULT_FILE

    
    inputs = ""
    while(inputs.upper() != "0"):
        print("NOTE!! If file is open, no additions can be added only viewing!")
        print("1: Add new row to spreadsheet? \n2: Update existing row? (Not Implemented) \n3: View Row Data? (Not Implemented) \n4: View Number of Applications \n0: Exit Program")
        inputs = input()
        if(inputs == "1"):
            addRow(filenames)
            print("Added Complete")
        elif(inputs == "2"):
            #updateRow(filenames)
            print("Update Completed")
        elif(inputs == "3"):
            #viewRow(filenames)
            print("Viewed Row")
        elif(inputs == "4"):
            print("Viewed Total Number")
            numAppsComp(filenames)

    return


if __name__ == "__main__":
    main()
    print("Completed")