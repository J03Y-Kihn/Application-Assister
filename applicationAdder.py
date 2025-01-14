
from openpyxl import load_workbook

DEFAULT_FILE = "Applications Applied.xlsx"
TITLE_ROW = 1   #change if your titles are not in this row
STARTING_COLUMN = 2 #change if your format is different
STARTING_ROW = 2    #change if your format is different

#add a row to the file at the highest unused location
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
    workbook.close()

#not implemented will update the input row
def updateRow(file):
    """
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
    workbook.close()
    """

#outputs all information in a given row
def viewRow(file):
    #load excel file
    workbook = load_workbook(file)
    #open workbook
    sheet = workbook.active

    #Determine if the input Company or Position is Valid
    inputs = ""
    print("Which Company or Position would you like to view?")
    inputs = input()
    posRow = getRow(file, inputs)
    while(not posRow):
        print("Row is not in file! Please double check your spelling and try again or type 'back' to return to menu.")
        inputs = input()


        if(inputs.lower() == "back"):
            return
        else:
            posRow = getRow(file, inputs)
        
    currColumn = STARTING_COLUMN
    for values in posRow:
        while(currColumn<sheet.max_column+1):
            print(sheet.cell(row= TITLE_ROW, column= currColumn).value + ":", sheet.cell(row= values, column = currColumn).value)
            currColumn+=1
        print("")
        currColumn = STARTING_COLUMN

    print("Number of applications found:", len(posRow))
    #save file
    workbook.save(file)
    workbook.close()

#returns the row number if a row in the file has the same name as position in select columns
def getRow(file, position):

    #load excel file
    workbook = load_workbook(file)
    #open workbook
    sheet = workbook.active

    rows = []

    companyCol = STARTING_COLUMN
    positionCol = STARTING_COLUMN+1
    currRow = STARTING_ROW

    currCompanyCell = sheet.cell(currRow, companyCol)
    currPositionCell = sheet.cell(currRow, positionCol)
    while(currCompanyCell.value and currPositionCell.value):
        if(currCompanyCell.value.lower() == position.lower() or currPositionCell.value.lower() == position.lower()):
            rows.append(currRow)

        currRow+=1
        currCompanyCell = sheet.cell(currRow, companyCol)
        currPositionCell = sheet.cell(currRow, positionCol)
        
        
    workbook.close()
    if(len(rows) > 0):
        return rows
    else:
        return False

#outputs the number of applications(rows) in the file
def numAppsComp(file):
    #load excel file
    workbook = load_workbook(file, data_only=True)
    #open workbook
    sheet = workbook.active
    #Print value in first cell, which is the number of rows written to (number of applications)
    print(sheet.max_row -1)
    workbook.close()


def main():
    default = False
    print("Welcome to the Application Adder!! \nIf adding to a file that is not", DEFAULT_FILE, "please type it now, otherwise hit enter.")
    x = input()
    filenames = ""
    if(x != ""):
        filenames = x + ".xlsx"
    else:
        filenames = DEFAULT_FILE

    
    inputs = ""
    while(inputs.upper() != "0"):
        print("NOTE!! If file is open, no additions can be added only viewing!")
        print("1: Add new row to spreadsheet? \n2: Update existing row? (Not Implemented) \n3: View Row Data? \n4: View Number of Applications \n0: Exit Program")
        inputs = input()
        if(inputs == "1"):
            addRow(filenames)
        elif(inputs == "2"):
            #updateRow(filenames)
            print("Update Completed")
        elif(inputs == "3"):
            viewRow(filenames)
        elif(inputs == "4"):
            numAppsComp(filenames)

    return


if __name__ == "__main__":
    main()
    print("Good Night!")
