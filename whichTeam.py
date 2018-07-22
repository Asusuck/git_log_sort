import openpyxl

def makeTeamData():
    #找出team資料表
    wb = openpyxl.load_workbook('gitLog.xlsx')
    sheet = wb.get_sheet_by_name('Count')
    
    #將team資料輸入tempArray
    tempArray=[]
    teamArray=[]
    
    for row_cell in sheet['A2':'B'+str(sheet.max_row-1)]:
        for cell in row_cell:
            """print(cell.value)"""
            tempArray.append(cell.value)
        teamArray.append(tempArray)
        tempArray=[]
    return teamArray


#比對function
def whichTeam(memberName):
    teamArray=makeTeamData()
    for team in teamArray:
        if team[1] in memberName:
            return team[0]
    return 'Others'

