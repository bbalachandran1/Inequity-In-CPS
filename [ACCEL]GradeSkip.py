import pandas as pd
from openpyxl import load_workbook
from ordered_set import OrderedSet
from datetime import datetime
from styleframe import StyleFrame
from math import trunc

print("Imports Successful")

#Done 
def read_data(number, filename):
#    filename = "datasets/TestData"


    df = pd.read_excel(filename + ".xlsx")
    newDf = df.values.tolist()

    dataList = []
    for i in range(len(newDf)):
        dataList.append(newDf[i])

    return dataList

#unique school list (Done)
def getAllSchools(data):

    schoolIDList = []
    schoolNameList = []
    for i in range(len(data)):
        schoolIDList.append(data[i][1])
        schoolNameList.append(data[i][2])

    schoolSet = OrderedSet(schoolIDList)
    schoolIDList = list(schoolSet)

    schoolNameSet = OrderedSet(schoolNameList)
    schoolNameList = list(schoolNameSet)

    schoolList = []
    for i in range(len(schoolIDList)):
        schoolList.append([schoolIDList[i],schoolNameList[i]])

    return schoolList



#Done
def seperateSchool(data):
    schoolList = getAllSchools(data)
    school = 0
    rowList = []
    for i in range(len(data)):
        if(schoolList[school][0] == data[i][1]): 
            rowList.append(data[i])
            
        else:
            schoolList[school].append(rowList)
            school+=1
            rowList = []
            rowList.append(data[i])

    schoolList[school].append(rowList)
    
    return schoolList


def seperateGrade(sortedList):
    for school in range(len(sortedList)):
        gradeList=[[],[],[],[],[],[],[],[],[]]
        for i in range(len(sortedList[school][2])):
            if(sortedList[school][2][i][4]== "K"):
                gradeList[0].append(sortedList[school][2][i])
            elif(sortedList[school][2][i][4]== "1"):
                gradeList[1].append(sortedList[school][2][i])
            elif(sortedList[school][2][i][4]== "2"):
                gradeList[2].append(sortedList[school][2][i])
            elif(sortedList[school][2][i][4]== "3"):
                gradeList[3].append(sortedList[school][2][i])
            elif(sortedList[school][2][i][4]== "4"):
                gradeList[4].append(sortedList[school][2][i])
            elif(sortedList[school][2][i][4]== "5"):
                gradeList[5].append(sortedList[school][2][i])
            elif(sortedList[school][2][i][4]== "6"):
                gradeList[6].append(sortedList[school][2][i])
            elif(sortedList[school][2][i][4]== "7"):
                gradeList[7].append(sortedList[school][2][i])
            elif(sortedList[school][2][i][4]== "8"):
                gradeList[8].append(sortedList[school][2][i])
            else:
                print("Buji Mess up")

        sortedList[school][2]=gradeList

    return sortedList

def seperateSubject(sortedList):

    #school, constant,grade, row #, column #
#    print(sortedList[1][2][1])
    for school in range(len(sortedList)):
        for grade in range(len(sortedList[school][2])):
            subjectList=[[],[]]
            for i in range(len(sortedList[school][2][grade])):
                if(sortedList[school][2][grade][i][5]== "Mathematics"):
                    subjectList[0].append(sortedList[school][2][grade][i])
                elif(sortedList[school][2][grade][i][5]== "Reading"):
                    subjectList[1].append(sortedList[school][2][grade][i])        
#                else:
#                    print("Value of "+str(sortedList[school][2][grade][i][5])+", at Row "+str(i))

            sortedList[school][2][grade]=subjectList

    return sortedList

def writeToExcel(finalMatrix, year):
    dfList = []
    for i in range(len(finalMatrix)):
        matrix = []
        col = ['School',
                   'Grade',
                   'Subject',
                   '# of Students in the Grade',
                   'Mean Score of Students in the Grade',
                   '90th Percentile Score of Grade Above',
                   '50th Percentile Score of 2 Grades Above',
                   '90th Percentile Score of 2 Grades Above',
                   '# of Students Scoring Above 90% 1 Grade Higher',
                   '# of Students Scoring Above 50% 2 Grades Higher',
                   '# of Students Scoring above 90% 2 Grades Higher',
                   '95th Percentile Score Nationally',
                   '# of Students Scoring above 95% Nationally']

        for a in range(len(finalMatrix[i])):
            for b in range(len(finalMatrix[i][a])):
                matrix.append(finalMatrix[i][a][b])
            matrix.append([])
            matrix.append(col)

        df = pd.DataFrame (matrix, columns = ['School',
                                              'Grade',
                                              'Subject',
                                              '# of Students in the Grade',
                                              'Mean Score of Students in the Grade',
                                              '90th Percentile Score of Grade Above',
                                              '50th Percentile Score of 2 Grades Above',
                                              '90th Percentile Score of 2 Grades Above',
                                              '# of Students Scoring Above 90% 1 Grade Higher',
                                              '# of Students Scoring Above 50% 2 Grades Higher',
                                              '# of Students Scoring above 90% 2 Grades Higher',
                                              '95th Percentile Score Nationally',
                                              '# of Students Scoring above 95% Nationally'])
        df = StyleFrame(df)
        dfList.append(df)

    time = datetime.today().strftime('%Y-%m-%d')
    filename = "analyzedData/Analysis- HighPerformingStudentsof"+year+"- "+time+".xlsx"

    with StyleFrame.ExcelWriter(filename) as writer:  
        dfList[0].to_excel(writer, sheet_name='Fall', index = False)
        dfList[1].to_excel(writer, sheet_name='Winter', index = False)
        dfList[2].to_excel(writer, sheet_name='Spring', index = False)
    
    
def calculateStatistics(sortedList, a, b, c, nationalList,season):
    matrix = []
    gradeAboveList = []
    grade2AboveList = []
    dataList = []
    
    bypass = "No"
    row = 0
    #nationalList has indices [x][y][z] where:
    #x = subject (0 for reading, 1 for math)
    #y = season (0 for fall, 1 for winter, 2 for spring)
    #z = grade (0 for K, 1 for 1st, ..., 8 for 8th)
    
    
    #a = school
    #b = grade
    #c = subject

    x = (c+1)%2
    nat90percentile = nationalList[x][season][b]

        
    for i in range(len(sortedList[a][2][b][c])):
        dataList.append(sortedList[a][2][b][c][i][7])

         #Works
    if(sortedList[a][2][b][c] == []):
        matrix.append(sortedList[a][1])

        for grade in range(9):
            if grade == b:
                if grade == 0:
                    grade = "K"
                else:
                    grade = int(grade)
                matrix.append(grade)
        if(c == 1):
            matrix.append("Reading")
        else:
            matrix.append("Mathematics")
            
    else:
        matrix.append(sortedList[a][2][b][c][0][2])
        matrix.append(sortedList[a][2][b][c][0][4])
        matrix.append(sortedList[a][2][b][c][0][5])
    

    if b<8:
        for i in range(len(sortedList[a][2][b+1][c])):
            gradeAboveList.append(sortedList[a][2][b+1][c][i][7])
   
    if b < 7:
        for i in range(len(sortedList[a][2][b+2][c])):
            grade2AboveList.append(sortedList[a][2][b+2][c][i][7])

    else:
        bypass = "Yes"

    if dataList == []:
        matrix.append("N/A")
        matrix.append("N/A")
    else:
        matrix.append(len(dataList))
        matrix.append(trunc(sum(dataList)/len(dataList))+1)

   
    if gradeAboveList == []:
        matrix.append("N/A")
        matrix.append("N/A")
        matrix.append("N/A")
        matrix.append("N/A")
        matrix.append("N/A")
        matrix.append("N/A")
        if dataList ==[]:
            matrix.append(nat90percentile)
            matrix.append("N/A")
        else:
            matrix.append(nat90percentile)
            matrix.append(countOccurrences(dataList, nat90percentile))

    elif grade2AboveList == []:
        if gradeAboveList == []:
            matrix.append("N/A")
        else:
            index90 = trunc(9*len(gradeAboveList)/10)
            percentile90 = sorted(gradeAboveList)[index90]

            matrix.append(percentile90)
            matrix.append("N/A")
            matrix.append("N/A")
                  
            matrix.append(countOccurrences(dataList, percentile90))
            matrix.append("N/A")
            matrix.append("N/A")

            if dataList ==[]:
                matrix.append(nat90percentile)
                matrix.append("N/A")
            else:
                matrix.append(nat90percentile)
                matrix.append(countOccurrences(dataList, nat90percentile))


    else:
        if bypass == "No":
            if b < 7:
                index90 = (trunc(9*len(gradeAboveList)/10))
                percentile90 = sorted(gradeAboveList)[index90]
                matrix.append(percentile90)

                mean = trunc(sum(grade2AboveList)/len(grade2AboveList))+1
                matrix.append(mean)

                index90Two = (trunc(9*len(grade2AboveList)/10))
                percentile90Two = sorted(grade2AboveList)[index90Two]
                matrix.append(percentile90Two)
                
                matrix.append(countOccurrences(dataList, percentile90))
                matrix.append(countOccurrences(dataList, mean))
                matrix.append(countOccurrences(dataList, percentile90Two))

                if dataList ==[]:
                    matrix.append(nat90percentile)
                    matrix.append("N/A")
                else:
                    matrix.append(nat90percentile)
                    matrix.append(countOccurrences(dataList, nat90percentile))

        elif bypass == "Yes":
            if b == 7:
                index90 = trunc(9*len(gradeAboveList)/10)
                percentile90 = sorted(gradeAboveList)[index90]

                matrix.append(percentile90)

                matrix.append("N/A")

                matrix.append("N/A")
                          
                matrix.append(countOccurrences(dataList, percentile90))

                matrix.append("N/A")

                matrix.append("N/A")

                if dataList ==[]:
                    matrix.append(nat90percentile)
                    matrix.append("N/A")
                else:
                    matrix.append(nat90percentile)
                    matrix.append(countOccurrences(dataList, nat90percentile))


            else:
                matrix.append("N/A")
                matrix.append("N/A")
                matrix.append("N/A")
                matrix.append("N/A")
                matrix.append("N/A")
                matrix.append("N/A")

                if dataList ==[]:
                    matrix.append(nat90percentile)
                    matrix.append("N/A")
                else:
                    matrix.append(nat90percentile)
                    matrix.append(countOccurrences(dataList, nat90percentile))


    return matrix


def countOccurrences(dataList, num):
    counter = 0
    for i in range(len(dataList)):
        if(dataList[i]>num):            #Could change to >= if needed
            counter += 1

    return counter

def runStatistics(sortedList,year, season):
    bigMatrix = []
    nationalList = readNationalAverage(year)
    for schools in range(len(sortedList)):
        schoolMatrix = []
        for i in range(len(sortedList[schools][2])):
            for k in range(len(sortedList[schools][2][i])):
                schoolMatrix.append(calculateStatistics(sortedList, schools, i, k, nationalList, season))

        bigMatrix.append(schoolMatrix)

    return bigMatrix

def readNationalAverage(year):
    year = "2015"
    df = pd.read_excel("datasets/"+year+ " NWEA MAP Student Norms 95 percentile RIT scores 211022.xlsx")
    
    newDf = df.values.tolist()
    subjectList = []
    seasonList = []
    gradeList = []

    subjectList = []
    for b in range(2):
        x = 5*b+2
        seasonList = []
        for p in range(3):
            row = x + p
            gradeList = []
            for i in range(9):
                gradeList.append(newDf[row][i+1])
            seasonList.append(gradeList)
        subjectList.append(seasonList)
    
    return subjectList

def generateGraphs(sortedList, year):
    writeMatrix = []
    for i in range(3):
        finalMatrix = runStatistics(sortedList[i],year, i)
        writeMatrix.append(finalMatrix)


    writeToExcel(writeMatrix, year)

def main(num):
    mainList = []
    for season in ["Fall", "Winter", "Spring"]:
        print("Working on "+num+" "+season)
        
        filename = "datasets/"+num+"/FOIA_REQ_NWEA_" + num + season
        testNumber = 0
        data = read_data(testNumber, filename)
        print("Data inserted into lists")
        
        sortedList = seperateSchool(data)
        print("List sorted by school")

        sortedList = seperateGrade(sortedList)
        print("List sorted by Grade")
        
        sortedList = seperateSubject(sortedList)
        print("List sorted by Subject \n")

        mainList.append(sortedList)

    year = num
    generateGraphs(mainList, year)
    print("Done with analysis of "+year)

yearList = ["2017", "2018", "2019", "2020"]
for i in yearList:
    main(i)
