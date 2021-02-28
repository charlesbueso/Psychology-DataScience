import xlrd
import xlsxwriter
import os
import re
import time
from pathlib import Path

start_time = time.time()
#necessary hardcoded variables
teachers11 = ['611','612','613','614','621','622','631','632','633','634','641','642','643','651','652','653','654','661','662','671','681','682','683','691','692','693','694','6101','6102','6103']
teachers12 = ['611','612','613','614','621','622','631','632','633','641','642','643','651','652','653','661','662','671','681','682','683','691','692','693','694','695','6101','6102']
teachers13 = ['611','612','613','614','621','622','631','632','633','634','641','642','643','652','653','654','661','662','671','681','682','683','691','692','693','694','6101','6102']
teachers14 = ['612','613','614','621','622','631','633','634','641','642','643','651','653','654','662','671','681','682','683','691','692','693','694','695','6101','6102']
teachers21 = ['695','2611','2612','2621','2622','2631','2633','2641','2642','2643','2651','2661','2662','6103']
teachers22 = ['695','2611','2612_2','2621_2','2622_2','2631','2633','2641','2642','2643','2651','2661','2662','6103']
teachers23 = ['695','2611','2612','2621','2622','2631','2633','2641','2642','2643','2651','2661','2662','6103']
teachers24 = ['695','2611','2612','2621','2622','2631','2633','2641','2642','2643','2651','2661','2662','6103']

tchTypeTokensPath = r'//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 15 - 16/Completed Teacher Transcripts/Chicago/Round 1/stat.frq.xlsx'
                  
listOfTypesTokensDaleR1 = ['//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 15 - 16/Round 1 Completed Transcripts/Final Round 1 All Transcripts 7_24_19/Dale.frq_Round 1.xlsx', r'//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 16 - 17/WTG 16-17 Audio/Completed Student Transcripts/Round 1/Round 1 Compiled Transcripts 8_5_19/Dale.frq.xlsx', r'//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 15 - 16/Boston Recordings/Finished Student Transcripts/Round 1/Round 1 7_25_19/Dale.frq.xlsx', r'//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 16 - 17/Boston 16-17/Completed Student Transcripts/Round 1/Compiled Round 1 8_5_19/Dale.frq.xlsx']  
listOfTypesTokensNoDaleR1 = ['//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 15 - 16/Round 1 Completed Transcripts/Final Round 1 All Transcripts 7_24_19/NonDalestat.frq.xlsx', r'//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 16 - 17/WTG 16-17 Audio/Completed Student Transcripts/Round 1/Round 1 Compiled Transcripts 8_5_19/NoDale.frq.xlsx', r'//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 15 - 16/Boston Recordings/Finished Student Transcripts/Round 1/Round 1 7_25_19/NoDale.frq.xlsx', r'//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 16 - 17/Boston 16-17/Completed Student Transcripts/Round 1/Compiled Round 1 8_5_19/NoDale.frq.xlsx']    
listOfTypesTokensDaleR2 = ['//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 15 - 16/Boston Recordings/Finished Student Transcripts/Round 2/Round 2 7_25_19/Dale.frq.xlsx']
listOfTypesTokensNoDaleR2 = ['//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 15 - 16/Boston Recordings/Finished Student Transcripts/Round 2/Round 2 7_25_19/NoDale.frq.xlsx']
listOfTypesTokensDaleR3 = ['//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 15 - 16/Boston Recordings/Finished Student Transcripts/Round 3/Round 3 7_25_19/Dale.frq.xlsx']
listOfTypesTokensNoDaleR3 = ['//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 15 - 16/Boston Recordings/Finished Student Transcripts/Round 3/Round 3 7_25_19/NoDale.frq.xlsx']
listOfTypesTokensDaleR4 = ['//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 15 - 16/Boston Recordings/Finished Student Transcripts/Round 4/Round 4 Compiled 8_6_19/Dale.frq.xlsx']
listOfTypesTokensNoDaleR4 = ['//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Chicago_AudiosComplete_6thGrade_WTG 15 - 16/Boston Recordings/Finished Student Transcripts/Round 4/Round 4 Compiled 8_6_19/NoDale.frq.xlsx']

#create master dict with teachers, students, headings, and paths.
masterDict = {}

#list of all pathnames for all rounds in year 1 and 2
pathnamesForRounds = []
for i in range(4):
    firstPart = '//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Syntax Coding Files/Year 1 Student Syntax Files/Round ' 
    secondPart = str(i+1)
    firstPart2 = '//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Syntax Coding Files/Year 2 Student Syntax Files/round '
    pathnamesForRounds.append(firstPart + secondPart)
    pathnamesForRounds.append(firstPart2 + secondPart)
newOrder = [0,2,4,6,1,3,5,7]
pathnamesForRounds = [pathnamesForRounds[i] for i in newOrder]

#Switch statement to choose teachers.
def whatTeacher(i):
    return {
            1: teachers11,
            2: teachers12,
            3: teachers13,
            4: teachers14,
            5: teachers21,
            6: teachers22,
            7: teachers23,
            8: teachers24
    }[i]

def TCHTypeToken(teacherCode):
    try:
        wb = xlrd.open_workbook(tchTypeTokensPath)
        sheet = wb.sheet_by_index(0)
        sheet.cell_value(0,0)
        for i in range(22):
            if teacherCode in (str((sheet.cell_value(i, 0)))):
                if (teacherCode[0] != '2') and ('2' in (str((sheet.cell_value(i, 0))))):
                    continue
                else:
                    thisType = str(sheet.cell_value(i,4219))
                    thisToken = str(sheet.cell_value(i,4220))
                    return str(int(float(thisType))) + '/' + str(int(float(thisToken))) 
    except:
        return '0-'
    

#TESTING PASSED
#print(TCHTypeToken('611'))
    
#returns list with 4 different values in that order.
failedPathsForCHITOTUT = []
def chimultuChitotutStumultuStutotut(path):
    numberOfClausesColumn = 0
    numberOfTierColumn = 0
    fourValueList = []
    CHIMULTU = 0
    CHITOTUT = 0
    STUMULTU = 0
    STUTOTUT = 0
    
    try:
        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(0)
        sheet.cell_value(0,0)
         
        for i in range(20):
            if 'number of' in str(sheet.cell_value(0,i).lower()):
                numberOfClausesColumn = i
                break
             
        for i in range(20):
            if 'tier' in str(sheet.cell_value(0,i).lower()):
                numberOfTierColumn = i
                break
    
        for i in range(1000):
            checkForEnd = str(sheet.cell_value(i,1))
            checkForEnd2 = str(sheet.cell_value(i,2))
            if ('end' not in (checkForEnd.lower())) and ('end' not in (checkForEnd2.lower())):
                clause = (sheet.cell_value(i+1,numberOfClausesColumn))
                try:
                    int(clause)
                except:
                    continue
                tier = str(sheet.cell_value(i+1,numberOfTierColumn))
                 
                if ('CHI' in tier) and (clause > 1):
                    CHIMULTU = CHIMULTU + 1
                if ('CHI' in tier):
                    CHITOTUT = CHITOTUT + clause
                if ('STU' in tier) and (clause > 1):
                    STUMULTU = STUMULTU + 1   
                if ('STU' in tier):
                    STUTOTUT = STUTOTUT + clause
            else:
                break
         
        fourValueList.extend([str(int(CHIMULTU)), str(int(CHITOTUT)), str(int(STUMULTU)), str(int(STUTOTUT))])
        return fourValueList          
         
    except:
        failedPathsForCHITOTUT.append(path)
        fourValueInvalidList = ['0-','0-','0-','0-']
        return fourValueInvalidList

#TESTING PASSED
#print(chimultuChitotutStumultuStutotut('//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Syntax Coding Files/Year 2 Student Syntax Files/round 4/2662_4/266210_4_HD.xlsx'))

#returns a value found in the different paths of the list as parameter
def childTypeToken(listOfPaths, studentCode):
    fileColumn = 0
    typesColumn = 0
    validList = []
    invalidList = ['0-','0-','0-','0-']
    
    try:
        for i in list(listOfPaths):
            wb = xlrd.open_workbook(i)
            sheet = wb.sheet_by_index(0)
            sheet.cell_value(0,0)
            #print('first for loop: ' + i)
    
            for g in range(sheet.ncols):
                #print('second for loop: ' + str(g))
                if "Types" in str(sheet.cell_value(0,sheet.ncols-g-1)):
                    typesColumn = sheet.ncols-g-1
                else:
                    continue
                
            for j in range(sheet.nrows):  
                if typesColumn == 0:
                    return '-0'
                else:
                    #print('third for loop: ' + str(j))
                    fileToSplit = str(sheet.cell_value(j,fileColumn))
                    revisedFile = ''
                    if '_' in fileToSplit:
                        if '.' in fileToSplit:
                            fileToSplit = fileToSplit.split('.','')
                            
                        revisedFile = fileToSplit.split('_')[0]
    
                        if studentCode == revisedFile:
                            validList.append(str(int(float(sheet.cell_value(j,typesColumn)))))
                            validList.append(str(int(float(sheet.cell_value(j,typesColumn+1)))))
                            validList.append(str(int(float(sheet.cell_value(j+1,typesColumn)))))
                            validList.append(str(int(float(sheet.cell_value(j+1,typesColumn+1)))))
                            return validList
                    else:
                        if '.' in fileToSplit:
                            fileToSplit = fileToSplit.split('.','')
                        #maybe check here for . in file
                        for f in range(len(studentCode)):
                            revisedFile = revisedFile + fileToSplit[f] 
                        if studentCode == revisedFile:
                            validList.append(str(int(float(sheet.cell_value(j,typesColumn)))))
                            validList.append(str(int(float(sheet.cell_value(j,typesColumn+1)))))
                            validList.append(str(int(float(sheet.cell_value(j+1,typesColumn)))))
                            validList.append(str(int(float(sheet.cell_value(j+1,typesColumn+1)))))
                            return validList 
                        else:
                            continue
        
        return invalidList        
    except:
        return invalidList       


#TESTING PASSED
#print(childTypeToken(listOfTypesTokensDaleR4, '69115', 'CHITYPED'))

#check for all 14 complex code headings, add them to 14 element list and return
def complexCode(path):
    tierColumn = 0
    complexCodeColumn = 0
    fourteenValuesList = []
    CHICO = 0
    CHIS1 = 0 
    CHIS2 = 0
    CHISC = 0
    CHIOC = 0
    CHISRC = 0
    CHIORC = 0
    STUCO = 0
    STUS1 = 0
    STUS2 = 0
    STUSC = 0
    STUOC  = 0
    STUSRC = 0
    STUORC = 0
    
    try:
        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(0)
        sheet.cell_value(0,0)
        
        for i in range(20):
             if 'tier' in (str(sheet.cell_value(0,i).lower())):
                 tierColumn = i
                 break
             
        for i in range(20):     
             if 'code' in (str(sheet.cell_value(0,i).lower())):
                 complexCodeColumn = i 
                 break
        
        for j in range(500):
            if ('end' not in (str(sheet.cell_value(j,1)).lower())) and ('end' not in (str(sheet.cell_value(j,2)).lower())):
                if 'CO' in str(sheet.cell_value(j,complexCodeColumn)):
                    if 'CHI' in str(sheet.cell_value(j,tierColumn)):
                        CHICO = CHICO + (str(sheet.cell_value(j,complexCodeColumn))).count('CO')
                    if 'STU' in str(sheet.cell_value(j,tierColumn)):
                        STUCO = STUCO + (str(sheet.cell_value(j,complexCodeColumn))).count('CO')   
                if 'S1' in str(sheet.cell_value(j,complexCodeColumn)):
                    if 'CHI' in str(sheet.cell_value(j,tierColumn)):
                        CHIS1 = CHIS1 + (str(sheet.cell_value(j,complexCodeColumn))).count('S1')
                    if 'STU' in str(sheet.cell_value(j,tierColumn)):
                        STUS1 = STUS1 + (str(sheet.cell_value(j,complexCodeColumn))).count('S1')    
                if 'S2' in str(sheet.cell_value(j,complexCodeColumn)):
                    if 'CHI' in str(sheet.cell_value(j,tierColumn)):
                        CHIS2 = CHIS2 + (str(sheet.cell_value(j,complexCodeColumn))).count('S2')
                    if 'STU' in str(sheet.cell_value(j,tierColumn)):
                        STUS2 = STUS2 + (str(sheet.cell_value(j,complexCodeColumn))).count('S2')
                if 'SC' in str(sheet.cell_value(j,complexCodeColumn)):
                    if 'CHI' in str(sheet.cell_value(j,tierColumn)):
                        CHISC = CHISC + (str(sheet.cell_value(j,complexCodeColumn))).count('SC')
                    if 'STU' in str(sheet.cell_value(j,tierColumn)):
                        STUSC = STUSC + (str(sheet.cell_value(j,complexCodeColumn))).count('SC')   
                if 'OC' in str(sheet.cell_value(j,complexCodeColumn)):
                    if 'CHI' in str(sheet.cell_value(j,tierColumn)):
                        CHIOC = CHIOC + (str(sheet.cell_value(j,complexCodeColumn))).count('OC')
                    if 'STU' in str(sheet.cell_value(j,tierColumn)):
                        STUOC = STUOC + (str(sheet.cell_value(j,complexCodeColumn))).count('OC')    
                if 'SRC' in str(sheet.cell_value(j,complexCodeColumn)):
                    if 'CHI' in str(sheet.cell_value(j,tierColumn)):
                        CHISRC = CHISRC + (str(sheet.cell_value(j,complexCodeColumn))).count('SRC')
                    if 'STU' in str(sheet.cell_value(j,tierColumn)):
                        STUSRC = STUSRC + (str(sheet.cell_value(j,complexCodeColumn))).count('SRC')    
                if 'ORC' in str(sheet.cell_value(j,complexCodeColumn)):
                    if 'CHI' in str(sheet.cell_value(j,tierColumn)):
                        CHIORC = CHIORC + (str(sheet.cell_value(j,complexCodeColumn))).count('ORC')
                    if 'STU' in str(sheet.cell_value(j,tierColumn)):
                        STUORC = STUORC + (str(sheet.cell_value(j,complexCodeColumn))).count('ORC')    
            else:
                break
            
        fourteenValuesList.extend([str(CHICO),str(CHIS1),str(CHIS2),str(CHISC),str(CHIOC),str(CHISRC),str(CHIORC),str(STUCO),str(STUS1),str(STUS2),str(STUSC),str(STUOC),str(STUSRC),str(STUORC)])
        return fourteenValuesList
    except:
        listInvalid = ['0-','0-','0-','0-','0-','0-','0-','0-','0-','0-','0-','0-','0-','0-']
        return listInvalid

#TESTING PASSED
#print(complexCode('//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Syntax Coding Files/Year 2 Student Syntax Files/round 3/2612_3/26124_RD.xlsx'))

#Create list of subdirectories on each round, list of teachers
dictOfSubDirs = {}
for i in pathnamesForRounds:
    currPath = Path(i)
    directoryContents = os.listdir(currPath)
    list1 = []
    list1 = currPath.parts
    a = str(list1[5])
    b = str([int(f) for f in a.split() if f.isdigit()])
    c = str(list1[6])
    d = str([int(f) for f in c.split() if f.isdigit()]) 
    indexOfDic = str(b+d)
    dictOfSubDirs[indexOfDic] = directoryContents

#check contains teachername and round path, create list of valid paths
validInRoundPaths = []
for i in range(8):
    listOfTeachers = whatTeacher(i+1)
    listOfKeys = list(dictOfSubDirs.keys())
    dictKey = listOfKeys[i]
    for k in range(len(listOfTeachers)):
            currentTeacher = listOfTeachers[k]
            for j in range(len(dictOfSubDirs[dictKey])):
                currentList = dictOfSubDirs[dictKey]
                currentString = currentList[j]
                if currentTeacher in currentString:
                    newPath = pathnamesForRounds[i] + "/" + currentString
                    validInRoundPaths.append(newPath)
                    if masterDict.__contains__(currentTeacher):
                        continue
                    else:
                        if '_' not in currentTeacher:
                            masterDict[currentTeacher] = {}
                            masterDict[currentTeacher]['Round 1'] = {}
                            masterDict[currentTeacher]['Round 2'] = {}
                            masterDict[currentTeacher]['Round 3'] = {}
                            masterDict[currentTeacher]['Round 4'] = {}
             
#create list of valid paths for compiled versions on each folder.
validCompiledPaths = []
for i in validInRoundPaths:
    directoryContents = []
    directoryContents = os.listdir(i)
    for j in directoryContents:
        if "Compiled" in j:
            validCompiledPaths.append(str(i + "/" + j))
        elif "compiled" in j:
            validCompiledPaths.append(str(i + "/" + j))
        continue

#Create list of lists of students without duplicates.
listOfStudentsWithDuplicates = []
listOfStudents = []
listOfTeachers = list(masterDict.keys())
listOfTeachers.sort()
#validCompiledPaths.remove('//ismb1.luc.edu/psychology/PsychResearchGamez/General/Sixth Grade LENA Study/Syntax Coding Files/Year 1 Student Syntax Files/Round 1/612_1/~$612_1_Compiled_CF.xlsx')

for i in validCompiledPaths:
    loc = (i)
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0,0)

    for j in range(sheet.nrows):
        currentCell = sheet.cell_value(j,0)
        if type(currentCell) == int or type(currentCell) == float:
            listOfStudentsWithDuplicates.append(currentCell)
            
    #add students to each round checking teachername contained in path 
            listOfPath = i.split('/')
            roundNo = listOfPath[9] 
            possibleTeacher = listOfPath[10]
            if '1' in roundNo:
                roundNo = 'Round 1'
            if '2' in roundNo:
                roundNo = 'Round 2'
            if '3' in roundNo:
                roundNo = 'Round 3'
            if '4' in roundNo:
                roundNo = 'Round 4'

            for g in listOfTeachers:
                if g in possibleTeacher:
                    splitValidString = (re.split('\_|\.',possibleTeacher))[0]
                    if g == splitValidString:
                        masterDict[g][roundNo][str(int(currentCell))] = {}
                    else:
                        continue

#eliminating duplicates in student list
[listOfStudents.append(x) for x in listOfStudentsWithDuplicates if x not in listOfStudents]
tempList = [int(i) for i in listOfStudents]
listOfUnsortedStudents = tempList
tempList.sort(reverse=True)
listOfStudents = []
listOfStudents = map(str, tempList)
listOfStudents = list(listOfStudents)

#Dictionary of key:student and values:paths
eliminatePaths = []
dictStudentPaths = {}
for i in validInRoundPaths:
    directoryContents = []
    directoryContents = os.listdir(i)
    for j in listOfStudents:
        for k in directoryContents:
            cleanK = str(k)
            cleanK.replace('.','')
            #k = k.replace('_','')
            if j in cleanK:
                    checkString = str(i + "/" + k)
                    if checkString not in validCompiledPaths:
                        if checkString not in eliminatePaths:
                            if dictStudentPaths.__contains__(j):
                                dictStudentPaths[j].append(str(i + "/" + k))
                                eliminatePaths.append(checkString)
                            else: 
                                dictStudentPaths[j] = [str(i + "/" + k)]
                                eliminatePaths.append(checkString)

                   
#looping through every student in masterdict and adding a new dictionary for each heading in masterspreadsheet (CHITOUT, STUMULTU...)
oldKeys = list(dictStudentPaths.keys())
for teacher in list(masterDict.keys()):
    for thisRound in list(masterDict[teacher].keys()):
        for student in list(masterDict[teacher][thisRound].keys()):
            masterDict[teacher][thisRound][student] = {'CHIMULTU': '','CHITOTUT': '', 'STUMULTU': '', 'STUTOTUT': '', 'CHITYPED': [], 'CHITOKD': [], 'STUTYPED': [], 'STUTOKD': [], 'CHITYPEND': [], 'CHITOKND': [], 'STUTYPEND': [], 'STUTOKND': [], 'CHICO': ''}
            for antiqueKey in oldKeys:
                if antiqueKey == student:
                    for p in dictStudentPaths[antiqueKey]:
                        if ('round 1' in p.lower()) and ('1' in thisRound):
                            masterDict[teacher][thisRound][student]['CHIMULTU'] = p
                            masterDict[teacher][thisRound][student]['CHITOTUT'] = p 
                            masterDict[teacher][thisRound][student]['STUMULTU'] = p
                            masterDict[teacher][thisRound][student]['STUTOTUT'] = p
                            masterDict[teacher][thisRound][student]['CHICO'] = p
                        if ('round 2' in p.lower()) and ('2' in thisRound):
                            masterDict[teacher][thisRound][student]['CHIMULTU'] = p
                            masterDict[teacher][thisRound][student]['CHITOTUT'] = p 
                            masterDict[teacher][thisRound][student]['STUMULTU'] = p
                            masterDict[teacher][thisRound][student]['STUTOTUT'] = p
                            masterDict[teacher][thisRound][student]['CHICO'] = p
                        if ('round 3' in p.lower()) and ('3' in thisRound):
                            masterDict[teacher][thisRound][student]['CHIMULTU'] = p
                            masterDict[teacher][thisRound][student]['CHITOTUT'] = p 
                            masterDict[teacher][thisRound][student]['STUMULTU'] = p
                            masterDict[teacher][thisRound][student]['STUTOTUT'] = p
                            masterDict[teacher][thisRound][student]['CHICO'] = p
                        if ('round 4' in p.lower()) and ('4' in thisRound):
                            masterDict[teacher][thisRound][student]['CHIMULTU'] = p
                            masterDict[teacher][thisRound][student]['CHITOTUT'] = p 
                            masterDict[teacher][thisRound][student]['STUMULTU'] = p
                            masterDict[teacher][thisRound][student]['STUTOTUT'] = p
                            masterDict[teacher][thisRound][student]['CHICO'] = p

                            
#adding student's types and tokens to masterdict depending on round
for teacher in list(masterDict.keys()):
    for thisRound in list(masterDict[teacher].keys()):
        for student in list(masterDict[teacher][thisRound].keys()):
            if '1' in thisRound: 
                masterDict[teacher][thisRound][student]['CHITYPED'] = listOfTypesTokensDaleR1
                masterDict[teacher][thisRound][student]['CHITOKD'] = listOfTypesTokensDaleR1
                masterDict[teacher][thisRound][student]['STUTYPED'] = listOfTypesTokensDaleR1
                masterDict[teacher][thisRound][student]['STUTOKD'] = listOfTypesTokensDaleR1
                masterDict[teacher][thisRound][student]['CHITYPEND'] = listOfTypesTokensNoDaleR1
                masterDict[teacher][thisRound][student]['CHITOKND'] = listOfTypesTokensNoDaleR1
                masterDict[teacher][thisRound][student]['STUTYPEND'] = listOfTypesTokensNoDaleR1
                masterDict[teacher][thisRound][student]['STUTOKND'] = listOfTypesTokensNoDaleR1
            if '2' in thisRound: 
                masterDict[teacher][thisRound][student]['CHITYPED'] = listOfTypesTokensDaleR2
                masterDict[teacher][thisRound][student]['CHITOKD'] = listOfTypesTokensDaleR2
                masterDict[teacher][thisRound][student]['STUTYPED'] = listOfTypesTokensDaleR2
                masterDict[teacher][thisRound][student]['STUTOKD'] = listOfTypesTokensDaleR2
                masterDict[teacher][thisRound][student]['CHITYPEND'] = listOfTypesTokensNoDaleR2
                masterDict[teacher][thisRound][student]['CHITOKND'] = listOfTypesTokensNoDaleR2
                masterDict[teacher][thisRound][student]['STUTYPEND'] = listOfTypesTokensNoDaleR2
                masterDict[teacher][thisRound][student]['STUTOKND'] = listOfTypesTokensNoDaleR2
            if '3' in thisRound: 
                masterDict[teacher][thisRound][student]['CHITYPED'] = listOfTypesTokensDaleR3
                masterDict[teacher][thisRound][student]['CHITOKD'] = listOfTypesTokensDaleR3
                masterDict[teacher][thisRound][student]['STUTYPED'] = listOfTypesTokensDaleR3
                masterDict[teacher][thisRound][student]['STUTOKD'] = listOfTypesTokensDaleR3
                masterDict[teacher][thisRound][student]['CHITYPEND'] = listOfTypesTokensNoDaleR3
                masterDict[teacher][thisRound][student]['CHITOKND'] = listOfTypesTokensNoDaleR3
                masterDict[teacher][thisRound][student]['STUTYPEND'] = listOfTypesTokensNoDaleR3
                masterDict[teacher][thisRound][student]['STUTOKND'] = listOfTypesTokensNoDaleR3
            if '4' in thisRound: 
                masterDict[teacher][thisRound][student]['CHITYPED'] = listOfTypesTokensDaleR4
                masterDict[teacher][thisRound][student]['CHITOKD'] = listOfTypesTokensDaleR4
                masterDict[teacher][thisRound][student]['STUTYPED'] = listOfTypesTokensDaleR4
                masterDict[teacher][thisRound][student]['STUTOKD'] = listOfTypesTokensDaleR4
                masterDict[teacher][thisRound][student]['CHITYPEND'] = listOfTypesTokensNoDaleR4
                masterDict[teacher][thisRound][student]['CHITOKND'] = listOfTypesTokensNoDaleR4
                masterDict[teacher][thisRound][student]['STUTYPEND'] = listOfTypesTokensNoDaleR4
                masterDict[teacher][thisRound][student]['STUTOKND'] = listOfTypesTokensNoDaleR4
                                

#start_time2 = time.time()
#print(childTypeToken(listOfTypesTokensDaleR1, '6111'))
#end_time2 = time.time()
#print("Runtime: {} minutes".format((end_time2 - start_time2)/60))  


#Creating master spreadsheet.

workbook = xlsxwriter.Workbook('CarlosTeacherStudentRubric.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A2', 'Order')
worksheet.write('B2', 'Site')
worksheet.write('C2', 'Cohort')
worksheet.write('D2', 'Student')
worksheet.write('E2', 'Teacher')
worksheet.write('F2', 'THCType')

worksheet.write('G1', 'ROUND 1')
worksheet.write('G2', 'CHIMULTU')
worksheet.write('H2', 'CHITOUT')
worksheet.write('I2', 'STUMULTU')
worksheet.write('J2', 'STUTOUT')
worksheet.write('K2', 'CHITYPED')
worksheet.write('L2', 'CHITOKD')
worksheet.write('M2', 'STUTYPED')
worksheet.write('N2', 'STUTOKD')
worksheet.write('O2', 'CHITYPEND')
worksheet.write('P2', 'CHITOKND')
worksheet.write('Q2', 'STUTYPEND')
worksheet.write('R2', 'STUTOKND')
worksheet.write('S2', 'CHICO')
worksheet.write('T2', 'CHIS1')
worksheet.write('U2', 'CHIS2')
worksheet.write('V2', 'CHISC')
worksheet.write('W2', 'CHIOC')
worksheet.write('X2', 'CHISRC')
worksheet.write('Y2', 'CHIORC')
worksheet.write('Z2', 'STUCO')
worksheet.write('AA2', 'STUS1')
worksheet.write('AB2', 'STUS2')
worksheet.write('AC2', 'STUSC')
worksheet.write('AD2', 'STUOC')
worksheet.write('AE2', 'STUSRC')
worksheet.write('AF2', 'STUORC')

worksheet.write('AG1', 'ROUND 2')
worksheet.write('AG2', 'CHIMULTU')
worksheet.write('AH2', 'CHITOUT')
worksheet.write('AI2', 'STUMULTU')
worksheet.write('AJ2', 'STUTOUT')
worksheet.write('AK2', 'CHITYPED')
worksheet.write('AL2', 'CHITOKD')
worksheet.write('AM2', 'STUTYPED')
worksheet.write('AN2', 'STUTOKD')
worksheet.write('AO2', 'CHITYPEND')
worksheet.write('AP2', 'CHITOKND')
worksheet.write('AQ2', 'STUTYPEND')
worksheet.write('AR2', 'STUTOKND')
worksheet.write('AS2', 'CHICO')
worksheet.write('AT2', 'CHIS1')
worksheet.write('AU2', 'CHIS2')
worksheet.write('AV2', 'CHISC')
worksheet.write('AW2', 'CHIOC')
worksheet.write('AX2', 'CHISRC')
worksheet.write('AY2', 'CHIORC')
worksheet.write('AZ2', 'STUCO')
worksheet.write('BA2', 'STUS1')
worksheet.write('BB2', 'STUS2')
worksheet.write('BC2', 'STUSC')
worksheet.write('BD2', 'STUOC')
worksheet.write('BE2', 'STUSRC')
worksheet.write('BF2', 'STUORC')

worksheet.write('BG1', 'ROUND 3')
worksheet.write('BG2', 'CHIMULTU')
worksheet.write('BH2', 'CHITOUT')
worksheet.write('BI2', 'STUMULTU')
worksheet.write('BJ2', 'STUTOUT')
worksheet.write('BK2', 'CHITYPED')
worksheet.write('BL2', 'CHITOKD')
worksheet.write('BM2', 'STUTYPED')
worksheet.write('BN2', 'STUTOKD')
worksheet.write('BO2', 'CHITYPEND')
worksheet.write('BP2', 'CHITOKND')
worksheet.write('BQ2', 'STUTYPEND')
worksheet.write('BR2', 'STUTOKND')
worksheet.write('BS2', 'CHICO')
worksheet.write('BT2', 'CHIS1')
worksheet.write('BU2', 'CHIS2')
worksheet.write('BV2', 'CHISC')
worksheet.write('BW2', 'CHIOC')
worksheet.write('BX2', 'CHISRC')
worksheet.write('BY2', 'CHIORC')
worksheet.write('BZ2', 'STUCO')
worksheet.write('CA2', 'STUS1')
worksheet.write('CB2', 'STUS2')
worksheet.write('CC2', 'STUSC')
worksheet.write('CD2', 'STUOC')
worksheet.write('CE2', 'STUSRC')
worksheet.write('CF2', 'STUORC')

worksheet.write('CG1', 'ROUND 4')
worksheet.write('CG2', 'CHIMULTU')
worksheet.write('CH2', 'CHITOUT')
worksheet.write('CI2', 'STUMULTU')
worksheet.write('CJ2', 'STUTOUT')
worksheet.write('CK2', 'CHITYPED')
worksheet.write('CL2', 'CHITOKD')
worksheet.write('CM2', 'STUTYPED')
worksheet.write('CN2', 'STUTOKD')
worksheet.write('CO2', 'CHITYPEND')
worksheet.write('CP2', 'CHITOKND')
worksheet.write('CQ2', 'STUTYPEND')
worksheet.write('CR2', 'STUTOKND')
worksheet.write('CS2', 'CHICO')
worksheet.write('CT2', 'CHIS1')
worksheet.write('CU2', 'CHIS2')
worksheet.write('CV2', 'CHISC')
worksheet.write('CW2', 'CHIOC')
worksheet.write('CX2', 'CHISRC')
worksheet.write('CY2', 'CHIORC')
worksheet.write('CZ2', 'STUCO')
worksheet.write('DA2', 'STUS1')
worksheet.write('DB2', 'STUS2')
worksheet.write('DC2', 'STUSC')
worksheet.write('DD2', 'STUOC')
worksheet.write('DE2', 'STUSRC')
worksheet.write('DF2', 'STUORC')

allTeachers = list(masterDict.keys())
columnCounter = 2
for teacher in allTeachers:
    studentsTeacherDup = []
    studentsTeacher = []
    
    if teacher == '612':
        try:
            studentsTeacher.remove('6124')
            studentsTeacher.remove('62113')
            studentsTeacher.remove('61220')
        except:
            pass
        
    studentsTeacherDup.extend(list(masterDict[teacher]['Round 1'].keys()))
    studentsTeacherDup.extend(list(masterDict[teacher]['Round 2'].keys()))
    studentsTeacherDup.extend(list(masterDict[teacher]['Round 3'].keys()))
    studentsTeacherDup.extend(list(masterDict[teacher]['Round 4'].keys()))
    [studentsTeacher.append(x) for x in studentsTeacherDup if x not in studentsTeacher] 
    
    for x in range(len(studentsTeacher)):
        worksheet.write(columnCounter,3, studentsTeacher[x])
        worksheet.write(columnCounter,4,teacher)
        columnCounter = columnCounter + 1

workbook.close()

workbook = xlsxwriter.Workbook('CarlosMasterSpreadsheet.xlsx')
worksheet = workbook.add_worksheet()

rowIteration = 2
rubric = r'C:/Users/papia/Documents/Loyola University Chicago/LUC Fourth Semester/Bizarre/Psychology Data Science project/CarlosTeacherStudentRubric.xlsx'
wb = xlrd.open_workbook(rubric)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0,0)
listOfUnvalidDictPaths = []

#main function to create masterspreadsheet iterating through MasterDict
for i in range(540):
    print('RUN' + str(i))
    columnIteration = 6
    try:
        if((columnIteration >= 6) and (columnIteration <= 31)):
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))    
            if valueNeeded == 'CHIMULTU':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 1'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    if 'Year 1' in path:
                        worksheet.write(rowIteration,2, 'Year 1')
                    if 'Year 2' in path:
                        worksheet.write(rowIteration,2, 'Year 2')
                    listOf4 = chimultuChitotutStumultuStutotut(path)
                    for b in listOf4:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1
                except:
                #listOfUnvalidDictPaths.append((str(sheet.cell_value(rowIteration,4))) + '  Round 1  ' + (str(sheet.cell_value(rowIteration,3))) +  ' ' + valueNeeded)
                    columnIteration = columnIteration + 4      
                    pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))
            if valueNeeded == 'CHITYPED':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 1'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    addThisList = childTypeToken(path, (str(sheet.cell_value(rowIteration,3))))
                    for b in addThisList:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1
                except:
                    columnIteration = columnIteration + 4
                    pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))
            if valueNeeded == 'CHITYPEND':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 1'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    addThisList = childTypeToken(path, (str(sheet.cell_value(rowIteration,3))))
                    for b in addThisList:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1
                except:
                    columnIteration = columnIteration + 4
                    pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))           
            if valueNeeded == 'CHICO':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 1'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    listOf14 = complexCode(path)
                    for b in listOf14:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1   
                except:
                    columnIteration = columnIteration + 14
                    pass
                 
                
        if((columnIteration >= 32) and (columnIteration <= 57)):
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))
            if valueNeeded == 'CHIMULTU':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 2'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    if 'Year 1' in path:
                        worksheet.write(rowIteration,2, 'Year 1')
                    if 'Year 2' in path:
                        worksheet.write(rowIteration,2, 'Year 2')
                    listOf4 = chimultuChitotutStumultuStutotut(path)
                    for b in listOf4:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1
                except:
                #listOfUnvalidDictPaths.append((str(sheet.cell_value(rowIteration,4))) + '  Round 2  ' + (str(sheet.cell_value(rowIteration,3))) +  ' ' + valueNeeded)
                    columnIteration = columnIteration + 4
                    pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))
            if valueNeeded == 'CHITYPED':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 2'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    addThisList = childTypeToken(path, (str(sheet.cell_value(rowIteration,3))))
                    for b in addThisList:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1
                except:
                    columnIteration = columnIteration + 4
                    pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))
            if valueNeeded == 'CHITYPEND':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 2'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    addThisList = childTypeToken(path, (str(sheet.cell_value(rowIteration,3))))
                    for b in addThisList:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1
                except:
                    columnIteration = columnIteration + 4
                    pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration))) 
            if valueNeeded == 'CHICO':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 2'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    listOf14 = complexCode(path)
                    for b in listOf14:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1   
                except:
                    columnIteration = columnIteration + 14
                    pass       
            
        if((columnIteration >= 58) and (columnIteration <= 82)):
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))
            if valueNeeded == 'CHIMULTU':
                try:
                        path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 3'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                        if 'Year 1' in path:
                            worksheet.write(rowIteration,2, 'Year 1')
                        if 'Year 2' in path:
                            worksheet.write(rowIteration,2, 'Year 2')
                        listOf4 = chimultuChitotutStumultuStutotut(path)
                        for b in listOf4:
                            worksheet.write(rowIteration,columnIteration,b)
                            columnIteration = columnIteration + 1
                except:
                    #listOfUnvalidDictPaths.append((str(sheet.cell_value(rowIteration,4))) + '  Round 2  ' + (str(sheet.cell_value(rowIteration,3))) +  ' ' + valueNeeded)
                        columnIteration = columnIteration + 4
                        pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))
            if valueNeeded == 'CHITYPED':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 3'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    addThisList = childTypeToken(path, (str(sheet.cell_value(rowIteration,3))))
                    for b in addThisList:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1
                except:
                    columnIteration = columnIteration + 4
                    pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))
            if valueNeeded == 'CHITYPEND':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 3'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    addThisList = childTypeToken(path, (str(sheet.cell_value(rowIteration,3))))
                    for b in addThisList:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1
                except:
                    columnIteration = columnIteration + 4
                    pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration))) 
            if valueNeeded == 'CHICO':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 3'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    listOf14 = complexCode(path)
                    for b in listOf14:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1   
                except:
                    columnIteration = columnIteration + 14
                    pass
                    
        if((columnIteration >= 83) and (columnIteration <= 107)):
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))
            if valueNeeded == 'CHIMULTU':
                try:
                        path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 4'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                        if 'Year 1' in path:
                            worksheet.write(rowIteration,2, 'Year 1')
                        if 'Year 2' in path:
                            worksheet.write(rowIteration,2, 'Year 2')
                        listOf4 = chimultuChitotutStumultuStutotut(path)
                        for b in listOf4:
                            worksheet.write(rowIteration,columnIteration,b)
                            columnIteration = columnIteration + 1
                except:
                    #listOfUnvalidDictPaths.append((str(sheet.cell_value(rowIteration,4))) + '  Round 2  ' + (str(sheet.cell_value(rowIteration,3))) +  ' ' + valueNeeded)
                        columnIteration = columnIteration + 4
                        pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))
            if valueNeeded == 'CHITYPED':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 4'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    addThisList = childTypeToken(path, (str(sheet.cell_value(rowIteration,3))))
                    for b in addThisList:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1
                except:
                    columnIteration = columnIteration + 4
                    pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration)))
            if valueNeeded == 'CHITYPEND':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 4'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    addThisList = childTypeToken(path, (str(sheet.cell_value(rowIteration,3))))
                    for b in addThisList:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1
                except:
                    columnIteration = columnIteration + 4
                    pass
            valueNeeded = (str(sheet.cell_value(1,columnIteration))) 
            if valueNeeded == 'CHICO':
                try:
                    path = masterDict[(str(sheet.cell_value(rowIteration,4)))]['Round 4'][(str(sheet.cell_value(rowIteration,3)))][valueNeeded]
                    listOf14 = complexCode(path)
                    for b in listOf14:
                        worksheet.write(rowIteration,columnIteration,b)
                        columnIteration = columnIteration + 1   
                except:
                    columnIteration = columnIteration + 14
                    pass
    except:
        print('Whoops! Something went wrong!')
        continue               
    
    rowIteration = rowIteration + 1

worksheet.write('A2', 'Order')
worksheet.write('B2', 'Site')
worksheet.write('C2', 'Cohort')
worksheet.write('D2', 'Student')
worksheet.write('E2', 'Teacher')
worksheet.write('F2', 'THCType')

worksheet.write('G1', 'ROUND 1')
worksheet.write('G2', 'CHIMULTU')
worksheet.write('H2', 'CHITOUT')
worksheet.write('I2', 'STUMULTU')
worksheet.write('J2', 'STUTOUT')
worksheet.write('K2', 'CHITYPED')
worksheet.write('L2', 'CHITOKD')
worksheet.write('M2', 'STUTYPED')
worksheet.write('N2', 'STUTOKD')
worksheet.write('O2', 'CHITYPEND')
worksheet.write('P2', 'CHITOKND')
worksheet.write('Q2', 'STUTYPEND')
worksheet.write('R2', 'STUTOKND')
worksheet.write('S2', 'CHICO')
worksheet.write('T2', 'CHIS1')
worksheet.write('U2', 'CHIS2')
worksheet.write('V2', 'CHISC')
worksheet.write('W2', 'CHIOC')
worksheet.write('X2', 'CHISRC')
worksheet.write('Y2', 'CHIORC')
worksheet.write('Z2', 'STUCO')
worksheet.write('AA2', 'STUS1')
worksheet.write('AB2', 'STUS2')
worksheet.write('AC2', 'STUSC')
worksheet.write('AD2', 'STUOC')
worksheet.write('AE2', 'STUSRC')
worksheet.write('AF2', 'STUORC')

worksheet.write('AG1', 'ROUND 2')
worksheet.write('AG2', 'CHIMULTU')
worksheet.write('AH2', 'CHITOUT')
worksheet.write('AI2', 'STUMULTU')
worksheet.write('AJ2', 'STUTOUT')
worksheet.write('AK2', 'CHITYPED')
worksheet.write('AL2', 'CHITOKD')
worksheet.write('AM2', 'STUTYPED')
worksheet.write('AN2', 'STUTOKD')
worksheet.write('AO2', 'CHITYPEND')
worksheet.write('AP2', 'CHITOKND')
worksheet.write('AQ2', 'STUTYPEND')
worksheet.write('AR2', 'STUTOKND')
worksheet.write('AS2', 'CHICO')
worksheet.write('AT2', 'CHIS1')
worksheet.write('AU2', 'CHIS2')
worksheet.write('AV2', 'CHISC')
worksheet.write('AW2', 'CHIOC')
worksheet.write('AX2', 'CHISRC')
worksheet.write('AY2', 'CHIORC')
worksheet.write('AZ2', 'STUCO')
worksheet.write('BA2', 'STUS1')
worksheet.write('BB2', 'STUS2')
worksheet.write('BC2', 'STUSC')
worksheet.write('BD2', 'STUOC')
worksheet.write('BE2', 'STUSRC')
worksheet.write('BF2', 'STUORC')

worksheet.write('BG1', 'ROUND 3')
worksheet.write('BG2', 'CHIMULTU')
worksheet.write('BH2', 'CHITOUT')
worksheet.write('BI2', 'STUMULTU')
worksheet.write('BJ2', 'STUTOUT')
worksheet.write('BK2', 'CHITYPED')
worksheet.write('BL2', 'CHITOKD')
worksheet.write('BM2', 'STUTYPED')
worksheet.write('BN2', 'STUTOKD')
worksheet.write('BO2', 'CHITYPEND')
worksheet.write('BP2', 'CHITOKND')
worksheet.write('BQ2', 'STUTYPEND')
worksheet.write('BR2', 'STUTOKND')
worksheet.write('BS2', 'CHICO')
worksheet.write('BT2', 'CHIS1')
worksheet.write('BU2', 'CHIS2')
worksheet.write('BV2', 'CHISC')
worksheet.write('BW2', 'CHIOC')
worksheet.write('BX2', 'CHISRC')
worksheet.write('BY2', 'CHIORC')
worksheet.write('BZ2', 'STUCO')
worksheet.write('CA2', 'STUS1')
worksheet.write('CB2', 'STUS2')
worksheet.write('CC2', 'STUSC')
worksheet.write('CD2', 'STUOC')
worksheet.write('CE2', 'STUSRC')
worksheet.write('CF2', 'STUORC')

worksheet.write('CG1', 'ROUND 4')
worksheet.write('CG2', 'CHIMULTU')
worksheet.write('CH2', 'CHITOUT')
worksheet.write('CI2', 'STUMULTU')
worksheet.write('CJ2', 'STUTOUT')
worksheet.write('CK2', 'CHITYPED')
worksheet.write('CL2', 'CHITOKD')
worksheet.write('CM2', 'STUTYPED')
worksheet.write('CN2', 'STUTOKD')
worksheet.write('CO2', 'CHITYPEND')
worksheet.write('CP2', 'CHITOKND')
worksheet.write('CQ2', 'STUTYPEND')
worksheet.write('CR2', 'STUTOKND')
worksheet.write('CS2', 'CHICO')
worksheet.write('CT2', 'CHIS1')
worksheet.write('CU2', 'CHIS2')
worksheet.write('CV2', 'CHISC')
worksheet.write('CW2', 'CHIOC')
worksheet.write('CX2', 'CHISRC')
worksheet.write('CY2', 'CHIORC')
worksheet.write('CZ2', 'STUCO')
worksheet.write('DA2', 'STUS1')
worksheet.write('DB2', 'STUS2')
worksheet.write('DC2', 'STUSC')
worksheet.write('DD2', 'STUOC')
worksheet.write('DE2', 'STUSRC')
worksheet.write('DF2', 'STUORC')

allTeachers = list(masterDict.keys())
columnCounter = 2
for teacher in allTeachers:
    studentsTeacherDup = []
    studentsTeacher = []
    
    if teacher == '612':
        try:
            studentsTeacher.remove('6124')
            studentsTeacher.remove('62113')
            studentsTeacher.remove('61220')
        except:
            pass
        
    studentsTeacherDup.extend(list(masterDict[teacher]['Round 1'].keys()))
    studentsTeacherDup.extend(list(masterDict[teacher]['Round 2'].keys()))
    studentsTeacherDup.extend(list(masterDict[teacher]['Round 3'].keys()))
    studentsTeacherDup.extend(list(masterDict[teacher]['Round 4'].keys()))
    [studentsTeacher.append(x) for x in studentsTeacherDup if x not in studentsTeacher] 
    
    for x in range(len(studentsTeacher)):
        worksheet.write(columnCounter,3, studentsTeacher[x])
        worksheet.write(columnCounter,4,teacher)
        columnCounter = columnCounter + 1

#print(listOfUnvalidDictPaths)
workbook.close()
end_time = time.time()
print("Runtime: {} minutes".format((end_time - start_time)/60))                
                
                        
                




    
