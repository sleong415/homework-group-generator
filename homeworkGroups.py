import pandas as pd
import xlsxwriter

headerFontSize = 14
textFontSize = 11
columnWidth = 36
numGroups = 8

# global variable to track running group num between campus and online sections
groupIndex = 0

'''
Retrieves students ONLY from excel sheet downloaded from canvas roster.
Sorts names by first name while leaving names in 'last, first' format.
'''
def getStudentRosterFromExcel(sheet):
    roster = pd.read_excel(f"sheets/{sheet}", usecols=['Name', 'Role'])
    names = roster[roster['Role'] == 'Student'].get('Name')
    return sorted(names, key=lambda name: name.split(', ')[1])

'''
Given both rosters and total number of TAs, calculate TA distribution between
online and campus section
'''
def calculateTADistribution(campusRoster, onlineRoster, totalTAs):
    campusSize = len(campusRoster)
    onlineSize = len(onlineRoster)

    print(f"campus size: {campusSize}, online size: {onlineSize}")

    totalStudents = campusSize + onlineSize
    percentCampus = campusSize / totalStudents
    percentOnline = onlineSize / totalStudents

    numCampusTAs = round(totalTAs * percentCampus)
    numOnlineTAs = round(totalTAs * percentOnline)

    print(f"campus tas: {numCampusTAs}, online tas: {numOnlineTAs}")

    return numCampusTAs, numOnlineTAs

'''
Given TA roster, count total TAs and return 2 arrays (new TAs and returning TAs). 
Make sure there is a newline between the 2 groupings (new TAs first, then returning TAs).

sheet = list of TA names
'''
def countAndReturnTAs(sheet):
    with open(f'sheets/{sheet}', "r") as file:
        newTAs = []
        returningTAs = []
        isNewTAs = True

        for line in file:
            name = line.strip()

            if name == '':
                isNewTAs = False
                continue

            if isNewTAs:
                newTAs.append(name)
            else:
                returningTAs.append(name)

        totalTAs = len(newTAs) + len(returningTAs)
        return totalTAs, newTAs, returningTAs

'''
Disitrbute TAs into groups of 3-4 with at least 1 returning TA
in each group. Returns a list of groups

newTAs = list of new TAs
returningTAs = list of returningTAs
'''
def createTAGroups(newTAs, returningTAs):
    groupings = []

    for _ in range(numGroups):
        groupings.append([])

    # add new TAs to every group
    for i, ta in enumerate(newTAs):
        groupings[i % numGroups].append(ta)

    groupings.reverse()     # put groups of 2 new TAs at the end

    # assign 2 returning TAs to each group
    for i in range(numGroups):
        groupings[i].append(returningTAs.pop())

        if len(returningTAs) >= 1:
            groupings[i].append(returningTAs.pop())

    return groupings

'''
Separate TA groups into campus/online groups given required numbers
for each section.

numCampusTAs = number of TAs distributed for campus section
numOnlineTAs = number of TAs distribruted for online section
'''
def separateGroups(groups, numCampusTAs, numOnlineTAs):
    campusGroup = []
    onlineGroup = []
    campusCount = 0
    onlineCount = 0

    # reverse list so groups of 3 are pulled first
    for group in groups[::-1]:
        groupSize = len(group)

        if onlineCount + groupSize <= numOnlineTAs:
            onlineGroup.append(group)
            onlineCount += groupSize
        else:
            campusGroup.append(group)
            campusCount += groupSize

    if campusCount != numCampusTAs or onlineCount != numOnlineTAs:
        errorMsg = f"The groups cannot be evenly distributed to match the required counts. Actual campus tas: {campusCount}. Actual online tas: {onlineCount}"
        raise ValueError(errorMsg)
    
    return campusGroup, onlineGroup

'''
Creates first sheet with all TA groups displayed.

campusTAs = list of campusTA groups
onlineTAs = list of onlineTA groups
'''
def createFrontSheet(workbook, campusTAs, onlineTAs):
    center = workbook.add_format({
        'align': 'center',
        'font_size': headerFontSize,
    })
    boldCenter = workbook.add_format({
        'bold': True,
        'align': 'center',
        'font_size': headerFontSize,
    })

    groupingSheet = workbook.add_worksheet("Groups")
    groupingSheet.set_column('B:Z', 26)

    groupingSheet.write(1, 0, "Groups:", boldCenter)

    row = 1
    col = 1

    # create main page with all TA groups
    for i in range(1, numGroups + 1):
        currRow = row

        # number header
        groupingSheet.write(currRow, col, i, boldCenter)

        # writes TA names
        if (i-1 < len(campusTAs)):
            for ta in campusTAs[i-1]:
                groupingSheet.write(currRow + 1, col, ta, center)
                currRow += 1
        else:
            for ta in onlineTAs[i-1 - len(campusTAs)]:
                groupingSheet.write(currRow + 1, col, ta, center)
                currRow += 1

        # new row for every 6 groups
        if (i % 6 == 0):
            row += 6
            col = 1
        else:
            col += 1
    
    print("main sheet created")

'''
Creates sheets for each group for one section (campus or online). Each sheet
has TA names and their assigned students.
'''
def createGroupSheets(workbook, taRoster, numTAs, studentRoster, text):
    global groupIndex

    boldCenter = workbook.add_format({
        'bold': True,
        'align': 'center',
        'font_size': headerFontSize,
    })

    center = workbook.add_format({
        'align': 'center',
        'font_size': textFontSize,
    })

    studentIndex = 0
    baseSize = len(studentRoster) // numTAs
    remainder = len(studentRoster) % numTAs

    for i in range(len(taRoster)):
        # groupExtra ensures that remainder students are distirbuted evenly per GROUP of 3-4 TAs not per TA
        # e.g. 1 extra student for the first 5 groups instead of 1 extra student for the first 5 TAs

        groupExtra = remainder // len(taRoster)

        if (remainder % len(taRoster) > i):
            groupExtra += 1

        groupSheet = workbook.add_worksheet(f"Group{groupIndex + 1}")

        if (len(taRoster[i]) == 3):
            groupSheet.set_column('A:C', columnWidth)
            groupSheet.set_column('E:E', 40)
            groupSheet.write(1, 4, f"ALL {text.upper()} STUDENTS", boldCenter)
        else:
            groupSheet.set_column('A:D', columnWidth)
            groupSheet.set_column('F:F', 40)
            groupSheet.write(1, 5, f"ALL {text.upper()} STUDENTS", boldCenter)

        col = 0
        for ta in taRoster[i]:
            row = 0

            # write ta name
            groupSheet.write(row, col, ta, boldCenter)
            row += 1

            # disitrbute 'remainder' students to each group rather than per TA for more even disitrbution
            groupSize = baseSize
            if (groupExtra > 0):
                groupSize += 1
                groupExtra -= 1

            # write students under ta
            studentGroup = studentRoster[studentIndex : studentIndex + groupSize]
            studentIndex += groupSize

            for student in studentGroup:
                groupSheet.write(row, col, student, center)
                row += 1

            col += 1 

        groupIndex += 1

    print(f"{text} group sheet created")

    
def main():
    # file variable names
    campusRosterFileName = 'CampusRoster.xlsx'
    onlineRosterFileName = 'OnlineRoster.xlsx'
    taRosterFileName = 'exampleTARoster.txt'
    outputFileName = 'HomeworkGroups.xlsx'

    # get and parse rosters
    campusRoster = getStudentRosterFromExcel(campusRosterFileName)
    onlineRoster = getStudentRosterFromExcel(onlineRosterFileName)

    # count TAs given TA roster
    totalTAs, newTAs, returningTAs = countAndReturnTAs(taRosterFileName)

    # number of TAs for campus/online section
    numCampusTAs, numOnlineTAs = calculateTADistribution(campusRoster, onlineRoster, totalTAs)

    # create TA groups
    groupingList = createTAGroups(newTAs, returningTAs)
    campusTAs, onlineTAs = separateGroups(groupingList, numCampusTAs, numOnlineTAs)
    campusTAs.reverse()
    onlineTAs.reverse()

    # create sheet
    workbook = xlsxwriter.Workbook(outputFileName)
    createFrontSheet(workbook, campusTAs, onlineTAs)
    createGroupSheets(workbook, campusTAs, numCampusTAs, campusRoster, "campus/hybrid")
    createGroupSheets(workbook, onlineTAs, numOnlineTAs, onlineRoster, "online")
    workbook.close()

if __name__ == "__main__":
    main()