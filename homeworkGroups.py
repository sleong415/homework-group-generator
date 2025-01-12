import pandas as pd
import xlsxwriter

headerFontSize = 14
textFontSize = 11
numGroups = 15

colorList = ['#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc', '#efefef']
bold_center_formats = []

'''
Retrieves students ONLY from excel sheet downloaded from canvas roster.
Sorts names by first name while leaving names in 'last, first' format.
'''
def getStudentRosterFromExcel(sheet):
    rosterSheet = pd.read_excel(f"sheets/{sheet}", usecols=['Name', 'Role'])
    names = rosterSheet[rosterSheet['Role'] == 'Student'].get('Name')
    roster = sorted(names, key=lambda name: name.split(', ')[1])

    print("num of students:", len(roster))

    return roster

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

        print("new TAs:", len(newTAs))
        print("returning TAs:", len(returningTAs))
        print("total TAs:", totalTAs)

        return totalTAs, newTAs, returningTAs

'''
Disitrbute TAs into groups of 3-4 with at least 1 returning TA
in each group. Returns a list of groups

newTAs = list of new TAs
returningTAs = list of returningTAs
'''
def createTAGroups(newTAs, returningTAs):
    groupings = []
    taCount = 0
    numTAs = len(newTAs) + len(returningTAs)

    for _ in range(numGroups):
        groupings.append([])

    # add new TAs to every group
    for i, ta in enumerate(newTAs):
        groupings[i % numGroups].append(ta)
        taCount += 1

    # assign 2 returning TAs to each group
    for i in range(numGroups):
        groupings[i].append(returningTAs.pop())
        taCount += 1

        if len(returningTAs) >= 1:
            groupings[i].append(returningTAs.pop())
            taCount += 1

    if (len(newTAs) < numGroups):
        # add one more to group w/ no new TAs
        for i in range(len(newTAs), numGroups):
            if len(returningTAs) >= 1:
                groupings[i].append(returningTAs.pop())
                taCount += 1

    # place remaining TAs to make groups of 4
    if (taCount < numTAs):
        for i in range(numGroups):
            if len(returningTAs) < 1:
                break
            groupings[i].append(returningTAs.pop())
            taCount += 1

    groupings.reverse()     # put groups of 4 at the end

    if (taCount != numTAs):
        raise Exception("Unsuccessful in evenly/correctly distributing TAs into groups")

    return groupings

def set_bold_center_bg_color(workbook, color):
    return workbook.add_format({
            'bold': True,
            'align': 'center',
            'font_size': headerFontSize,
            'bg_color': color
        })

def set_center_bg_color(workbook, color):
    return workbook.add_format({
            'align': 'center',
            'font_size': headerFontSize,
            'bg_color': color
        })

'''
Creates first sheet with all TA groups displayed.
'''
def createFrontSheet(workbook, taList):
    boldCenter = workbook.add_format({
        'bold': True,
        'align': 'center',
        'font_size': headerFontSize
    })
      
    groupingSheet = workbook.add_worksheet("Groups")
    groupingSheet.set_column('B:G', 26)

    groupingSheet.write(1, 0, "Groups:", boldCenter)
    groupingSheet.set_column('A:A', 10)

    global bold_center_formats
    row = 1
    col = 1

    # create main page with all TA groups
    for i in range(1, numGroups + 1):
        currRow = row

        group_bg_color = colorList[(i-1) % len(colorList)]
        group_num_format = set_bold_center_bg_color(workbook, group_bg_color)
        bold_center_formats.append(group_num_format)

        ta_name_format = set_center_bg_color(workbook, group_bg_color)

        # number header
        groupingSheet.write(currRow, col, i, group_num_format)

        # writes TA names
        for ta in taList[i-1]:
            currRow += 1
            groupingSheet.write(currRow, col, ta, ta_name_format)

        if (len(taList[i-1]) == 3):
            groupingSheet.write(currRow + 1, col, None, ta_name_format)

        # new row for every 6 groups
        if (i % 6 == 0):
            row += 6
            col = 1
        else:
            col += 1
    
    print("main sheet created")

'''
Creates sheets for each group. Each sheet has TA names and their assigned students.
'''
def createGroupSheets(workbook, taRoster, numTAs, studentRoster):
    center = workbook.add_format({
        'align': 'center',
        'font_size': textFontSize,
    })

    groupIndex = 0
    studentIndex = 0
    baseSize = len(studentRoster) // numTAs
    remainder = len(studentRoster) % numTAs

    print(f"about {baseSize} students per TA")

    for i in range(len(taRoster)):
        # groupExtra ensures that remainder students are distirbuted evenly per GROUP of 3-4 TAs not per TA
        # e.g. 1 extra student for the first 5 groups instead of 1 extra student for the first 5 TAs

        groupExtra = remainder // len(taRoster)

        if (remainder % len(taRoster) > i):
            groupExtra += 1

        groupSheet = workbook.add_worksheet(f"Group{groupIndex + 1}")

        column_range = 'A:C' if len(taRoster[i]) == 3 else 'A:D'
        groupSheet.set_column(column_range, 36)
       
        global bold_center_formats
        col = 0
        for ta in taRoster[i]:
            row = 0

            # write ta name
            groupSheet.write(row, col, ta, bold_center_formats[i])
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

    print("group sheets created")

def main():
    # file variable names
    rosterFileName = 'Roster.xlsx'
    taRosterFileName = 'taRoster.txt'
    outputFileName = 'HomeworkGroups.xlsx'

    # get and parse rosters
    roster = getStudentRosterFromExcel(rosterFileName)

    # count TAs given TA roster
    totalTAs, newTAs, returningTAs = countAndReturnTAs(taRosterFileName)

    # create TA groups
    groupingList = createTAGroups(newTAs, returningTAs)

    # create sheet
    workbook = xlsxwriter.Workbook(outputFileName)
    createFrontSheet(workbook, groupingList)
    createGroupSheets(workbook, groupingList, totalTAs, roster)
    workbook.close()

if __name__ == "__main__":
    main()