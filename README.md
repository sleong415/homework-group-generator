## How to Run Homework Grouping Script
1. Download campus and online student roster from Canvas
2. Set up teaching assistant (TA) roster
    * Each TA name should be listed on a new line.
    * New TAs are listed first, followed by returning TAs.
    * List of new TAs and returning TAs should be separated by an empty line.
    * `exampleTARoster.txt` is an example of how it should be set up.
3. Update name variables at the top of `main()` to reflect file names
4. Place student and TA rosters under the `sheets` folder
5. Run `python .\homeworkGrouops.py` in the root directory
6. View output under root directory

</br>
NOTE: this script still works if you only have **one section**. Check out "Running Script With 1 Class Section" below.

## Running Script With 1 Class Section
Follow the set-up steps above the same way without downloading the online roster. Modify `main()` to look like the following.

```
def main():
    # get and parse rosters
    campusRoster = getStudentRosterFromExcel('CampusRoster.xlsx')
    onlineRoster = []

    # count TAs given TA roster
    # totalTAs, newTAs, returningTAs = countAndReturnTAs('taRoster.txt')
    totalTAs, newTAs, returningTAs = countAndReturnTAs('exampleTARoster.txt')

    # number of TAs for campus/online section
    numCampusTAs, numOnlineTAs = calculateTADistribution(campusRoster, onlineRoster, totalTAs)

    # create TA groups
    groupingList = createTAGroups(newTAs, returningTAs)
    campusTAs, onlineTAs = separateGroups(groupingList, numCampusTAs, numOnlineTAs)
    campusTAs.reverse()

    # create sheet
    workbook = xlsxwriter.Workbook('HomeworkGroups.xlsx')
    createFrontSheet(workbook, campusTAs, onlineTAs)
    createGroupSheets(workbook, campusTAs, numCampusTAs, campusRoster, "campus/hybrid")
    workbook.close()
```