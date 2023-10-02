import pdfplumber
import pandas as pd
from datetime import timedelta, datetime
import re
# from ics import Calendar, Event

class Timetable:
    def __init__(self):
        self.__df = None  # the table

        self.__weekDict = {}  # key: week. value: (name, locType, weeksTuple, dayRow, timeColumn).
        self.__dayDict = {}  # key: day of week row number. value: 0 (Monday), 1 (Tuesday), etc.
        # do not need a timeDict because time column number = timedelta / 30 min.

        self.__startTime = None

        self.extractTable()
        # self.outputTable(self.__df)
        self.regularizeDayWeek()
        self.identifyTime()
        self.readSchedule()
        self.createEvents()

    def extractTable(self):  # convert pdf into pandas dataframe
        with pdfplumber.open('./Timetable.pdf') as pdf:
            page = pdf.pages[0]
            data = page.extract_table()
            self.__df = pd.DataFrame(data[1:], columns=data[0])

    def outputTable(self, dataframe):  # for previewing in Excel
        dataframe.to_excel('output.xlsx', index=False)

    def regularizeDayWeek(self):  # for the day of week column, fill None with the currect day of week; map day row number to Mon, Tue in number etc.
        # select a cell: self.__df.iloc[(row # from 0 excl. title, column # from 0)]
        # print(self.__df)  # preview
        dayWeekColumn = self.__df.iloc[:, 0]  # first column incl. title, type: 'pandas.core.series.Series'
        # self.outputTable(dayWeekColumn)
        nullDayOfWeek = list(dayWeekColumn[dayWeekColumn.isnull()].index)
        while len(nullDayOfWeek) != 0:
            if nullDayOfWeek == [0]:  # first row is null, nothing to copy from - unlikely but need prevention
                break
            nullDayOfWeek = list(dayWeekColumn[dayWeekColumn.isnull()].index)  # the Nones in days of week
            for nullDay in nullDayOfWeek:
                dayWeekColumn.iloc[nullDay] = dayWeekColumn.iloc[nullDay - 1]
        # print(self.__df)
        # for each cell in the dayWeekColumn: key = row number, value = 0-6 for Mon-Sun
        dayIsWhichDayOfWeek = ('Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun')
        for rowNumber in range(len(dayWeekColumn)):
            if dayWeekColumn.iloc[rowNumber] in dayIsWhichDayOfWeek:
                self.__dayDict[rowNumber] = dayIsWhichDayOfWeek.index(dayWeekColumn.iloc[rowNumber])
        # print(self.__dayDict)

    def identifyTime(self):  # input starting week and the Monday, work out the weeks
        startDay = input("Please input the starting Monday YYYY-MM-DD. e.g. 2023-08-07: ") or '2023-08-07'
        startWeek = int(input("Please input the starting week number. e.g. 1: ") or 1)
        startTime = int(input("Please input the start time of your schedule, e.g. 8: ") or 8)
        startDaySeparated = startDay.split('-')
        self.__startTime = datetime(int(startDaySeparated[0]), int(startDaySeparated[1]), int(startDaySeparated[2]), startTime, 0)
        # print(self.__startTime)

    def readSchedule(self):
        courseID = -1  # to separate different courses
        coursesColumn = self.__df.iloc[:, 1:]  # all courses
        columnCount = len(coursesColumn.columns)  # number of columns
        coursesList = list(coursesColumn.fillna('').stack())  # fillna prevents empty cells to be deleted
        # print(coursesList)
        # self.outputTable(coursesColumn)
        for index in range(len(coursesList)):
            if coursesList[index] != '':  # if there is course in this cell
                # deal with course details like:
                # CS5222/L  -- name
                # WGB_110 - CS LabLecture  -- location & type, aka locType
                # Wks: 24-33, 36-37  -- weeks
                courseID += 1
                courseDetails = coursesList[index]
                for subIndex in range(len(courseDetails) - 1):
                    if courseDetails[subIndex].islower() and courseDetails[subIndex + 1].isupper():  # to separate 2 words like: LabLecture
                        courseDetails = courseDetails[:(subIndex + 1)] + ' ' + courseDetails[(subIndex + 1):]
                    if courseDetails[subIndex] == ' ' and courseDetails[subIndex + 1] == ' ':  # get rid of excess space symbols
                        courseDetails = courseDetails[:subIndex] + courseDetails[(subIndex + 1):]
                # parse course Details into name, loction + type, weeks
                courseDetailsPattern = '^(.+?)\\n(.+?)\\nWks:.*?(\d.+)$'  # parse the course details
                courseDetailsResult = re.findall(courseDetailsPattern, courseDetails)
                name  = courseDetailsResult[0][0]
                locType = courseDetailsResult[0][1]
                weeks = courseDetailsResult[0][2].replace(' ', '').split(',')  # will be: ['24-33', '36-37']
                # print(name, locType, weeks)
                # deal with course weeks
                weeksList = []
                for weekSlice in weeks:
                    if (re.findall('(\d+)-\d+', weekSlice)):  # if is like \d-\d, not only \d
                        startWeek = int(re.findall('(\d+)-\d+', weekSlice)[0])
                        endWeek = int(re.findall('\d+-(\d+)', weekSlice)[0])
                        weeksList.append((startWeek, endWeek))
                        # print(weeks, startWeek, endWeek)
                    else:
                        weeksList.append(int(weekSlice))
                weeksTuple = tuple(weeksList)  # to save time when querying using dict
                # print(courseDetails, '\n', weeksTuple)
                dayRow = index // columnCount
                timeColumn = index % columnCount
                # add to __weekDict
                self.__weekDict[courseID] = (name, locType, weeksTuple, dayRow, timeColumn)
                # each entry like: 25: ('CS5018/L', 'WGB_107 Lecture', ((24, 33), (36, 37)), 8, 10)
        # print(self.__weekDict)

    def createEvents(self):
        uid = -1
        ical_string = '''BEGIN:VCALENDAR
VERSION:2.0
PRODID:ucctimetable
BEGIN:VTIMEZONE
TZID:Europe/Dublin
BEGIN:STANDARD
DTSTART:20231029T020000
TZOFFSETFROM:+0100
TZOFFSETTO:+0000
TZNAME:GMT
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:20240331T010000
TZOFFSETFROM:+0000
TZOFFSETTO:+0100
TZNAME:IST
END:DAYLIGHT
END:VTIMEZONE\n'''
        # cal = Calendar()
        for course in self.__weekDict.keys():
            for weekSlice in self.__weekDict[course][2]:  # should be a tuple or a rowNumber
                # event = Event()
                uid += 1
                name = self.__weekDict[course][0]
                location = self.__weekDict[course][1]
                # duration = "PT1H"
                if isinstance(weekSlice, tuple):  # if is a week interval
                    firstDateTime = self.__startTime + timedelta(
                    weeks = weekSlice[0] - 1,
                    days = self.__dayDict[self.__weekDict[course][3]],
                    hours = self.__weekDict[course][4] * 0.5
                    )
                    beginTime = firstDateTime.strftime('%Y%m%dT%H%M%SZ')
                    endTime = (firstDateTime + timedelta(hours=1)).strftime('%Y%m%dT%H%M%SZ')

                    ical_string += f'''BEGIN:VEVENT
LOCATION:{location}
SUMMARY:{name}
DTSTART;TZID=Europe/Dublin:{beginTime}
DTEND;TZID=Europe/Dublin:{endTime}
UID:{uid}
RRULE:FREQ=WEEKLY;COUNT={weekSlice[1] - weekSlice[0] + 1}
END:VEVENT
'''
                else:  # is single week
                    onlyDateTime = self.__startTime + timedelta(
                    weeks = weekSlice - 1,
                    days = self.__dayDict[self.__weekDict[course][3]],
                    hours = self.__weekDict[course][4] * 0.5
                    )
                    beginTime = onlyDateTime.strftime('%Y%m%dT%H%M%SZ')
                    endTime = (onlyDateTime + timedelta(hours=1)).strftime('%Y%m%dT%H%M%SZ')

                    ical_string += f'''BEGIN:VEVENT
LOCATION:{location}
SUMMARY:{name}
DTSTART;TZID=Europe/Dublin:{beginTime}
DTEND;TZID=Europe/Dublin:{endTime}
UID:{uid}
END:VEVENT
'''
            # cal.events.add(event)
        #
        # # Save the calendar string to a file
        ical_string += f'END:VCALENDAR'
        with open('course_schedule.ics', 'w') as f:
            f.writelines(ical_string)

        print("Calendar with recurring course schedule created.")


if __name__ == '__main__':
    t = Timetable()
