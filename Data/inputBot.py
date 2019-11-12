import pandas as pd
import xlrd, xlwt

from classes import courseInfo as courseInfo
from classes import lessonInfo as lessonInfo
from classes import meatingInfo as meatingInfo
from classes import classroomInfo as classroomInfo

courseList = []
classroomList = []

doubleCode2Course = {}
crn2Course = {}

def createCourseList(subjectName, CRN, Building, Room, enrolment, weekdays, beginTime, endTime, doubleCoded, PROP) :
    # Chcek given course is exists
    if CRN in crn2Course:  # if exists
        if doubleCoded in doubleCode2Course:  # if it is doubleCoded course

            newLesson = crn2Course[CRN]
            meetingList = newLesson.getMeetingList()
            lessonList = newLesson.getCrnList()

            lessonCRN = lessonList[0].getcrn()
            courseItem = lessonList[0].getCourseItem()

            if lessonCRN == CRN:

                for i in range(len(weekdays)):
                    if weekdays[i] != "":
                        day = weekdays[i]

                        meetingInformation = meatingInfo(Building, Room, day, beginTime, endTime, courseItem)
                        meetingList.append(meetingInformation)
                        newLesson.setMeetingList(meetingList)

        elif doubleCoded not in doubleCode2Course:  # if it isn't doubleCoded course
            # Add new meeting
            newLesson = crn2Course[CRN]

            meetingList = newLesson.getMeetingList()
            courseItem = meetingList[0].getCourseItem()

            for i in range(len(weekdays)):
                if weekdays[i] != "":
                    day = weekdays[i]

                    # Create meeting information list
                    meetingInformation = meatingInfo(Building, Room, day, beginTime, endTime, courseItem)
                    meetingList.append(meetingInformation)
                    newLesson.setMeetingList(meetingList)

    elif CRN not in crn2Course:  # if does not  exists
        # Chcek given course is doublecoded
        if doubleCoded in doubleCode2Course:  # if it is doubleCoded course

            newCourse = doubleCode2Course[doubleCoded]

            newEnrolment = newCourse.getTotalEnrolment()
            newEnrolment = newEnrolment + enrolment
            newCourse.setTotalEnrolment(newEnrolment)
            lessonList = newCourse.getCrnList()
            courseItem = lessonList[0].getCourseItem()

            # Create lesson information
            lessonInformation = lessonInfo(subjectName, CRN, enrolment, courseItem)

            lessonList.append(lessonInformation)
            newCourse.setCrnList(lessonList)

            crn2Course[CRN] = newCourse
            doubleCode2Course[doubleCoded] = newCourse

        elif doubleCoded not in doubleCode2Course:  # if it isn't doubleCoded course

            lessonList = []
            meetingList = []

            # Create course information and unite lessonlist and meetinglist
            courseInformation = courseInfo(enrolment, PROP, doubleCoded)
            # Create lesson information list
            lessonInformation = lessonInfo(subjectName, CRN, enrolment, len(courseList))
            lessonList.append(lessonInformation)

            for i in range(len(weekdays)):#if the given day is is not empty then add it to the lesson's meeting info object

                if weekdays[i] != "":
                    day = weekdays[i]

                    # Create meeting information list
                    meetingInformation = meatingInfo(Building, Room, day, beginTime, endTime, len(courseList))
                    # meatingInfo object then put in the meeting list
                    meetingList.append(meetingInformation)

            courseInformation.setMeetingList(meetingList)
            courseInformation.setCrnList(lessonList)

            courseList.append(courseInformation)
            #courseInformation is registered to crn2Course dictionary via its CRN
            crn2Course[CRN] = courseInformation

            if (doubleCoded != ""):  # if courrse has a doublecode
                '#courseInformation is registered to doubleCode2Course ,aside of crn2Course for the sake of easier access, dictionary via its CRN'
                doubleCode2Course[doubleCoded] = courseInformation

def lesseonParse(fileName, tearm):
    # Open file and read lines
    data_file = open(fileName, 'r')
    content = data_file.readlines()
    # Read content line by line
    for line in content:
        # Parse lines
        subjCode = line[0: line.find(",")]
        line = line[line.find(",") + 1:]

        courseNumber = line[0: line.find(",")]
        line = line[line.find(",") + 1:]

        sectionNumber = line[0: line.find(",")]
        line = line[line.find(",") + 1:]

        CRN = int(float(line[0: line.find(",")]))
        print(CRN)
        line = line[line.find(",") + 1:]

        Building = line[0: line.find(",")]
        line = line[line.find(",") + 1:]
        if (Building == ""):
            Building = "Null"

        Room = line[0: line.find(",")]
        line = line[line.find(",") + 1:]
        if (Room == ""):
            Room = "Null"

        enrolment = int(float(line[0: line.find(",")]))
        line = line[line.find(",") + 1:]

        dayM = line[0: line.find(",")]
        line = line[line.find(",") + 1:]

        dayT = line[0: line.find(",")]
        line = line[line.find(",") + 1:]

        dayW = line[0: line.find(",")]
        line = line[line.find(",") + 1:]

        dayR = line[0: line.find(",")]
        line = line[line.find(",") + 1:]

        dayF = line[0: line.find(",")]
        line = line[line.find(",") + 1:]

        dayS = line[0: line.find(",")]
        line = line[line.find(",") + 1:]

        weekdays = [dayM, dayT, dayW, dayR, dayF, dayS]

        beginTime = int(float(line[0: line.find(",")]))
        line = line[line.find(",") + 1:]
        if (beginTime == ""):
            beginTime = "Null"

        endTime = int(float( line[0: line.find(",")]))
        line = line[line.find(",") + 1:]
        if (endTime == ""):
            endTime = "Null"

        doubleCoded = line[0: line.find(",")]
        line = line[line.find(",") + 1:]
        '''       
         if (doubleCoded == ""):
            doubleCoded = "00"
        '''

        PROP = line[0: line.find(",")]
        line = line[line.find(",") + 1:]
        # if prop is not Null // if there is a special request
        if (PROP != ""):
            PROP = PROP.split(";")
            newprop = []
            for index in PROP:
                newprop.append(index.lstrip())  # " PC"  To seperate this i used lstrip to strip from lefth
            PROP = newprop

        # Make New SubjectName
        subjectName = subjCode + " - " + courseNumber + " - " + sectionNumber + " - " + str(CRN)
        # startTime and endTime 900 or smilar
        if (beginTime != "Null" or endTime != "Null"):
            if (beginTime == 900 or beginTime == 1000 or beginTime == 1100 or beginTime == 1200
                    or beginTime == 1300 or beginTime == 1400 or beginTime == 1500 or beginTime == 1600
                    or beginTime == 1700 or beginTime == 1800 or beginTime == 1900 or beginTime == 2000
                    or beginTime == 2100 or beginTime == 2200 or beginTime == 2300):
                print("In", tearm, "Course Name :\n", subjectName, "starts at", beginTime)
                beginTime = beginTime - 60
                print(" we asume that will start at", beginTime)
            if (endTime == 900 or endTime == 1000 or endTime == 1100 or endTime == 1200
                    or endTime == 1300 or endTime == 1400 or endTime == 1500 or endTime == 1600
                    or endTime == 1700 or endTime == 1800 or endTime == 1900 or endTime == 2000
                    or endTime == 2100 or endTime == 2200 or endTime == 2300):
                print("In", tearm, "Course Name :\n", subjectName, "ends at", endTime)
                endTime = endTime + 30
                print(" we asume that will ends at", endTime)

        # if day is weakend or building is empty or room is empty or building of campus
        # or beginning time is empty or ending time is empty dont added to the list
        if (Building != "Null" and Building != "KCC" and Building != "UC"
                and Room != "Null" and Room != "G013-14" and Room != "CAFE"
                and dayS != "S" and beginTime != "Null" and endTime != "Null"
                and subjCode !="CIP") :
            createCourseList(subjectName, CRN, Building, Room, enrolment, weekdays, beginTime, endTime, doubleCoded, PROP)

        elif (Building == "Null" and Building == "KCC" and Building == "UC"
                and Room == "Null" and Room == "G013-14" and Room == "CAFE"
                and dayS == "S" and beginTime == "Null" and endTime == "Null"
                and subjCode == "CIP") :
            print("In Term", tearm, "Subject Name :", subjectName, "in building", Building,
                  "in room", Room, "day", dayS, "starts at", beginTime, "ends at", endTime)

def classroomParse(fileName):

    data_file = open(fileName, 'r')
    content = data_file.readlines()

    #Take Features
    properties = []
    for line  in content :
        lineSplit = line.split(",")
        featureCode = lineSplit[5]
        properties.append(featureCode)
    #Delete dublicate features
    unicProperties = set(properties)
    #construct a dictionary
    propertyDictionary = {}
    while unicProperties:#took all possible properties as set
        propertyDictionary[unicProperties.pop()] = False

    classRoompropertyDictionary = {}
    prewRoom = ""
    for line in content :
        lineSplit = line.split(",")
        building = lineSplit[0]
        room = lineSplit[1]
        des = lineSplit[2]
        classroomType = lineSplit[3]
        capacity = lineSplit[4]
        featureCode = lineSplit[5]
        classRoom = building + room
        if (prewRoom == classRoom) :
            #aynıyasa üstüne eklenen propertyler güncellencek
            classRoompropertyDictionary[featureCode] = True
        else :
            prewRoom = classRoom
            classRoompropertyDictionary = propertyDictionary.copy()
            '#class obj yarat listeye ekle'
            classRoompropertyDictionary[featureCode] = True
            classroomInfos = classroomInfo(building, room, des, classroomType, int(capacity), classRoompropertyDictionary)
            classroomList.append(classroomInfos)

def printAll() :
    # Print full courselist
    for i in courseList :
        crnList = i.getCrnList()
        meetinglist = i.getMeetingList()
        print("Lesson :")
        for j in crnList:
            print(j.getSubjName())
        print("Meeting :")
        for k in meetinglist:
            print("k Meeting:", ":", k.getBuilding() , ",",k.getRoom() ,",",k.getDay() ,",",k.getBeginTime() , ",",k.getEndTime())
    #pint ClassRoomLİst
    print("ClassRoomList:")
    for i in classroomList  :
        print("ClassRoom :", i.getClassBuilding() , ",",i.getClassRoom(), i.getClassCapacity())

def statistic() :

    new, sumclass, fullall = 0, 0, 0
    m, t, w, r, f, arr = [], [], [], [], [], []
    count100and200, count0and50, count51and100, count201and300 = 0, 0, 0, 0

    for k in classroomList:
        classroomCapacity = k.getClassCapacity()
        sumclass += classroomCapacity
        arr.append(classroomCapacity)
        if 0 < classroomCapacity < 51:
            count0and50 += 1
        elif 50 < classroomCapacity < 101:
            count51and100 += 1
        elif 100 < classroomCapacity < 201:
            count100and200 += 1
        elif 200 < classroomCapacity < 300:
            count201and300 += 1
        elif 301 < classroomCapacity < 400:
            new += 1
    print("capacities:", set(arr))
    print("Number of classrooms capasities between 0 and 50:", count0and50)
    print("Number of classrooms capasities between 51 and 100:", count51and100)
    print("Number of classrooms capasities between 101 and 200:", count100and200)
    print("Number of classrooms capasities between 201 and 300:", count201and300)
    print("Number of classrooms capasities between 301 and 400:", new)
    print("Total number of classrooms:", count0and50 + count51and100 + count100and200 + count201and300 + new)

    new = 0
    count100and200, count0and50, count51and100, count201and300 = 0, 0, 0, 0
    for i in courseList:
        CourseEnrolment = i.getTotalEnrolment()
        if 0 < CourseEnrolment < 51:
            count0and50 += 1
        elif 50 < CourseEnrolment < 101:
            count51and100 += 1
        elif 100 < CourseEnrolment < 201:
            count100and200 += 1
        elif 200 < CourseEnrolment < 300:
            count201and300 += 1
        elif 301 < CourseEnrolment < 400:
            new += 1
        fullall += i.getTotalEnrolment()

        meetings = i.getMeetingList()

        for ik in meetings:
            ikl = ik.getDay()
            if ikl == "M":
                m.append(i.getTotalEnrolment())
            elif ikl == "T":
                t.append(i.getTotalEnrolment())
            elif ikl == "W":
                w.append(i.getTotalEnrolment())
            elif ikl == "R":
                r.append(i.getTotalEnrolment())
            elif ikl == "F":
                f.append(i.getTotalEnrolment())

    print("Number of class capasities between 0 and 50:", count0and50)
    print("Number of class capasities between 51 and 100:", count51and100)
    print("Number of class capasities between 101 and 200:", count100and200)
    print("Number of class capasities between 201 and 300:", count201and300)
    print("Number of class capasities between 301 and 400:", new)
    print("Total number of class:", count0and50 + count51and100 + count100and200 + count201and300 + new)

    summ, sumt, sumw, sumr, sumf = 0, 0, 0, 0, 0
    for im, it, iw, ir, iff in zip(m, t, w, r, f):
        summ += im
        sumt += it
        sumw += iw
        sumr += ir
        sumf += iff

        print(summ, sumt, sumw, sumr, sumf)
        print("full:", summ + sumt + sumw + sumr + sumf, "fullall:", fullall)
        print(len(courseList), len(m) + len(t) + len(w) + len(r) + len(f))
        print("classroom:", sumclass)

def findclass(CourseEnrolment, courseProps, clas, i):
    print("findclass says hi")
    for k in classroomList:
        classroomCapacity = k.getClassCapacity()
        classprops = k.getClassFeatures()

        if CourseEnrolment <= classroomCapacity:
            if courseProps != "":
                numerOfProps = len(courseProps)
                count = 0
                for props in courseProps:

                    if classprops[props] is True:
                        clas[classroomList.index(k)] = 1
                        count += 1
                if count == numerOfProps:
                    clas[classroomList.index(k)] = 1

            else:  # if courseprops == "Null":
                clas[classroomList.index(k)] = 1
        clas[-1] = i.getCrnList()[0].getcrn()
    return clas

def makeAitAndCij () :
    cIJ = []  #course to classroom
    aITm, aITt, aITw, aITr, aITf = [], [], [], [], [] #course to time
    '# Construct aIT list'
    time = []
    hoursStart = [840, 940, 1040, 1140, 1240, 1340, 1440, 1540, 1640, 1740, 1840, 1940, 2040, 2140, 2240]
    hoursFinish = [930, 1030, 1130, 1230, 1330, 1430, 1530, 1630, 1730, 1830, 1930, 2030, 2130, 2230, 2330]
    '# Create Time array'
    for i in range(0, 15):
        time.append(0)
    time.append(0)  # crn

    QI, dim, dit, diw, dir, dif = [], [], [], [], [], []
    # Create cIj
    clasRooms, classRoomName = [], []
    cijm, cijt, cijw, cijr, cijf = [], [], [], [], []  # course to time
    for i in classroomList:
        clasRooms.append(0)
        classRoomName.append(i.getClassName())
    clasRooms.append(0)  # crn
    cijm.append(classRoomName)
    cijt.append(classRoomName)
    cijw.append(classRoomName)
    cijr.append(classRoomName)
    cijf.append(classRoomName)

    for i in courseList:
        '#for ait'
        meetinglist = i.getMeetingList()
        newTime = time.copy()
        '#for cij'
        CourseEnrolment = i.getTotalEnrolment()
        courseProps = i.getPROP()
        clas = clasRooms.copy()
        # print(courseProps, type(courseProps), len(courseProps))

        for k in meetinglist:

            LessonEnd = hoursFinish.index(k.getEndTime())
            LessonBegin = hoursStart.index(k.getBeginTime())

            newTime[LessonBegin] = 1
            newTime[LessonEnd] = 1

            for j in range(LessonBegin, LessonEnd):
                newTime[j] = 1
            newTime[-1] = i.getCrnList()[0].getcrn()

            if k.getDay() is "M":
                aITm.append(newTime)
                clas = findclass(CourseEnrolment, courseProps, clas, i)
                cijm.append(clas)
                dim.append(CourseEnrolment)
            elif k.getDay() is "T":
                aITt.append(newTime)
                clas = findclass(CourseEnrolment, courseProps, clas, i)
                cijt.append(clas)
                dit.append(CourseEnrolment)
            elif k.getDay() is "W":
                aITw.append(newTime)
                clas = findclass(CourseEnrolment, courseProps, clas, i)
                cijw.append(clas)
                diw.append(CourseEnrolment)
            elif k.getDay() is "R":
                aITr.append(newTime)
                clas = findclass(CourseEnrolment, courseProps, clas, i)
                cijr.append(clas)
                dir.append(CourseEnrolment)
            elif k.getDay() is "F":
                aITf.append(newTime)
                clas = findclass(CourseEnrolment, courseProps, clas, i)
                cijf.append(clas)
                dif.append(CourseEnrolment)

    for classRoomCapacity in classroomList:
        QI.append(classRoomCapacity.getClassCapacity())

    # Write them to excel file

    newaITm = pd.DataFrame(aITm)
    newaITt = pd.DataFrame(aITt)
    newaITw = pd.DataFrame(aITw)
    newaITr = pd.DataFrame(aITr)
    newaITf = pd.DataFrame(aITf)

    newcijm = pd.DataFrame(cijm)
    newcijt = pd.DataFrame(cijt)
    newcijw = pd.DataFrame(cijw)
    newcijr = pd.DataFrame(cijr)
    newcijf = pd.DataFrame(cijf)

    newdim = pd.DataFrame(dim)
    newdit= pd.DataFrame(dit)
    newdiw = pd.DataFrame(diw)
    newdir = pd.DataFrame(dir)
    newdif = pd.DataFrame(dif)

    newQI = pd.DataFrame(QI)

    with pd.ExcelWriter('outputdaily.xlsx') as writer:
        newaITm.to_excel(writer, "aITm", header = False, index = False)
        newaITt.to_excel(writer, "aITt", header = False, index = False)
        newaITw.to_excel(writer, "aITw", header = False, index = False)
        newaITr.to_excel(writer, "aITr", header = False, index = False)
        newaITf.to_excel(writer, "aITf", header = False, index = False)

        newcijm.to_excel(writer, "cijm", header = False, index = False)
        newcijt.to_excel(writer, "cijt", header = False, index = False)
        newcijw.to_excel(writer, "cijw", header = False, index = False)
        newcijr.to_excel(writer, "cijr", header = False, index = False)
        newcijf.to_excel(writer, "cijf", header = False, index = False)

        newdim.to_excel(writer, "dim", header = False, index = False)
        newdit.to_excel(writer, "dit", header = False, index = False)
        newdiw.to_excel(writer, "diw", header = False, index = False)
        newdir.to_excel(writer, "dir", header = False, index = False)
        newdif.to_excel(writer, "dif", header = False, index = False)

        newQI.to_excel(writer, "QI", header=False, index=False)

def objectifFunction() :
    #kapasit-emnrolment
    for i in courseList :
        enrolment = i.getTotalEnrolment()

def printToExcel() :
    print("hi, PrintToExcel")
    my_workbook = xlwt.Workbook()
    my_sheet = my_workbook.add_sheet("My Sheet", True)

    '#Titles'
    my_sheet.write(0, 0, "Term Code")
    my_sheet.write(0, 1, "Subj Code")
    my_sheet.write(0, 2, "Crse Numb")
    my_sheet.write(0, 3, "Section Numb")
    my_sheet.write(0, 4, "CRN")
    my_sheet.write(0, 5, "Building")
    my_sheet.write(0, 6, "Room")
    my_sheet.write(0, 7, "SSBSECT_ENRL")
    my_sheet.write(0, 8, "MON")
    my_sheet.write(0, 9, "TUE")
    my_sheet.write(0, 10, "WED")
    my_sheet.write(0, 11, "THU")
    my_sheet.write(0, 12, "FRI")
    my_sheet.write(0, 13, "SAT")
    my_sheet.write(0, 14, "Begin Time")
    my_sheet.write(0, 15, "End Time")
    my_sheet.write(0, 16, "Double Coded")
    my_sheet.write(0, 17, "PROP")

    index = 1
    days = ["M", "T", "W", "R", "F"]
    for course in courseList :
        crnList = course.getCrnList()
        meetinglist = course.getMeetingList()
        for lesson in crnList:
            subjectnames = lesson.getSubjName().split(" - ")
            for meeting in meetinglist:

                my_sheet.write(index, 0, "201701")
                my_sheet.write(index, 1, subjectnames[0])
                my_sheet.write(index, 2, subjectnames[1])
                my_sheet.write(index, 3, subjectnames[2])
                my_sheet.write(index, 4, subjectnames[3])
                my_sheet.write(index, 5, meeting.getBuilding())
                my_sheet.write(index, 6, meeting.getRoom())
                my_sheet.write(index, 7, lesson.getEnrolment())

                my_sheet.write(index, days.index(meeting.getDay()) + 8, meeting.getDay())

                my_sheet.write(index, 14, meeting.getBeginTime())
                my_sheet.write(index, 15, meeting.getEndTime())
                my_sheet.write(index, 16, course.getDoubleCoded())

                my_sheet.write(index, 17, ';'.join(course.getPROP()))
                index += 1

    my_workbook.save("solution.xls")
    print("End Of Me!!!...")

def solutions() :

    print("solutions")
    file_location = "Data/write.xlsx"
    print("Hi, solutions")
    solutionWorkBook = xlrd.open_workbook(file_location)

    solutionSheet = solutionWorkBook.sheet_by_index(0)

    for row  in range(solutionSheet.nrows) :
        for colum in range(solutionSheet.ncols - 1) :
            value = int(solutionSheet.cell_value(row, colum))
            if value == 1 :
                CourseCRN = solutionSheet.cell_value(row, solutionSheet.ncols - 1)
                course = crn2Course[CourseCRN]

                classroom = classroomList[colum]

                for neeting in course.getMeetingList():
                    neeting.setBuilding(classroom.getClassBuilding())
                    neeting.setRoom(classroom.getClassRoom())

    print("calling ... printToExcel")
    printToExcel()
#Main function

# For Rooms

# Read file from excel
file_location = "Data/derslik_new.xlsx"
#file_locationClassRooms = input("please enter file location of the classrooms excel file")
data_file = pd.read_excel(file_location)

data_file.to_csv("derslik.csv", header = False, index = False)
classroomParse("derslik.csv")

#For courses
file_location = "Data/dersler_new.xlsx"
#file_locationlessons = input("please enter file location of the lessons excel file")
data_file = pd.read_excel(file_location)

#Seperate terms
term201701 = data_file.loc[data_file.TermCode == 201701]
term201702 = data_file.loc[data_file.TermCode == 201702]
term201801 = data_file.loc[data_file.TermCode == 201801]
term201802 = data_file.loc[data_file.TermCode == 201802]
#Drop TermCode colum
term201701 = term201701.drop('TermCode', axis = 1)
term201702 = term201702.drop('TermCode', axis = 1)
term201801 = term201801.drop('TermCode', axis = 1)
term201802 = term201802.drop('TermCode', axis = 1)

#Make excel file to csv file
# header = False = Drops the header of colums
# header = False = Drops the index of rows
term201701.to_csv('term201701.csv', header = False, index = False)
crn2Course.clear()
courseList.clear()
doubleCode2Course.clear()
lesseonParse("term201701.csv", "201701")
makeAitAndCij()
#objectifFunction()
#solutions()
'''
term201702.to_csv("term201702.csv", header = False, index = False)
crn2Course.clear()
courseList.clear()
doubleCode2Course.clear()
#lesseonParse("term201702.csv, "201702")

term201801.to_csv("term201801.csv", header = False, index = False)
crn2Course.clear()
courseList.clear()
doubleCode2Course.clear()
#lesseonParse("term201801.csv, "201801")

term201802.to_csv("term201802.csv", header = False, index = False)
crn2Course.clear()
courseList.clear()
doubleCode2Course.clear()
#lesseonParse("term201802.csv, "201802")
'''
print("Before assignment:")
#printAll()
solutions()
print("After assignment:")
#printAll()
#print("Printing Statistics..")
#statistic()

print("done")
