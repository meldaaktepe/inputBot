import pandas as pd
import xlrd, xlwt

import tkinter as tk
from tkinter import filedialog

from classes import courseInfo as courseInfo
from classes import lessonInfo as lessonInfo
from classes import meatingInfo as meatingInfo
from classes import classroomInfo as classroomInfo

courseList = []
classroomList = []

doubleCode2Course = {}
crn2Course = {}

courseCapacity = {}
studiAndLab = {}

requsite_classdic = {}

'#For UI'
'Something here'


def createCourseList(subjectName, CRN, Building, Room, enrolment, weekdays, beginTime, endTime, doubleCoded, PROP, clas) :
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
            courseInformation = courseInfo(enrolment, PROP, doubleCoded, clas)
            # Create lesson information list
            lessonInformation = lessonInfo(subjectName, CRN, enrolment, len(courseList))
            lessonList.append(lessonInformation)

            for i in range(len(weekdays)):
                if weekdays[i] != "":
                    day = weekdays[i]

                    # Create meeting information list
                    meetingInformation = meatingInfo(Building, Room, day, beginTime, endTime, len(courseList))
                    meetingList.append(meetingInformation)

            courseInformation.setMeetingList(meetingList)
            courseInformation.setCrnList(lessonList)

            courseList.append(courseInformation)

            crn2Course[CRN] = courseInformation

            if (doubleCoded != ""):  # if courrse has a doublecode
                '''
                Ne yapcağını açıkla /// dictionary e ekle
                '''
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
        line = line[line.find(",") + 1:]
        #print("CRN: ", CRN, type(CRN))

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

        beginTime = line[0: line.find(",")]
        line = line[line.find(",") + 1:]
        if (beginTime == ""):
            beginTime = "Null"
        else :
            beginTime = int(float(beginTime))
        endTime = line[0: line.find(",")]
        line = line[line.find(",") + 1:]
        if (endTime == ""):
            endTime = "Null"
        else :
            endTime = int(float(endTime))

        doubleCoded = line[0: line.find(",")]
        line = line[line.find(",") + 1:]

        PROP = line[0: line.find(",")]
        line = line[line.find(",") + 1:]
        # if prop is not Null // if there is a special request
        if (PROP != ""):
            PROP = PROP.split(";")
            newprop = []
            for index in PROP:
                newprop.append(index.lstrip())  # " PC"  To seperate this i used lstrip to strip from lefth
            PROP = newprop
        requsite_class = line

        if (requsite_class != "\n"):
            requsite_class = requsite_class.split(";")
            requsite_class[-1] = requsite_class[-1].strip("\n")
        else :
            requsite_class = ""

        # Make New SubjectName
        subjectName = subjCode + " - " + courseNumber + " - " + sectionNumber + " - " + str(CRN)
        # startTime and endTime 900 or smilar
        if (beginTime != "Null" or endTime != "Null"):
            if (beginTime == 900 or beginTime == 1000 or beginTime == 1100 or beginTime == 1200
                    or beginTime == 1300 or beginTime == 1400 or beginTime == 1500 or beginTime == 1600
                    or beginTime == 1700 or beginTime == 1800 or beginTime == 1900 or beginTime == 2000
                    or beginTime == 2100 or beginTime == 2200 or beginTime == 2300):
                #print("In", tearm, "Course Name :", subjectName, "starts at", beginTime)
                beginTime = beginTime - 60
                #print(", we asume that will start at", beginTime)

            elif (beginTime == 930 or beginTime == 1030 or beginTime == 1130 or beginTime == 1230
                  or beginTime == 1330 or beginTime == 1430 or beginTime == 1530 or beginTime == 1630
                  or beginTime == 1730 or beginTime == 1830 or beginTime == 1930 or beginTime == 2030
                  or beginTime == 2130 or beginTime == 2230) :
                #print("In", tearm, "Course Name :", subjectName, "starts at", beginTime)
                beginTime = beginTime + 10
                #print(", we asume that will start at", beginTime)

            elif (beginTime == 950 or beginTime == 1050 or beginTime == 1150 or beginTime == 1250
                  or beginTime == 1350 or beginTime == 1450 or beginTime == 1550 or beginTime == 1650
                  or beginTime == 1750 or beginTime == 1850 or beginTime == 1950 or beginTime == 2050
                  or beginTime == 2150 or beginTime == 2250) :
                #print("In", tearm, "Course Name :", subjectName, "starts at", beginTime)
                beginTime = beginTime - 10
                #print(", we asume that will start at", beginTime)

            if (endTime == 900 or endTime == 1000 or endTime == 1100 or endTime == 1200
                    or endTime == 1300 or endTime == 1400 or endTime == 1500 or endTime == 1600
                    or endTime == 1700 or endTime == 1800 or endTime == 1900 or endTime == 2000
                    or endTime == 2100 or endTime == 2200 or endTime == 2300):
                #print("In", tearm, "Course Name :", subjectName, "ends at", endTime)
                endTime = endTime + 30
                #print(", we asume that will ends at", endTime)

            elif (endTime == 950 or endTime == 1050 or endTime == 1150 or endTime == 1250
                    or endTime == 1350 or endTime == 1450 or endTime == 1550 or endTime == 1650
                    or endTime == 1750 or endTime == 1850 or endTime == 1950 or endTime == 2050
                    or endTime == 2150 or endTime == 2250 or endTime == 2350):
                #print("In", tearm, "Course Name :", subjectName, "ends at", endTime)
                endTime = endTime + 40
                #print(", we asume that will ends at", endTime)
            elif (endTime == 940 or endTime == 1040 or endTime == 1140 or endTime == 1240
                  or endTime == 1340 or endTime == 1440 or endTime == 1540 or endTime == 1640
                  or endTime == 1740 or endTime == 1840 or endTime == 1940 or endTime == 2040
                  or endTime == 2140 or endTime == 2240 or endTime == 2340):
                #print("In", tearm, "Course Name :", subjectName, "ends at", endTime)
                endTime = endTime - 10
                #print(", we asume that will ends at", endTime)

        else :
            alfa=0
                #print("In", tearm, "Course Name :", subjectName, "starts at", beginTime, ", ends at", endTime)

        # if day is weakend or building is empty or room is empty or building of campus
        # or beginning time is empty or ending time is empty dont added to the list

        if dayS != "S" and (Building != "Null" and Building != "KCC" and Building != "UC"
                and Room != "Null" and Room != "G013-14" and Room != "CAFE"
                and beginTime != "Null" and endTime != "Null"
                and subjCode !="CIP" and (Building + Room) not in studiAndLab) :
            if "FRT" in requsite_class or "FR" in requsite_class:
                requsite_classdic[subjectName] = (Building + Room)
            createCourseList(subjectName, CRN, Building, Room, enrolment, weekdays, beginTime, endTime, doubleCoded, PROP, requsite_class)

        elif dayS == "S" or ((Building == "Null" or Building == "KCC" or Building == "UC")
                or (Room == "Null" or Room == "G013-14" or Room == "CAFE")
                 or beginTime == "Null" or endTime == "Null"
                or subjCode == "CIP" or (Building + Room) in studiAndLab) :
            if (Building + Room) in studiAndLab :
                #print("In Term", tearm, "Subject Name :", subjectName, "in building", Building,"in room", Room, "day", weekdays, "starts at", beginTime, "ends at", endTime, "in", studiAndLab[Building + Room])
                alfa=7
            else :
                #print("In Term", tearm, "Subject Name :", subjectName, "in building", Building,"in room", Room, "day", weekdays, "starts at", beginTime, "ends at", endTime)
                alfa=7

def classroomParse(file):

    '# Take Features'
    properties = []
    for _ in range(len(file)):
        properties.append(file["RDEF_CODE"][_])
    '# Delete dublicate features'
    unicProperties = set(properties)
    # construct a dictionary'
    propertyDictionary =  dict.fromkeys(unicProperties , False)

    classRoompropertyDictionary = {}
    prewRoom = ""

    file['CAPACITY'].astype(int)

    for _ in range((len(file))):
        building = file["BLDG"][_]
        room = file["ROOM"][_]
        des = file["CLASS_DESC"][_]
        type = file["CLASS_TYPE"][_]
        capacity = file["CAPACITY"][_]
        feature_code = file["RDEF_CODE"][_]
        # feature_des = file["RDEF_DESC"][_]

        if type != "LAB" and type != "STUD":
            classRoom = building + room
            if prewRoom == classRoom:
                '#aynıyasa property güncelllencek'
                classRoompropertyDictionary[feature_code] = True
            else:
                prewRoom = classRoom
                courseCapacity[classRoom] = capacity
                classRoompropertyDictionary = propertyDictionary.copy()
                classRoompropertyDictionary[feature_code] = True
                '#class obj yarat listeye ekle'
                classroomInfos = classroomInfo(building, room, des, type, capacity, classRoompropertyDictionary)
                classroomList.append(classroomInfos)
        else:
            studiAndLab[building + room] = type
            courseCapacity[classRoom] = capacity

def printAll() :
    # Print full courselist
    for i in courseList :
        crnList = i.getCrnList()
        meetinglist = i.getMeetingList()
        #print("Lesson :")
        for j in crnList:
            alfa = 0
            #print(j.getSubjName())
            #print("Meeting :")

        for k in meetinglist:
    #print("k Meeting:", ":", k.getBuilding() , ",",k.getRoom() ,",",k.getDay() ,",",k.getBeginTime() , ",",k.getEndTime())
    #pint ClassRoomLİst
    #print("ClassRoomList:")
            alfa=0
    for i in classroomList  :
        alfa = 0
       #print("ClassRoom :", i.getClassBuilding() , ",",i.getClassRoom(), i.getClassCapacity())

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
    #print("capacities:", set(arr))
    #print("Number of classrooms capasities between 0 and 50:", count0and50)
    #print("Number of classrooms capasities between 51 and 100:", count51and100)
    #print("Number of classrooms capasities between 101 and 200:", count100and200)
    #print("Number of classrooms capasities between 201 and 300:", count201and300)
    #print("Number of classrooms capasities between 301 and 400:", new)
    #print("Total number of classrooms:", count0and50 + count51and100 + count100and200 + count201and300 + new)

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

    #print("Number of class capasities between 0 and 50:", count0and50)
    #print("Number of class capasities between 51 and 100:", count51and100)
    #print("Number of class capasities between 101 and 200:", count100and200)
    #print("Number of class capasities between 201 and 300:", count201and300)
    #print("Number of class capasities between 301 and 400:", new)
    #print("Total number of class:", count0and50 + count51and100 + count100and200 + count201and300 + new)

    summ, sumt, sumw, sumr, sumf = 0, 0, 0, 0, 0
    for im, it, iw, ir, iff in zip(m, t, w, r, f):
        summ += im
        sumt += it
        sumw += iw
        sumr += ir
        sumf += iff

        #print(summ, sumt, sumw, sumr, sumf)
        #print("full:", summ + sumt + sumw + sumr + sumf, "fullall:", fullall)
        #print(len(courseList), len(m) + len(t) + len(w) + len(r) + len(f))
        #print("classroom:", sumclass)

def findclass(courseProps, clas, course):
    #print("findclass says hi")
    for k in classroomList:
        classprops = k.getClassFeatures()
        requsite_class = course.getclas()
        if requsite_class == "" or "FT" in requsite_class :
            if courseProps != "":
                numerOfProps = len(courseProps)
                count, classcount = 0, 0
                for props in courseProps:
                    if props in classprops :
                        if classprops[props] == True:
                            clas[classroomList.index(k)] = 1
                            count += 1
                        elif classprops[props] == False:
                            alfa = 0
                            #print("Course,",course.getCrnList()[0].getSubjName(), "course Prop:", props, " does not exist in Classroom props in classromm", k.getClassName())
                    else:
                        alfa = 0
                        #print("Course,", course.getCrnList()[0].getSubjName(), "asks for:", str(courseProps) + ". However classroom:", k.getClassName(), "doesn't have these properties.")
                if count == numerOfProps:
                    clas[classroomList.index(k)] = 1
            else: #if courseprops == "Null":
                clas[classroomList.index(k)] = 1

        else: #if  requsite_class != "":
            coursenamelist = []
            for i in course.getCrnList():
                coursenamelist.append((i.getSubjName()))
            for lessonname in coursenamelist:
                if lessonname in requsite_classdic:
                    if requsite_classdic[lessonname] == k.getClassName():
                        clas[classroomList.index(k)] = 1

    return clas

def makeAitAndCij (term) :
    ait, aITm, aITt, aITw, aITr, aITf = [], [], [], [], [], []
    QI, di, dim, dit, diw, dir, dif = [], [], [], [], [], [], []
    clasRooms, classRoomName = [], []
    cij, cijm, cijt, cijw, cijr, cijf = [], [], [], [], [], []  # course to time
    daily_time, weekly_time = [], []

    '# Construct aIT list'
    days = ["M", "T", "W", "R", "F"]
    hoursStart = [840, 940, 1040, 1140, 1240, 1340, 1440, 1540, 1640, 1740, 1840, 1940, 2040, 2140, 2240]
    hoursFinish = [930, 1030, 1130, 1230, 1330, 1430, 1530, 1630, 1730, 1830, 1930, 2030, 2130, 2230, 2330]
    '# Create Time array'
    for _ in range(0, 15):
        daily_time.append(0)
    for _ in range(0, 15 * 5):
        weekly_time.append(0)

    daily_time.append(0)  # crn
    daily_time.append(0)  # name
    daily_time.append(0)  # begintime
    daily_time.append(0)  # endtime

    weekly_time.append(0)  # crn
    weekly_time.append(0)  # name
    weekly_time.append(0)  # begintime
    weekly_time.append(0)  # endtime

    '# Create cIj and QI'
    for i in classroomList:
        clasRooms.append(0)
        classRoomName.append(i.getClassName())
        QI.append(i.getClassCapacity())
    clasRooms.append(0)  # crn
    clasRooms.append(0)  # name
    cij.append(classRoomName)
    cijm.append(classRoomName)
    cijt.append(classRoomName)
    cijw.append(classRoomName)
    cijr.append(classRoomName)
    cijf.append(classRoomName)

    for i in courseList:
        '#for ait'
        meetinglist = i.getMeetingList()
        '#for cij'
        CourseEnrolment = i.getTotalEnrolment()
        courseProps = i.getPROP()

        new_weekly_time = weekly_time.copy()

        clas_weekly = clasRooms.copy()
        clas_weekly = findclass(courseProps, clas_weekly, i)
        clas_weekly[-2] = i.getCrnList()[0].getcrn()
        clas_weekly[-1] = i.getCrnList()[0].getSubjName()

        # print(courseProps, type(courseProps), len(courseProps))

        for k in meetinglist:

            new_daily_time = daily_time.copy()
            LessonDay = days.index(k.getDay())

            LessonEnd = hoursFinish.index(k.getEndTime())
            LessonBegin = hoursStart.index(k.getBeginTime())

            startIndex = (LessonDay * 15) + (LessonBegin)
            finishIndex = (LessonDay * 15) + (LessonEnd)

            new_daily_time[LessonBegin] = 1
            new_daily_time[LessonEnd] = 1

            new_weekly_time[startIndex] = 1
            new_weekly_time[finishIndex] = 1

            for daily in range(LessonBegin, LessonEnd):
                new_daily_time[daily] = 1
            for weekly in range(startIndex, finishIndex):
                new_weekly_time[weekly] = 1

            new_daily_time[-4] = i.getCrnList()[0].getcrn()
            new_daily_time[-3] = i.getCrnList()[0].getSubjName()
            new_daily_time[-2] = k.getBeginTime()
            new_daily_time[-1] = k.getEndTime()

            new_weekly_time[-4] = i.getCrnList()[0].getcrn()
            new_weekly_time[-3] = i.getCrnList()[0].getSubjName()
            new_weekly_time[-2] = k.getBeginTime()
            new_weekly_time[-1] = k.getEndTime()

            clas_daily = clas_weekly.copy()
            clas_daily[-2] = i.getCrnList()[0].getcrn()
            clas_daily[-1] = i.getCrnList()[0].getSubjName()

            if k.getDay() is "M":
                aITm.append(new_daily_time)
                cijm.append(clas_daily)
                dim.append(CourseEnrolment)
            elif k.getDay() is "T":
                aITt.append(new_daily_time)
                cijt.append(clas_daily)
                dit.append(CourseEnrolment)
            elif k.getDay() is "W":
                aITw.append(new_daily_time)
                cijw.append(clas_daily)
                diw.append(CourseEnrolment)
            elif k.getDay() is "R":
                aITr.append(new_daily_time)
                cijr.append(clas_daily)
                dir.append(CourseEnrolment)
            elif k.getDay() is "F":
                aITf.append(new_daily_time)
                cijf.append(clas_daily)
                dif.append(CourseEnrolment)
        ait.append(new_weekly_time)
        cij.append(clas_weekly)
        di.append(CourseEnrolment)

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
    newdit = pd.DataFrame(dit)
    newdiw = pd.DataFrame(diw)
    newdir = pd.DataFrame(dir)
    newdif = pd.DataFrame(dif)

    newaIT = pd.DataFrame(ait)
    newcij = pd.DataFrame(cij)
    newdi = pd.DataFrame(di)

    newQI = pd.DataFrame(QI)

    excelName = "Output/outputweekly" + term + ".xlsx"
    with pd.ExcelWriter(excelName) as writer_weekly:
        newaIT.to_excel(writer_weekly, "aIT", header=False, index=False)
        newcij.to_excel(writer_weekly, "cij", header=False, index=False)
        newdi.to_excel(writer_weekly, "di", header=False, index=False)
        newQI.to_excel(writer_weekly, "QI", header=False, index=False)

    excelName = "Output/outputdaily" + term + ".xlsx"
    with pd.ExcelWriter(excelName) as writer:
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

        newQI.to_excel(writer, "QI", header = False, index = False)
    #print("End of Make-ait amd cij")

def objectifFunction(term) :
    #print(term)
    # kapasit-emnrolment
    #print(courseCapacity.keys())
    sumtotal, sumM, sumt, sumw, sumr, sumf = 0, 0, 0, 0, 0, 0
    for i in courseList:
        enrolment = i.getTotalEnrolment()
        meetinglist = i.getMeetingList()
        for meeting in meetinglist :
            coursename = meeting.getname()
            if coursename in courseCapacity :
                if meeting.getDay()  is "M":
                    sumM += (courseCapacity[coursename] - enrolment)
                    sumtotal += (courseCapacity[coursename] - enrolment)
                elif meeting.getDay()  is "T":
                    sumt += (courseCapacity[coursename] - enrolment)
                    sumtotal += (courseCapacity[coursename] - enrolment)
                elif meeting.getDay()  is "W":
                    sumw += (courseCapacity[coursename] - enrolment)
                    sumtotal += (courseCapacity[coursename] - enrolment)
                elif meeting.getDay()  is "R":
                    sumr += (courseCapacity[coursename] - enrolment)
                    sumtotal += (courseCapacity[coursename] - enrolment)
                elif meeting.getDay()  is "F":
                    sumf += (courseCapacity[coursename] - enrolment)
                    sumtotal += (courseCapacity[coursename] - enrolment)
            else :
                #print("Course Name: ", coursename)
                alfa = 0
    #print(sumtotal, sumM, sumt, sumw, sumr, sumf)

def printToExcel() :
    #print("hi, PrintToExcel")
    my_workbook = xlwt.Workbook()
    my_sheet = my_workbook.add_sheet("My Sheet",True)

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
    #print("End Of Me!!!...")

def solutions() :

    #print("solutions")
    file_location = "Data/write.xlsx"
    #print("Hi, solutions")
    solutionWorkBook = xlrd.open_workbook(file_location)

    solutionSheet = solutionWorkBook.sheet_by_index(0)

    for row  in range(solutionSheet.nrows) :
        for colum in range(solutionSheet.ncols - 1) :
            value = int(solutionSheet.cell_value(row, colum))
            if value == 1 :
                CourseCRN = int(solutionSheet.cell_value(row, solutionSheet.ncols - 1))
                course = crn2Course[CourseCRN]

                classroom = classroomList[colum]

                for neeting in course.getMeetingList():
                    neeting.setBuilding(classroom.getClassBuilding())
                    neeting.setRoom(classroom.getClassRoom())

    #("calling ... printToExcel")
    printToExcel()


def classparse(file) :
    alfa = 0
    #print(file.columns)
    #print("ssubj", file.loc[0])


#Main function

# For Rooms

# Read file from excel
file_location = "Data/derslik_20190410.xlsx"
#file_locationClassRooms = input("please enter file location of the classrooms excel file")
data_file = pd.read_excel(file_location)

#data_file.to_csv("derslik.csv", header = False, index = False)
classroomParse(data_file)

#For courses
file_location = "Data/dersler_20191108.xlsx"
#file_locationlessons = input("please enter file location of the lessons excel file")
data_file = pd.read_excel(file_location)

#Seperate terms
'''
term201701 = data_file.loc[data_file["Term Code"] == 201701]
term201702 = data_file.loc[data_file["Term Code"] == 201702]
term201801 = data_file.loc[data_file["Term Code"] == 201801]
term201802 = data_file.loc[data_file["Term Code"] == 201802]
#Drop TermCode colum
term201701 = term201701.drop('Term Code', axis = 1)
term201702 = term201702.drop('Term Code', axis = 1)
term201801 = term201801.drop('Term Code', axis = 1)
term201802 = term201802.drop('Term Code', axis = 1)
'''

term201901 = data_file.loc[data_file["Term Code"] == 201901]
term201902 = data_file.loc[data_file["Term Code"] == 201902]
term201901 = term201901.drop('Term Code', axis = 1)
term201902 = term201902.drop('Term Code', axis = 1)

#Make excel file to csv file
# header = False = Drops the header of colums
# header = False = Drops the index of rows

#classparse(term201701)

term201901.to_csv('term201701.csv', header = False, index = False)


lesseonParse("term201701.csv", "201901")
makeAitAndCij("201901")
#print(courseList[])

for a in classroomList:
    alfa=0
    #print(a)
for b in courseList:#b is course info
    #print(b.getTotalEnrolment()  )
    #for x in b.getclas():#this is list type
        #print("getclas() FR/FRT/FT")

    #print("ders props",b.getPROP())
    for x in b.getCrnList():
        #print(x.getSubjName()+" is on ")
        for c in b.getMeetingList():
            #print(c.getname())
            for d in classroomList:
                if d.getClassName() == c.getname():
                    sayaç=True
                    toplam=""
                    for all in b.getPROP():
                        if all == 'VID' or d.getClassFeatures()[all] !=True  :
                            sayaç=False
                            toplam=all+";"+toplam
                    if sayaç:
                        #print(" room has given features:")
                        alpha=0
                    elif sayaç==False :
                        print(x.getSubjName() + " is on " + c.getname() + " this room lacks :" + toplam)
                        print("Course demands "+str(b.getTotalEnrolment())+" enrollment")
                    #print("all features of classroom:")
                    #print(d.getClassFeatures())
        #print(x.getEnrolment())


objectifFunction("201901")
#solutions()

crn2Course.clear()
courseList.clear()
doubleCode2Course.clear()
'''
term201702.to_csv("term201702.csv", header = False, index = False)
crn2Course.clear()
courseList.clear()
doubleCode2Course.clear()
lesseonParse("term201701.csv", "201702")
makeAitAndCij("201702")
objectifFunction("20702")

term201801.to_csv("term201801.csv", header = False, index = False)
crn2Course.clear()
courseList.clear()
doubleCode2Course.clear()
lesseonParse("term201801.csv", "201801")
makeAitAndCij("201801")
objectifFunction()

term201802.to_csv("term201802.csv", header = False, index = False)
crn2Course.clear()
courseList.clear()
doubleCode2Course.clear()
lesseonParse("term201701.csv", "201802")
makeAitAndCij("201802")
objectifFunction()
'''
'''
print("Before assignment:")
#printAll()
#solutions()
print("After assignment:")
#printAll()
#print("Printing Statistics..")
statistic()
'''
print("done")
