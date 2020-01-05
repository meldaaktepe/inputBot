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
requsite_class_witmissprops = []

report_file = open("Reports/report_file.txt", 'w')

class_file_path, course_file_path, opl_file_path = "", "", ""
cuurentterm = ""

def createCourseList(subjectName, CRN, Building, Room, enrolment, weekdays, beginTime, endTime, doubleCoded, PROP, req_classroom) :
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

                        meetingInformation = meatingInfo(Building, Room, day, beginTime, endTime, courseItem, req_classroom)
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
                    meetingInformation = meatingInfo(Building, Room, day, beginTime, endTime, courseItem, req_classroom)
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

            for i in range(len(weekdays)):
                if weekdays[i] != "":
                    day = weekdays[i]

                    # Create meeting information list
                    meetingInformation = meatingInfo(Building, Room, day, beginTime, endTime, len(courseList), req_classroom)
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

def lesseonParse(file):

    for i in range(len(file)):

        subjCode = file.loc[i]['Subj Code']
        courseNumber = file.loc[i]['Crse Numb']
        sectionNumber = file.loc[i]['Section Numb']
        CRN = file.loc[i]['CRN']
        Building = file.loc[i]['Building']
        Room = file.loc[i]['Room']
        enrolment = file.loc[i]['SSBSECT_ENRL']
        dayM = file.loc[i]['MON']
        dayT = file.loc[i]['TUE']
        dayW = file.loc[i]['WED']
        dayR = file.loc[i]['THU']
        dayF = file.loc[i]['FRI']
        dayS = file.loc[i]['SAT']
        beginTime = file.loc[i]['Begin Time']
        endTime = file.loc[i]['End Time']
        doubleCoded = file.loc[i]['Double Coded']
        PROP = file.loc[i]['PROP']
        requsite_class = file.loc[i]['FORCETIMEROOM']

        CRN = int(float(CRN))
        if pd.isnull(file.loc[i, 'Building']):
            Building = "Null"
        if pd.isnull(file.loc[i, 'Room']):
            Room = "Null"
        enrolment = int(float(enrolment))
        if pd.isnull(file.loc[i, 'MON']):
            dayM = ""
        if pd.isnull(file.loc[i, 'TUE']):
            dayT = ""
        if pd.isnull(file.loc[i, 'WED']):
            dayW = ""
        if pd.isnull(file.loc[i, 'THU']):
            dayR = ""
        if pd.isnull(file.loc[i, 'FRI']):
            dayF = ""
        if pd.isnull(file.loc[i, 'SAT']):
            dayS = ""
        weekdays = [dayM, dayT, dayW, dayR, dayF, dayS]
        if pd.isnull(file.loc[i, 'Begin Time']):
            beginTime = "Null"
        else :
            beginTime = int(float(beginTime))
        if pd.isnull(file.loc[i, 'End Time']):
            endTime = "Null"
        else:
            endTime = int(float(endTime))
        if pd.isnull(file.loc[i, 'Double Coded']):
            doubleCoded = ""
        if pd.isnull(file.loc[i, 'PROP']):
            PROP = ""

        # if prop is not Null // if there is a special request
        if (PROP != ""):
            PROP = PROP.split(";")
            newprop = []
            for index in PROP:
                newprop.append(index.lstrip())  # " PC"  To seperate this i used lstrip to strip from lefth
            PROP = newprop

        if pd.isnull(file.loc[i, 'FORCETIMEROOM']):
            requsite_class = ""
        else:
            requsite_class = requsite_class.split(";")
            requsite_class[-1] = requsite_class[-1].strip("\n")

        # Make New SubjectName
        subjectName = subjCode + " - " + courseNumber + " - " + sectionNumber + " - " + str(CRN)
        # startTime and endTime 900 or smilar
        if (beginTime != "Null" or endTime != "Null"):
            if (beginTime == 900 or beginTime == 1000 or beginTime == 1100 or beginTime == 1200
                    or beginTime == 1300 or beginTime == 1400 or beginTime == 1500 or beginTime == 1600
                    or beginTime == 1700 or beginTime == 1800 or beginTime == 1900 or beginTime == 2000
                    or beginTime == 2100 or beginTime == 2200 or beginTime == 2300):
                begin = [beginTime]
                beginTime = beginTime - 60
                begin.append(beginTime)
                report = "Course Name :" + subjectName + " starts at " + str(begin[0]) + ", we asume that will start at " + str(begin[1]) + "\n"
                report_file.write(report)

            elif (beginTime == 930 or beginTime == 1030 or beginTime == 1130 or beginTime == 1230
                  or beginTime == 1330 or beginTime == 1430 or beginTime == 1530 or beginTime == 1630
                  or beginTime == 1730 or beginTime == 1830 or beginTime == 1930 or beginTime == 2030
                  or beginTime == 2130 or beginTime == 2230) :
                begin = [beginTime]
                beginTime = beginTime + 10
                begin.append(beginTime)
                report = "Course Name :" + subjectName + " starts at " + str(begin[0]) + ", we asume that will start at " + str(begin[1]) + "\n"
                report_file.write(report)

            elif (beginTime == 950 or beginTime == 1050 or beginTime == 1150 or beginTime == 1250
                  or beginTime == 1350 or beginTime == 1450 or beginTime == 1550 or beginTime == 1650
                  or beginTime == 1750 or beginTime == 1850 or beginTime == 1950 or beginTime == 2050
                  or beginTime == 2150 or beginTime == 2250) :
                begin = [beginTime]
                beginTime = beginTime - 10
                begin.append(beginTime)
                report = "Course Name :" + subjectName + " starts at " + str(begin[0]) + ", we asume that will start at " + str(begin[1]) + "\n"
                report_file.write(report)

            if (endTime == 900 or endTime == 1000 or endTime == 1100 or endTime == 1200
                    or endTime == 1300 or endTime == 1400 or endTime == 1500 or endTime == 1600
                    or endTime == 1700 or endTime == 1800 or endTime == 1900 or endTime == 2000
                    or endTime == 2100 or endTime == 2200 or endTime == 2300):
                end = [endTime]
                endTime = endTime + 30
                end.append(endTime)
                report = "Course Name :" + subjectName + " ends at " + str(end[0]) + ", we asume that will end at " + str(end[1]) + "\n"
                report_file.write(report)

            elif (endTime == 950 or endTime == 1050 or endTime == 1150 or endTime == 1250
                    or endTime == 1350 or endTime == 1450 or endTime == 1550 or endTime == 1650
                    or endTime == 1750 or endTime == 1850 or endTime == 1950 or endTime == 2050
                    or endTime == 2150 or endTime == 2250 or endTime == 2350):
                end = [endTime]
                endTime = endTime + 80
                end.append(endTime)
                report = "Course Name :" + subjectName + " ends at " + str(end[0]) + ", we asume that will end at " + str(end[1]) + "\n"
                report_file.write(report)
            elif (endTime == 940 or endTime == 1040 or endTime == 1140 or endTime == 1240
                  or endTime == 1340 or endTime == 1440 or endTime == 1540 or endTime == 1640
                  or endTime == 1740 or endTime == 1840 or endTime == 1940 or endTime == 2040
                  or endTime == 2140 or endTime == 2240 or endTime == 2340):
                end = [endTime]
                endTime = endTime - 10
                end.append(endTime)
                report = "Course Name :" + subjectName + " ends at " + str(end[0]) + ", we asume that will end at " + str(end[1]) + "\n"
                report_file.write(report)


        else :
            report = "Course Name :" + subjectName + " starts at " + str(beginTime) + ", ends at " + str(endTime) + "\n"
            report_file.write(report)

        # if day is weakend or building is empty or room is empty or building of campus
        # or beginning time is empty or ending time is empty dont added to the list

        if dayS != "S" and (Building != "Null" and Building != "KCC" and Building != "UC"
                and Room != "Null" and Room != "G013-14" and Room != "CAFE"
                and beginTime != "Null" and endTime != "Null"
                and subjCode !="CIP" and (Building + Room) not in studiAndLab) :
            if "FRT" in requsite_class or "FR" in requsite_class:
                if subjectName in requsite_classdic:
                    classes = []
                    temp = requsite_classdic[subjectName]
                    if type(temp) is str :
                        classes.append(temp)
                    elif type(temp) is list :
                        for i in temp:
                            classes.append(i)
                    classes.append((Building + Room))
                    requsite_classdic[subjectName] = classes
                elif subjectName not in requsite_classdic:
                    requsite_classdic[subjectName] = (Building + Room)
            createCourseList(subjectName, CRN, Building, Room, enrolment, weekdays, beginTime, endTime, doubleCoded, PROP, requsite_class)

        elif dayS == "S" or ((Building == "Null" or Building == "KCC" or Building == "UC")
                or (Room == "Null" or Room == "G013-14" or Room == "CAFE")
                 or beginTime == "Null" or endTime == "Null"
                or subjCode == "CIP" or (Building + Room) in studiAndLab) :
            if (Building + Room) in studiAndLab :
                report = "Subject Name :" + subjectName + ", in building " + Building + " in room " + Room + "day " + str(weekdays) + " starts at " + str(beginTime) + " ends at " + str(endTime) + "in " + studiAndLab[Building + Room] + "\n"
                report_file.write(report)
            else :
                report = "Subject Name : " + subjectName + " in building " + Building + " in room " + Room + "day " + str(weekdays) + " starts at " + str(beginTime) + " ends at " + str(endTime) + "\n"
                report_file.write(report)

def classroomParse(file):

    '#Take Features'
    properties = []
    for _ in range(len(file)):
        properties.append(file["RDEF_CODE"][_])
    '# Delete dublicate features'
    unicProperties = set(properties)
    '# construct a dictionary'
    propertyDictionary = dict.fromkeys(unicProperties, False)

    classRoompropertyDictionary = {}
    prewRoom = ""

    file['CAPACITY'].astype(int)

    for _ in range((len(file))):
        building = file["BLDG"][_]
        room = str(file["ROOM"][_])
        des = file["CLASS_DESC"][_]
        room_type = file["CLASS_TYPE"][_]
        capacity = file["CAPACITY"][_]
        feature_code = file["RDEF_CODE"][_]
        # feature_des = file["RDEF_DESC"][_]

        if room_type != "LAB" and room_type != "STUD":
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
                classroomInfos = classroomInfo(building, room, des, room_type, capacity, classRoompropertyDictionary)
                classroomList.append(classroomInfos)
        else:
            studiAndLab[building + room] = room_type
            courseCapacity[classRoom] = capacity

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
        print(i.getPROP())
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

def findclass(courseProps, clas, course, requsite_class):
    #print("findclass says hi")
    for k in classroomList:
        classprops = k.getClassFeatures()
        if requsite_class == "" or "FT" in requsite_class :
            if courseProps != "":
                numerOfProps = len(courseProps)
                count = 0
                for props in courseProps:
                    if props in classprops :
                        if classprops[props] == True:
                            #clas[classroomList.index(k)] = 1
                            count += 1
                        elif classprops[props] == False:
                            report = "Course, " + course.getCrnList()[0].getSubjName() + " course Prop:" + props + " does not exist in Classroom props in classromm"  + k.getClassName() + "\n"
                            report_file.write(report)
                    else:
                        report =  "Course, " + course.getCrnList()[0].getSubjName(), " asks for:" + str(courseProps) + ". However classroom:" + k.getClassName() + " doesn't have these properties." + "\n"
                        report_file.write(report)
                if count == numerOfProps:
                    clas[classroomList.index(k)] = 1
            else: #if courseprops == "Null":
                clas[classroomList.index(k)] = 1
        elif "FRT" in requsite_class or "FR" in requsite_class:
            coursenamelist = []
            for i in course.getCrnList():
                coursenamelist.append((i.getSubjName()))
            for lessonname in coursenamelist:
                if lessonname in requsite_classdic:
                    classnames = requsite_classdic[lessonname]
                    if type(classnames) is str:
                        if classnames == k.getClassName():
                            clas[classroomList.index(k)] = 1
                    elif type(classnames) is list:
                        for classname in classnames:
                            if classname == k.getClassName():
                                clas[classroomList.index(k)] = 1

    return clas

def findclass2(building, room, clas):
    for k in classroomList:
        classname = (building + room)
        if classname == k.getClassName():
            clas[classroomList.index(k)] = 1
    return clas

def makeAitAndCij(term) :
    temp = out_text.get(1.0, "end")
    temp += "\n" + "makeAitAndCij"
    out_text.delete(1.0, "end")
    out_text.insert(1.0, temp)

    aITm, aITt, aITw, aITr, aITf = [], [], [], [], []
    QI, dim, dit, diw, dir, dif = [], [], [], [], [], []
    clasRooms, classRoomName = [], []
    cijm, cijt, cijw, cijr, cijf = [], [], [], [], []  # course to time
    daily_time = []

    '# Construct aIT list'
    hoursStart = [840, 940, 1040, 1140, 1240, 1340, 1440, 1540, 1640, 1740, 1840, 1940, 2040, 2140, 2240]
    hoursFinish = [930, 1030, 1130, 1230, 1330, 1430, 1530, 1630, 1730, 1830, 1930, 2030, 2130, 2230, 2330]
    '# Create Time array'
    for _ in range(0, 15):
        daily_time.append(0)

    daily_time.append(0)  # crn
    daily_time.append(0)  # name
    daily_time.append(0)  # begintime
    daily_time.append(0)  # endtime

    '# Create cIj and QI'
    for i in classroomList:
        clasRooms.append(0)
        classRoomName.append(i.getClassName())
        QI.append(i.getClassCapacity())
    clasRooms.append(0)  # crn
    clasRooms.append(0)  # name

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

        for k in meetinglist:

            new_daily_time = daily_time.copy()

            LessonEnd = hoursFinish.index(k.getEndTime())
            LessonBegin = hoursStart.index(k.getBeginTime())

            new_daily_time[LessonBegin] = 1
            new_daily_time[LessonEnd] = 1

            for daily in range(LessonBegin, LessonEnd):
                new_daily_time[daily] = 1

            new_daily_time[-4] = i.getCrnList()[0].getcrn()
            new_daily_time[-3] = i.getCrnList()[0].getSubjName()
            new_daily_time[-2] = k.getBeginTime()
            new_daily_time[-1] = k.getEndTime()

            clas_daily = clasRooms.copy()
            clas_daily = findclass(courseProps, clas_daily, i, k.getReq_classroom())
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

    newQI = pd.DataFrame(QI)
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
    temp = out_text.get(1.0, "end")
    temp += "\n" + "End of Make-ait amd cij"
    out_text.delete(1.0, "end")
    out_text.insert(1.0, temp)

def objectifFunction(term) :
    print(term)
    # kapasit-emnrolment
    print(courseCapacity.keys())
    sumtotal, sumM, sumt, sumw, sumr, sumf = 0, 0, 0, 0, 0, 0
    for i in courseList:
        enrolment = i.getTotalEnrolment()
        meetinglist = i.getMeetingList()
        for meeting in meetinglist :
            coursename = meeting.getname()
            if coursename in courseCapacity :
                result = (courseCapacity[coursename] - enrolment)
                if meeting.getDay()  is "M":
                    sumM += result
                elif meeting.getDay()  is "T":
                    sumt += result
                elif meeting.getDay()  is "W":
                    sumw += result
                elif meeting.getDay()  is "R":
                    sumr += result
                elif meeting.getDay()  is "F":
                    sumf += result
                sumtotal += result
            else :
                print("Course Name: ", coursename)

    print(sumtotal, sumM, sumt, sumw, sumr, sumf)

def printToExcel(cuurentterm) :

    temp = out_text.get(1.0, "end")
    temp += "\n" + "hi, PrintToExcel"
    out_text.delete(1.0, "end")
    out_text.insert(1.0, temp)

    my_workbook = xlwt.Workbook()
    my_sheet = my_workbook.add_sheet("Solutions",True)

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
    my_sheet.write(0, 18, "FORCETIMEROOM")

    index = 1
    days = ["M", "T", "W", "R", "F"]
    for course in courseList :
        crnList = course.getCrnList()
        meetinglist = course.getMeetingList()
        for lesson in crnList:
            subjectnames = lesson.getSubjName().split(" - ")
            for meeting in meetinglist:

                my_sheet.write(index, 0, cuurentterm)
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
                my_sheet.write(index, 18, ';'.join(meeting.getReq_classroom()))
                index += 1

    my_workbook.save("solution.xls")
    temp = out_text.get(1.0, "end")
    temp += "\n" + "End of PrintToExcel, You can find your output in solutions.xls"
    out_text.delete(1.0, "end")
    out_text.insert(1.0, temp)

def solutions(cuurentterm) :

    temp = out_text.get(1.0, "end")
    temp += "\n" + "Hi, solutions"
    out_text.delete(1.0, "end")
    out_text.insert(1.0, temp)
    solutionWorkBook = xlrd.open_workbook(opl_file_path)

    solutions = []

    solutionSheetM = solutionWorkBook.sheet_by_name("M")
    solutionSheetT = solutionWorkBook.sheet_by_name("T")
    solutionSheetW = solutionWorkBook.sheet_by_name("W")
    solutionSheetR = solutionWorkBook.sheet_by_name("R")
    solutionSheetF = solutionWorkBook.sheet_by_name("F")

    solutions.append(solutionSheetM)
    solutions.append(solutionSheetT)
    solutions.append(solutionSheetW)
    solutions.append(solutionSheetR)
    solutions.append(solutionSheetF)

    count = 1
    for solutionSheet in solutions:
        for row  in range(1, solutionSheet.nrows) :
            for colum in range(solutionSheet.ncols - 16) :
                value = int(solutionSheet.cell_value(row, colum))
                if value == 1 :
                    CourseCRN = int(solutionSheet.cell_value(row, solutionSheet.ncols - 17))
                    course = crn2Course[CourseCRN]
                    classroom = classroomList[colum]

                    for meeting in course.getMeetingList():
                        if count == 1:
                            if meeting.getDay() == "M":
                                meeting.setBuilding(classroom.getClassBuilding())
                                meeting.setRoom(classroom.getClassRoom())
                        if count == 2:
                            if meeting.getDay() == "T":
                                meeting.setBuilding(classroom.getClassBuilding())
                                meeting.setRoom(classroom.getClassRoom())
                        if count == 3:
                            if meeting.getDay() == "W":
                                meeting.setBuilding(classroom.getClassBuilding())
                                meeting.setRoom(classroom.getClassRoom())
                        if count == 4:
                            if meeting.getDay() == "R":
                                meeting.setBuilding(classroom.getClassBuilding())
                                meeting.setRoom(classroom.getClassRoom())
                        if count == 5:
                            if meeting.getDay() == "F":
                                meeting.setBuilding(classroom.getClassBuilding())
                                meeting.setRoom(classroom.getClassRoom())
        count += 1

    temp = out_text.get(1.0, "end")
    temp += "\n" + "calling ... printToExcel"
    out_text.delete(1.0, "end")
    out_text.insert(1.0, temp)
    printToExcel(cuurentterm)

def find_missingProps():
    report_file.write("find missing props")
    count = 0
    for i in courseList:
        proplist = []
        courseProps = i.getPROP()
        meatings = i.getMeetingList()
        for meeting in meatings:
            if meeting.getReq_classroom() is "FT" or meeting.getReq_classroom() is "":
                for j in classroomList:
                    if meeting.getname() == j.getClassName():
                        classprops = j.getClassFeatures()
                        for prop in courseProps:
                            if prop in classprops:
                                if  classprops[prop] == True:
                                    requsite_class_witmissprops.append(courseList.index(i))
                                    proplist.append(prop)
                                else:
                                    count += 1
                                    report = i.getCrnList()[0].getSubjName() + " wants " + prop + " but, class " + j.getClassName() + "'s feature list doesn't have this prop"
                                    report_file.write(report)
        proplist = list(set(proplist))
        if proplist != []:
            i.setPROP(proplist)

'#Main function'
def main(cuurentterm):
    '#For Rooms'
    data_file = pd.read_excel(class_file_path)
    classroomParse(data_file)

    '#For courses'
    data_file = pd.read_excel(course_file_path)

    '#Seperate terms'
    term = data_file.loc[data_file["Term Code"] == int(cuurentterm)]
    new_term = term.drop('Term Code', axis = 1)
    new_term = new_term.reset_index(drop=True)

    crn2Course.clear()
    courseList.clear()
    doubleCode2Course.clear()

    lesseonParse(new_term)
    find_missingProps()
    makeAitAndCij(cuurentterm)

'#For UI'
ui = tk.Tk()
ui.title("Course-Classroom Assignment")

def select_class_file():
    global class_file_path
    class_file_path = filedialog.askopenfilename()
    new_path = class_file_path.rfind("/")
    new_path = class_file_path[new_path + 1 : ]
    class_file_label["text"] = str(new_path)
def select_course_file():
    global course_file_path
    course_file_path= filedialog.askopenfilename()
    new_path = course_file_path.rfind("/")
    new_path = course_file_path[new_path + 1:]
    cource_file_label["text"] = str(new_path)
def select_opl_file():
    global opl_file_path
    opl_file_path = filedialog.askopenfilename()
    new_path = opl_file_path.rfind("/")
    new_path = opl_file_path[new_path + 1:]
    opl_label["text"] = str(new_path)

def null_check():

    if class_file_label["text"] == "" or cource_file_label["text"] == "":
        if class_file_label["text"] == "":
            out_text.insert(1.0, "please chose class File")
        else:
            out_text.delete(1.0, "end")
            out_text.insert(1.0, "please chose course File")
    elif class_file_label["text"] != "" and cource_file_label["text"] != "":
        global cuurentterm
        cuurentterm = term_text.get(1.0, "end")
        if cuurentterm.find('\n') != -1:
            cuurentterm = cuurentterm.strip('\n')
        main(cuurentterm)
def opl_check():
    if opl_label["text"] == "":
        text = out_text.get(1.0, "end")
        out_text.insert(len(text) + 1, "please chose opl output File")
    else:
        solutions(cuurentterm)

canvas = tk.Canvas(ui, height=500, width=600)
canvas.pack()
'For class'
class_button = tk.Button(height='2', width='17', text="Choose classroom file", command=select_class_file)
class_button.pack()
class_button.place(relx="0.050", rely="0.050")

class_file_label = tk.Label(height='2', width='25', bg='white', font=("Times New Roman", "8"))
class_file_label.pack()
class_file_label.place(relx="0.30", rely="0.055")

'For cource'
course_button = tk.Button(height='2', width='17', text="Choose course file", command=select_course_file)
course_button.pack()
course_button.place(relx="0.050", rely="0.15")

cource_file_label = tk.Label(height='2', width='25', bg='white', font=("Times New Roman", "8"))
cource_file_label.pack()
cource_file_label.place(relx="0.30", rely="0.15")

'For importing opl'
opl_button = tk.Button(height='2', width='17', text="Import Solutions", command=select_opl_file)
opl_button.pack()
opl_button.place(relx="0.050", rely="0.25")

opl_label = tk.Label(height='2', width='25', bg='white', font=("Times New Roman", "8"))
opl_label.pack()
opl_label.place(relx="0.30", rely="0.25")

'For term'
Choose_term_label = tk.Label(height='2', width='10', text="Choose term", font=("Times New Roman", "10"))
Choose_term_label.pack()
Choose_term_label.place(relx="0.6", rely="0.050")

term_text = tk.Text(height='1', width='10')
term_text.pack()
term_text.place(relx='0.75', rely='0.050')

'For assignment'
assign_button = tk.Button(height='2', width='24', text="Assign", command=null_check)
assign_button.pack()
assign_button.place(relx="0.6", rely="0.15")

'For solutions'
solutions_button = tk.Button(height='2', width='24', text="Print Solutions", command=opl_check)
solutions_button.pack()
solutions_button.place(relx="0.6", rely="0.25")

'For out text'
out_text = tk.Text(height = '15', width = '64', bg ='white')
out_text.pack()
out_text.place(relx='0.050', rely='0.4')

ui.resizable(False, False)

ui.mainloop()