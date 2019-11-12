class courseInfo:
    def __init__(self, totalEnrolment, PROP, doubleCoded, clas):
        self.totalEnrolment = totalEnrolment
        self.PROP = PROP
        self.doubleCoded = doubleCoded
        self.clas = clas

    def getTotalEnrolment(self):
        return self.totalEnrolment

    def getPROP(self):
        return self.PROP

    def getDoubleCoded(self):
        return self.doubleCoded

    def getMeetingList(self):
        return self.meetingList

    def getCrnList(self):
        return self.crnList

    def getclas(self):
        return self.clas

    def setTotalEnrolment(self, totalEnrolment):
        self.totalEnrolment = totalEnrolment

    def setPROP(self, PROP):
        self.PROP = PROP

    def setDoubleCoded(self, doubleCoded):
        self.doubleCoded = doubleCoded

    def setMeetingList(self, meetingList):
        self.meetingList = meetingList

    def setCrnList(self, crnList):
        self.crnList = crnList

    def setclas(self, clas):
        self.clas = clas


class lessonInfo:
    def __init__(self, subjName, crn, enrolment, courseItem):
        self.subjName = subjName
        self.crn = crn
        self.enrolment = enrolment
        self.courseItem = courseItem

    def getSubjName(self):
        return self.subjName

    def getcrn(self):
        return self.crn

    def getEnrolment(self):
        return self.enrolment

    def getCourseItem(self):
        return self.courseItem


class meatingInfo:
    def __init__(self, Building, Room, day, beginTime, endTime, courseItem):
        self.Building = Building
        self.Room = Room
        self.day = day
        self.beginTime = beginTime
        self.endTime = endTime
        self.courseItem = courseItem

    def getBuilding(self):
        return self.Building

    def getRoom(self):
        return self.Room

    def getDay(self):
        return self.day

    def getBeginTime(self):
        return self.beginTime

    def getEndTime(self):
        return self.endTime

    def getCourseItem(self):
        return self.courseItem

    def getname(self):
        return self.Building + self.Room


class classroomInfo:
    def __init__(self, building, room, des, classroomType, capacity, features):
        self.building = building
        self.room = room
        self.des = des
        self.classroomType = classroomType
        self.capacity = capacity
        self.features = features

    def getClassBuilding(self):
        return self.building

    def getClassRoom(self):
        return self.room

    def getClassDescription(self):
        return self.des

    def getClassclassroomType(self):
        return self.classroomType

    def getClassCapacity(self):
        return self.capacity

    def getClassFeatures(self):
        return self.features

    def getClassName(self):
        ClassName = self.building + self.room
        return ClassName
