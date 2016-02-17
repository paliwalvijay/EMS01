from Tkinter import *
from openpyxl import *
from tkFileDialog import askopenfilename
from operator import *
from datetime import datetime

class Invigilator:
  def __init__(self,email="",name="",noOfExams=0,courses = []):
    self.email = email
    self.name = name
    self.noOfExams = noOfExams
    self.courses = courses

class Resource:
  def __init__(self,answerSheets= 0,graphPapers = 0,extraSheets = 0,tag=0):
    self.answerSheets = answerSheets
    self.graphPapers = graphPapers
    self.extraSheets = extraSheets
    self.tag = tag

class UI:
  def verifyTimeTable():
    return 0
 
  def generateAttendanceSheet():
    return 0

  def getAttendanceSheet():
    return 0

  def mailSeatingPlan():
    return 0

  def notifyMakeupExam():
    return 0

  def updateMakeupDB():
    return 0

  def getInvigilatorDuty():
    return 0

  def sendExamGuidelines():
    return 0

class TimeTable :
  def __init__(self,examList=[],roomList=[]):
    self.examList = examList
    self.roomList = roomList

  def readTimeTable(self,ttFile):
    self.ttBook  = load_workbook(ttFile)
    self.ttSheet = self.ttBook.active
    self.noExams = self.ttSheet['A1'].value
    i = 0;
    for i in range (0,self.noExams):
      rowNo = str(i+2)
      self.examCode = self.ttSheet['A'+rowNo].value
      self.examName = self.ttSheet['B'+rowNo].value
      self.examTime = datetime.strptime(self.ttSheet['C'+rowNo].value, '%d-%m-%Y %I:%M%p')
      self.noOfStudents = self.ttSheet['D'+rowNo].value
      self.exam = Exam(courseTitle = self.examName, courseCode = self.examCode, examTime = self.examTime, noOfStudents = self.noOfStudents)
      self.examList.append(self.exam)
    for i in range (0,self.noExams):
      print self.examList[i].courseTitle,self.examList[i].courseCode,self.examList[i].examTime,self.examList[i].noOfStudents
    return 0

  def readRoomList(self,rlFile):
    self.rlBook  = load_workbook(rlFile)
    self.rlSheet = self.rlBook.active
    self.noRooms = self.rlSheet['A1'].value
    i = 0;
    for i in range (0,self.noRooms):
      rowNo = str(i+2)
      self.roomNo = self.rlSheet['A'+rowNo].value
      self.rows = self.rlSheet['B'+rowNo].value
      self.columns = self.rlSheet['C'+rowNo].value
      self.room = Room(roomNo = self.roomNo, rows = self.rows, columns = self.columns)
      self.roomList.append(self.room)
    for i in range (0,self.noRooms):
     print self.roomList[i].roomNo,self.roomList[i].rows,self.roomList[i].columns
    return 0

  def verifyTimeTable(self):
    self.capacity = 0
    self.totalStudents = 0
    for self.room in self.roomList: 
      self.capacity = self.capacity + (self.room.rows)*(self.room.columns)
    self.newList = sorted(self.examList, key=attrgetter('examTime'), reverse=False)
    self.prevTime = None
    self.slotStudents = 0
    self.valid = 1
    print "Total capacity is : ",self.capacity
    for self.exam in self.newList:
      if(self.exam.examTime == self.prevTime):
        self.slotStudents = self.slotStudents + self.exam.noOfStudents
        #print self.exam.examTime,self.slotStudents
      else:
        if(self.slotStudents > self.capacity):
          print " Error : At slot",self.prevTime,"the total students sitting for exam is",self.slotStudents,"which exceeds capacity",self.capacity
          self.valid = 0
        self.slotStudents = self.exam.noOfStudents
      self.prevTime = self.exam.examTime
    if(self.slotStudents > self.capacity):
      print " Error : At slot",self.prevTime,"the total students sitting for exam is",self.slotStudents,"which exceeds capacity",self.capacity
      self.valid = 0
    #print self.newList[1].examTime,self.newList[2].examTime;
    print "Time-table validity is : ",self.valid,"  ('1' if valid, '0' if invalid)"
    return self.valid

  def getTimeTable(self):
    return 0

class Exam:
  def __init__(self,courseTitle,courseCode,examTime,noOfStudents):
    self.courseTitle = courseTitle
    self.courseCode = courseCode
    self.examTime = examTime
    self.noOfStudents = noOfStudents

class Student:
  def __init__(self,name,rollNo,courseList,email):
    self.name = name
    self.rollNo = rollNo
    self.courseList = courseList  #List of course codes
    self.email = email

class Course:
  def __init__(self,courseTitle,courseCode,studentList,instructor):
    self.courseTitle = courseTitle
    self.courseCode = courseCode
    self.studentList = studentList
    self.instructor = instructor

class Faculty:
  def __init__(self,name,email,courseList):
    self.name = name
    self.email = email
    self.courseList = courseList

  def getCourseList():
    return courseList

  def setCourseList(courseList):
    courseList = courseList

class Room:
  def __init__(self,roomNo,rows,columns,studentList = [[]]):
    self.roomNo = roomNo
    self.rows = rows
    self.columns = columns
    self.studentList = studentList

class Example(Frame):
  def __init__(self,parent):
    Frame.__init__(self,parent)
    self.parent = parent
    print "Inside Example __init__"
    self.initUI()

  def initUI(self):
    self.parent.title("Browse files " )
    fileSelectLBL = Label(self.parent, text = "Please select Time-table file : ")
    fileSelectLBL.pack()
    self.lbl = Label(self.parent,text="Please browse files: ")
    self.lbl.pack()
    print "Creating button now"
    self.ttbutton = Button(self.parent,text = "Browse Time-table", command = self.load_tt)
    self.ttbutton.pack(side="left")
    self.rlbutton = Button(self.parent,text = "Browse Rooms-List", command = self.load_rl)
    self.rlbutton.pack(side="left")
    self.subbutton = Button(self.parent,text = "Submit", command = self.submit)
    self.subbutton.pack(side="left")
    self.f1 = ""
    self.rl = ""

  def load_tt(self, ftypes = None):
    ftypes = [("Excel files","*.xlsx")]
    self.f1 = askopenfilename(filetypes = ftypes)

  def load_rl(self, ftypes = None):
    ftypes = [("Excel files","*.xlsx")]
    self.rl = askopenfilename(filetypes = ftypes)

  def submit(self):
    print self.f1
    print self.rl
    if ((self.f1=="") or (self.rl=="")):
      self.lbl.config(text="Cannot open input files are not provided.")
      self.lbl.pack()
    else :
      self.lbl.config(text="Getting data")
      tt = TimeTable()
      tt.readTimeTable(ttFile = self.f1)
      tt.readRoomList(rlFile = self.rl)
    print "Got !"
    tt.verifyTimeTable()

def main():
  window = Tk()
  print "inside main"
  ex = Example(window)
  window.geometry("900x400")
  window.title("Exam Management Software")
  #i1 = Invigilator(name= "Vijay", email = "paliwal.2@iitj.ac.in",noOfExams = 2 , courses = ["Physics","Maths","COA"])
  #i2 = Invigilator(name = "Dinesh")
  #print i1.name," ",i1.email," ",i1.noOfExams," ",i1.courses;
  #print i2.name," ",i2.email," ",i2.noOfExams," ",i2.courses;
  window.mainloop()

main()
