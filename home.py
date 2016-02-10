from Tkinter import *
from tkFileDialog import askopenfilename

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

class Example(Frame):
  def __init__(self,parent):
    Frame.__init__(self,parent)
    self.parent = parent
    print "Inside Example __init__"
    self.initUI()

  def initUI(self):
    self.parent.title("Browse files " )
    self.fileSelectLBL = Label(self, text = "Please select Time-table file : ")
    self.fileSelectLBL.pack()
    print "Creating button now"
    self.bbutton = Button(self.parent,text = "Browse Time-table", command = self.load_file)
    self.bbutton.pack(side="left")

  def load_file(self, ftypes = None):
    ftypes = [("Excel files","*.xlsx")]
    f1 = askopenfilename(filetypes = ftypes)
    if f1 != "":
      print f1

def main():
  window = Tk()
  print "inside main"
  ex = Example(window)
  window.geometry("300x400")
  window.title("Exam Management Software")
  #i1 = Invigilator(name= "Vijay", email = "paliwal.2@iitj.ac.in",noOfExams = 2 , courses = ["Physics","Maths","COA"])
  #i2 = Invigilator(name = "Dinesh")
  #print i1.name," ",i1.email," ",i1.noOfExams," ",i1.courses;
  #print i2.name," ",i2.email," ",i2.noOfExams," ",i2.courses;
  window.mainloop()

main()
