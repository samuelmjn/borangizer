from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf
from termcolor import colored
import random

template = "template.xls"

workbook = copy(open_workbook(template, formatting_info=True))
sheet = workbook.get_sheet(0)


class BorangInfo:
    def __init__(self, name, homeGroup, classCode, faculty, students):
        self.name = name
        self.homeGroup = homeGroup
        self.classCode = classCode
        self.faculty = faculty
        self.students = students

print('''                                                          
(  _`\                                   _                    
| (_) )   _    _ __   _ _   ___     __  (_) ____    __   _ __ 
|  _ <' /'_`\ ( '__)/'_` )/' _ `\ /'_ `\| |(_  ,) /'__`\( '__)
| (_) )( (_) )| |  ( (_| || ( ) |( (_) || | /'/_ (  ___/| |   
(____/'`\___/'(_)  `\__,_)(_) (_)`\__  |(_)(____)`\____)(_)   
                                 ( )_) |                      
                                  \___/'                      
    '''
)
print(colored("filling up borang never be this ez", "yellow"))

newBorangInfo = BorangInfo("", "", "", "", [])
newBorangInfo.name = input("Masukan nama: ")
newBorangInfo.homeGroup = input("Masukan HG anda: ")
newBorangInfo.classCode = input("Masukan kode kelas: ")
newBorangInfo.faculty = input("Masukan fakultas anda: ")
print("=" * 20, "Masukan nama anggota kelompok (tidak termasuk anda). Tekan (enter) untuk mengakhiri")
while True:
    newStudent = input("Nama anggota: ")
    if newStudent == "":
        break
    elif len(newBorangInfo.students) > 8:
        print(colored("Sorry, maximum students amount allowed are 8!", "red"))
    newBorangInfo.students.append(newStudent)

print("=" * 20)


sheet.write(11, 2, newBorangInfo.name)
sheet.write(12, 2, newBorangInfo.homeGroup)
sheet.write(13, 2, newBorangInfo.classCode)
sheet.write(14, 2, newBorangInfo.faculty)

start_column = 20
for student in newBorangInfo.students:
    summation = 0
    sheet.write(start_column, 1, student)
    for i in range(1,5):
        score = random.choice([4, 5])
        summation += score
        sheet.write(start_column, i+1, score)

    mean = summation / 4
    sheet.write(start_column, 6, mean)
    start_column += 1

filename = "BorangB1" + "_HG" + newBorangInfo.homeGroup + "_" + newBorangInfo.name + ".xls"
print(colored("Borang berhasil dibuat! Silahkan buka " + filename, "red"))
workbook.save(filename)