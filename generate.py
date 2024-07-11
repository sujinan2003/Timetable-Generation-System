import random
import gspread
import time

from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope) 

# เชื่อมต่อ Google Sheets API
client = gspread.authorize(credentials)

# ดึงข้อมูล TimeSlot จากไฟล์ Main
mainFile = client.open('Main')
timeSlotSheet = mainFile.worksheet('TimeSlot')
timeSlots = [cell.value for cell in timeSlotSheet.range('C3:C14')]

# ดึงข้อมูลห้องเรียนจากไฟล์ Main
roomSheet = mainFile.worksheet('Room')
rooms = [cell.value for cell in roomSheet.range('G3:G') if cell.value]

# ดึงข้อมูล Curriculum จากไฟล์ Curriculum
curriculumFile = client.open('Curriculum')
curriculum = []
for sheet in curriculumFile.worksheets():
    curriculum.extend(sheet.get_all_records())

# ดึงข้อมูล Open Course จากไฟล์ Open Course
openCourseFile = client.open('Open Course')
openCourseSheets = openCourseFile.worksheets()

courses = []
for sheet in openCourseSheets:
    course_codes = [cell.value for cell in sheet.range('A2:A') if cell.value]
    sections = [cell.value for cell in sheet.range('B2:B') if cell.value]
    teachers = [cell.value for cell in sheet.range('C2:C') if cell.value]
    for course_code, section, teacher in zip(course_codes, sections, teachers):
        courses.append({
            'รหัสวิชา': course_code,
            'เซคเรียน': section,
            'อาจารย์': teacher
        })
    
        # เพิ่มการหน่วงเวลา
        time.sleep(1)  # หน่วงเวลา 1 วินาที

# แสดงผลลัพธ์ที่ดึงมา
#print("Time Slots:", timeSlots)
#print("Rooms:", rooms)
#print("Curriculum:", curriculum)
#print("Courses:", courses)