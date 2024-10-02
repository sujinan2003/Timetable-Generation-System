import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope) 
client = gspread.authorize(credentials)

# เปิดไฟล์ Main
mainFile = client.open('Main')

# เลือกชีท Curriculum
curriculumSheet = mainFile.worksheet('Curriculum')
curriculumData = [row[1] for row in curriculumSheet.get_all_values()[2:] if row[1]]


newCurriculumFile = client.open('Curriculum_General Education Program')

# สร้างชีทใหม่สำหรับแต่ละหลักสูตร
for i, courseID in enumerate(curriculumData):
    newSheetName = f'{courseID}'
    newSheet = newCurriculumFile.add_worksheet(title=newSheetName, rows='20', cols='12') 
    headers = ['รหัสวิชา', 'ชื่อวิชา', 'หมวดหมู่รายวิชา', 'หน่วยกิต (ทฤษฎี-ปฏิบัติ-ศึกษาด้วยตนเอง)', 'คาบเรียน (บรรยาย)', 'คาบเรียน (ปฏิบัติ)', 'วันเรียนบรรยาย', 'คาบบรรยาย(เริ่ม)', 'คาบบรรยาย(จบ)', 'วันเรียนปฎิบัติ', 'คาบบรรปฎิบัติ(เริ่ม)', 'คาบปฎิบัติ(จบ)']
    newSheet.insert_row(headers, index=1)

    # หน่วงเวลา 1 วินาที
    time.sleep(1)  

    # ใช้ batch update แทนการอัปเดตทีละเซลล์
    blank_data = [[''] * len(headers) for _ in range(19)]
    newSheet.update('A2:L20', blank_data)

print("Success!")
