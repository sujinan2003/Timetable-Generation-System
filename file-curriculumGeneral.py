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
curriculumData = curriculumSheet.range('B3:B10')

newCurriculumFile = client.open('Curriculum_General Education Program')

# สร้างชีทใหม่สำหรับแต่ละหลักสูตร
for i, cell in enumerate(curriculumData):
    courseID = cell.value
    newSheetName = f'{courseID}'
    
    # เพิ่มชีทใหม่ด้วยจำนวนแถว 20 และจำนวนคอลัมน์ 12
    newSheet = newCurriculumFile.add_worksheet(title=newSheetName, rows='20', cols='12')
    
    headers = ['รหัสวิชา', 'ชื่อวิชา', 'หมวดหมู่รายวิชา', 'หน่วยกิต (ทฤษฎี-ปฏิบัติ-ศึกษาด้วยตนเอง)', 'คาบเรียน (บรรยาย)', 'คาบเรียน (ปฏิบัติ)', 'วันเรียนบรรยาย', 'คาบบรรยาย(เริ่ม)', 'คาบบรรยาย(จบ)', 'วันเรียนปฎิบัติ', 'คาบบรรปฎิบัติ(เริ่ม)', 'คาบปฎิบัติ(จบ)']
    newSheet.insert_row(headers, index=1)

    # หน่วงเวลา 1 วินาที
    time.sleep(1)

    # ใช้ batch update แทนการอัปเดตทีละเซลล์
    blank_data = [[''] * len(headers) for _ in range(19)]  # 19 แถวแทนช่วงจากแถว 2 ถึง 20
    newSheet.update('A2:L20',blank_data)

print("Success!")