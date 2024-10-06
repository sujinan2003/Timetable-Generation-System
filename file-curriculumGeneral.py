import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope) 
client = gspread.authorize(credentials)


mainFile = client.open('Main')

#เลือกชีท Curriculum ไปดึงชื่อ สาขา
curriculumSheet = mainFile.worksheet('Curriculum')
curriculumData = [row[1] for row in curriculumSheet.get_all_values()[2:] if row[1]] # ดึงข้อมูลในแถวB ตั้งแต่ index 2 (คอลัม3)

# ใช้ไฟล์ Curriculum_General Education Program เพื่อสร้างตาราง
newCurriculumFile = client.open('Curriculum_General Education Program')

# สร้างชีทใหม่สำหรับแต่ละหลักสูตร
for i, courseID in enumerate(curriculumData):
    newSheetName = f'{courseID}'
    newSheet = newCurriculumFile.add_worksheet(title=newSheetName, rows='20', cols='12') 
    headers = ['รหัสวิชา', 'ชื่อวิชา', 'หมวดหมู่รายวิชา', 'หน่วยกิต (บรรยาย-ปฏิบัติ-ศึกษาด้วยตนเอง)', 'คาบเรียน (บรรยาย)', 'คาบเรียน (ปฏิบัติ)', 'วันเรียนบรรยาย', 'คาบบรรยาย(เริ่ม)', 'คาบบรรยาย(จบ)', 'วันเรียนปฎิบัติ', 'คาบบรรปฎิบัติ(เริ่ม)', 'คาบปฎิบัติ(จบ)']
    newSheet.insert_row(headers, index=1)

    # เวลา 1 วิ
    time.sleep(1)  

    # ใช้ batch update แทนการอัปเดตทีละเซลล์
    blank_data = [[''] * len(headers) for _ in range(19)]
    newSheet.update('A2:L20', blank_data)

print("Success!")

# ไฟล์สำหรับหลักสูตรวิชา ศึกษาทั่วไป 
# หน้าตาไฟล์เหมือนกับไฟล์ Curriculum
