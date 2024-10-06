import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope) 
client = gspread.authorize(credentials)

mainFile = client.open('Main')

# เลือกชีท Curriculum ใน Main ไปดึงชื่อ สาขา
curriculumSheet = mainFile.worksheet('Curriculum')
curriculumData = [row[1] for row in curriculumSheet.get_all_values()[2:] if row[1]] # ดึงข้อมูลในแถวB ตั้งแต่ index 2 (คอลัม3)


# ใช้ไฟล์ Curriculum เพื่อสร้างตาราง
newCurriculumFile = client.open('Curriculum')

# สร้างชีทใหม่สำหรับแต่ละหลักสูตร
for i, courseID in enumerate(curriculumData):
    newSheetName = f'{courseID}'
    newSheet = newCurriculumFile.add_worksheet(title=newSheetName, rows='10', cols='6') # มี 6 หัวข้อ
    headers = ['รหัสวิชา', 'ชื่อวิชา', 'หมวดหมู่รายวิชา', 'หน่วยกิต (บรรยาย-ปฏิบัติ-ศึกษาด้วยตนเอง)', 'คาบเรียน (บรรยาย)', 'คาบเรียน (ปฏิบัติ)']
    newSheet.insert_row(headers, index=1)

    # หน่วง 1 วิ
    time.sleep(1)  

    # ใช้ batch update แทนการอัปเดตทีละเซลล์
    blank_data = [[''] * len(headers) for _ in range(2, 11)]
    newSheet.update('A2:F11', blank_data)

print("Success!")

# ไฟล์สำหรับหลักสูตรวิชาในคณะ 