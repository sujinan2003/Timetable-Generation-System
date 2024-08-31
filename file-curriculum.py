import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope) 
client = gspread.authorize(credentials)

#เปิดไฟล์ Main
mainFile = client.open('Main')

#เลือกชีท Curriculum
curriculumSheet = mainFile.worksheet('Curriculum')
curriculumData = curriculumSheet.range('B3:B10')

newCurriculumFile = client.open('Curriculum')

#สร้างชีทใหม่สำหรับแต่ละหลักสูตร
for i, cell in enumerate(curriculumData):
    courseID = cell.value
    newSheetName = f'{courseID}'
    newSheet = newCurriculumFile.add_worksheet(title=newSheetName, rows='10', cols='6') #มี6หัวข้อ
    headers = ['รหัสวิชา', 'ชื่อวิชา', 'หมวดหมู่รายวิชา', 'หน่วยกิต (ทฤษฎี-ปฏิบัติ-ศึกษาด้วยตนเอง)', 'คาบเรียน (บรรยาย)', 'คาบเรียน (ปฏิบัติ)']
    newSheet.insert_row(headers, index=1)

    #หน่วงเวลา 1 วิ
    time.sleep(1)  

    #ใช้ batch update แทนการอัปเดตทีละเซลล์
    blank_data = [[''] * len(headers) for _ in range(2, 11)]
    newSheet.update('A2:F11', blank_data)


print("Success!")
