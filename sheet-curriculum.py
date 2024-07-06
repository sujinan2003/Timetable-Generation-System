import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope) 

#เชื่อมต่อ Google Sheets API
client = gspread.authorize(credentials)

#เปิดไฟล์ Main
mainFile = client.open('Main')

#เลือกชีท 'Curriculum'
curriculumSheet = mainFile.worksheet('Curriculum')

curriculumData = curriculumSheet.range('B3:B10')

#สร้างไฟล์ Curriculum ใหม่
newCurriculumFile = client.open('Curriculum')

#สร้างชีทใหม่สำหรับแต่ละหลักสูตร
for i, cell in enumerate(curriculumData):
    courseID = cell.value
    newSheetName = f'{courseID}'
    newSheet = newCurriculumFile.add_worksheet(title=newSheetName, rows='10', cols='6') #มี6หัวข้อ
    headers = ['รหัสวิชา', 'ชื่อวิชา', 'หมวดหมู่รายวิชา', 'หน่วยกิต (ทฤษฎี-ปฏิบัติ-ศึกษาด้วยตนเอง)', 'คาบเรียน (บรรยาย)', 'คาบเรียน (ปฏิบัติ)']
    newSheet.insert_row(headers, index=1)

    #เพิ่มข้อมูลลงในชีทใหม่
    for col in range(len(headers)):
        newSheet.update_cell(2, col + 1, '')  #ใส่ข้อมูลเป็นช่องว่างเพื่อเตรียมสำหรับข้อมูล

print("Success!")
