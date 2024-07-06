import gspread
from oauth2client.service_account import ServiceAccountCredentials

# กำหนดขอบเขตของข้อมูลที่ต้องการใช้
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)  # เปลี่ยนชื่อ data.json เป็นชื่อไฟล์ของคุณ

#เชื่อมต่อกับ Google Sheets API
client = gspread.authorize(credentials)

#เปิดไฟล์ Main
mainFile = client.open('Main')

#เลือกชีท 'Students'
studentSheet = mainFile.worksheet('Students')

#ดึงข้อมูลจำนวนชั้นปี จาก D3
numGrades = int(studentSheet.acell('D3').value)

#ดึงข้อมูลชื่อสาขาทั้งหมด D11:D18
branchNamesRange = studentSheet.range('D11:D18')
branchNames = [cell.value for cell in branchNamesRange if cell.value]

#เปิด 'Open Course'
openCourseFile = client.open('Open Course')

#สร้างชีทใหม่ แต่ละชื่อสาขาและชั้นปี
for grade in range(1, numGrades + 1):
    for branch_name in branchNames:
        newSheetName = f'{branch_name}_Y{grade}'
        newSheet = openCourseFile.add_worksheet(title=newSheetName , rows='10', cols='3')
        headers = ['รหัสวิชา', 'เซคเรียน', 'รหัสอาจารย์']
        newSheet.insert_row(headers, index=1)

#เพิ่มข้อมูลลงในชีทใหม่
    for col in range(len(headers)):
        newSheet.update_cell(2, col + 1, '')  #ใส่ข้อมูลเป็นช่องว่างเพื่อเตรียมสำหรับข้อมูล

print("Success!")
