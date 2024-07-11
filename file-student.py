import gspread
from oauth2client.service_account import ServiceAccountCredentials

# กำหนดขอบเขตของข้อมูลที่ต้องการใช้
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)

# เชื่อมต่อกับ Google Sheets API
client = gspread.authorize(credentials)

# เปิดไฟล์ Main
mainFile = client.open('Main')

# เลือกชีท 'Students'
studentSheet = mainFile.worksheet('Students')

# ดึงข้อมูลจำนวนชั้นปี จาก D3
numGrades = int(studentSheet.acell('D3').value)

# ดึงข้อมูลชื่อสาขาทั้งหมด D11:D18
branchNamesRange = studentSheet.range('D11:D18')
branchNames = [cell.value for cell in branchNamesRange if cell.value]

# เปิดไฟล์ Students
studentFile = client.open('Student')

# สร้างชีทใหม่ แต่ละชื่อสาขาและชั้นปี
for grade in range(1, numGrades + 1):
    for branch_name in branchNames:
        newSheetName = f'{branch_name}_Y{grade}'
        newSheet = studentFile.add_worksheet(title=newSheetName, rows='9', cols='13')
        header1 = ['', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']
        header2 = ['วัน/เวลา', '08:00 - 09:00', '09:00 - 10:00', '10:00 - 11:00', '11:00 - 12:00', '12:00 - 13:00', '13:00 - 14:00', '14:00 - 15:00', '15:00 - 16:00', '16:00 - 17:00', '17:00 - 18:00', '18:00 - 19:00', '19:00 - 20:00']
        headerDay = ['จันทร์', 'อังคาร', 'พุธ' , 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']
        newSheet.insert_row(header1, index=1)
        newSheet.insert_row(header2, index=2)
        for i, day in enumerate(headerDay, start=3):
            newSheet.update_cell(i, 1, day)

        # เพิ่มข้อมูลลงในชีทใหม่
        blank_data = [[''] * len(header1) for _ in range(3, 11)]  # ปรับจำนวนแถวตามต้องการ
        newSheet.update('A3:M10', blank_data)  # ปรับช่วงข้อมูลตามต้องการ

print("Success!")
