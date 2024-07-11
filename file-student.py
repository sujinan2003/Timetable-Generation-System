import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)

client = gspread.authorize(credentials)

mainFile = client.open('Main')
studentSheet = mainFile.worksheet('Students')

#ดึงข้อมูลจำนวนชั้นปี จาก D3
numGrades = int(studentSheet.acell('D3').value)

#ดึงข้อมูลชื่อสาขาทั้งหมด D11:D18
branchNamesRange = studentSheet.range('D11:D18')
branchNames = [cell.value for cell in branchNamesRange if cell.value]

#เปิดไฟล์ Student
studentFile = client.open('Student')

#สร้างชีทใหม่ แต่ละชื่อสาขาและชั้นปี
for grade in range(1, numGrades + 1):
    for branch_name in branchNames:
        newSheetName = f'{branch_name}_Y{grade}'
        newSheet = studentFile.add_worksheet(title=newSheetName, rows='9', cols='13')
        header1 = ['', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']
        header2 = ['วัน/เวลา', '08:00 - 09:00', '09:00 - 10:00', '10:00 - 11:00', '11:00 - 12:00', '12:00 - 13:00', '13:00 - 14:00', '14:00 - 15:00', '15:00 - 16:00', '16:00 - 17:00', '17:00 - 18:00', '18:00 - 19:00', '19:00 - 20:00']
        headerDay = ['จันทร์', 'อังคาร', 'พุธ' , 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']

        #ใช้ batch update สำหรับการเพิ่มข้อมูลลงในชีทใหม่
        batch_data = []
        batch_data.append({'range': 'A1:M1', 'values': [header1]})
        batch_data.append({'range': 'A2:M2', 'values': [header2]})
        for i, day in enumerate(headerDay, start=3):
            batch_data.append({'range': f'A{i}', 'values': [[day]]})

        #เพิ่มข้อมูลลงในชีทใหม่
        newSheet.batch_update(batch_data)

        #หน่วงเวลา 1 วิ 
        time.sleep(1) 

print("Success!")
