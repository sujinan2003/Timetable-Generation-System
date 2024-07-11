import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# กำหนดขอบเขตของข้อมูลที่ต้องการใช้
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)

client = gspread.authorize(credentials)

mainFile = client.open('Main')
teacherSheet = mainFile.worksheet('Teachers')

#ดึงข้อมูลชื่อ Teachers
branchNamesRange = teacherSheet.range('C3:C')
branchNames = [cell.value for cell in branchNamesRange if cell.value]

#เปิดไฟล์ Teacher
teacherFile = client.open('Teacher')

for branch_name in branchNames:
    newSheetName = f'{branch_name}'
    newSheet = teacherFile.add_worksheet(title=newSheetName, rows='9', cols='13')
    header1 = ['', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']
    header2 = ['วัน/เวลา', '08:00 - 09:00', '09:00 - 10:00', '10:00 - 11:00', '11:00 - 12:00', '12:00 - 13:00', '13:00 - 14:00', '14:00 - 15:00', '15:00 - 16:00', '16:00 - 17:00', '17:00 - 18:00', '18:00 - 19:00', '19:00 - 20:00']
    headerDay = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']

    #เตรียมข้อมูล for batch update
    batch_data = []
    batch_data.append({'range': 'A1:M1', 'values': [header1]})
    batch_data.append({'range': 'A2:M2', 'values': [header2]})
    for i, day in enumerate(headerDay, start=3):
        batch_data.append({'range': f'A{i}', 'values': [[day]]})
        
    newSheet.batch_update(batch_data)
    
    time.sleep(1)

print("Success!")
