import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# กำหนดขอบเขตของข้อมูลที่ต้องการใช้
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)

client = gspread.authorize(credentials)

mainFile = client.open('Main')
teacherSheet = mainFile.worksheet('Room')
timeSlotSheet = mainFile.worksheet('TimeSlot')

# ดึงข้อมูลชื่อ Teachers
branchNamesRange = teacherSheet.range('C3:C')
branchNames = [cell.value for cell in branchNamesRange if cell.value]

# ดึงข้อมูลช่วงเวลา
periodRange = timeSlotSheet.range('D3:D14')
period = [cell.value for cell in periodRange if cell.value]

timeSlotStartRange = timeSlotSheet.range('D3:D14')
timeSlotStart = [cell.value for cell in timeSlotStartRange if cell.value]

timeSlotEndRange = timeSlotSheet.range('E3:E14')
timeSlotEnd = [cell.value for cell in timeSlotEndRange if cell.value]

# สร้าง header2 จากช่วงเวลา
header2 = ['วัน/เวลา'] + [f'{start} - {end}' for start, end in zip(timeSlotStart, timeSlotEnd)]

# เปิดไฟล์ Teacher
teacherFile = client.open('Room')

for branch_name in branchNames:
    newSheetName = f'{branch_name}'
    newSheet = teacherFile.add_worksheet(title=newSheetName, rows='9', cols='13')
    header1 = [''] + [str(i) for i in range(1, len(period) + 1)]
    headerDay = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']

    # เตรียมข้อมูล for batch update
    batch_data = []
    batch_data.append({'range': 'A1:M1', 'values': [header1]})
    batch_data.append({'range': 'A2:M2', 'values': [header2]})
    for i, day in enumerate(headerDay, start=3):
        batch_data.append({'range': f'A{i}', 'values': [[day]]})

    newSheet.batch_update(batch_data)

    time.sleep(1)

print("Success!")
