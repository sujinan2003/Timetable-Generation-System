import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# กำหนดขอบเขตของข้อมูลที่ต้องการใช้
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)

client = gspread.authorize(credentials)

mainFile = client.open('Main')
teacherSheet = mainFile.worksheet('Teachers')
timeSlotSheet = mainFile.worksheet('TimeSlot')

# ดึงรหัส Teachers
idTeacherRange = teacherSheet.range('C3:C')
idTeacher = [cell.value for cell in idTeacherRange if cell.value]

# ดึงชื่อ Teachers
nameTeacherRange = teacherSheet.range('D3:D')
nameTeacher = [cell.value for cell in nameTeacherRange if cell.value]

# ดึงข้อมูลช่วงเวลา
periodRange = timeSlotSheet.range('D3:D14')
period = [cell.value for cell in periodRange if cell.value]

timeSlotStartRange = timeSlotSheet.range('D3:D14')
timeSlotStart = [cell.value for cell in timeSlotStartRange if cell.value]

timeSlotEndRange = timeSlotSheet.range('E3:E14')
timeSlotEnd = [cell.value for cell in timeSlotEndRange if cell.value]

# สร้าง header2 จากช่วงเวลา
headerTime = ['วัน/เวลา'] + [f'{start} - {end}' for start, end in zip(timeSlotStart, timeSlotEnd)]


# Pair idTeacher and nameTeacher
teacher_data = zip(idTeacher, nameTeacher)

# เปิดไฟล์ Teacher
teacherFile = client.open('Teacher')

for branch_id, teacher_name in teacher_data:
    newSheetName = f'{branch_id}'
    newSheet = teacherFile.add_worksheet(title=newSheetName, rows='10', cols='13')
    headerTitle = [f'ตารางสอน {teacher_name}']
    headerPeriod = [''] + [str(i) for i in range(1, len(period) + 1)]
    headerDay = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์'] 

    # เตรียมข้อมูล for batch update
    batch_data = []
    batch_data.append({'range': 'A1:M1', 'values': [headerTitle]})
    batch_data.append({'range': 'A2:M2', 'values': [headerPeriod]})
    batch_data.append({'range': 'A3:M3', 'values': [headerTime]})
    for i, day in enumerate(headerDay, start=4):
        if i <= 10:  # Ensure the range is within the limits of the sheet
            batch_data.append({'range': f'A{i}:A{i}', 'values': [[day]]}) 

    newSheet.batch_update(batch_data)

    time.sleep(1)

print("Success!")
