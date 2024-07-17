import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# กำหนดขอบเขตของข้อมูลที่ต้องการใช้
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)

client = gspread.authorize(credentials)

mainFile = client.open('Main')
teacherSheet = mainFile.worksheet('Teachers')

# ดึงรหัส Teachers
idTeacherRange = teacherSheet.range('C3:C')
idTeacher = [cell.value for cell in idTeacherRange if cell.value]

# ดึงชื่อ Teachers
nameTeacherRange = teacherSheet.range('D3:D')
nameTeacher = [cell.value for cell in nameTeacherRange if cell.value]

# Pair idTeacher and nameTeacher
teacher_data = zip(idTeacher, nameTeacher)

# เปิดไฟล์ Teacher
teacherFile = client.open('Teacher')

for branch_id, teacher_name in teacher_data:
    newSheetName = f'{branch_id}'
    newSheet = teacherFile.add_worksheet(title=newSheetName, rows='10', cols='13')
    headerTitle = [f'ตารางสอน {teacher_name}']
    headerPeriod = ['', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']
    headerTime = ['วัน/เวลา', '08:00 - 09:00', '09:00 - 10:00', '10:00 - 11:00', '11:00 - 12:00', '12:00 - 13:00', '13:00 - 14:00', '14:00 - 15:00', '15:00 - 16:00', '16:00 - 17:00', '17:00 - 18:00', '18:00 - 19:00', '19:00 - 20:00']
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
