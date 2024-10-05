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
teacherRangeID = teacherSheet.range('C3:C')
teacherID = [cell.value for cell in teacherRangeID if cell.value]

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

# กำหนด header สำหรับช่วงเวลาและคาบ
headerPeriod = ['คาบ'] + [str(i) for i in range(1, len(period) + 1)]
headerTime = ['เวลา'] + [f'{start} - {end}' for start, end in zip(timeSlotStart, timeSlotEnd)]

# Pair teacherID and nameTeacher
teacherData = zip(teacherID, nameTeacher)

# เปิดไฟล์ Teacher
teacherFile = client.open('Teacher')

for branch_id, teacher_name in teacherData:
    newSheetName = f'{branch_id}'
    newSheet = teacherFile.add_worksheet(title=newSheetName, rows='200', cols='13')
    headerTitle = [f'ตารางสอน {teacher_name}']
    header = ['คาบ', 'เวลา', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']

# ใช้ batch update เพื่อเพิ่มข้อมูลลงในชีทใหม่
    batch_data = []
    row_start = 1  # เริ่มที่แถวแรก
    
    # เพิ่ม Header (คาบ, เวลา, วันต่างๆ)
    batch_data.append({'range': f'A{row_start}:I{row_start}', 'values': [headerTitle]})
    row_start += 1
    batch_data.append({'range': f'A{row_start}:I{row_start}', 'values': [header]})
    row_start += 1 

    # เพิ่มข้อมูลหัวข้อคาบ
    periodData = [[headerPeriod[i], headerTime[i]] + ['' for _ in range(7)] for i in range(1, len(headerPeriod))]
    batch_data.append({'range': f'A{row_start}:I{row_start + len(periodData) - 1}', 'values': periodData})
    row_start += len(periodData)

    # เพิ่มข้อมูลลงในชีทใหม่
    newSheet.batch_update(batch_data)

    time.sleep(1)

print("Success!")
