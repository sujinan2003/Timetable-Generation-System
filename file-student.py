import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# เชื่อมต่อกับ Google Sheets API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)
client = gspread.authorize(credentials)

# เปิดไฟล์และชีทต่างๆ
mainFile = client.open('Main')
studentSheet = mainFile.worksheet('Students')
timeSlotSheet = mainFile.worksheet('TimeSlot')

# ดึงข้อมูลช่วงเวลา
periodRange = timeSlotSheet.range('D3:D14')
period = [cell.value for cell in periodRange if cell.value]

timeSlotStartRange = timeSlotSheet.range('D3:D14')
timeSlotStart = [cell.value for cell in timeSlotStartRange if cell.value]

timeSlotEndRange = timeSlotSheet.range('E3:E14')
timeSlotEnd = [cell.value for cell in timeSlotEndRange if cell.value]

# ดึงข้อมูลจำนวนชั้นปี จากเซลล์ D3
numGrades = int(studentSheet.acell('D3').value)

# ดึงข้อมูลชื่อสาขาทั้งหมด จากช่วงเซลล์ D11:D18
branchNamesRange = studentSheet.range('D11:D18')
branchNames = [cell.value for cell in branchNamesRange if cell.value]

# กำหนด header สำหรับช่วงเวลาและคาบ
headerPeriod = ['คาบ'] + [str(i) for i in range(1, len(period) + 1)]
headerTime = ['เวลา'] + [f'{start} - {end}' for start, end in zip(timeSlotStart, timeSlotEnd)]

# เปิดไฟล์ที่เกี่ยวข้อง
openCourseFile = client.open('Open Course2')
studentFile = client.open('Student')

# สร้างชีทใหม่สำหรับแต่ละชื่อสาขาและชั้นปี
for grade in range(1, numGrades + 1):
    for branch_name in branchNames:
        newSheetName = f'{branch_name}_Y{grade}'
        
        # สร้างชีทใหม่ในไฟล์ 'Student'
        newSheet = studentFile.add_worksheet(title=newSheetName, rows='200', cols='13')
    
        # ดึงข้อมูลจำนวนเซคเรียนจากไฟล์ Open Course
        openCourseSheet = openCourseFile.worksheet(f'{branch_name}_Y{grade}')
        sectionNames = list(set(openCourseSheet.col_values(1)[1:]))  # กำจัดข้อมูลซ้ำ
        
        # กำหนด Headers และข้อมูลในแต่ละ Section
        header = ['คาบ', 'เวลา', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']

        # ใช้ batch update เพื่อเพิ่มข้อมูลลงในชีทใหม่
        batch_data = []
        row_start = 1  # เริ่มที่แถวแรก
        
        for section in sectionNames:
            # เพิ่มหัวตารางเรียนสำหรับแต่ละ section
            batch_data.append({'range': f'A{row_start}:I{row_start}', 'values': [[f'ตารางเรียน {section}']]})
            row_start += 1  # ขยับแถวเริ่มต้นสำหรับ header

            # เพิ่ม Header (คาบ, เวลา, วันต่างๆ)
            batch_data.append({'range': f'A{row_start}:I{row_start}', 'values': [header]})
            row_start += 1  # ขยับแถวเริ่มต้นสำหรับข้อมูลช่วงเวลา

            # เพิ่มข้อมูลหัวข้อคาบ
            period_data = [[headerPeriod[i], headerTime[i]] + ['' for _ in range(7)] for i in range(1, len(headerPeriod))]
            batch_data.append({'range': f'A{row_start}:I{row_start + len(period_data) - 1}', 'values': period_data})
            row_start += len(period_data)

        # เพิ่มข้อมูลลงในชีทใหม่
        newSheet.batch_update(batch_data)
        
        time.sleep(1)  # รอ 1 วินาทีเพื่อป้องกันปัญหาการร้องขอ API เกิน

print("Success!")
