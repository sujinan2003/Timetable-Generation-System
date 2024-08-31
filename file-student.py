import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)

client = gspread.authorize(credentials)
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

#ดึงข้อมูลจำนวนชั้นปี จากเซลล์ D3
numGrades = int(studentSheet.acell('D3').value)

#ดึงข้อมูลชื่อสาขาทั้งหมด จากช่วงเซลล์ D11:D18
branchNamesRange = studentSheet.range('D11:D18')
branchNames = [cell.value for cell in branchNamesRange if cell.value]

# สร้าง header2 จากช่วงเวลา
header2 = ['วัน/เวลา'] + [f'{start} - {end}' for start, end in zip(timeSlotStart, timeSlotEnd)]

openCourseFile = client.open('Open Course')

studentFile = client.open('Student')

#สร้างชีทใหม่สำหรับแต่ละชื่อสาขาและชั้นปี
for grade in range(1, numGrades + 1):
    for branch_name in branchNames:
        newSheetName = f'{branch_name}_Y{grade}'
        #สร้างชีทใหม่ในไฟล์ 'Student'
        newSheet = studentFile.add_worksheet(title=newSheetName, rows='200', cols='13')  # ปรับจำนวนแถวให้พอสำหรับหลายตาราง
    
        # ดึงข้อมูลจำนวนเซคเรียนจากไฟล์ Open Course
        openCourseSheet = openCourseFile.worksheet(f'{branch_name}_Y{grade}')
        sectionNames = openCourseSheet.col_values(1)[1:]  # สมมติว่าชื่อเซคอยู่ที่คอลัมน์ A เริ่มจากแถวที่ 2
        
        #กำจัดชื่อเซคเรียนที่ซ้ำกัน
        uniqueSectionNames = []
        seen = set()
        for section in sectionNames:
            if section not in seen:
                uniqueSectionNames.append(section)
                seen.add(section)
        
        header1 = [''] + [str(i) for i in range(1, len(period) + 1)]
        headerDay = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']

        #ใช้ batch update เพื่อเพิ่มข้อมูลลงในชีทใหม่
        batch_data = []
        row_start = 1  # เริ่มที่แถวแรก
        for section in uniqueSectionNames:
            batch_data.append({'range': f'A{row_start}:M{row_start}', 'values': [[f'ตารางเรียน {section}']]})
            row_start += 1  #ขยับแถวเริ่มต้นสำหรับ header

            batch_data.append({'range': f'A{row_start}:M{row_start}', 'values': [header1]})
            batch_data.append({'range': f'A{row_start + 1}:M{row_start + 1}', 'values': [header2]})
            for i, day in enumerate(headerDay, start=row_start + 2):
                batch_data.append({'range': f'A{i}', 'values': [[day]]})
            row_start += 10  #ขยับแถวเริ่มต้นสำหรับตารางถัดไป

        #เพิ่มข้อมูลลงในชีทใหม่
        newSheet.batch_update(batch_data)
        
        time.sleep(1)

print("Success!")