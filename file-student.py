import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)

client = gspread.authorize(credentials)
mainFile = client.open('Main')
studentSheet = mainFile.worksheet('Students')

#ดึงข้อมูลจำนวนชั้นปี จากเซลล์ D3
numGrades = int(studentSheet.acell('D3').value)

#ดึงข้อมูลชื่อสาขาทั้งหมด จากช่วงเซลล์ D11:D18
branchNamesRange = studentSheet.range('D11:D18')
branchNames = [cell.value for cell in branchNamesRange if cell.value]

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
        sectionNames = openCourseSheet.col_values(2)[1:]  # สมมติว่าชื่อเซคอยู่ที่คอลัมน์ B เริ่มจากแถวที่ 2
        
        #กำจัดชื่อเซคเรียนที่ซ้ำกัน
        uniqueSectionNames = []
        seen = set()
        for section in sectionNames:
            if section not in seen:
                uniqueSectionNames.append(section)
                seen.add(section)
        
        header1 = ['', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']
        header2 = ['วัน/เวลา', '08:00 - 09:00', '09:00 - 10:00', '10:00 - 11:00', '11:00 - 12:00', '12:00 - 13:00', '13:00 - 14:00', '14:00 - 15:00', '15:00 - 16:00', '16:00 - 17:00', '17:00 - 18:00', '18:00 - 19:00', '19:00 - 20:00']
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
