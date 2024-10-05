import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)
client = gspread.authorize(credentials)

studentFile = client.open('Student')
generateFile = client.open('Generate')

def prepare_data_for_updates(generate_sheet, student_sheet):
    batch_data = []
    
    # ดึงข้อมูลตารางเรียนทั้งหมดในชีท "Student"
    all_data = student_sheet.get_all_values()

    # วนลูปข้อมูลจากไฟล์ Generate (เช่น เซคเรียน, รหัสวิชา, วันเรียน, คาบ ฯลฯ)
    for course in generate_sheet[1:]:  # ข้ามแถวหัวตาราง
        section = course[0]
        course_code = course[1]
        id_teacher = course[2]
        room = course[3]
        course_type = course[4]
        day = course[5]
        start_period = int(course[6])
        end_period = int(course[7])

        # หาแถวของเซคที่ต้องการใน Student
        start_row = None
        for row_idx, row in enumerate(all_data):
            if len(row) > 0 and row[0] == f"ตารางเรียน {section}":  # ตรวจสอบหัวตารางที่มีชื่อเซค
                start_row = row_idx + 2  # สมมติว่าตารางเริ่มต้นที่แถวที่ 3 หลังหัวตาราง
                break
        
        if start_row is None:
            continue  # ถ้าไม่เจอ section ในตาราง Student ข้ามไป
        
        # วันเรียนจะอยู่ในคอลัมน์ไหน (สมมติว่าเริ่มที่คอลัมน์ C)
        days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']
        
        if day in days:
            col = days.index(day) + 2  # เริ่มที่คอลัมน์ C2
            
            # วนลูปใส่ข้อมูลวิชาลงในตารางสำหรับช่วงคาบที่ระบุ
            for period in range(start_period, end_period + 1):
                row = start_row + period - 1  # แถวที่ต้องการใส่ข้อมูล
                if row < len(all_data) and not all_data[row][col]:  # ถ้าไม่มีข้อมูลในแถวนี้และแถวไม่เกินขอบเขต
                    batch_data.append({
                        'range': f'{chr(65 + col)}{row + 1}',  # A1-based range
                        'values': [[f'{course_code}\n{course_type}\n{room}\n{id_teacher}']]
                    })
    
    return batch_data

# ฟังก์ชันอัปเดตตารางเรียนในชีท Student
def update_student_sheet(sheet, batch_data):
    if batch_data:
        sheet.batch_update([{
            'range': data['range'],
            'values': data['values']
        } for data in batch_data])



def main():
    # รายชื่อสาขาที่ต้องการจัดการ
    branchNames = ['CS', 'IT', 'SE', 'AAI']
    
    for branch in branchNames:  # วนลูปแต่ละสาขา
        for grade in range(1, 5):  # Y1 ถึง Y4
            genSheetName = f'Gen_Y{grade}'  # ชื่อชีทในไฟล์ Generate
            try:
                generateSheet = generateFile.worksheet(genSheetName)
            except gspread.exceptions.WorksheetNotFound:
                print(f"Sheet '{genSheetName}' not found in Generate file.")
                continue
            
            # ดึงข้อมูลจากชีทในไฟล์ Generate
            openCourseData = generateSheet.get_all_values()

            # ตรวจสอบการมีอยู่ของชีทใน Student ที่ต้องการอัปเดต
            studentSheetName = f'{branch}_Y{grade}'  # ตัวอย่าง: ชื่อชีทใน Student ตามสาขา
            try:
                studentSheet = studentFile.worksheet(studentSheetName)
            except gspread.exceptions.WorksheetNotFound:
                print(f"Sheet '{studentSheetName}' not found in Student file.")
                continue

            # เตรียมข้อมูลสำหรับการอัปเดตตาราง
            batch_data = prepare_data_for_updates(openCourseData, studentSheet)
            
            # อัปเดตตารางเรียนใน Student
            update_student_sheet(studentSheet, batch_data)
            
            print(f"Updated sheet: {studentSheetName} for branch {branch} and grade {grade}")
            time.sleep(5)  
            
if __name__ == "__main__":
    main()

