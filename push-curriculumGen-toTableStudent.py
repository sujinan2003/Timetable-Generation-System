import gspread
from oauth2client.service_account import ServiceAccountCredentials

# เชื่อมต่อกับ Google Sheets API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)
client = gspread.authorize(credentials)

# เปิดไฟล์และชีทต่างๆ
openCourseFile = client.open('Open Course2')
curriculumFile = client.open('Curriculum_General Education Program')
studentFile = client.open('Student')

# ดึงข้อมูลจากทุกชีทใน Open Course2
openCourseData = {}
for sheet in openCourseFile.worksheets():
    openCourseData[sheet.title] = sheet.get_all_values()

# ดึงข้อมูลจากทุกชีทใน Curriculum_General Education Program
curriculumData = {}
for sheet in curriculumFile.worksheets():
    curriculumData[sheet.title] = sheet.get_all_values()

# ดึงชื่อสาขาและชั้นปี
branchNames = ['CS', 'IT', 'SE', 'AAI']  # แก้ไขให้ตรงกับชื่อสาขาจริง
numGrades = 4  # จำนวนชั้นปี

def prepare_data_for_updates(open_course_data, curriculum_data, section_names, student_sheet):
    batch_data = []

    # ดึงข้อมูลตารางเรียนทั้งหมดในชีท "Student"
    all_data = student_sheet.get_all_values()

    # วนลูปสำหรับแต่ละเซคเรียนที่ต้องการอัปเดต
    for section in section_names:
        # ค้นหาหัวตารางที่มีชื่อเซคที่ต้องการในชีท "Student"
        for row_idx in range(len(all_data)):
            if len(all_data[row_idx]) > 0 and all_data[row_idx][0] == f"ตารางเรียน {section}":  # ตรวจสอบหัวตาราง
                # ดึงข้อมูลตารางของเซคที่ตรงกัน
                start_row = row_idx + 3  # สมมติว่าตารางเริ่มต้นที่แถวที่ 3 หลังหัวตาราง
                period_data = all_data[start_row:start_row + 12]  # ตัวอย่างการดึง 12 แถวตารางเรียนของเซค

                # วนลูปสำหรับวิชาที่มีใน Open Course
                for course in open_course_data:
                    if len(course) < 2:
                        continue

                    course_section = course[0]
                    course_code = course[1]
                    course_credit = course[5]  # หน่วยกิตอยู่ในคอลัมน์ E 
                    
                    # ตรวจสอบให้แน่ใจว่า course_section ตรงกับ section ที่กำลังอัปเดต
                    if course_section == section:
                        # ค้นหาข้อมูลวิชาใน Curriculum
                        matching_courses = [row for row in curriculum_data if len(row) > 0 and row[0] == course_code]
                        
                        for match in matching_courses:
                            if len(match) < 12:
                                continue

                            lecture_days = match[6]  # วันเรียนบรรยาย
                            lecture_start = int(match[7])  # คาบบรรยาย(เริ่ม)
                            lecture_end = int(match[8])  # คาบบรรยาย(จบ)
                            practical_days = match[9]  # วันเรียนปฏิบัติ
                            practical_start = int(match[10])  # คาบปฏิบัติ(เริ่ม)
                            practical_end = int(match[11])  # คาบปฏิบัติ(จบ)

                            # เตรียมข้อมูลอัปเดตในรูปแบบ batch
                            days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']
                            for day in lecture_days.split(','):
                                day = day.strip()
                                if day in days:
                                    col = days.index(day) + 2  # ตัวอย่างเริ่มคอลัมน์ที่ c
                                    for period in range(lecture_start, lecture_end + 1):
                                        row = start_row + period - 1  # เพิ่มแถวตามที่ต้องการ
                                        batch_data.append({'range': f'{chr(65 + col)}{row}', 'values': [[f'{course_code}\n{course_credit}']]})

                            for day in practical_days.split(','):
                                day = day.strip()
                                if day in days:
                                    col = days.index(day) + 1
                                    for period in range(practical_start, practical_end + 1):
                                        row = start_row + period - 1
                                        batch_data.append({'range': f'{chr(65 + col)}{row}', 'values': [[f'{course_code}\n{course_credit}']]})
    
    return batch_data

def update_student_sheet(sheet, batch_data):
    sheet.batch_update(batch_data)

def main():
    # ตรวจสอบและจัดการข้อมูล
    for grade in range(1, numGrades + 1):
        for branch_name in branchNames:
            newSheetName = f'{branch_name}_Y{grade}'
            
            # ตรวจสอบการมีอยู่ของชีท
            try:
                studentSheet = studentFile.worksheet(newSheetName)
            except gspread.exceptions.WorksheetNotFound:
                print(f"Sheet '{newSheetName}' not found in Student file.")
                continue

            # ดึงข้อมูลจำนวนเซคเรียนจากไฟล์ Open Course
            openCourseSheetName = f'{branch_name}_Y{grade}'
            openCourseSheet = openCourseFile.worksheet(openCourseSheetName)
            sectionNames = list(set(openCourseSheet.col_values(1)[1:]))  # กำจัดข้อมูลซ้ำ

            curriculum_sheet_name = f'{branch_name}_{64 + grade - 1}'  # ชื่อชีทที่ต้องการดึงข้อมูลใน Curriculum
            
            if curriculum_sheet_name in curriculumData:  # ตรวจสอบว่าชีทมีอยู่หรือไม่
                batch_data = prepare_data_for_updates(
                    openCourseData.get(openCourseSheetName, []), 
                    curriculumData[curriculum_sheet_name], 
                    sectionNames, 
                    studentSheet
                )
                
                # ทำการ batch update ข้อมูลตารางเรียน
                if batch_data:
                    print(f"Updating sheet: {newSheetName} for branch {branch_name} grade {grade}")
                    update_student_sheet(studentSheet, batch_data)
            else:
                print(f"Sheet '{curriculum_sheet_name}' not found in Curriculum_General Education Program.")
    
    print("Success!")
    
if __name__ == "__main__":
    main()
