import gspread
import time
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

def prepare_data_for_updates(openCourseData, curriculumData, sectionName, studentSheet):
    batch_data = []
    incomplete_sections = []

    # ดึงข้อมูลตารางเรียนทั้งหมดในชีท "Student"
    all_data = studentSheet.get_all_values()

    # วนลูปสำหรับแต่ละเซคเรียนที่ต้องการอัปเดต
    for section in sectionName:
        section_complete = True  # ตั้งค่าเริ่มต้นว่า section นี้สมบูรณ์

        # ค้นหาหัวตารางที่มีชื่อเซคที่ต้องการในชีท "Student"
        for row_idx in range(len(all_data)):
            if len(all_data[row_idx]) > 0 and all_data[row_idx][0] == f"ตารางเรียน {section}":  # ตรวจสอบหัวตาราง
                # ดึงข้อมูลตารางของเซคที่ตรงกัน
                start_row = row_idx + 3  # สมมติว่าตารางเริ่มต้นที่แถวที่ 3 หลังหัวตาราง
                period_data = all_data[start_row:start_row + 12]  # ตัวอย่างการดึง 12 แถวตารางเรียนของเซค

                # วนลูปสำหรับวิชาที่มีใน Open Course
                for course in openCourseData:
                    if len(course) < 2:
                        continue

                    courseSection = course[0]
                    courseID = course[1]
                    courseType = course[4] 
                    courseCredit = course[5]  
                    
                    # ตรวจสอบให้แน่ใจว่า courseSection ตรงกับ section ที่กำลังอัปเดต
                    if courseSection == section:
                        # ค้นหาข้อมูลวิชาใน Curriculum
                        matching_courses = [row for row in curriculumData if len(row) > 0 and row[0] == courseID]
                        
                        if not matching_courses:
                            section_complete = False  # ข้อมูลไม่ครบในส่วนของ curriculum
                            incomplete_sections.append(section)
                            continue  # ถ้าไม่มีข้อมูลวิชาใน Curriculum ข้ามไป

                        for match in matching_courses:
                            if len(match) < 12:
                                section_complete = False  # ข้อมูลใน match ไม่ครบถ้วน
                                incomplete_sections.append(section)
                                continue

                            lectureDay = match[6]  # วันเรียนบรรยาย
                            lectureStart = int(match[7])  # คาบบรรยาย(เริ่ม)
                            lectureEnd = int(match[8])  # คาบบรรยาย(จบ)
                            labDay = match[9]  # วันเรียนปฏิบัติ
                            labStart = int(match[10])  # คาบปฏิบัติ(เริ่ม)
                            labEnd = int(match[11])  # คาบปฏิบัติ(จบ)

                            # เตรียมข้อมูลอัปเดตในรูปแบบ batch
                            days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']
                            for day in lectureDay.split(','):
                                day = day.strip()
                                if day in days:
                                    col = days.index(day) + 2  # ตัวอย่างเริ่มคอลัมน์ที่ C
                                    for period in range(lectureStart, lectureEnd + 1):
                                        row = start_row + period - 1  # เพิ่มแถวตามที่ต้องการ
                                        if not all_data[row][col]:  # ตรวจสอบว่าแถวนี้ว่างหรือไม่
                                            batch_data.append({'range': f'{chr(65 + col)}{row}', 'values': [[f'{courseID}\n{courseCredit}\n{courseType}']]})

                            for day in labDay.split(','):
                                day = day.strip()
                                if day in days:
                                    col = days.index(day) + 2 #+1
                                    for period in range(labStart, labEnd + 1):
                                        row = start_row + period - 1
                                        if not all_data[row][col]:  # ตรวจสอบว่าแถวนี้ว่างหรือไม่
                                            batch_data.append({'range': f'{chr(65 + col)}{row}', 'values': [[f'{courseID}\n{courseCredit}\n{courseType}']]})
        
        if not section_complete:  # หากข้อมูลยังไม่ครบ
            incomplete_sections.append(section)

    return batch_data, incomplete_sections

def update_studentSheet(sheet, batch_data):
    if batch_data:
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

            curriculumSheetName = f'{branch_name}'  # ชื่อชีทที่ต้องการดึงข้อมูลใน Curriculum
            
            if curriculumSheetName in curriculumData:  # ตรวจสอบว่าชีทมีอยู่หรือไม่
                batch_data, incomplete_sections = prepare_data_for_updates(
                    openCourseData.get(openCourseSheetName, []), 
                    curriculumData[curriculumSheetName], 
                    sectionNames, 
                    studentSheet
                )
                
                # ทำการ batch update ข้อมูลตารางเรียน
                update_studentSheet(studentSheet, batch_data)

                if incomplete_sections:
                    print(f"Incomplete data for sections: {', '.join(set(incomplete_sections))}. Please check the data and run again.")
                else:
                    print(f"Updating sheet: {newSheetName} for branch {branch_name} grade {grade}")
                
                time.sleep(5)
            else:
                print(f"Sheet '{curriculumSheetName}' not found in Curriculum_General Education Program.")
    
    print("Success!")

if __name__ == "__main__":
    main()
