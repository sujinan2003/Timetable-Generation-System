import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# กำหนดขอบเขตของข้อมูลที่ต้องการใช้
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)

client = gspread.authorize(credentials)

# เปิดไฟล์ Curriculum, Curriculum_General Education Program, และ Open Course
curriculumFile = client.open('Curriculum')
generalEdFile = client.open('Curriculum_General Education Program')
openCourseFile = client.open('Open Course')

# ดึงข้อมูลจากทุกชีทในไฟล์ Curriculum
curriculumSheets = curriculumFile.worksheets()
curriculumData = []
for sheet in curriculumSheets:
    curriculumData.extend(sheet.get_all_records())

# ดึงข้อมูลจากทุกชีทในไฟล์ Curriculum_General Education Program
generalEdSheets = generalEdFile.worksheets()
generalEdData = []
for sheet in generalEdSheets:
    generalEdData.extend(sheet.get_all_records())

# ดึงข้อมูลจากทุกชีทในไฟล์ Open Course
openCourseSheets = openCourseFile.worksheets()
openCourseData = {}
for sheet in openCourseSheets:
    sheet_data = sheet.get_all_records()
    openCourseData[sheet.title] = sheet_data  # แยกข้อมูลตามชื่อชีท เช่น CS_Y1

# เปิดไฟล์ Open Course2
openCourseFile2 = client.open('Open Course2')

# สร้างชีทใหม่หรือเลือกชีทที่มีอยู่แล้วใน Open Course2 สำหรับแต่ละปีการศึกษาและสาขา
for sheet_name, records in openCourseData.items():
    try:
        # ตรวจสอบว่าชีทมีอยู่แล้วหรือไม่ ถ้ามีแล้วไม่ต้องสร้างใหม่
        if sheet_name not in [sheet.title for sheet in openCourseFile2.worksheets()]:
            newSheet = openCourseFile2.add_worksheet(title=sheet_name, rows='100', cols='8')
            headers = ['เซคเรียน', 'รหัสวิชา', 'รหัสอาจารย์', 'ชื่อวิชา', 'หมวดหมู่รายวิชา', 'หน่วยกิต (ทฤษฎี-ปฏิบัติ-ศึกษาด้วยตนเอง)', 'คาบเรียน (บรรยาย)', 'คาบเรียน (ปฏิบัติ)']
            newSheet.insert_row(headers, index=1)
        else:
            newSheet = openCourseFile2.worksheet(sheet_name)

        # เพิ่มข้อมูลในชีทที่มีอยู่แล้ว
        for record in records:
            course_id = record['รหัสวิชา']
            # ค้นหาข้อมูลที่ตรงกันใน Curriculum และ General Education Program
            matching_records = [r for r in curriculumData + generalEdData if r['รหัสวิชา'] == course_id]

            if not matching_records:
                continue  # ถ้าไม่มีข้อมูลตรงกันในไฟล์ Curriculum หรือ General Education Program ให้ข้ามไป

            # จับคู่ข้อมูลจาก Curriculum/General Education Program
            matched_record = matching_records[0]  # ใช้ข้อมูลที่ตรงกัน (เลือกอันแรก)
            section = record['เซคเรียน']
            teacher_id = record['รหัสอาจารย์']
            course_name = matched_record['ชื่อวิชา']
            category = matched_record['หมวดหมู่รายวิชา']
            credits = matched_record['หน่วยกิต (ทฤษฎี-ปฏิบัติ-ศึกษาด้วยตนเอง)']
            lecture_hours = matched_record['คาบเรียน (บรรยาย)']
            practice_hours = matched_record['คาบเรียน (ปฏิบัติ)']

            # เพิ่มข้อมูลในชีทใหม่
            data = [section, course_id, teacher_id, course_name, category, credits, lecture_hours, practice_hours]
            newSheet.append_row(data)  # ใช้ append_row เพื่อเพิ่มแถวใหม่
            
            # หน่วงเวลาเพื่อหลีกเลี่ยง Rate Limit
            time.sleep(1)

    except Exception as e:
        print(f"Error processing sheet {sheet_name}: {e}")

print("Success!")
