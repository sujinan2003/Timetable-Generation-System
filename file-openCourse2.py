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

# ฟังก์ชันตรวจสอบว่าข้อมูลมีอยู่แล้วหรือไม่
def data_exists(sheet, section, course_id, type_class):
    existing_data = sheet.get_all_records()
    for record in existing_data:
        if record['เซคเรียน'] == section and record['รหัสวิชา'] == course_id and record['ประเภท (บรรยาย/ปฏิบัติ)'] == type_class:
            return True
    return False

# ลิสต์สำหรับเก็บข้อมูลเซคชั่นที่ไม่ครบ
incomplete_sections = []

# สร้างชีทใหม่หรือเลือกชีทที่มีอยู่แล้วใน Open Course2 สำหรับแต่ละปีการศึกษาและสาขา
for sheet_name, records in openCourseData.items():
    try:
        # ตรวจสอบว่าชีทมีอยู่แล้วหรือไม่ ถ้ามีแล้วไม่ต้องสร้างใหม่
        if sheet_name not in [sheet.title for sheet in openCourseFile2.worksheets()]:
            newSheet = openCourseFile2.add_worksheet(title=sheet_name, rows='100', cols='8')
            headers = ['เซคเรียน', 'รหัสวิชา', 'ชื่อวิชา', 'หมวดหมู่รายวิชา', 'หน่วยกิต (ทฤษฎี-ปฏิบัติ-ศึกษาด้วยตนเอง)', 'จำนวนชั่วโมง', 'ประเภท (บรรยาย/ปฏิบัติ)', 'รหัสอาจารย์']
            newSheet.insert_row(headers, index=1)
        else:
            newSheet = openCourseFile2.worksheet(sheet_name)

        # ตรวจสอบข้อมูลในชีท
        for record in records:
            course_id = record['รหัสวิชา']
            # ค้นหาข้อมูลที่ตรงกันใน Curriculum และ General Education Program
            matching_records = [r for r in curriculumData + generalEdData if r['รหัสวิชา'] == course_id]

            if not matching_records:
                continue  # ถ้าไม่มีข้อมูลตรงกันในไฟล์ Curriculum หรือ General Education Program ให้ข้ามไป

            # ข้อมูลจาก Curriculum/General Education Program
            matched_record = matching_records[0]  # ใช้ข้อมูลที่ตรงกัน (เลือกอันแรก)
            section = record['เซคเรียน']
            course_name = matched_record['ชื่อวิชา']
            category = matched_record['หมวดหมู่รายวิชา']
            credits = matched_record['หน่วยกิต (ทฤษฎี-ปฏิบัติ-ศึกษาด้วยตนเอง)']
            teacher_id = ''  

            # สำหรับคาบเรียนบรรยาย
            lecture_hours = matched_record['คาบเรียน (บรรยาย)']
            if lecture_hours and not data_exists(newSheet, section, course_id, 'บรรยาย'):
                data = [section, course_id, course_name, category, credits, lecture_hours, 'บรรยาย', teacher_id]
                newSheet.append_row(data)  # เพิ่มแถวที่มีข้อมูลคาบเรียนบรรยาย
            elif not lecture_hours:
                incomplete_sections.append((section, course_id))  # เพิ่มเซคชั่นที่ไม่ครบข้อมูลบรรยาย

            # สำหรับคาบเรียนปฏิบัติ
            practice_hours = matched_record['คาบเรียน (ปฏิบัติ)']
            if practice_hours and not data_exists(newSheet, section, course_id, 'ปฏิบัติ'):
                data = [section, course_id, course_name, category, credits, practice_hours, 'ปฏิบัติ', teacher_id]
                newSheet.append_row(data)  # เพิ่มแถวที่มีข้อมูลคาบเรียนปฏิบัติ
            elif not practice_hours:
                incomplete_sections.append((section, course_id))  # เพิ่มเซคชั่นที่ไม่ครบข้อมูลปฏิบัติ

            time.sleep(3)

    except Exception as e:
        print(f"Error processing sheet {sheet_name}: {e}")

# แสดงผลเซคชั่นที่ยังไม่ครบ
if incomplete_sections:
    print("Incomplete sections found:")
    for section in incomplete_sections:
        print(f"Section: {section[0]}, Course ID: {section[1]}")

print("Success!")
