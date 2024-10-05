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

# ดึงข้อมูลจากทุกชีทในไฟล์ Curriculum (วิชาในคณะ)
curriculumSheets = curriculumFile.worksheets()
curriculumData = []
for sheet in curriculumSheets:
    curriculumData.extend(sheet.get_all_records())
 
# ดึงข้อมูลจากทุกชีทในไฟล์ Curriculum_General Education Program (วิชาศึกษาทั่วไป)
generalEdSheets = generalEdFile.worksheets()
generalEdData = []
for sheet in generalEdSheets:
    generalEdData.extend(sheet.get_all_records())

# ดึงข้อมูลจากทุกชีทในไฟล์ Open Course (วิชาที่แต่ละเซคต้องเรียน ในเทอมนั้น ๆ)
openCourseSheets = openCourseFile.worksheets()
openCourseData = {}
for sheet in openCourseSheets:
    sheet_data = sheet.get_all_records()
    openCourseData[sheet.title] = sheet_data  # แยกข้อมูลตามชื่อชีท เช่น CS_Y1

# เปิดไฟล์ Open Course2 
openCourseFile2 = client.open('Open Course2')

# ฟังก์ชันตรวจสอบว่าข้อมูลมีอยู่แล้วหรือไม่
def data_exists(sheet, section, courseID, courseType):
    existing_data = sheet.get_all_records()
    for record in existing_data:
        if record['เซคเรียน'] == section and record['รหัสวิชา'] == courseID and record['ประเภท (บรรยาย/ปฏิบัติ)'] == courseType:
            return True
    return False

# ลิสต์สำหรับเก็บข้อมูลเซคชั่นที่ไม่ครบ
incompleteSections = []

# สร้างชีทใหม่หรือเลือกชีทที่มีอยู่แล้วใน Open Course2 สำหรับแต่ละปีการศึกษาและสาขา
for sheetName, records in openCourseData.items():
    try:
        # ตรวจสอบว่าชีทมีอยู่แล้วหรือไม่ ถ้ามีแล้วไม่ต้องสร้างใหม่
        if sheetName not in [sheet.title for sheet in openCourseFile2.worksheets()]:
            newSheet = openCourseFile2.add_worksheet(title=sheetName, rows='100', cols='8')
            headers = ['เซคเรียน', 'รหัสวิชา', 'ชื่อวิชา', 'หมวดหมู่รายวิชา', 'หน่วยกิต (ทฤษฎี-ปฏิบัติ-ศึกษาด้วยตนเอง)', 'จำนวนชั่วโมง', 'ประเภท (บรรยาย/ปฏิบัติ)', 'รหัสอาจารย์']
            newSheet.insert_row(headers, index=1)
        else:
            newSheet = openCourseFile2.worksheet(sheetName)

        # ตรวจสอบข้อมูลในชีท
        for record in records:
            courseID = record['รหัสวิชา']
            # ค้นหาข้อมูลที่ตรงกันใน Curriculum และ General Education Program
            matching_records = [r for r in curriculumData + generalEdData if r['รหัสวิชา'] == courseID]

            if not matching_records:
                continue  # ถ้าไม่มีข้อมูลตรงกันในไฟล์ Curriculum หรือ General Education Program ให้ข้ามไป

            # ข้อมูลจาก Curriculum/General Education Program
            matched_record = matching_records[0]  # ใช้ข้อมูลที่ตรงกัน (เลือกอันแรก)
            section = record['เซคเรียน']
            courseName = matched_record['ชื่อวิชา']
            category = matched_record['หมวดหมู่รายวิชา']
            credits = matched_record['หน่วยกิต (ทฤษฎี-ปฏิบัติ-ศึกษาด้วยตนเอง)']
            teacherID = ''  

            # สำหรับคาบเรียนบรรยาย
            lecturePeriod = matched_record['คาบเรียน (บรรยาย)']
            if lecturePeriod and not data_exists(newSheet, section, courseID, 'บรรยาย'):
                data = [section, courseID, courseName, category, credits, lecturePeriod, 'บรรยาย', teacherID]
                newSheet.append_row(data)  # เพิ่มแถวที่มีข้อมูลคาบเรียนบรรยาย
            elif not lecturePeriod:
                incompleteSections.append((section, courseID))  # เพิ่มเซคชั่นที่ไม่ครบข้อมูลบรรยาย

            # สำหรับคาบเรียนปฏิบัติ
            labPeriod = matched_record['คาบเรียน (ปฏิบัติ)']
            if labPeriod and not data_exists(newSheet, section, courseID, 'ปฏิบัติ'):
                data = [section, courseID, courseName, category, credits, labPeriod, 'ปฏิบัติ', teacherID]
                newSheet.append_row(data)  # เพิ่มแถวที่มีข้อมูลคาบเรียนปฏิบัติ
            elif not labPeriod:
                incompleteSections.append((section, courseID))  # เพิ่มเซคชั่นที่ไม่ครบข้อมูลปฏิบัติ

            time.sleep(5)

    except Exception as e:
        print(f"Error sheet {sheetName}: {e}")

# แสดงผลเซคชั่นที่ยังไม่ครบ
if incompleteSections:
    print("Incomplete sections found:")
    for section in incompleteSections:
        print(f"Section: {section[0]}, Course ID: {section[1]}")

print("Success!")
