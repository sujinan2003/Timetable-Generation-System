import random
import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials

# Setup Google Sheets connection
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)
client = gspread.authorize(credentials)

# ฟังก์ชันสำหรับการร้องขอข้อมูลจาก Google Sheets
def get_sheet_data(sheet, range_name):
    try:
        data = sheet.get(range_name)
    except gspread.exceptions.APIError as e:
        print(f"APIError: {e}")
        handle_api_error(e)
        data = retry_request(sheet.get, 3, 5, range_name)
    return data

# ฟังก์ชันสำหรับการอัปเดตข้อมูลใน Google Sheets
def update_sheet_data(sheet, range_name, values):
    try:
        sheet.update(range_name, values)
    except gspread.exceptions.APIError as e:
        print(f"APIError: {e}")
        handle_api_error(e)
        retry_request(sheet.update, 3, 5, range_name, values)

# ฟังก์ชันสำหรับจัดการข้อผิดพลาดจาก API
def handle_api_error(exception):
    if 'Quota exceeded' in str(exception):
        print("Quota exceeded. Waiting before retrying.")
        time.sleep(60)  # รอ 60 วินาทีก่อนทำการร้องขอใหม่
    else:
        print(f"An unexpected error occurred: {exception}")

# ฟังก์ชันสำหรับจัดการการทดลองใหม่
def retry_request(request_func, retries=3, delay=5, *args, **kwargs):
    for attempt in range(retries):
        try:
            return request_func(*args, **kwargs)
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            time.sleep(delay)
    raise Exception("All retry attempts failed.")

# ฟังก์ชันสำหรับการจัดการคำร้องขอแบบ Batch
def batch_update(sheet, updates):
    requests = []
    for update in updates:
        range_name = update['range']
        values = update['values']
        requests.append({
            "updateCells": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": range_name[0],
                    "endRowIndex": range_name[1],
                    "startColumnIndex": range_name[2],
                    "endColumnIndex": range_name[3]
                },
                "rows": [{"values": [{"userEnteredValue": {"stringValue": value}} for value in row]} for row in values],
                "fields": "userEnteredValue"
            }
        })
    body = {
        "requests": requests
    }
    sheet.spreadsheet.batchUpdate(body)
    
    
def write_timetable_to_sheet(timetable, sheet_name):
    try:
        # เปิดไฟล์ Google Sheets ที่ชื่อ 'Generate'
        generateFile = client.open('Generate')
        
        # ดึงชื่อของชีททั้งหมดในไฟล์ 'Generate'
        sheetNames = [sheet.title for sheet in generateFile.worksheets()]
        
        # ตรวจสอบว่าชื่อชีทที่ต้องการเขียนมีอยู่หรือไม่
        if sheet_name not in sheetNames:
            # ถ้ายังไม่มี ชื่อชีทนั้นให้สร้างชีทใหม่ด้วยจำนวนแถวและคอลัมน์ที่กำหนด
            generateFile.add_worksheet(title=sheet_name, rows="100", cols="20")
        
        # เข้าถึงชีทที่ต้องการเขียนข้อมูล
        sheet = generateFile.worksheet(sheet_name)
        
        # ล้างข้อมูลในชีทก่อนเพื่อไม่ให้มีข้อมูลเก่าปนอยู่
        sheet.clear()
        
        data = []  # สร้างลิสต์ที่จะเก็บข้อมูลที่จะเขียนลงชีท
        # สร้างหัวข้อของตาราง
        header = [
            'เซคเรียน', 'รหัสวิชา', 'อาจารย์', 'ห้องเรียน',
            'ประเภทวิชา', 'วันเรียน', 'คาบ (เริ่ม)', 'คาบ (จบ)'
        ]
        data.append(header)  # เพิ่มหัวข้อในลิสต์ข้อมูล
        
        # วนลูปผ่านตารางเวลาที่มีอยู่ใน timetable.schedule
        for schedule in timetable.schedule:
            # สร้างแถวข้อมูลจากแต่ละ schedule
            row = [
                schedule.get('เซคเรียน', ''),  # รับค่าเซคเรียน ถ้าไม่มีให้ใช้ค่าว่าง
                schedule.get('รหัสวิชา', ''),   # รับค่ารหัสวิชา
                schedule.get('อาจารย์', ''),     # รับค่าอาจารย์
                schedule.get('ห้องเรียน', ''),   # รับค่าห้องเรียน
                schedule.get('ประเภทวิชา', ''), # รับค่าประเภทวิชา
                schedule.get('วันเรียน', ''),    # รับค่าวันเรียน
                schedule.get('คาบ (เริ่ม)', ''),  # รับค่าคาบเริ่มต้น
                schedule.get('คาบ (จบ)', '')     # รับค่าคาบจบ
            ]
            data.append(row)  # เพิ่มแถวข้อมูลในลิสต์ข้อมูล
        
        # ใช้ batch_update เพื่อเพิ่มประสิทธิภาพในการเรียก API
        # กำหนดช่วงเซลล์ที่จะอัปเดตข้อมูล
        cell_range = f'A1:{chr(65 + len(header) - 1)}{len(data)}'
        # อัปเดตข้อมูลทั้งหมดในชีท
        sheet.update(range_name=cell_range, values=data)
        
        print("Success!")  # แสดงข้อความเมื่อสำเร็จ
    
    except Exception as e:
        # แสดงข้อความเมื่อเกิดข้อผิดพลาด
        print(f"An error occurred: {e}")
        

def load_data_from_main(client):
    
    mainFile = client.open('Main')
    
    timeSlotSheet = mainFile.worksheet('TimeSlot')
    
    # ดึงค่าจากคอลัมน์ที่ 3 (C) ของชีท 'TimeSlot' และข้ามแถวหัวตาราง
    timeSlots = timeSlotSheet.col_values(3)[2:]  # Skip header row

    # เข้าถึงชีท 'Room' ในไฟล์ 'Main'
    roomSheet = mainFile.worksheet('Room')
    
    # ดึงข้อมูลทั้งหมดในชีท 'Room'
    room_range = roomSheet.get_all_values()
    
    # สร้างลิสต์ของห้องเรียน โดยดึงค่าจากคอลัมน์ที่ 3 (C) และข้ามแถวหัวตาราง
    rooms = [row[2] for row in room_range[2:] if row[2]]  # Only include non-empty values

    # ส่งค่าที่ได้กลับ (timeSlots และ rooms)
    return timeSlots, rooms


def load_room_types(client):
    # เปิดไฟล์ Google Sheets ที่ชื่อ 'Main'
    mainFile = client.open('Main')
    
    # เข้าถึงชีท 'Room' ในไฟล์ 'Main'
    roomSheet = mainFile.worksheet('Room')
    
    # ดึงช่วงข้อมูลชื่อห้องจากคอลัมน์ C เริ่มจากแถวที่ 3
    room_names_range = roomSheet.range('C3:C')  # ดึงชื่อห้องจากคอลัมน์ C
    
    # ดึงช่วงข้อมูลประเภทห้องจากคอลัมน์ G เริ่มจากแถวที่ 3
    room_types_range = roomSheet.range('G3:G')  # ดึงประเภทห้องจากคอลัมน์ G
    
    # สร้างดิกชันนารีสำหรับเก็บชื่อห้องและประเภทห้อง
    room_types = {}
    
    # วนลูปผ่านชื่อห้องและประเภทห้องพร้อมกัน
    for room_name, room_type in zip(room_names_range, room_types_range):
        # ตรวจสอบว่าทั้งชื่อห้องและประเภทห้องมีค่า (ไม่ว่าง)
        if room_name.value and room_type.value:
            # เพิ่มชื่อห้องและประเภทห้องในดิกชันนารี
            room_types[room_name.value] = room_type.value
    
    # ส่งกลับดิกชันนารีของชื่อห้องและประเภทห้อง
    return room_types


def load_courses_curriculum(client):
    
    # สร้างลิสต์เพื่อเก็บข้อมูลหลักสูตร
    curriculum = []
    
    try:
        # เปิดไฟล์ Google Sheets ที่ชื่อ 'Open Course2'
        openCourseFile = client.open('Open Course2')
        
        # ดึงชีททั้งหมดในไฟล์ 'Open Course2'
        openCourseSheets = openCourseFile.worksheets()

        # วนลูปผ่านชีทแต่ละชีท
        for sheet in openCourseSheets:
            # เพิ่มเงื่อนไขเพื่อเลือกเฉพาะชีทปี 1
            if not sheet.title.endswith('_Y4'):
                continue  # ข้ามชีทที่ไม่ใช่ปี 1

            # ดึงข้อมูลทั้งหมดในชีท
            data = sheet.get_all_values()
            for row in data[1:]:  # ข้ามแถวหัวตาราง
                section = row[0]  # เซคเรียน
                course_id = row[1]  # รหัสวิชา
                course_name = row[2]  # ชื่อวิชา
                category = row[3]  # หมวดหมู่
                credits = row[4]  # หน่วยกิต
                hours = row[5]  # จำนวนชั่วโมง
                course_type = row[6]  # ประเภทวิชา
                teacher = row[7]  # อาจารย์

                # ตรวจสอบว่าหมวดหมู่เป็น "ศึกษาทั่วไป" และข้ามถ้าใช่
                if category == 'ศึกษาทั่วไป':
                    continue

                # ตรวจสอบและแปลงจำนวนชั่วโมงเป็นจำนวนเต็ม
                try:
                    hours = int(hours)
                except ValueError:
                    print(f"Invalid data for hours: {hours}")  # แจ้งข้อมูลไม่ถูกต้อง
                    continue

                # ตรวจสอบว่ามีข้อมูลรหัสวิชา เซคเรียน และอาจารย์
                if course_id and section and teacher is not None:
                    # เพิ่มข้อมูลหลักสูตรลงในลิสต์ curriculum
                    curriculum.append({
                        'รหัสวิชา': course_id,
                        'เซคเรียน': section,
                        'อาจารย์': teacher,
                        'จำนวนชั่วโมง': hours,
                        'ประเภทวิชา': course_type
                    })

    except Exception as e:
        print(f"Failed to load courses curriculum: {e}")  # แจ้งข้อผิดพลาดถ้าเกิดขึ้น
    
    # ส่งกลับลิสต์ curriculum ที่เก็บข้อมูลหลักสูตร
    return curriculum


def load_teacher_availability(client):
    # สร้างดิกชันนารีเพื่อเก็บข้อมูลความพร้อมของอาจารย์
    teacher_availability = {}
    try:
        # เปิดไฟล์ Google Sheets ที่ชื่อ 'Teacher'
        teacherFile = client.open('Teacher')
        
        # ดึงชีททั้งหมดในไฟล์ 'Teacher'
        teacherSheets = teacherFile.worksheets()

        # วนลูปผ่านชีทแต่ละชีท
        for sheet in teacherSheets:
            # ข้ามชีทที่ชื่อว่า 'คำอธิบาย'
            if sheet.title == 'คำอธิบาย':
                continue
            
            teacher_id = sheet.title  # ใช้ชื่อชีทเป็นรหัสอาจารย์
            data = sheet.get_all_values()  # ดึงข้อมูลทั้งหมดในชีท

            # ตรวจสอบว่าข้อมูลมีพอหรือไม่
            if not data or len(data) < 3:
                print(f"Insufficient data in sheet {teacher_id}.")
                continue  # ข้ามชีทถ้ามีข้อมูลไม่เพียงพอ

            # ดึงแถวหัวตารางและช่วงเวลา
            headers = data[1]  # แถวที่ 2 เป็นหัวตาราง: คาบ, เวลา, จันทร์, อังคาร, ..., อาทิตย์
            periods = [row[0] for row in data[2:]]  # ดึงคาบจากคอลัมน์ A (เริ่มจาก A3)

            # สร้างดิกชันนารีสำหรับเก็บความพร้อม
            availability = {}

            # วนลูปผ่านแต่ละคาบ
            for idx, period in enumerate(periods):
                period_number = period  # e.g., 1, 2, 3, ...
                availability[period_number] = {}  # สร้างดิกชันนารีซ้อนสำหรับแต่ละคาบ

                # วนลูปผ่านสถานะความพร้อมในแต่ละวัน
                for day_index, availability_status in enumerate(data[idx + 2][2:]):
                    day = headers[day_index + 2]  # ดึงชื่อวันจากหัวตาราง

                    try:
                        # เก็บเฉพาะถ้าสถานะความพร้อมเป็น 1 (สามารถสอน)
                        if int(availability_status) == 1:
                            availability[period_number][day] = 1
                    except ValueError:
                        continue  # ข้ามไปถ้าค่าไม่ใช่จำนวนเต็มหรือไม่ถูกต้อง

            # ถ้ามีความพร้อมของอาจารย์ ให้เพิ่มลงใน teacher_availability
            if availability:
                teacher_availability[teacher_id] = {'availability': availability}

    except Exception as e:
        print(f"Failed to load teacher availability: {e}")  # แจ้งข้อผิดพลาดถ้าเกิดขึ้น

    return teacher_availability  # ส่งกลับดิกชันนารีความพร้อมของอาจารย์


# ตรวจสอบว่าในตารางสอน นักเรียน มีคาบไหนถูกจอง/ว่าง  เฉพาะของปี1
def check_timetable_student(client):
    
    # สร้างดิกชันนารีเพื่อเก็บข้อมูลความพร้อมของนักเรียน
    student_availability = {}
    
    try:
       
        studentFile = client.open('Student')
      
        studentSheets = studentFile.worksheets()
        
        # วนลูปผ่านชีทแต่ละชีท
        for sheet in studentSheets:
            # เพิ่มเงื่อนไขเพื่อเลือกเฉพาะชีทปี 4
            if not sheet.title.endswith('_Y4'):
                continue

            # ดึงข้อมูลทั้งหมดในชีท
            data = sheet.get_all_values()

            # ข้ามชีทถ้ามีข้อมูลน้อยเกินไป
            if len(data) < 3:
                continue  

            row_idx = 0  # เริ่มต้นที่แถวแรก
            while row_idx < len(data):
                # ค้นหาจุดเริ่มต้นของตารางใหม่
                if len(data[row_idx]) > 0 and "ตารางเรียน" in data[row_idx][0]:
                    section_name = data[row_idx][0].replace("ตารางเรียน ", "").strip()  # ดึงชื่อเซคเรียน
                    row_idx += 1  # ข้ามไปยังแถวหัวข้อ (คาบ, เวลา, วัน)
                    
                    if row_idx >= len(data):
                        break
                    
                    headers = data[row_idx]  # เก็บหัวตาราง
                    row_idx += 1  # ข้ามไปยังข้อมูลคาบแรก
                    
                    availability = {}  # สร้างดิกชันนารีสำหรับความพร้อม
                    
                    # อ่านข้อมูลทีละคาบ (12 คาบ)
                    for _ in range(12):
                        if row_idx >= len(data) or len(data[row_idx]) < 2:
                            break  # ข้ามหากไม่มีข้อมูลในแถวปัจจุบัน
                        
                        period = data[row_idx][0]  # คาบเรียน
                        availability[period] = {}  # สร้างดิกชันนารีสำหรับคาบนี้
                        
                        # วนลูปผ่านวันในคาบ
                        for day_index, slot_data in enumerate(data[row_idx][2:]):
                            if day_index + 2 < len(headers):
                                day = headers[day_index + 2]  # ดึงชื่อวันจากหัวตาราง
                                slot_data = slot_data.strip()
                                # เก็บสถานะการจองห้อง
                                availability[period][day] = 'ถูกจอง' if slot_data else 'ว่าง'
                        
                        row_idx += 1  # ไปยังแถวถัดไป
                    
                    # บันทึกความพร้อมของเซคเรียนนี้
                    student_availability[section_name] = {'availability': availability}
                else:
                    row_idx += 1  # ไปยังแถวถัดไป

    except Exception as e:
        print(f"Failed: {e}")  # แจ้งข้อผิดพลาดถ้าเกิดขึ้น

    return student_availability  # ส่งกลับดิกชันนารีความพร้อมของนักเรียน



def check_timetable_student_generateFile(client):
    studentGen_availability = {}  

    try:
        
        gen_student_file = client.open('Generate')
        sheets = gen_student_file.worksheets()

        # วนลูปเพื่อเลือก sheet ที่ลงท้ายด้วย 'Gen_Y1'
        for sheet in sheets:
            if not sheet.title.endswith('Gen_Y1') and not sheet.title.endswith('Gen_Y2') and not sheet.title.endswith('Gen_Y3'):
                continue  

            #ดึงข้อมูลทั้งหมดใน sheet
            data = sheet.get_all_values()
            
            
            for row in data[1:]: #เริ่มตั้งแต่ แถวที่2เป้นต้นไป
                section = row[0] #วันอยู่ในคอลัมน์ที่ 1
                course_code = row[1]
                teacher_id = row[2]
                room = row[3]
                course_type = row[4]
                day = row[5]
                start_period = row[6]
                end_period = row[7]  
                  
                
                # เพิ่มข้อมูลที่เช็คเข้าไปใน dictionary
                if section not in studentGen_availability:
                    studentGen_availability[section] = []
                
                # เก็บข้อมูลแต่ละรายการตาม section
                studentGen_availability[section].append({
                    "เซคเรียน": section,
                    "รหัสวิชา": course_code ,
                    "อาจารย์": teacher_id,
                    "ห้องเรียน": room,
                    "ประเภทวิชา": course_type,
                    "วันเรียน": day,
                    "คาบ (เริ่ม)": start_period,
                    "คาบ (จบ)": end_period
                })
    
    except Exception as e:
        print(f"Failed: {e}")
    
    return studentGen_availability



class TimeTable:
    def __init__(self, timeSlots, rooms, room_types, curriculum, teacher_availability,student_availability,studentGen_availability):
        self.timeSlots = timeSlots
        self.rooms = rooms
        self.room_types = room_types
        self.curriculum = curriculum
        self.teacher_availability = teacher_availability 
        self.student_availability = student_availability
        self.studentGen_availability = studentGen_availability
        self.schedule = []
        self.fitness = 0

    def initialize(self):
        days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์']
        courses = self.curriculum.copy()
        random.shuffle(courses)
        
        for course in courses:
            course_type = course['ประเภทวิชา']
            if not self.add_course_to_schedule(course, days, course_type):
                print(f"ไม่สามารถจัดตารางให้กับวิชา {course['รหัสวิชา']} สำหรับ {course['เซคเรียน']}")
        
        self.calculate_fitness()
    
    
    def add_course_to_schedule(self, course, days, course_type):
        num_hours = course['จำนวนชั่วโมง']
        teacher_id = course['อาจารย์']
        section = course['เซคเรียน']
        max_attempts = 50  # เพิ่มจำนวนครั้งในการลองจัดตาราง

        for attempt in range(max_attempts):
            day = random.choice(days)
            start_period = random.choice(self.timeSlots)
            end_period = self.calculate_end_period(start_period, num_hours)
            room = self.get_available_room(course_type)


            if int(end_period) - int(start_period) + 1 != num_hours:
                continue

            if not self.check_teacher_availability(teacher_id, day, start_period):
                continue
            
            
            if not self.check_student_availability(section, day, start_period, end_period):
                continue
            

            if self.check_schedule_conflict(day, start_period, end_period, room, section):
                continue
            
            if self.check_student_generateFile_availability(day, start_period, end_period, room, teacher_id):
                continue

            self.schedule.append({
                'เซคเรียน': course['เซคเรียน'],
                'รหัสวิชา': course['รหัสวิชา'],
                'อาจารย์': teacher_id,
                'ห้องเรียน': room,
                'ประเภทวิชา': course_type,
                'วันเรียน': day,
                'คาบ (เริ่ม)': start_period,
                'คาบ (จบ)': end_period
            })
            print(f"Successfully scheduled {course['รหัสวิชา']} for section {section}")
            return True

        print(f"ไม่สามารถจัดตารางให้กับวิชา {course['รหัสวิชา']} ได้หลังจากลองทำ {max_attempts} ครั้ง")
        return False
        

    def check_teacher_availability(self, teacher_id, day, period):
        # ตรวจสอบว่า ID ของอาจารย์อยู่ในข้อมูลความพร้อมหรือไม่
        if teacher_id not in self.teacher_availability:
            return False  # ถ้าไม่มีข้อมูล ให้คืนค่า False

        # ดึงข้อมูลความพร้อมของอาจารย์
        availability = self.teacher_availability[teacher_id]['availability']
        
        # ตรวจสอบความพร้อมในคาบเรียนและวันที่กำหนด
        return availability.get(period, {}).get(day, 0) == 1  # คืนค่า True ถ้าอาจารย์พร้อม, False ถ้าไม่พร้อม

    
    def check_student_availability(self, section, day, start_period, end_period):
        # ตรวจสอบว่าเซคชั่นนักเรียนอยู่ในข้อมูลความพร้อมหรือไม่
        if section in self.student_availability:
            # ดึงข้อมูลความพร้อมของนักเรียนในเซคชั่นนั้น
            availability = self.student_availability[section]['availability']

            # แปลง start_period และ end_period เป็นตัวเลข
            start_num = int(start_period)
            end_num = int(end_period)

            # ตรวจสอบทุกคาบในช่วงเวลาที่ต้องการจัด
            for period_num in range(start_num, end_num + 1):
                period = str(period_num)  # แปลงกลับเป็น string เพื่อใช้เป็น key
                if period in availability and day in availability[period]:
                    if availability[period][day] == 'ถูกจอง':
                        return False  # ถ้าพบว่ามีคาบใดคาบหนึ่งถูกจองแล้ว ให้คืนค่า False

        return True  # ถ้าทุกคาบว่างหรือไม่มีข้อมูล ให้คืนค่า True
    
    # หาห้องที่เหมาะกับประเภทของวิชา 
    def get_available_room(self, course_type):
        # สร้างรายการห้องที่ตรงกับประเภทวิชาที่กำหนด
        available_rooms = [room for room in self.rooms if self.get_room_type_for_room(room) == course_type]
        
        # คืนห้องเรียนที่เลือกแบบสุ่มถ้ามีห้องว่าง
        return random.choice(available_rooms) if available_rooms else None  # ถ้าไม่มีห้องว่าง คืนค่า None

    # ตรวจสอบประเภทของห้องตามชื่อห้องที่กำหนด
    def get_room_type_for_room(self, room):
        return self.room_types.get(room, None)  # คืนค่าประเภทห้อง ถ้าไม่พบคืนค่า None

    
    def check_schedule_conflict(self, day, start_period, end_period, room, section):
        for entry in self.schedule:
            # ตรวจสอบว่ารายการอยู่ในวันเดียวกันและเป็นเซคเรียนเดียวกัน
            if entry['วันเรียน'] == day and entry['เซคเรียน'] == section or entry['ห้องเรียน'] == room:
                # ตรวจสอบการชนกันของเวลา
                if (int(start_period) <= int(entry['คาบ (จบ)']) and int(end_period) >= int(entry['คาบ (เริ่ม)']) or 
                    int(entry['คาบ (เริ่ม)']) <= int(end_period) and int(entry['คาบ (จบ)']) >= int(start_period)
                    ):
                    return True  # พบการชนกัน
        return False  # ไม่มีการชนกัน
    
    
    # คำนวณช่วงเวลาสิ้นสุดของคาบเรียนจากช่วงเวลาเริ่มต้นและจำนวนคาบที่กำหนด
    def calculate_end_period(self, start_period, num_periods):
        start_index = self.timeSlots.index(start_period)  # ค้นหาตำแหน่งเริ่มต้นใน timeSlots
        end_index = min(start_index + num_periods - 1, len(self.timeSlots) - 1)  # คำนวณตำแหน่งสิ้นสุด
        return self.timeSlots[end_index]  # คืนค่าช่วงเวลาสิ้นสุด
    
    
    def check_student_generateFile_availability(self, day, start_period, end_period, room, teacher_id):
        for section, entries in self.studentGen_availability.items():
            for entry in entries:
                if entry['วันเรียน'] == day:
                    
                    if entry['ห้องเรียน'] == room or entry['อาจารย์'] == teacher_id:
                        if (start_period <= entry['คาบ (จบ)'] and end_period >= entry['คาบ (เริ่ม)']) or \
                        (entry['คาบ (เริ่ม)'] <= end_period and entry['คาบ (จบ)'] >= start_period):
                            return True  # มีการชนกันของตารางปี 1
                        
                    # เพิ่มเงื่อนไข อาจารย์ วันเรียน ห้องเรียน คาบเรียน ห้ามตรงกันในวันนั้นๆ
                    if entry['อาจารย์'] == teacher_id and entry['วันเรียน'] == day and \
                       ((start_period <= entry['คาบ (จบ)'] and end_period >= entry['คาบ (เริ่ม)']) or \
                        (entry['คาบ (เริ่ม)'] <= end_period and entry['คาบ (จบ)'] >= start_period)):
                        return True  
                    
                    if entry['ห้องเรียน'] == room and entry['วันเรียน'] == day and \
                       ((start_period <= entry['คาบ (จบ)'] and end_period >= entry['คาบ (เริ่ม)']) or \
                        (entry['คาบ (เริ่ม)'] <= end_period and entry['คาบ (จบ)'] >= start_period)):
                        return True  
                    
        return False  # ไม่มีการชนกัน


    def calculate_fitness(self):
        # Basic fitness calculation (higher is better)
        self.fitness = len(self.schedule)


# ฟังก์ชันวัดคุณภาพ (Fitness Function)
def fitness(timetable):
    score = 0

    # ตรวจสอบการชนกันของอาจารย์
    teacher_schedule = {}
    for entry in timetable.schedule:
        teacher = entry['อาจารย์']
        day = entry['วันเรียน']
        start_period = entry['คาบ (เริ่ม)']
        end_period = entry['คาบ (จบ)']

        # สร้างตารางเรียนสำหรับอาจารย์แต่ละคน
        if teacher not in teacher_schedule:
            teacher_schedule[teacher] = {}
        if day not in teacher_schedule[teacher]:
            teacher_schedule[teacher][day] = []

        
        # ตรวจสอบว่ามีการชนกันของคาบเรียนหรือไม่
        for period in teacher_schedule[teacher][day]:
            if not (end_period <= period['start_period'] or start_period >= period['end_period']):
                score -= 10  # หักคะแนนเมื่อมีการชนกัน
                break
        
        
        teacher_schedule[teacher][day].append({
            'start_period': start_period,
            'end_period': end_period
        })

    
    # ตรวจสอบการชนกันของห้องเรียน
    room_schedule = {}
    for entry in timetable.schedule:
        room = entry['ห้องเรียน']
        day = entry['วันเรียน']
        start_period = entry['คาบ (เริ่ม)']
        end_period = entry['คาบ (จบ)']

        if room not in room_schedule:
            room_schedule[room] = {}
        if day not in room_schedule[room]:
            room_schedule[room][day] = []

        for period in room_schedule[room][day]:
            if not (end_period <= period['start_period'] or start_period >= period['end_period']):
                score -= 5  # หักคะแนนเมื่อมีการชนกันของห้องเรียน
                break

        room_schedule[room][day].append({
            'start_period': start_period,
            'end_period': end_period
        })

    # เพิ่มคะแนนถ้าไม่ชนกัน
    score += 100
    
    # ตรวจสอบความถูกต้องของชั่วโมงเรียน
    for course in timetable.schedule:
        hours_assigned = len(range(int(course['คาบ (เริ่ม)']), int(course['คาบ (จบ)']) + 1))
        course_hours = next((c['จำนวนชั่วโมง'] for c in timetable.curriculum if c['รหัสวิชา'] == course['รหัสวิชา']), 0)
        
        if hours_assigned != course_hours:
            score -= 5  # หักคะแนนเมื่อชั่วโมงเรียนไม่ตรงตามที่กำหนด

    return score


# ฟังก์ชันสร้างประชากรเริ่มต้น
def generate_initial_population(size, time_slots, classrooms, sections, teachers):
    population = []
    for _ in range(size):
        timetable = []
        for section in sections:
            for slot in time_slots:
                course = random.choice(sections[section])  # สุ่มเลือกคอร์สสำหรับแต่ละ section
                room = random.choice(classrooms)           # สุ่มห้องเรียน
                teacher = random.choice(teachers)          # สุ่มครู
                timetable.append((section, slot, course, room, teacher))
        population.append(timetable)
    return population

# ฟังก์ชันเลือกพ่อแม่ (Selection)
def selection(population):
    # เลือกตารางเวลาที่ดีที่สุด 2 ตารางมาเป็นพ่อแม่
    sorted_population = sorted(population, key=lambda timetable: fitness(timetable), reverse=True)
    return sorted_population[:2]

# ฟังก์ชันผสมพันธุ์ (Crossover)
def crossover(parent1, parent2):
    # รวมตารางเวลาของพ่อแม่สองคน
    crossover_point = random.randint(0, len(parent1)-1)
    child1 = parent1[:crossover_point] + parent2[crossover_point:]
    child2 = parent2[:crossover_point] + parent1[crossover_point:]
    return child1, child2

# ฟังก์ชันกลายพันธุ์ (Mutation)
def mutation(timetable, classrooms, teachers):
    mutation_point = random.randint(0, len(timetable)-1)
    section, slot, course, _, _ = timetable[mutation_point]
    timetable[mutation_point] = (section, slot, course, random.choice(classrooms), random.choice(teachers))
    return timetable

# ฟังก์ชันหลักสำหรับอัลกอริทึมเจเนติก
def genetic_algorithm(time_slots, classrooms, sections, teachers, generations=100, population_size=10):
    population = generate_initial_population(population_size, time_slots, classrooms, sections, teachers)

    for generation in range(generations):
        # คำนวณค่าฟิตเนสของแต่ละ timetable
        fitness_values = [fitness(timetable) for timetable in population]
        
        # ปริ้นค่าฟิตเนสของประชากรในแต่ละ generation
        print(f"Generation {generation}: Fitness values: {fitness_values}")
        
        # ปริ้นค่าฟิตเนสที่ดีที่สุดในแต่ละ generation
        best_timetable = max(population, key=lambda timetable: fitness(timetable))
        print(f"Best fitness in generation {generation}: {fitness(best_timetable)}")
        
        # เลือกพ่อแม่ (Selection)
        parents = selection(population)

 
        new_population = []
        for _ in range(population_size // 2):  # ผสมพันธุ์ประชากรครึ่งหนึ่ง
            parents = selection(population)
            child1, child2 = crossover(parents[0], parents[1])
            new_population.append(mutation(child1, classrooms, teachers))
            new_population.append(mutation(child2, classrooms, teachers))
        
        population = new_population
    

    # หาตารางเวลาที่ดีที่สุดหลังจากครบทุก generation
    best_timetable = max(population, key=lambda timetable: fitness(timetable))

    return best_timetable


def run():
    random.seed(42)  # ตั้งค่า seed สำหรับ random

    timeSlots, rooms = load_data_from_main(client)
    room_types = load_room_types(client)
    curriculum = load_courses_curriculum(client)
    teacher_availability = load_teacher_availability(client)
    student_availability = check_timetable_student(client)
    studentGen_availability = check_timetable_student_generateFile(client)

    best_timetable = TimeTable(timeSlots, rooms, room_types, curriculum, teacher_availability, student_availability, studentGen_availability)
    best_timetable.initialize()
    
    print(f"Best fitness in generation : {fitness(best_timetable)}")

    write_timetable_to_sheet(best_timetable, 'Gen_Y4')

if __name__ == '__main__':
    run()
