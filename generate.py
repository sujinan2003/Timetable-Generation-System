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
    

# Write timetable to Google Sheets
def write_timetable_to_sheet(timetable, sheet_name):
    try:
        generateFile = client.open('Generate')
        sheet_names = [sheet.title for sheet in generateFile.worksheets()]
        
        if sheet_name not in sheet_names:
            generateFile.add_worksheet(title=sheet_name, rows="100", cols="20")
        
        sheet = generateFile.worksheet(sheet_name)
        sheet.clear()
        
        data = []
        header = [
            'เซคเรียน', 'รหัสวิชา', 'อาจารย์', 'ห้องเรียน',
            'ประเภทวิชา', 'วันเรียน', 'คาบ (เริ่ม)', 'คาบ (จบ)'
        ]
        data.append(header)
        
        for schedule in timetable.schedule:
            row = [
                schedule.get('เซคเรียน', ''),
                schedule.get('รหัสวิชา', ''),
                schedule.get('อาจารย์', ''),
                schedule.get('ห้องเรียน', ''),
                schedule.get('ประเภทวิชา', ''),
                schedule.get('วันเรียน', ''),
                schedule.get('คาบ (เริ่ม)', ''),
                schedule.get('คาบ (จบ)', '')
            ]
            data.append(row)
        
        # Use batch_update to minimize API calls
        cell_range = f'A1:{chr(65 + len(header) - 1)}{len(data)}'
        sheet.update(range_name=cell_range, values= data)
        
        print("Success!")
    
    except Exception as e:
        print(f"An error occurred: {e}")

# Example for reading data
def load_data_from_main(client):
    mainFile = client.open('Main')
    timeSlotSheet = mainFile.worksheet('TimeSlot')
    timeSlots = timeSlotSheet.col_values(3)[2:]  # Skip header row

    roomSheet = mainFile.worksheet('Room')
    room_range = roomSheet.get_all_values()
    rooms = [row[2] for row in room_range[2:] if row[2]]

    return timeSlots, rooms


# Load room types from Main file
def load_room_types(client):
    mainFile = client.open('Main')
    roomSheet = mainFile.worksheet('Room')
    room_names_range = roomSheet.range('C3:C')  # ดึงชื่อห้องจากคอลัมน์ C
    room_types_range = roomSheet.range('G3:G')  # ดึงประเภทห้องจากคอลัมน์ G
    room_types = {}
    for room_name, room_type in zip(room_names_range, room_types_range):
        if room_name.value and room_type.value:
            room_types[room_name.value] = room_type.value
    return room_types

# Load courses and curriculum data
def load_courses_curriculum(client):
    curriculum = []
    try:
        openCourseFile = client.open('Open Course2')
        openCourseSheets = openCourseFile.worksheets()

        for sheet in openCourseSheets:
            data = sheet.get_all_values()
            for row in data[1:]:  # Skip header row
                section = row[0]
                course_code = row[1]
                course_name = row[2]
                category = row[3]
                credits = row[4]
                hours = row[5]
                course_type = row[6]
                teacher = row[7]

                # Check if the category is "ศึกษาทั่วไป" and skip if it is
                if category == 'ศึกษาทั่วไป':
                    continue

                # Validate and convert to integer
                try:
                    hours = int(hours)
                except ValueError:
                    print(f"Invalid data for hours: {hours}")
                    continue

                if course_code and section and teacher is not None:
                    curriculum.append({
                        'รหัสวิชา': course_code,
                        'เซคเรียน': section,
                        'อาจารย์': teacher,
                        'จำนวนชั่วโมง': hours,
                        'ประเภทวิชา': course_type
                    })

    except Exception as e:
        print(f"Failed to load courses curriculum: {e}")
    
    return curriculum

# Load teacher availability data
def load_teacher_availability(client):
    teacher_availability = {}
    try:
        teacher_file = client.open('Teacher')
        sheets = teacher_file.worksheets()

        for sheet in sheets:
            if sheet.title == 'คำอธิบาย':
                continue
            
            teacher_id = sheet.title
            data = sheet.get_all_values()

            if not data or len(data) < 3:
                print(f"Insufficient data in sheet {teacher_id}.")
                continue

            # Extract headers and periods
            headers = data[1]  # Headers: คาบ, เวลา, จันทร์, อังคาร, ..., อาทิตย์
            periods = [row[0] for row in data[2:]]  # Periods from A3:A

            # Initialize the dictionary for storing availability
            availability = {}

            # Iterate through each period
            for idx, period in enumerate(periods):
                period_number = period  # e.g., 1, 2, 3, ...
                availability[period_number] = {}  # Create a nested dictionary for each period

                # Iterate through each day's availability status
                for day_index, availability_status in enumerate(data[idx + 2][2:]):
                    day = headers[day_index + 2]  # Mapping day names

                    try:
                        # Store only if the availability status is 1 (available to teach)
                        if int(availability_status) == 1:
                            availability[period_number][day] = 1
                    except ValueError:
                        continue  # Ignore if it's not an integer or valid value

            
            if availability:
                teacher_availability[teacher_id] = {'availability': availability}

    except Exception as e:
        print(f"Failed to load teacher availability: {e}")

    return teacher_availability

# Check student timetable availability
def check_timetable_student(client):
    student_availability = {}

    try:
        student_file = client.open('Student')
        sheets = student_file.worksheets()

        for sheet in sheets:
            if sheet.title == 'คำอธิบาย':
                continue

            major_year = sheet.title
            data = sheet.get_all_values()

            if len(data) < 3:
                continue  # ข้ามชีทถ้ามีข้อมูลน้อยเกินไป

            headers = data[1]  # แถวที่สองเป็นหัวข้อวัน
            row_idx = 0  # เริ่มต้นตรวจสอบที่แถวที่ 0

            while row_idx < len(data):
                # ตรวจสอบว่าข้อมูลในแถวนี้มีคำว่า 'ตารางเรียน' หรือไม่
                if len(data[row_idx]) > 0 and "ตารางเรียน" in data[row_idx][0]:
                    # ดึงชื่อเซคออกมา
                    section_name = data[row_idx][0].replace("ตารางเรียน ", "").strip()
                    row_idx += 1  # ขยับไปที่แถว header คาบและวัน

                    if row_idx + 12 >= len(data):
                        break  # ถ้าข้อมูลคาบไม่ครบให้หยุดการทำงาน

                    # ตรวจสอบว่ามีข้อมูลคาบในแต่ละเซคหรือไม่
                    periods = [row[0] for row in data[row_idx:row_idx + 12] if len(row) > 0]  # ดึงข้อมูลคาบ 12 คาบ
                    availability = {}

                    # วนลูปผ่านคาบ
                    for period_idx, period in enumerate(periods):
                        period_number = period
                        availability[period_number] = {}

                        # วนลูปผ่านแต่ละวัน
                        for day_index, availability_status in enumerate(data[row_idx + period_idx][2:]):
                            if day_index + 2 >= len(headers):  # ตรวจสอบว่า index ของ headers ยังไม่เกินขอบเขต
                                continue

                            day = headers[day_index + 2]  # วันในสัปดาห์
                            availability_status = availability_status.strip()  # ตัดช่องว่างที่ไม่จำเป็นออก

                            if availability_status == '' or availability_status == ' ':
                                availability[period_number][day] = 'ว่าง'
                            else:
                                availability[period_number][day] = 'ถูกจอง'

                    # ถ้ามี section_name อยู่แล้วใน student_availability ให้รวมข้อมูลใหม่เข้ากับข้อมูลเก่า
                    if section_name in student_availability:
                        student_availability[section_name]['availability'].update(availability)
                    else:
                        student_availability[section_name] = {'availability': availability}

                    # ขยับ row_idx ไปยังตารางของเซคถัดไป
                    row_idx += 12  # ขยับแถวเพื่อไปยังเซคถัดไป (12 แถวสำหรับข้อมูลคาบ)

                else:
                    # ถ้าไม่เจอ 'ตารางเรียน' ให้ขยับไปยังแถวถัดไป
                    row_idx += 1

    except Exception as e:
        print(f"Failed: {e}")

    return student_availability




    
# TimeTable class for managing schedule
class TimeTable:
    def __init__(self, timeSlots, rooms, room_types, curriculum, teacher_availability, student_availability):
        self.timeSlots = timeSlots
        self.rooms = rooms
        self.room_types = room_types
        self.curriculum = curriculum
        self.teacher_availability = teacher_availability  
        self.student_availability = student_availability
        self.schedule = []
        self.fitness = 0

    def initialize(self):
        days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์']
        # Generate lectures first
        for course in self.curriculum:
            if course['ประเภทวิชา'] == 'บรรยาย':
                self.add_course_to_schedule(course, days, 'บรรยาย')
        
        # Then generate practicals
        for course in self.curriculum:
            if course['ประเภทวิชา'] == 'ปฏิบัติ':
                self.add_course_to_schedule(course, days, 'ปฏิบัติ')
        self.calculate_fitness()

    def add_course_to_schedule(self, course, days, course_type):
        num_hours = course['จำนวนชั่วโมง']
        teacher_id = course['อาจารย์']
        day = random.choice(days)
        start_period = random.choice(self.timeSlots)
        end_period = self.calculate_end_period(start_period, num_hours)
        room = self.get_available_room(course_type)

        if not self.check_teacher_availability(teacher_id, day, start_period):
            return

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
        
    def check_teacher_availability(self, teacher_id, day, period):
        if teacher_id not in self.teacher_availability:
            return False
        availability = self.teacher_availability[teacher_id]['availability']
        return availability.get(period, {}).get(day, 0) == 1
    
    def check_student_availability(self, client, section, day, period):
        student_availability = check_timetable_student(client, [period])
        for (section_name, slot, day_index), status in student_availability.items():
            if section_name == section and day_index == day:
                return status
        return None

    def calculate_end_period(self, start_period, num_periods):
        start_index = self.timeSlots.index(start_period)
        end_index = min(start_index + num_periods - 1, len(self.timeSlots) - 1)
        return self.timeSlots[end_index]

    def get_available_room(self, course_type):
        available_rooms = [room for room in self.rooms if self.get_room_type_for_room(room) == course_type]
        return random.choice(available_rooms) if available_rooms else None

    def get_room_type_for_room(self, room):
        return self.room_types.get(room, None)

    def calculate_fitness(self):
        # Basic fitness calculation (higher is better)
        self.fitness = len(self.schedule)


# ตัวอย่างฟังก์ชันวัดคุณภาพ (Fitness Function)
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
                score -= 10  # หักคะแนนเมื่อมีการชนกันของห้องเรียน
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

    # ตรวจสอบว่าตารางเรียนมีความต่อเนื่อง ไม่ว่างมากเกินไป
    continuous_schedules = 0
    for day in ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์']:
        periods_in_day = [entry for entry in timetable.schedule if entry['วันเรียน'] == day]
        periods_in_day.sort(key=lambda x: x['คาบ (เริ่ม)'])

        for i in range(1, len(periods_in_day)):
            if periods_in_day[i]['คาบ (เริ่ม)'] == periods_in_day[i-1]['คาบ (จบ)'] + 1:
                continuous_schedules += 1
    
    # ให้คะแนนเมื่อมีการจัดคาบเรียนต่อเนื่อง
    score += continuous_schedules * 5

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

# Example function for running the timetable generation
def run():
    timeSlots, rooms = load_data_from_main(client)
    room_types = load_room_types(client)
    curriculum = load_courses_curriculum(client)
    teacher_availability = load_teacher_availability(client)
    student_availability = check_timetable_student(client)

    print(student_availability)
    
  #  best_timetable = TimeTable(timeSlots, rooms, room_types, curriculum, teacher_availability, student_availability)
  #  best_timetable.initialize()

  #  write_timetable_to_sheet(best_timetable, 'Generate')

if __name__ == '__main__':
    run()