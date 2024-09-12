import random
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Setup Google Sheets connection
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)
client = gspread.authorize(credentials)

# Write timetable to Google Sheets
def write_timetable_to_sheet(timetable, sheet_name):
    try:
        generateFile = client.open('Generate')
        if sheet_name not in [sheet.title for sheet in generateFile.worksheets()]:
            generateFile.add_worksheet(title=sheet_name, rows="100", cols="20")

        sheet = generateFile.worksheet(sheet_name)
        sheet.clear()

        data = []
        header = [
            'วันเรียนบรรยาย', 'คาบบรรยาย(เริ่ม)', 'คาบบรรยาย(จบ)', 'ห้องเรียนบรรยาย',
            'วันเรียนปฏิบัติ', 'คาบปฏิบัติ(เริ่ม)', 'คาบปฏิบัติ(จบ)', 'ห้องเรียนปฏิบัติ',
            'เซคเรียน', 'รหัสวิชา', 'อาจารย์'
        ]
        data.append(header)

        for schedule in timetable.schedule:
            row = [
                schedule.get('วันเรียนบรรยาย', ''),
                schedule.get('คาบบรรยาย(เริ่ม)', ''),
                schedule.get('คาบบรรยาย(จบ)', ''),
                schedule.get('ห้องเรียนบรรยาย', ''),
                
                schedule.get('วันเรียนปฏิบัติ', ''),
                schedule.get('คาบปฏิบัติ(เริ่ม)', ''),
                schedule.get('คาบปฏิบัติ(จบ)', ''),
                schedule.get('ห้องเรียนปฏิบัติ', ''),
                
                schedule.get('เซคเรียน', ''),
                schedule.get('รหัสวิชา', ''),
                schedule.get('อาจารย์', ''),
            ]
            data.append(row)

        # Use range_name and values parameters
        sheet.update('A1', data)
        print("Success!")
    
    except Exception as e:
        print(f"An error occurred: {e}")

# Load time slots and rooms from Main file
def load_data_from_main(client):
    mainFile = client.open('Main')
    timeSlotSheet = mainFile.worksheet('TimeSlot')
    timeSlots = [cell.value for cell in timeSlotSheet.range('C3:C14')]

    roomSheet = mainFile.worksheet('Room')
    rooms = [cell.value for cell in roomSheet.range('C3:C') if cell.value]

    return timeSlots, rooms

# Load room types from Main file
def load_room_types(client):
    mainFile = client.open('Main')
    roomSheet = mainFile.worksheet('Room')
    room_types = [cell.value for cell in roomSheet.range('G3:G') if cell.value]
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
                teacher = row[2]
                category = row[4]
                lectures = row[6]
                practicals = row[7]

                # Check if the category is "ศึกษาทั่วไป" and skip if it is
                if category == 'ศึกษาทั่วไป':
                    continue

                # Validate and convert to integer
                try:
                    lectures = int(lectures)
                    practicals = int(practicals)
                except ValueError:
                    print(f"Invalid data for lectures or practicals: {lectures}, {practicals}")
                    continue

                if course_code and section and teacher is not None:
                    curriculum.append({
                        'รหัสวิชา': course_code,
                        'เซคเรียน': section,
                        'อาจารย์': teacher,
                        'จำนวนคาบบรรยาย': lectures,
                        'จำนวนคาบปฏิบัติ': practicals
                    })

    except Exception as e:
        print(f"Failed to load courses curriculum: {e}")
    
    return curriculum

def load_teacher_availability(client):
    teacher_availability = {}
    try:
        teacher_file = client.open('Teacher')
        sheets = teacher_file.worksheets()

        for sheet in sheets:
            if sheet.title == 'คำแนะนำ':
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
    
# TimeTable class for managing schedule
class TimeTable:
    def __init__(self, timeSlots, rooms, room_types, curriculum, teacher_availability):
        self.timeSlots = timeSlots
        self.rooms = rooms
        self.room_types = room_types
        self.curriculum = curriculum
        self.teacher_availability = teacher_availability  # เพิ่มการรับข้อมูลความพร้อมของอาจารย์
        self.schedule = []
        self.fitness = 0

    def initialize(self):
        days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์']
        for course in self.curriculum:
            num_lectures = course.get('จำนวนคาบบรรยาย', 0)
            num_practicals = course.get('จำนวนคาบปฏิบัติ', 0)
            teacher_id = course.get('อาจารย์', '')

            # Initialize variables
            day_lectures = start_period_lectures = end_period_lectures = room_lectures = ''
            day_practicals = start_period_practicals = end_period_practicals = room_practicals = ''

            if num_lectures > 0:
                day_lectures = random.choice(days)
                start_period_lectures = random.choice(self.timeSlots)
                end_period_lectures = self.calculate_end_period(start_period_lectures, num_lectures)
                room_lectures = self.get_available_room('บรรยาย')

                # ตรวจสอบความพร้อมของอาจารย์สำหรับคาบบรรยาย
                if not self.check_teacher_availability(teacher_id, day_lectures, start_period_lectures):
                    continue  # ข้ามการเพิ่มตารางเรียนถ้าอาจารย์ไม่พร้อมสอน

            if num_practicals > 0:
                day_practicals = random.choice(days)
                start_period_practicals = random.choice(self.timeSlots)
                end_period_practicals = self.calculate_end_period(start_period_practicals, num_practicals)
                room_practicals = self.get_available_room('ปฏิบัติ')

                # ตรวจสอบความพร้อมของอาจารย์สำหรับคาบปฏิบัติ
                if not self.check_teacher_availability(teacher_id, day_practicals, start_period_practicals):
                    continue  # ข้ามการเพิ่มตารางเรียนถ้าอาจารย์ไม่พร้อมสอน

            # Add schedule entry
            self.schedule.append({
                'วันเรียนบรรยาย': day_lectures,
                'คาบบรรยาย(เริ่ม)': start_period_lectures,
                'คาบบรรยาย(จบ)': end_period_lectures,
                'ห้องเรียนบรรยาย' : room_lectures,
                
                'วันเรียนปฏิบัติ': day_practicals,
                'คาบปฏิบัติ(เริ่ม)': start_period_practicals,
                'คาบปฏิบัติ(จบ)': end_period_practicals,
                'ห้องเรียนปฏิบัติ' : room_practicals,
                
                'เซคเรียน': course['เซคเรียน'],
                'รหัสวิชา': course['รหัสวิชา'],
                'อาจารย์': teacher_id,
                
            })

        self.calculate_fitness()
        # Print fitness value after initialization
        print(f"Fitness Value: {self.fitness}")

    def check_teacher_availability(self, teacher_id, day, period):
        if teacher_id not in self.teacher_availability:
            return False
        
        availability = self.teacher_availability[teacher_id]['availability']
        
        if period not in availability:
            return False
        
        if day not in availability[period]:
            return False
        
        return availability[period][day] == 1


    def calculate_end_period(self, start_period, num_periods):
        start_index = self.timeSlots.index(start_period)
        end_index = min(start_index + num_periods - 1, len(self.timeSlots) - 1)
        return self.timeSlots[end_index]

    def get_available_room(self, course_type):
        available_rooms = [room for room in self.rooms if self.get_room_type_for_room(room) == course_type]
        if available_rooms:
            return random.choice(available_rooms)
        return random.choice(self.rooms)  # fallback to any room if no matching room found

    def get_room_type_for_room(self, room):
        try:
            index = self.rooms.index(room)
            return self.room_types[index] if index < len(self.room_types) else 'บรรยาย'
        except ValueError:
            return 'บรรยาย'

    def calculate_fitness(self):
        conflicts = 0
        
        for i in range(len(self.schedule)):
            for j in range(i + 1, len(self.schedule)):
                # Check conflicts for lectures
                if (self.schedule[i]['คาบบรรยาย(เริ่ม)'] == self.schedule[j]['คาบบรรยาย(เริ่ม)'] and 
                    self.schedule[i]['วันเรียนบรรยาย'] == self.schedule[j]['วันเรียนบรรยาย']):
                    if (self.schedule[i]['ห้องเรียนบรรยาย'] == self.schedule[j]['ห้องเรียนบรรยาย'] or
                        self.schedule[i]['อาจารย์'] == self.schedule[j]['อาจารย์']):
                        conflicts += 1
                
                # Check conflicts for practicals
                if (self.schedule[i]['คาบปฏิบัติ(เริ่ม)'] == self.schedule[j]['คาบปฏิบัติ(เริ่ม)'] and 
                    self.schedule[i]['วันเรียนปฏิบัติ'] == self.schedule[j]['วันเรียนปฏิบัติ']):
                    if (self.schedule[i]['ห้องเรียนปฏิบัติ'] == self.schedule[j]['ห้องเรียนปฏิบัติ'] or
                        self.schedule[i]['อาจารย์'] == self.schedule[j]['อาจารย์']):
                        conflicts += 1

        self.fitness = -conflicts


    def crossover(self, partner):
        midpoint = random.randint(0, len(self.schedule) - 1)
        child = TimeTable(self.timeSlots, self.rooms, self.room_types, self.curriculum)
        child.schedule = self.schedule[:midpoint] + partner.schedule[midpoint:]
        child.calculate_fitness()
        return child

    def mutate(self, mutation_rate):
        days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์']
        for schedule in self.schedule:
            if random.random() < mutation_rate:
                # Randomly choose between lecture and practical
                if random.random() < 0.5:
                    schedule['วันเรียนบรรยาย'] = random.choice(days)
                    schedule['คาบบรรยาย(เริ่ม)'] = random.choice(self.timeSlots)
                    schedule['คาบบรรยาย(จบ)'] = self.calculate_end_period(
                        schedule['คาบบรรยาย(เริ่ม)'],
                        len(schedule['คาบบรรยาย(เริ่ม)'])
                    )
                    schedule['ห้องเรียนบรรยาย'] = self.get_available_room('บรรยาย')
                else:
                    schedule['วันเรียนปฏิบัติ'] = random.choice(days)
                    schedule['คาบปฏิบัติ(เริ่ม)'] = random.choice(self.timeSlots)
                    schedule['คาบปฏิบัติ(จบ)'] = self.calculate_end_period(
                        schedule['คาบปฏิบัติ(เริ่ม)'],
                        len(schedule['คาบปฏิบัติ(เริ่ม)'])
                    )
                    schedule['ห้องเรียนปฏิบัติ'] = self.get_available_room('ปฏิบัติ')

        self.calculate_fitness()



def genetic_algorithm(client):
    timeSlots, rooms = load_data_from_main(client)
    room_types = load_room_types(client)
    curriculum = load_courses_curriculum(client)
    teacher_availability = load_teacher_availability(client)

    # Initialize population
    population_size = 10
    generations = 100
    mutation_rate = 0.01

    # สร้างประชากรเริ่มต้น
    population = [TimeTable(timeSlots, rooms, room_types, curriculum, teacher_availability) for _ in range(population_size)]
    for timetable in population:
        timetable.initialize()

    # Evolve the population
    for generation in range(generations):
        # คัดเลือกประชากรที่มีความฟิตมากที่สุด 50% มาเพื่อการผสมพันธุ์
        population.sort(key=lambda x: x.fitness, reverse=True)
        print(f"Generation {generation}: Best Fitness = {population[0].fitness}")

        # หากได้ผลลัพธ์ที่ไม่มีความขัดแย้ง (fitness = 0) หยุดการรัน
        if population[0].fitness == 0:
            print("Optimal solution found!")
            break

        # คัดเลือกครึ่งหนึ่งของประชากรที่มีความฟิตมากที่สุด
        parents = population[:population_size // 2]

        # สร้างลูกใหม่จากการผสมพันธุ์
        children = []
        for _ in range(population_size // 2):
            parent1 = random.choice(parents)
            parent2 = random.choice(parents)
            child = parent1.crossover(parent2)
            child.mutate(mutation_rate)
            children.append(child)

        # ประกอบประชากรใหม่จากครึ่งหนึ่งของพ่อแม่และลูกใหม่
        population = parents + children

    # ผลลัพธ์ที่ดีที่สุดหลังจากสิ้นสุดการวิวัฒนาการ
    best_timetable = population[0]

    # บันทึกผลลัพธ์ที่ดีที่สุดไปยัง Google Sheets
    write_timetable_to_sheet(best_timetable, "Best Timetable")



genetic_algorithm(client)
