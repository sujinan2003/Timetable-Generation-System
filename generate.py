import random
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

# ตั้งค่าการเชื่อมต่อกับ Google Sheets โดยใช้ OAuth2 credentials
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)
client = gspread.authorize(credentials)

def write_timetable_to_sheet(timetable, sheet_name):
    """
    ฟังก์ชันนี้ใช้เพื่อเขียนตารางเรียนไปยังชีทใน Google Sheets
    """
    try:
        # เชื่อมต่อกับไฟล์ Google Sheets ชื่อ 'Generate'
        generateFile = client.open('Generate')

        # ตรวจสอบว่าชีทที่ระบุมีอยู่แล้วหรือไม่ ถ้าไม่มีก็เพิ่มชีทใหม่
        if sheet_name not in [sheet.title for sheet in generateFile.worksheets()]:
            generateFile.add_worksheet(title=sheet_name, rows="100", cols="20")

        # เข้าถึงชีทที่ระบุ
        sheet = generateFile.worksheet(sheet_name)
        sheet.clear()  # ล้างข้อมูลเดิมในชีท

        # เตรียมข้อมูลสำหรับ DataFrame
        data = []
        header = ['คาบเรียน', 'เซคเรียน', 'รหัสวิชา', 'อาจารย์', 'ห้องเรียน', 'วัน', 'ประเภทวิชา', 'ประเภทห้องเรียน']

        # สร้างพจนานุกรมเพื่อเก็บข้อมูลที่เห็นแล้ว
        seen_records = {}

        # ตรวจสอบตารางเรียนและเก็บข้อมูลที่ไม่ซ้ำกัน
        for schedule in timetable.schedule:
            key = (schedule['รหัสวิชา'], schedule['เซคเรียน'], schedule['ประเภทวิชา'])
            if key not in seen_records:
                seen_records[key] = {
                    'คาบเรียน': schedule['คาบเรียน'],
                    'เซคเรียน': schedule['เซคเรียน'],
                    'รหัสวิชา': schedule['รหัสวิชา'],
                    'อาจารย์': schedule['อาจารย์'],
                    'ห้องเรียน': schedule['ห้องเรียน'],
                    'วัน': schedule['วัน'],
                    'ประเภทวิชา': schedule['ประเภทวิชา'],
                    'ประเภทห้องเรียน': schedule['ประเภทห้องเรียน']
                }
            # เพิ่มข้อมูลที่เห็นแล้วลงใน data
            data.append(seen_records[key])

        # สร้าง DataFrame จากข้อมูล
        df = pd.DataFrame(data, columns=header)

        # แปลง DataFrame เป็นรายการของรายการ
        data_list = df.values.tolist()
        data_list.insert(0, header)  # เพิ่มหัวเรื่องที่ด้านบน

        # อัปเดตข้อมูลลงใน Google Sheet
        sheet.update('A1', data_list)

        print("สำเร็จ!")  # แสดงข้อความเมื่อเขียนข้อมูลสำเร็จ

    except Exception as e:
        print(f"เกิดข้อผิดพลาด: {e}")  # แสดงข้อความข้อผิดพลาดหากมีข้อผิดพลาด

def adjust_period_format(period):
    """
    ปรับฟอร์แมตของคาบเรียนให้เป็นสองหลักเสมอ
    """
    try:
        return f"{int(period):02d}"  # แปลงคาบเรียนให้เป็นสองหลัก
    except ValueError:
        return period  # คืนค่าต้นฉบับหากการแปลงล้มเหลว

def load_data_from_main(client):
    """
    โหลดคาบเรียนและห้องจากไฟล์หลัก
    """
    mainFile = client.open('Main')
    timeSlotSheet = mainFile.worksheet('TimeSlot')
    timeSlots = [cell.value for cell in timeSlotSheet.range('C3:C14') if cell.value.isdigit() and 1 <= int(cell.value) <= 12]

    roomSheet = mainFile.worksheet('Room')
    rooms = [cell.value for cell in roomSheet.range('C3:C') if cell.value]

    return timeSlots, rooms

def load_room_types(client):
    """
    โหลดประเภทห้องเรียนจากไฟล์หลัก
    """
    mainFile = client.open('Main')
    roomSheet = mainFile.worksheet('Room')
    room_types = [cell.value for cell in roomSheet.range('G3:G') if cell.value]
    return room_types

def load_courses_curriculum(client):
    """
    โหลดข้อมูลหลักสูตรและวิชาจากไฟล์ Curriculum และ Open Course
    """
    curriculumFile = client.open('Curriculum')
    curriculum = []
    openCourseFile = client.open('Open Course')
    openCourseSheets = openCourseFile.worksheets()

    for sheet in openCourseSheets:
        data = sheet.get_all_values()
        for row in data[1:]:  # ข้ามแถวหัวเรื่อง
            course_code = row[0]
            section = row[1]
            teacher = row[2]
            if course_code and section and teacher:
                curriculum.append({
                    'รหัสวิชา': course_code,
                    'เซคเรียน': section,
                    'อาจารย์': teacher
                })

    return curriculum

class TimeTable:
    """
    คลาสสำหรับจัดการตารางเรียนและทำงานกับข้อมูลตารางเรียน
    """
    def __init__(self, timeSlots, rooms, room_types, curriculum):
        self.timeSlots = timeSlots
        self.rooms = rooms
        self.room_types = room_types
        self.curriculum = curriculum
        self.schedule = []  # รายการของตารางเรียน
        self.fitness = 0    # คะแนนความเหมาะสมของตารางเรียน

    def initialize(self):
        """
        สร้างตารางเรียนเบื้องต้นโดยใช้ข้อมูลหลักสูตร
        """
        days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์']
        for course in self.curriculum:
            lecture_slots = self.get_lecture_slots(course['รหัสวิชา'])
            practical_slots = self.get_practical_slots(course['รหัสวิชา'])

            # จัดตารางเรียนสำหรับการบรรยาย
            for _ in range(lecture_slots):
                timeSlot = self.get_consecutive_slot()
                room = self.get_available_room('บรรยาย')
                day = random.choice(days)
                self.schedule.append({
                    'คาบเรียน': timeSlot,
                    'เซคเรียน': course['เซคเรียน'],
                    'รหัสวิชา': course['รหัสวิชา'],
                    'อาจารย์': course['อาจารย์'],
                    'ห้องเรียน': room,
                    'วัน': day,
                    'ประเภทวิชา': 'บรรยาย',
                    'ประเภทห้องเรียน': self.get_room_type_for_course('บรรยาย')
                })

            # จัดตารางเรียนสำหรับการปฏิบัติ
            for _ in range(practical_slots):
                timeSlot = self.get_consecutive_slot()
                room = self.get_available_room('ปฏิบัติ')
                day = random.choice(days)
                self.schedule.append({
                    'คาบเรียน': timeSlot,
                    'เซคเรียน': course['เซคเรียน'],
                    'รหัสวิชา': course['รหัสวิชา'],
                    'อาจารย์': course['อาจารย์'],
                    'ห้องเรียน': room,
                    'วัน': day,
                    'ประเภทวิชา': 'ปฏิบัติ',
                    'ประเภทห้องเรียน': self.get_room_type_for_course('ปฏิบัติ')
                })
        self.calculate_fitness()

    def get_lecture_slots(self, course_code):
        """
        กำหนดจำนวนคาบเรียนสำหรับการบรรยายของรหัสวิชา
        """
        return 2  # ตัวอย่าง: จำนวนคาบเรียนบรรยาย

    def get_practical_slots(self, course_code):
        """
        กำหนดจำนวนคาบเรียนสำหรับการปฏิบัติของรหัสวิชา
        """
        return 3  # ตัวอย่าง: จำนวนคาบเรียนปฏิบัติ

    def get_consecutive_slot(self):
        """
        เลือกคาบเรียนที่ถูกต้อง (1 ถึง 12)
        """
        valid_slots = [slot for slot in self.timeSlots if int(slot) >= 1 and int(slot) <= 12]
        return random.choice(valid_slots)

    def get_available_room(self, course_type):
        """
        เลือกห้องเรียนที่ว่างตามประเภทวิชา
        """
        available_rooms = [room for room in self.rooms if self.get_room_type_for_room(room) == course_type]
        if available_rooms:
            return random.choice(available_rooms)
        return random.choice(self.rooms)  # หากไม่พบห้องที่ตรงตามประเภท

    def get_room_type_for_room(self, room):
        """
        คืนประเภทห้องเรียนของห้องเรียนที่ระบุ
        """
        index = self.rooms.index(room)
        return self.room_types[index] if index < len(self.room_types) else 'บรรยาย'

    def get_room_type_for_course(self, course_type):
        """
        คืนประเภทห้องเรียนที่เหมาะสมตามประเภทวิชา
        """
        return course_type  # สมมติว่า ประเภทห้องเรียนตรงกับประเภทวิชา

    def calculate_fitness(self):
        """
        คำนวณคะแนนความเหมาะสมของตารางเรียน โดยตรวจสอบความขัดแย้ง
        """
        conflicts = 0
        for i in range(len(self.schedule)):
            for j in range(i + 1, len(self.schedule)):
                if (self.schedule[i]['คาบเรียน'] == self.schedule[j]['คาบเรียน'] and
                    self.schedule[i]['วัน'] == self.schedule[j]['วัน'] and
                    self.schedule[i]['ห้องเรียน'] == self.schedule[j]['ห้องเรียน']):
                    conflicts += 1
        self.fitness = 1 / (1 + conflicts)  # คะแนนความเหมาะสม

    def crossover(self, other):
        """
        สร้างลูกจากการครอสโอเวอร์ระหว่างตารางเรียนสองชุด
        """
        crossover_point = random.randint(1, len(self.schedule) - 1)
        child_schedule = self.schedule[:crossover_point] + other.schedule[crossover_point:]
        child = TimeTable(self.timeSlots, self.rooms, self.room_types, self.curriculum)
        child.schedule = child_schedule
        child.calculate_fitness()
        return child

    def mutate(self, mutation_rate):
        """
        ดัดแปลงตารางเรียนโดยการสุ่มเปลี่ยนค่า
        """
        valid_slots = [slot for slot in self.timeSlots if int(slot) >= 1 and int(slot) <= 12]
        for idx in range(len(self.schedule)):
            if random.random() < mutation_rate:
                self.schedule[idx]['คาบเรียน'] = random.choice(valid_slots)
                self.schedule[idx]['ห้องเรียน'] = random.choice(self.rooms)
                self.schedule[idx]['วัน'] = random.choice(['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์'])
        self.calculate_fitness()

def genetic_algorithm(timeSlots, rooms, room_types, curriculum, generations, population_size, mutation_rate):
    """
    ฟังก์ชันหลักของ Genetic Algorithm สำหรับการสร้างตารางเรียนที่เหมาะสม
    """
    # สร้างประชากรเริ่มต้น
    population = [TimeTable(timeSlots, rooms, room_types, curriculum) for _ in range(population_size)]
    for timetable in population:
        timetable.initialize()

    # ดำเนินการตามจำนวนรุ่นที่กำหนด
    for generation in range(generations):
        # จัดอันดับประชากรตามคะแนนความเหมาะสม
        population.sort(key=lambda x: x.fitness, reverse=True)
        print(f"Generation {generation} Best Fitness: {population[0].fitness}")

        # เก็บพ่อพันธุ์แม่พันธุ์ที่ดีที่สุด
        new_population = population[:2]

        # สร้างประชากรใหม่โดยการครอสโอเวอร์และการกลายพันธุ์
        while len(new_population) < population_size:
            parent1 = random.choice(population[:10])
            parent2 = random.choice(population[:10])
            child = parent1.crossover(parent2)
            child.mutate(mutation_rate)
            new_population.append(child)

        population = new_population

    # ค้นหาตารางเรียนที่ดีที่สุด
    population.sort(key=lambda x: x.fitness, reverse=True)
    best_timetable = population[0]
    return best_timetable

def main():
    """
    ฟังก์ชันหลักในการทำงานของโปรแกรม
    """
    # โหลดข้อมูลจาก Google Sheets
    timeSlots, rooms = load_data_from_main(client)
    room_types = load_room_types(client)
    curriculum = load_courses_curriculum(client)

    # กำหนดพารามิเตอร์ของ Genetic Algorithm
    generations = 100
    population_size = 10
    mutation_rate = 0.01

    # เรียกใช้งาน Genetic Algorithm และเขียนตารางเรียนที่ดีที่สุดลงใน Google Sheets
    best_timetable = genetic_algorithm(timeSlots, rooms, room_types, curriculum, generations, population_size, mutation_rate)
    write_timetable_to_sheet(best_timetable, 'Generated Timetable')

if __name__ == "__main__":
    main()
