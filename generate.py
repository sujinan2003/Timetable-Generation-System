import random
import gspread
import time

from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)

client = gspread.authorize(credentials)

#เขียนข้อมูลลงใน Google Sheets
def write_timetable_to_sheet(timetable, sheet_name):
    try:
        #เปิดไฟล์
        generateFile = client.open('Generate')
        #ตรวจสอบว่ามีชีทที่ต้องการหรือยัง ถ้ายังไม่มีให้สร้างใหม่
        if sheet_name not in [sheet.title for sheet in generateFile.worksheets()]:
            generateFile.add_worksheet(title=sheet_name, rows="100", cols="20")

        sheet = generateFile.worksheet(sheet_name)

        #ลบข้อมูลทั้งหมดในชีทก่อนเขียน
        sheet.clear()

        #กำหนดข้อมูลที่จะเขียน
        data = []
        header = ['คาบเรียน', 'เซคเรียน', 'รหัสวิชา', 'อาจารย์', 'ห้องเรียน', 'วัน']
        data.append(header)
        for schedule in timetable.schedule:
            row = [
                schedule['คาบเรียน'],
                schedule['เซคเรียน'],
                schedule['รหัสวิชา'],
                schedule['อาจารย์'],
                schedule['ห้องเรียน'],
                schedule['วัน']
            ]
            data.append(row)

        # เขียนข้อมูลลงใน Google Sheets
        sheet.update('A1', data)

        print(f"Success!")
    
    except Exception as e:
        print(f"An error occurred: {e}")


#ดึงข้อมูล TimeSlot และห้องเรียนจากไฟล์ Main มาเก็บไว้
def load_data_from_main(client):
    mainFile = client.open('Main')
    timeSlotSheet = mainFile.worksheet('TimeSlot')
    timeSlots = [cell.value for cell in timeSlotSheet.range('C3:C14')]

    roomSheet = mainFile.worksheet('Room')
    rooms = [cell.value for cell in roomSheet.range('C3:C') if cell.value]

    return timeSlots, rooms

#ดึงข้อมูล Open Course และ Curriculum จากไฟล์
def load_courses_curriculum(client):
    curriculumFile = client.open('Curriculum')
    curriculum = []
    openCourseFile = client.open('Open Course')
    openCourseSheets = openCourseFile.worksheets()

    for sheet in openCourseSheets:
        data = sheet.get_all_values()
        for row in data[1:]:  # Skip header row
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

#กำหนดคลาส TimeTable สำหรับจัดการตารางเรียน
class TimeTable:
    def __init__(self, timeSlots, rooms, curriculum):
        self.timeSlots = timeSlots
        self.rooms = rooms
        self.curriculum = curriculum
        self.schedule = []
        self.fitness = 0

    def initialize(self):
        days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์']
        for course in self.curriculum:
            timeSlot = random.choice(self.timeSlots)
            room = random.choice(self.rooms)
            day = random.choice(days)
            self.schedule.append({
                'คาบเรียน': timeSlot,
                'เซคเรียน': course['เซคเรียน'],
                'รหัสวิชา': course['รหัสวิชา'],
                'อาจารย์': course['อาจารย์'],
                'ห้องเรียน': room,
                'วัน': day
            })
        self.calculate_fitness()

    def calculate_fitness(self):
        conflicts = 0
        for i in range(len(self.schedule)):
            for j in range(i + 1, len(self.schedule)):
                if self.schedule[i]['คาบเรียน'] == self.schedule[j]['คาบเรียน'] and self.schedule[i]['วัน'] == self.schedule[j]['วัน']:
                    if self.schedule[i]['ห้องเรียน'] == self.schedule[j]['ห้องเรียน'] or self.schedule[i]['อาจารย์'] == self.schedule[j]['อาจารย์']:
                        conflicts += 1
        self.fitness = -conflicts

    def crossover(self, partner):
        midpoint = random.randint(0, len(self.schedule) - 1)
        child = TimeTable(self.timeSlots, self.rooms, self.curriculum)
        child.schedule = self.schedule[:midpoint] + partner.schedule[midpoint:]
        child.calculate_fitness()
        return child

    def mutate(self, mutation_rate):
        days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์']
        for i in range(len(self.schedule)):
            if random.random() < mutation_rate:
                self.schedule[i]['คาบเรียน'] = random.choice(self.timeSlots)
                self.schedule[i]['ห้องเรียน'] = random.choice(self.rooms)
                self.schedule[i]['วัน'] = random.choice(days)
        self.calculate_fitness()

#สร้างประชากรเริ่มต้น
def create_initial_population(pop_size, timeSlots, rooms, curriculum):
    population = []
    for _ in range(pop_size):
        timetable = TimeTable(timeSlots, rooms, curriculum)
        timetable.initialize()
        population.append(timetable)
    return population

#เลือก perents for cossover
def select_parents(population):
    fitness_sum = sum([timetable.fitness for timetable in population])
    if fitness_sum == 0:
        return random.choice(population)  #กรณีไม่มีค่า fitness ที่ไม่ใช่ 0 ให้สุ่มเลือกค่าใน population มาแทน
    pick = random.uniform(0, fitness_sum)
    current = 0
    for timetable in population:
        current += timetable.fitness
        if current > pick:
            return timetable
    return random.choice(population)  #กรณีไม่พบตารางเรียนที่มีค่า fitness มากกว่า pick ให้สุ่มเลือกค่าใน population มาแทน

#ค้นหาตารางเรียนที่ดีที่สุด
def genetic_algorithm(pop_size, timeSlots, rooms, curriculum, generations, mutation_rate):
    population = create_initial_population(pop_size, timeSlots, rooms, curriculum)
    for generation in range(generations):
        new_population = []
        for _ in range(pop_size):
            parent1 = select_parents(population)
            parent2 = select_parents(population)
            child = parent1.crossover(parent2)
            child.mutate(mutation_rate)
            new_population.append(child)
        population = sorted(new_population, key=lambda x: x.fitness, reverse=True)
        print(f'Generation {generation} Best Fitness: {population[0].fitness}')
    return population[0]

#กำหนดค่าเริ่มต้น
pop_size = 50
generations = 100
mutation_rate = 0.01

#ดึงข้อมูลจาก Google Sheets
timeSlots, rooms = load_data_from_main(client)
curriculum = load_courses_curriculum(client)

#run Genetic
best_timetable = genetic_algorithm(pop_size, timeSlots, rooms, curriculum, generations, mutation_rate)

#print("Best Schedule: ", best_timetable.schedule)
write_timetable_to_sheet(best_timetable, 'Best Timetable')