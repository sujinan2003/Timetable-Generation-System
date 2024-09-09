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
            'วันเรียนบรรยาย', 'คาบบรรยาย(เริ่ม)', 'คาบบรรยาย(จบ)', 
            'วันเรียนปฏิบัติ', 'คาบปฏิบัติ(เริ่ม)', 'คาบปฏิบัติ(จบ)', 
            'เซคเรียน', 'รหัสวิชา', 'อาจารย์', 'ห้องเรียน'
        ]
        data.append(header)

        for schedule in timetable.schedule:
            row = [
                schedule.get('วันเรียนบรรยาย', ''),
                schedule.get('คาบบรรยาย(เริ่ม)', ''),
                schedule.get('คาบบรรยาย(จบ)', ''),
                schedule.get('วันเรียนปฏิบัติ', ''),
                schedule.get('คาบปฏิบัติ(เริ่ม)', ''),
                schedule.get('คาบปฏิบัติ(จบ)', ''),
                schedule.get('เซคเรียน', ''),
                schedule.get('รหัสวิชา', ''),
                schedule.get('อาจารย์', ''),
                schedule.get('ห้องเรียน', ''),
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
            teacher_id = sheet.title
            data = sheet.get_all_values()

            if not data or len(data) < 3:
                print(f"Insufficient data in sheet {teacher_id}.")
                continue

            # Extract headers and periods
            headers = data[1]  # Headers: คาบ, เวลา, จันทร์, อังคาร, ..., อาทิตย์
            periods = [row[0] for row in data[2:]]  # Periods from A3:A

            # Initialize availability dictionary
            availability = {}
            for idx, period in enumerate(periods):
                period_number = period  # e.g., 1, 2, 3, ...
                availability[period_number] = {}
                for day_index, availability_status in enumerate(data[idx + 2][2:]):
                    day = headers[day_index + 2]  # Mapping day names
                    try:
                        availability[period_number][day] = int(availability_status)  # Convert to integer
                    except ValueError:
                        availability[period_number][day] = 0  # Default to 0 if conversion fails

            teacher_availability[teacher_id] = {'availability': availability}

    except Exception as e:
        print(f"Failed to load teacher availability: {e}")

    return teacher_availability


def check_teacher_availability(teacher_availability, teacher_id, day, period):
    # ตรวจสอบว่ามีอาจารย์ตาม ID ที่ระบุหรือไม่
    if teacher_id not in teacher_availability:
        print(f"Teacher ID {teacher_id} not found.")
        return False
    
    # ดึงข้อมูลความพร้อมในการสอนของอาจารย์
    availability = teacher_availability[teacher_id]['availability']

    # ตรวจสอบว่ามีข้อมูลความพร้อมสำหรับช่วงเวลาและวันที่ระบุหรือไม่
    if period not in availability:
        print(f"Period {period} not found in teacher {teacher_id}'s schedule.")
        return False
    
    if day not in availability[period]:
        print(f"Day {day} not found in teacher {teacher_id}'s schedule.")
        return False
    
    # ตรวจสอบสถานะความพร้อม
    status = availability[period][day]
    if status == 1:
        print(f"Teacher {teacher_id} is available on {day} during period {period}.")
        return True
    else:
        print(f"Teacher {teacher_id} is not available on {day} during period {period}.")
        return False



# TimeTable class for managing schedule
class TimeTable:
    def __init__(self, timeSlots, rooms, room_types, curriculum):
        self.timeSlots = timeSlots
        self.rooms = rooms
        self.room_types = room_types
        self.curriculum = curriculum
        self.schedule = []
        self.fitness = 0

    def initialize(self):
        days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์']
        for course in self.curriculum:
            num_lectures = course.get('จำนวนคาบบรรยาย', 0)
            num_practicals = course.get('จำนวนคาบปฏิบัติ', 0)

            # Initialize variables
            day_lectures = start_period_lectures = end_period_lectures = room_lectures = ''
            day_practicals = start_period_practicals = end_period_practicals = room_practicals = ''

            if num_lectures > 0:
                day_lectures = random.choice(days)
                start_period_lectures = random.choice(self.timeSlots)
                end_period_lectures = self.calculate_end_period(start_period_lectures, num_lectures)
                room_lectures = self.get_available_room('บรรยาย')

            if num_practicals > 0:
                day_practicals = random.choice(days)
                start_period_practicals = random.choice(self.timeSlots)
                end_period_practicals = self.calculate_end_period(start_period_practicals, num_practicals)
                room_practicals = self.get_available_room('ปฏิบัติ')

            # Add schedule entry
            self.schedule.append({
                'วันเรียนบรรยาย': day_lectures,
                'คาบบรรยาย(เริ่ม)': start_period_lectures,
                'คาบบรรยาย(จบ)': end_period_lectures,
                'วันเรียนปฏิบัติ': day_practicals,
                'คาบปฏิบัติ(เริ่ม)': start_period_practicals,
                'คาบปฏิบัติ(จบ)': end_period_practicals,
                'เซคเรียน': course['เซคเรียน'],
                'รหัสวิชา': course['รหัสวิชา'],
                'อาจารย์': course['อาจารย์'],
                'ห้องเรียน': room_lectures if num_lectures > 0 else room_practicals
            })

        self.calculate_fitness()
        # Print fitness value after initialization
        print(f"Fitness Value: {self.fitness}")

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
                if (self.schedule[i]['คาบบรรยาย(เริ่ม)'] == self.schedule[j]['คาบบรรยาย(เริ่ม)'] and 
                    self.schedule[i]['วันเรียนบรรยาย'] == self.schedule[j]['วันเรียนบรรยาย']):
                    if (self.schedule[i]['ห้องเรียน'] == self.schedule[j]['ห้องเรียน'] or
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
                if random.random() < 0.5:
                    schedule['วันเรียนบรรยาย'] = random.choice(days)
                else:
                    schedule['วันเรียนปฏิบัติ'] = random.choice(days)
                schedule['คาบบรรยาย(เริ่ม)'] = random.choice(self.timeSlots)
                schedule['คาบบรรยาย(จบ)'] = self.calculate_end_period(
                    schedule['คาบบรรยาย(เริ่ม)'],
                    len(schedule['คาบบรรยาย(เริ่ม)'])
                )
                schedule['ห้องเรียน'] = self.get_available_room('บรรยาย')
        self.calculate_fitness()

# Main function to run the genetic algorithm
def genetic_algorithm(client):
    timeSlots, rooms = load_data_from_main(client)
    room_types = load_room_types(client)
    curriculum = load_courses_curriculum(client)

    # Initialize population
    population_size = 10
    generations = 100
    mutation_rate = 0.01

    population = [TimeTable(timeSlots, rooms, room_types, curriculum) for _ in range(population_size)]
    for timetable in population:
        timetable.initialize()

    # Evolve the population
    for generation in range(generations):
        # Sort by fitness (higher is better)
        population.sort(key=lambda x: x.fitness, reverse=True)

        # Print best fitness
        print(f"Generation {generation}: Best Fitness = {population[0].fitness}")

        # Selection and Crossover
        next_generation = [population[0]]  # Elitism - keep best
        for _ in range(1, population_size):
            parent1, parent2 = random.choices(population[:5], k=2)  # Select top 5 for mating
            child = parent1.crossover(parent2)
            next_generation.append(child)

        # Mutation
        for timetable in next_generation[1:]:  # Skip the best one
            timetable.mutate(mutation_rate)

        # Replace old population with new generation
        population = next_generation

    # Save the best result to Google Sheets
    best_timetable = population[0]
    write_timetable_to_sheet(best_timetable, 'Final_Timetable')

def main():
    # ตัวอย่างการเรียกใช้ฟังก์ชันตรวจสอบความพร้อม
    teacher_availability = load_teacher_availability(client)
    
    # ตรวจสอบความพร้อมของอาจารย์
    teacher_id = "T101"  # รหัสอาจารย์ตัวอย่าง
    day = "พฤหัส"  # วันตัวอย่าง
    period = 1  # ช่วงเวลาตัวอย่าง (คาบตัวอย่าง)

    available = check_teacher_availability(teacher_availability, teacher_id, day, period)
    if available:
        print(f"Teacher {teacher_id} is available for scheduling.")
    else:
        print(f"Teacher {teacher_id} is not available for scheduling.")

if __name__ == "__main__":
    main()

