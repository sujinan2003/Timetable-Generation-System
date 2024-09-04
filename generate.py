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
        sheet.update(range_name='A1', values=data)
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
                lectures = row[6]
                practicals = row[7]

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

    def get_room_type_for_course(self, course_type):
        return course_type  # Assuming room types match course types directly

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
                schedule['คาบบรรยาย(จบ)'] = self.calculate_end_period(schedule['คาบบรรยาย(เริ่ม)'], 1)
                schedule['คาบปฏิบัติ(เริ่ม)'] = random.choice(self.timeSlots)
                schedule['คาบปฏิบัติ(จบ)'] = self.calculate_end_period(schedule['คาบปฏิบัติ(เริ่ม)'], 1)

# Example Usage
def main():
    timeSlots, rooms = load_data_from_main(client)
    room_types = load_room_types(client)
    curriculum = load_courses_curriculum(client)

    timetable = TimeTable(timeSlots, rooms, room_types, curriculum)
    timetable.initialize()
    write_timetable_to_sheet(timetable, 'Sample_Timetable')

if __name__ == "__main__":
    main()
