import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# ตั้งค่าการเข้าถึง Google Sheets
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)
client = gspread.authorize(credentials)

def write_timetable_to_student_sheet():
    try:
        # เปิดไฟล์ Generate และดึงข้อมูล
        generate_file = client.open('Generate')
        generate_sheet = generate_file.worksheet('Best Timetable')
        timetable_data = generate_sheet.get_all_records()

        # ดึงข้อมูลเซคเรียนจากเซลล์ B2:B ในไฟล์ Generate
        sections = generate_sheet.col_values(2)[1:]  # ข้ามหัวข้อในแถวที่ 1

        # เปิดไฟล์ Student
        student_file = client.open('Student')
        student_sheets = student_file.worksheets()

        for student_sheet in student_sheets:
            if student_sheet.title == 'Sheet1':
                continue  # ข้ามชีทที่ชื่อ 'Sheet1'

            seen_sections = set()

            for section in sections:
                if section in seen_sections:
                    continue

                # ตรวจสอบว่า section_name มีอยู่ใน sheet ปัจจุบันหรือไม่
                section_name_cell = student_sheet.find(section)
                if not section_name_cell:
                    print(f"Section {section} not found in sheet {student_sheet.title}.")
                    continue
                
                seen_sections.add(section)
                section_name_row = section_name_cell.row

                # ลบข้อมูลในช่วงที่ต้องการเขียน แต่ไม่ลบคอลัมน์วัน (A)
                for row in range(section_name_row + 3, section_name_row + 10):  # แถวที่ 3 ถึง 9 หลัง section name row
                    cell_start = gspread.utils.rowcol_to_a1(row, 2)
                    cell_end = gspread.utils.rowcol_to_a1(row, 13)
                    student_sheet.batch_clear([f'{cell_start}:{cell_end}'])
                    time.sleep(1)  # หน่วงเวลาระหว่างการลบข้อมูล

                # สร้างข้อมูลที่ต้องการเขียนในรูปแบบ batch_update
                day_to_row = {'จันทร์': section_name_row + 3, 'อังคาร': section_name_row + 4, 'พุธ': section_name_row + 5, 'พฤหัสบดี': section_name_row + 6, 'ศุกร์': section_name_row + 7}
                batch_data = []
                for schedule in timetable_data:
                    if schedule['เซคเรียน'] != section:
                        continue
                    
                    class_period = schedule['คาบเรียน']
                    day = schedule['วัน']
                    course_info = f"{schedule['รหัสวิชา']}\n{schedule['อาจารย์']}\n{schedule['ห้องเรียน']}"

                    # หาตำแหน่งของข้อมูลในตาราง
                    day_row_index = day_to_row[day]
                    column_index = int(class_period) + 1  # Offset due to header columns

                    cell_range = gspread.utils.rowcol_to_a1(day_row_index, column_index)
                    batch_data.append({
                        'range': cell_range,
                        'values': [[course_info]]
                    })

                if batch_data:
                    # เพิ่มข้อมูลลงในชีทใหม่
                    student_sheet.batch_update(batch_data)
                    time.sleep(5)  # หน่วงเวลาระหว่างการเรียกใช้ API

        print("Data has been successfully written to the 'Student' file.")
    
    except Exception as e:
        print(f"An error occurred: {str(e)}")


def write_timetable_to_room_sheet():
    try:
        # เปิดไฟล์ Generate และดึงข้อมูล
        generate_file = client.open('Generate')
        generate_sheet = generate_file.worksheet('Best Timetable')
        timetable_data = generate_sheet.get_all_records()

        # ดึงข้อมูลห้องเรียนจากเซลล์ E2:E ในไฟล์ Generate
        rooms = generate_sheet.col_values(5)[1:]  # ข้ามหัวข้อในแถวที่ 1
        
        # เปิดไฟล์ Room
        room_file = client.open('Room')
        room_sheets = room_file.worksheets()

        # ใช้ชุดข้อมูลที่บันทึกห้องเรียนที่ถูกจองแล้ว
        booked_rooms = {}

        for schedule in timetable_data:
            room = schedule['ห้องเรียน']
            day = schedule['วัน']
            class_period = schedule['คาบเรียน']

            if room not in booked_rooms:
                booked_rooms[room] = {}
            if day not in booked_rooms[room]:
                booked_rooms[room][day] = set()
            booked_rooms[room][day].add(class_period)

        # ตรวจสอบและอัพเดตข้อมูลในชีทต่าง ๆ
        for room_sheet in room_sheets:
            room_sheet_title = room_sheet.title

            if room_sheet_title not in booked_rooms:
                continue

            # ลบข้อมูลในช่วงที่ต้องการเขียน แต่ไม่ลบคอลัมน์วัน (A)
            for row in range(3, 9):  
                cell_start = gspread.utils.rowcol_to_a1(row, 2)
                cell_end = gspread.utils.rowcol_to_a1(row, 13)
                room_sheet.batch_clear([f'{cell_start}:{cell_end}'])
                
            # สร้างข้อมูลที่ต้องการเขียนในรูปแบบ batch_update
            batch_data = []
            day_to_row = {'จันทร์': 3, 'อังคาร': 4, 'พุธ': 5, 'พฤหัสบดี': 6, 'ศุกร์': 7}

            for day, periods in booked_rooms[room_sheet_title].items():
                day_row_index = day_to_row.get(day)
                if not day_row_index:
                    continue
                
                for period in periods:
                    column_index = int(period) + 1  # คอลัมน์เริ่มต้นที่ 2
                
                    cell_range = gspread.utils.rowcol_to_a1(day_row_index, column_index)
                    batch_data.append({
                        'range': cell_range,
                        'values': [['ถูกจองแล้ว']]
                    })
                    
            if batch_data:
                # เพิ่มข้อมูลลงในชีทใหม่
                room_sheet.batch_update(batch_data)
                time.sleep(2)  # หน่วงเวลาระหว่างการเรียกใช้ API

        print("Data has been successfully written to the 'Room' file.")
    
    except Exception as e:
        print(f"An error occurred: {e}")

def write_timetable_to_teacher_sheet():
    try:
        # Open the Generate file and fetch data
        generate_file = client.open('Generate')
        generate_sheet = generate_file.worksheet('Best Timetable')
        timetable_data = generate_sheet.get_all_records()

        # Open the Teacher file
        teacher_file = client.open('Teacher')
        teacher_sheets = teacher_file.worksheets()

        # Create a dictionary to store teacher schedules
        teacher_schedules = {}
        for schedule in timetable_data:
            teacher = schedule['อาจารย์']
            day = schedule['วัน']
            class_period = schedule['คาบเรียน']
            course_info = f"{schedule['รหัสวิชา']}\n{schedule['ห้องเรียน']}"

            if teacher not in teacher_schedules:
                teacher_schedules[teacher] = {}
            if day not in teacher_schedules[teacher]:
                teacher_schedules[teacher][day] = set()
            teacher_schedules[teacher][day].add((class_period, course_info))

        # Update each teacher sheet
        for teacher_sheet in teacher_sheets:
            teacher_sheet_title = teacher_sheet.title
            if teacher_sheet_title not in teacher_schedules:
                continue

            # Read teacher availability
            availability = teacher_sheet.get_all_values()
            day_to_row = {'จันทร์': 4, 'อังคาร': 5, 'พุธ': 6, 'พฤหัสบดี': 7, 'ศุกร์': 8}

            # Prepare data for batch_update
            batch_data = []
            highlight_cells = []

            for day, periods in teacher_schedules[teacher_sheet_title].items():
                day_row_index = day_to_row.get(day)
                if not day_row_index:
                    continue
                
                for period, course_info in periods:
                    column_index = int(period) + 1  # Column starts at 2
                    cell_value = availability[day_row_index-1][column_index-1]
                    
                    if cell_value != '1':
                        # Highlight cell in red if teacher is not available
                        highlight_cells.append({
                            'range': gspread.utils.rowcol_to_a1(day_row_index, column_index),
                            'format': {'backgroundColor': {'red': 1.0, 'green': 0.0, 'blue': 0.0}}
                        })

                    cell_range = gspread.utils.rowcol_to_a1(day_row_index, column_index)
                    batch_data.append({
                        'range': cell_range,
                        'values': [[course_info]]
                    })
                    
            if batch_data:
                # Write data to the teacher sheet
                teacher_sheet.batch_update(batch_data)
                time.sleep(2)  # Rate limit handling

            if highlight_cells:
                # Apply conditional formatting for highlighting
                teacher_sheet.format(highlight_cells)
                time.sleep(2)  # Rate limit handling

        print("Data has been successfully written to the 'Teacher' file.")
    
    except Exception as e:
        print(f"An error occurred: {e}")
# เรียกใช้ฟังก์ชันเพื่อเขียนข้อมูลลงในไฟล์ Student
#write_timetable_to_student_sheet()
#write_timetable_to_room_sheet()
write_timetable_to_teacher_sheet()