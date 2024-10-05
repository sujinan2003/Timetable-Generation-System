import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials

# Google Sheets API authentication
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)
client = gspread.authorize(credentials)


generateFile = client.open('Generate')
teacherFile = client.open('Teacher')


def main():
    # ดึงรายชื่อของทุกชีทในไฟล์ Generate
    generate_sheets = generateFile.worksheets()
    
    # ดึงข้อมูลจากทุกชีทใน Generate
    generate_data = []
    for sheet in generate_sheets:
        generate_data.extend(sheet.get_all_values()[1:])  # ข้ามแถวแรก (header)
    
    # วนลูปผ่านทุกชีทในไฟล์ Teacher
    for teacher_sheet in teacherFile.worksheets():
        batch_data_teacher = []
        teacher = teacher_sheet.title
        
        # เตรียมข้อมูลสำหรับอัปเดต
        for course in generate_data:
            if len(course) > 7:  # ตรวจสอบว่ามีข้อมูลครบถ้วน
                section = course[0]
                course_code = course[1]
                id_teacher = course[2]
                id_room = course[3]
                course_type = course[4]
                day = course[5]
                start_period = int(course[6])
                end_period = int(course[7])

                # ถ้า id_teacher ตรงกับชื่อชีทใน Teacher
                if id_teacher == teacher:
                    days = ['จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์', 'อาทิตย์']
                    if day in days:
                        col = days.index(day) + 2  # เริ่มที่คอลัมน์ C

                        for period in range(start_period, end_period + 1):
                            row = period + 1  # เนื่องจากข้อมูลเริ่มต้นที่แถว 3 (คาบ 1 จะตรงกับแถวที่ 3)
                            batch_data_teacher.append({
                                'range': f'{chr(65 + col)}{row + 1}',
                                'values': [[f'{course_code}\n{course_type}\n{id_room}\n{id_teacher}\n{section}']]
                            })

        # อัปเดตข้อมูลในชีท Teacher
        if batch_data_teacher:
            teacher_sheet.batch_update([{
                'range': data['range'],
                'values': data['values']
            } for data in batch_data_teacher])
            print(f"Update Sheet: {teacher_sheet.title} ")
        else:
            print(f"No data in sheet: {teacher_sheet.title}")
        
        time.sleep(5)

    print("Success!")

if __name__ == "__main__":
    main()