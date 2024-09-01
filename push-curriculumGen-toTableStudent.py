import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

# ตั้งค่าและเชื่อมต่อกับ Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("path/to/your/credentials.json", scope)
client = gspread.authorize(creds)

# เปิดไฟล์ Google Sheets
curriculum_sheet = client.open("Curriculum_General Education Program").sheet1
open_course_sheet = client.open("Open Course2").sheet1
student_sheet = client.open("Student").sheet1

# ดึงข้อมูลจากไฟล์ Google Sheets
curriculum_data = curriculum_sheet.get_all_records()
open_course_data = open_course_sheet.get_all_records()

# แปลงข้อมูลเป็น DataFrame ของ Pandas
curriculum_df = pd.DataFrame(curriculum_data)
open_course_df = pd.DataFrame(open_course_data)

# ประมวลผลข้อมูล (ตัวอย่างการรวมข้อมูลตามรหัสวิชา)
merged_df = pd.merge(curriculum_df, open_course_df, on="รหัสวิชา")

# สร้างตารางสอน (ตัวอย่างการสร้าง DataFrame สำหรับตารางสอน)
timetable_df = pd.DataFrame({
    "รหัสวิชา": merged_df["รหัสวิชา"],
    "ชื่อวิชา": merged_df["ชื่อวิชา"],
    "วันเรียน": merged_df["วันเรียนบรรยาย"],
    "คาบเรียน": merged_df["คาบบรรยาย(เริ่ม)"].astype(str) + "-" + merged_df["คาบบรรยาย(จบ)"].astype(str)
})

# อัปเดตข้อมูลในไฟล์ Student
def update_student_sheet(start_row, start_col, timetable_df):
    for index, row in timetable_df.iterrows():
        student_sheet.update_cell(start_row + index, start_col, row["รหัสวิชา"])  # อัปเดตรหัสวิชา
        student_sheet.update_cell(start_row + index, start_col + 1, row["ชื่อวิชา"])  # อัปเดตชื่อวิชา
        student_sheet.update_cell(start_row + index, start_col + 2, row["วันเรียน"])  # อัปเดตวันเรียน
        student_sheet.update_cell(start_row + index, start_col + 3, row["คาบเรียน"])  # อัปเดตคาบเรียน

# อัปเดตข้อมูลในช่วงเซลล์ที่ระบุ
update_student_sheet(4, 2, timetable_df.iloc[:7])  # อัปเดตเซคแรก (B4:M10)
update_student_sheet(15, 2, timetable_df.iloc[7:14])  # อัปเดตเซคที่สอง (B15:M21)

print("อัปเดตตารางสอนเรียบร้อยแล้ว!")
