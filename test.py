import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Google Sheets API authentication
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('data.json', scope)
client = gspread.authorize(credentials)

# เปิดไฟล์ Google Sheets
generateFile = client.open('Generate')

# ฟังก์ชันตรวจสอบการทับซ้อนของตารางเรียน
def check_overlap(df):
    df['คาบ (เริ่ม)'] = pd.to_numeric(df['คาบ (เริ่ม)'])
    df['คาบ (จบ)'] = pd.to_numeric(df['คาบ (จบ)'])
    overlaps = []

    for i, row1 in df.iterrows():
        for j, row2 in df.iterrows():
            if i >= j:
                continue
            if (row1['วันเรียน'] == row2['วันเรียน'] and
                row1['ห้องเรียน'] == row2['ห้องเรียน'] and
                row1['คาบ (เริ่ม)'] < row2['คาบ (จบ)'] and
                row1['คาบ (จบ)'] > row2['คาบ (เริ่ม)']):
                overlaps.append((row1['เซคเรียน'], row2['เซคเรียน'], row1['วันเรียน'], row1['ห้องเรียน']))

    return overlaps

# ตรวจสอบการทับซ้อนในแต่ละชีท
for sheet in generateFile.worksheets():
    sheet_name = sheet.title
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    
    overlaps = check_overlap(df)
    if overlaps:
        print(f'พบการทับซ้อนในชีท {sheet_name}:')
        for overlap in overlaps:
            print(f'เซคเรียน {overlap} และ {overlap} ทับซ้อนกันในวัน {overlap} ที่ห้อง {overlap}')
    else:
        print(f'ไม่พบการทับซ้อนในชีท {sheet_name}')

'''
import pandas as pd

# อ่านข้อมูลจากไฟล์ Excel
file_path = 'Generate.xlsx'
sheets = pd.read_excel(file_path, sheet_name=None)

# ฟังก์ชันตรวจสอบการทับซ้อนของตารางเรียน
def check_overlap(df):
    df['คาบ (เริ่ม)'] = pd.to_numeric(df['คาบ (เริ่ม)'])
    df['คาบ (จบ)'] = pd.to_numeric(df['คาบ (จบ)'])
    overlaps = []

    for i, row1 in df.iterrows():
        for j, row2 in df.iterrows():
            if i >= j:
                continue
            if (row1['วันเรียน'] == row2['วันเรียน'] and
                row1['ห้องเรียน'] == row2['ห้องเรียน'] and
                row1['คาบ (เริ่ม)'] < row2['คาบ (จบ)'] and
                row1['คาบ (จบ)'] > row2['คาบ (เริ่ม)']):
                overlaps.append((row1['เซคเรียน'], row2['เซคเรียน'], row1['วันเรียน'], row1['ห้องเรียน']))

    return overlaps

# ตรวจสอบการทับซ้อนในแต่ละชีท
for sheet_name, df in sheets.items():
    overlaps = check_overlap(df)
    if overlaps:
        print(f'พบการทับซ้อนในชีท {sheet_name}:')
        for overlap in overlaps:
            print(f'เซคเรียน {overlap} และ {overlap} ทับซ้อนกันในวัน {overlap} ที่ห้อง {overlap}')
    else:
        print(f'ไม่พบการทับซ้อนในชีท {sheet_name}')



'''