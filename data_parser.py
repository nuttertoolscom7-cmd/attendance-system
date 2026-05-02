import pandas as pd
import openpyxl
import os
import io

def load_data(file_source):
    try:
        # Handle file-like objects (like Streamlit UploadedFile) or string paths
        if hasattr(file_source, 'read'):
            file_content = file_source.read()
            source_for_pd = io.BytesIO(file_content)
            source_for_ox = io.BytesIO(file_content)
            file_source.seek(0) # Reset just in case
        else:
            if not os.path.exists(file_source):
                return None
            with open(file_source, 'rb') as f:
                file_content = f.read()
            source_for_pd = io.BytesIO(file_content)
            source_for_ox = io.BytesIO(file_content)

        excel_file = pd.ExcelFile(source_for_pd)
        wb = openpyxl.load_workbook(source_for_ox, data_only=True)
    except Exception as e:
        print(f"Error loading file: {e}")
        return None
    
    sheets = excel_file.sheet_names
    
    relevant_sheets = ['ต.ค 67-ก.ย 68', 'ต.ค 68-ก.ย 69 (2)']
    all_details = []
    
    for sheet_name in relevant_sheets:
        if sheet_name not in sheets:
            continue
            
        df = pd.read_excel(source_for_pd, sheet_name=sheet_name)
        source_for_pd.seek(0) # Reset for next iteration if needed
        ws = wb[sheet_name]
        
        # Identify columns
        name_col = df.columns[0]
        month_col = 'เดือน'
        day_cols = df.columns[2:33] # Columns 1 to 31
        remark_col = 'หมายเหตุ'
        
        budget_year = '2568' if '67-ก.ย 68' in sheet_name else '2569'
        
        # Mapping for symbols
        type_map = {'ป': 'ลาป่วย', 'ก': 'ลากิจ', 'พ': 'ลาพักผ่อน', 'ส': 'สาย', 'ป.': 'ลาป่วย'}

        for idx, row in df.iterrows():
            name = row[name_col]
            month = row[month_col]
            
            if pd.isna(month) or month in ['Aเดือน', 'รวม', 'เดือน']:
                continue
            if pd.isna(name) or 'Aวันที่' in str(name):
                continue
                
            remarks = str(row[remark_col]) if pd.notna(row[remark_col]) else ""
            
            # Check each day
            for day_idx, day_val in enumerate(day_cols):
                val = row[day_val]
                if pd.notna(val) and str(val).strip() in type_map:
                    leave_type = type_map[str(val).strip()]
                    
                    # Determine row and col in Excel
                    # Pandas index 0 corresponds to Excel Row 2 (Header is Row 1)
                    excel_row = idx + 2
                    # Day columns start at index 2 in Pandas, which is Column 3 in Excel (C)
                    excel_col = day_idx + 3
                    
                    cell = ws.cell(row=excel_row, column=excel_col)
                    is_half = False
                    
                    # Check background color for yellow highlight (half-day)
                    fill = cell.fill
                    if fill and fill.start_color and getattr(fill.start_color, 'rgb', None):
                        color = str(fill.start_color.rgb)
                        # FFFFFF00 is standard yellow, check for this or similar variations
                        if color == 'FFFFFF00' or 'FFFF00' in color:
                            is_half = True
                    
                    all_details.append({
                        'Name': name,
                        'Month': month,
                        'Day': day_idx + 1,
                        'Type': leave_type,
                        'BudgetYear': budget_year,
                        'Remarks': remarks,
                        'IsHalf': is_half
                    })
                    
    return pd.DataFrame(all_details)

if __name__ == '__main__':
    data = load_data('/Users/mr.o./Downloads/1.ตารางวันลา (ตั้งแต่ งบ 68 เป็นต้นไป).xlsx')
    if data is not None:
        print(data[data['IsHalf'] == True][['Name', 'Month', 'Day', 'Type', 'IsHalf']])
