import pandas as pd

# Create sample data with Korean text
data = {
    '이름': ['김철수', '이영희', '박민준'],
    '직업': ['개발자', '디자이너', '관리자'],
    '부서': ['기술팀', '디자인팀', '인사팀'],
    '메모': ['열심히 일합니다', '창의적입니다', '책임감이 강합니다']
}

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel with some formatting
with pd.ExcelWriter('test_korean.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='직원정보', index=False)
    
    # Get the workbook and the worksheet
    workbook = writer.book
    worksheet = writer.sheets['직원정보']
    
    # Add some basic formatting
    for column in worksheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

print("Test Excel file created successfully!")
