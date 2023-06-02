import openpyxl

def extract_row_elements(filepath, sheetname, row_index, output_filepath):
    # 加载Excel文件
    workbook = openpyxl.load_workbook(filepath)
    # 选择工作表
    sheet = workbook[sheetname]
    # 获取指定行的所有单元格
    row = sheet[row_index]
    # 创建文本文件
    with open(output_filepath, 'w') as output_file:
        # 逐个写入每个单元格的值，每个值占据一行
        for cell in row:
            output_file.write(str(cell.value) + '\n')

    print("提取完成！")

# 示例用法
excel_filepath = 'path/to/your/excel/file.xlsx'
output_filepath = 'path/to/output/file.txt'
sheet_name = 'Sheet1'
row_number = 2  # 第2行，索引从1开始

extract_row_elements(excel_filepath, sheet_name, row_number, output_filepath)
