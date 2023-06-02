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

# 获取用户输入的Excel文件路径
excel_filepath = input("请输入Excel文件路径：")
# 加载Excel文件
workbook = openpyxl.load_workbook(excel_filepath)
# 获取所有工作表的名称
sheet_names = workbook.sheetnames

print("备选的工作表：")
for i, sheet_name in enumerate(sheet_names):
    print(f"{i+1}. {sheet_name}")

# 获取用户选择的工作表索引
selected_sheet_index = int(input("请输入要处理的工作表的索引：")) - 1
# 根据索引获取工作表名称
selected_sheet_name = sheet_names[selected_sheet_index]

# 获取用户选择的行数
selected_row_index = int(input("请输入要转换的行数："))

# 指定输出文件路径
output_filepath = 'path/to/output/file.txt'

extract_row_elements(excel_filepath, selected_sheet_name, selected_row_index, output_filepath)
