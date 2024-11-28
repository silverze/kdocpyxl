from kdocpyxl.read_data import WpsCloudExcelReader

"""
云文档自动化测试文件：https://cvte.kdocs.cn/l/cgyglnLWT6hv
"""

WEBHOOK_URL = "https://cvte.kdocs.cn/api/v3/ide/file/cgyglnLWT6hv/script/V2-4PSSwY3aqnGOlxKwp0u2BB/sync_task"
API_TOKEN = "2VbsT0H5cjj1H4sDVWx4Uz"

reader = WpsCloudExcelReader(WEBHOOK_URL, API_TOKEN)

# 测试 read_cell 方法
cell_value = reader.read_cell("sheet1", 7, 5)
print(cell_value)

# 测试 read_row 方法
row_values = reader.read_row("sheet1", 4)
print(row_values)

# 测试 read_column 方法
column_values = reader.read_column("sheet1", 5)
print(column_values)

# 测试 get_sheet_rows_and_columns 方法
rows_and_columns = reader.get_sheet_rows_and_columns("sheet1")
print(rows_and_columns)