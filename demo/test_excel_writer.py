from kdocpyxl.write_data import WpsCloudExcelWriter

"""
云文档自动化测试文件：https://cvte.kdocs.cn/l/cgyglnLWT6hv
"""
WEBHOOK_URL = "https://cvte.kdocs.cn/api/v3/ide/file/cgyglnLWT6hv/script/V2-4PSSwY3aqnGOlxKwp0u2BB/sync_task"
API_TOKEN = "328aBVqsuEePnyCGqkFIRX"

# 创建 WpsCloudExcelWriter 实例
writer = WpsCloudExcelWriter(WEBHOOK_URL, API_TOKEN)

# 测试 write_cell 方法
ret1 = writer.write_cell("sheet2", 1, 1, "测试数据")
print(ret1)

# 测试 write_row 方法
values = ["测试数据1", "测试数据2", "测试数据3"]
ret2 = writer.write_row("sheet2", 2, values)
print(ret2)

# 测试 write_column 方法
values = ["测试数据4", "测试数据5", "测试数据6"]
ret3 = writer.write_column("sheet2", 3, values)
print(ret3)