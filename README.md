# kdocpyxl

kdocpyxl 是一个用于读写金山文档云 Excel 的 Python 库。

## 功能

- 从金山云表格中读取或写入数据

## 安装

使用 pip 安装 kdocpyxl：

```bash
pip install kdocpyxl
```

## 使用方法
[AirScript 脚本令牌](https://airsheet.wps.cn/docs/apitoken/intro.html)

### 1. 创建AirScript脚本
- 登录金山云表格。
- 点击界面顶部的【效率】标签。
- 在下拉菜单中选择【高级开发】。
- 点击【AirScript脚本编辑器】开始创建一个新的AirScript脚本。

### 2. 编写AirScript代码
- 在脚本编辑器中，将本项目源码中的 `airscript.js` 代码复制并粘贴到编辑器中。
- 确保代码无误后，点击保存。

### 3. 获取Webhook URL和API Token
- 在脚本编辑器中，找到并复制生成的 `webhook_url`，这是外部应用触发脚本的URL。
- 为脚本创建一个 `api_token`，这是一个安全令牌，用于验证外部请求的合法性。

### 4. 配置脚本参数
- 在你的代码中，设置 `webhook_url` 和 `api_token` 这两个参数，以便脚本能够正确地与金山云表格交互。

## 贡献
欢迎贡献代码，你可以通过提交 Pull Request 或者通过 Issue 讨论新功能或者报告 bug。

## 许可证
本项目采用 MIT 许可证。