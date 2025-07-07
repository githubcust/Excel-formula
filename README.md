# Excel-GS

## 说明
自己使用的一个小工具。

## 项目简介

Excel-GS 是一个基于 VBA 的 Excel 插件，集成了 OpenAI（或兼容 API）接口，能够根据用户自然语言描述自动生成 Excel 公式，并支持历史记录管理。适用于需要提升公式编写效率的办公用户。

---

## 主要功能

- **自然语言生成公式**：输入描述，自动生成对应的 Excel 公式并写入选中单元格。
- **历史记录管理**：自动保存输入历史，支持历史公式预览、查阅和删除。
- **自定义输入窗体**：友好的窗体界面，支持历史选择与新输入。
- **配置灵活**：API 地址、Key、模型等参数均可通过 `config.ini` 配置，无需修改代码。
- **防乱码处理**：采用 ADODB.Stream 方式，确保中文响应内容不乱码。

---

## 文件结构

- `openaiapi.bas`：主模块，包含主逻辑、历史管理、API 调用等全部核心代码。
- `UserForm1.frm/.frx`：自定义输入窗体及其资源文件。
- `config.ini`：放置于用户目录（如 `C:\Users\用户名\config.ini`），用于存储 API 配置。

---

## 安装与使用

1. **导入模块和窗体**
   - 在 Excel VBA 编辑器中，导入 `openaiapi.bas` 和 `UserForm1.frm`。
   - 引入 [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) 的 `JsonConverter.bas`，并引用 `Microsoft Scripting Runtime`。

2. **配置 API 信息**
   - 在 `C:\Users\你的用户名\config.ini` 新建配置文件，内容示例：
     ```
     [openai]
     url=https://api.openai.com/v1/chat/completions
     apikey=sk-xxxxxx
     model=gpt-3.5-turbo
     ```
   - 支持自定义 API 地址和模型。

3. **使用方法**
   - 在 Excel 中选中目标单元格，运行 `GenerateExcelFormula` 宏。
   - 在弹出的窗体中输入需求描述或选择历史记录，点击确定即可自动生成公式。

---

## 注意事项

- 需保证网络畅通，API Key 有效。
- 历史记录文件默认保存在 `C:\Users\用户名\default.ini`，最多保存 50 条。
- 若遇到“用户定义类型未定义”，请在 VBA 工程中引用 `Microsoft Scripting Runtime`。
- 仅支持返回以 `=` 开头的标准 Excel 公式，AI 返回内容若不合法会有错误提示。

---

## 常见问题

- **中文乱码**：已采用 ADODB.Stream 以 UTF-8 解码，确保响应内容正常显示。
- **API 报错**：请检查 `config.ini` 配置、API Key 是否正确，或查看网络连接。
- **历史记录异常**：请确保 `default.ini` 文件可读写，或尝试删除后自动重建。

---

## 致谢

- [OpenAI](https://openai.com/)
- [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)

---

## License

本项目仅供学习与交流，禁止用于商业用途。API Key 请妥善保管，避免
