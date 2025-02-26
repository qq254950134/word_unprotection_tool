# word_unprotection_tool
一个简单的工具，可以去除word文档的保护。 "A simple tool to remove protection from Word documents."

# Word文档保护解除工具

## 简介
Word文档保护解除工具是一个简单易用的应用程序，旨在帮助用户解除Word文档的保护。该工具支持批量处理多个Word文件，并提供友好的用户界面，方便用户选择文件和设置输出选项。

## 特性
- 支持选择单个或多个Word文件
- 支持选择包含Word文件的文件夹
- 自定义输出目录和文件名格式
- 处理DOCX和DOC文件
- 显示处理日志和进度条
- 友好的图形用户界面
- 
## 成品（仅打包windows）
[打包了成品](https://github.com/qq254950134/word_unprotection_tool/releases/tag/v1.0)

[界面预览](https://i.miji.bid/2025/02/26/86eceab4a2f5943916c18f4a25ff4d95.png)

## 安装
1. 确保您的计算机上已安装Python 3.x。
2. 下载或克隆本项目。
3. 安装所需的依赖库：
   ```bash
   pip install tk
   pip install pywin32  # 仅在处理DOC文件时需要
   ```

## 使用方法
1. 运行程序：
   ```bash
   python word_unprotection_tool.py
   ```
2. 在界面中，点击“选择Word文件”或“选择文件夹”按钮，选择要处理的Word文件。
3. 选择输出设置，可以选择与源文件相同目录或自定义输出目录。
4. 设置输出文件名格式（使用“原文件名”表示原始文件名）。
5. 点击“开始转换”按钮，程序将开始处理文件。
6. 处理完成后，您可以在日志区域查看处理结果。

## 注意事项
- 该工具在处理DOC文件时需要安装`pywin32`库。
- 请确保您有权限访问和修改所选的Word文件。

## 贡献
欢迎任何形式的贡献！如果您发现了bug或有改进建议，请提交issue或pull request。

## 许可证
本项目采用MIT许可证，详情请参阅LICENSE文件。

