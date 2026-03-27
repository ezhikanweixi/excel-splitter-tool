# Excel分类导出工具

Windows 桌面应用程序，用于将 Excel 文件按指定列分类导出为多个文件。

## 功能
- 选择 Excel 文件（.xlsx/.xls）
- 选择分类列
- 按类别自动导出为多个 Excel 文件

## 使用方法

1. 从 [Releases](https://github.com/ezhikanweixi/excel-splitter-tool/releases) 下载最新版本
2. 解压并运行 `Excel分类导出工具.exe`

## 开发

### 本地运行
```bash
pip install -r requirements.txt
python excel_splitter.py
```

### 构建 Windows 版本
推送到 main 分支后，GitHub Actions 会自动构建 Windows 可执行文件。
