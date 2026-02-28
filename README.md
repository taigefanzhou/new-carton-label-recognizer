# 箱唛识别工具 - Windows独立版

## 功能
- 选择图片或文件夹导入
- 自动识别白色标签区域内容
- 提取：酒店名、产品名、数量、箱号
- 自动生成Excel装箱清单
- 无需安装Python，双击exe运行

## 识别规则
白色标签格式：
```
[酒店名称]
[产品名称]：[数量]pcs
NO: [箱号]
```

## 打包说明
使用PyInstaller + EasyOCR，OCR模型自动下载

## 依赖
easyocr
pillow
pandas
openpyxl
pyinstaller
