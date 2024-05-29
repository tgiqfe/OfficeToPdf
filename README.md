# OfficeToPdf

## 環境セットアップ

```dos
cd ConvertScript
python -m venv venv
venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install pywin32
```

## 使用方法

ExcelファイルをPDFに
```dos
cd ConvertScript
venv\Scripts\python convert_excel.py <Excelファイルのパス>
```

WordファイルをPDFに
```dos
cd ConvertScript
venv\Scripts\python convert_word.py <Wordファイルのパス>
```

PowerPointファイルをPDFに
```dos
cd ConvertScript
venv\Scripts\python convert_powerpoint.py <PowerPointファイルのパス>
```
