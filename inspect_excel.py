import pandas as pd
try:
    xl = pd.ExcelFile('c:\\DSB\\monitoreos 2025.xlsx')
    print("Hojas encontradas:", xl.sheet_names)
except Exception as e:
    print("Error:", e)
