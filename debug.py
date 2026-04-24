import dashboard as db
df = db.load_data('datos_pgn.xlsx')
try:
    print('Columns:', df.columns)
    print('Sistemas:', df['Sistema'].unique())
    print('Alarmas counts:', df['Alarma'].value_counts())
    print('Fechas max:', df['Fecha'].max())
except Exception as e:
    print('Error:', df)
