import pandas as pd

def load_data(file):
    try:
        diccionario_hojas = pd.read_excel(file, sheet_name=None)
        hojas_validas = []
        for nombre_hoja, df_hoja in diccionario_hojas.items():
            df_hoja.columns = df_hoja.columns.astype(str).str.strip().str.lower()
            renombres = {}
            for col in df_hoja.columns:
                if 'fecha' in col: renombres[col] = 'Fecha'
                elif 'hora' in col: renombres[col] = 'Hora'
                elif 'aplicativo' in col: renombres[col] = 'Sistema'
                elif 'inconveniente' in col: renombres[col] = 'Inconveniente'
                elif 'comentario' in col: renombres[col] = 'Comentario'
            df_hoja = df_hoja.rename(columns=renombres)
            if 'Fecha' in df_hoja.columns and 'Sistema' in df_hoja.columns:
                hojas_validas.append(df_hoja)
        if not hojas_validas:
            return 'No se detectaron columnas de Fecha o Aplicativo en ninguna hoja.'
        df = pd.concat(hojas_validas, ignore_index=True)
        df['Sistema'] = df['Sistema'].fillna('N/A').astype(str).str.strip()
        if 'Inconveniente' in df.columns:
            df['Inconveniente'] = df['Inconveniente'].fillna('OK').astype(str).str.strip()
            df['Alarma'] = df['Inconveniente'].apply(lambda x: 'No' if str(x).upper() in ['OK', 'N/A', 'NA', 'NINGUNO', 'NAN', ''] else 'Sí')
        else:
            df['Alarma'] = df['Sistema'].apply(lambda x: 'No' if str(x).upper() in ['N/A', 'NA', 'NO', 'NAN', ''] else 'Sí')
        
        # Guardemos la lista REAL de sistemas antes de reemplazarlos por SIN ALARMA
        real_sistemas = df[df['Sistema'].str.upper().isin(['N/A', 'NA', 'NO', 'NAN', 'OK', '']) == False]['Sistema'].unique()
        
        df['Sistema'] = df['Sistema'].apply(lambda x: 'SIN ALARMA' if str(x).upper() in ['N/A', 'NA', 'NO', 'NAN', 'OK', ''] else x)
        meses = {'enero': 'jan', 'febrero': 'feb', 'marzo': 'mar', 'abril': 'apr', 'mayo': 'may', 'junio': 'jun', 'julio': 'jul', 'agosto': 'aug', 'septiembre': 'sep', 'setiembre': 'sep', 'octubre': 'oct', 'noviembre': 'nov', 'diciembre': 'dec'}
        fechas_texto = df['Fecha'].dt.strftime('%Y-%m-%d') if pd.api.types.is_datetime64_any_dtype(df['Fecha']) else df['Fecha'].astype(str).str.lower()
        fechas_texto = fechas_texto.str.replace(' de ', ' ', regex=False).str.replace('-', ' ', regex=False).str.replace('/', ' ', regex=False)
        for es, en in meses.items():
            fechas_texto = fechas_texto.str.replace(r'\b'+es+r'\b', en, regex=True)
        df['Fecha_dt'] = pd.to_datetime(fechas_texto, errors='coerce')
        # print unparsable dates
        unparsable = df[df['Fecha_dt'].isna()]['Fecha'].unique()
        print('Unparsable dates:', unparsable)
        df = df.dropna(subset=['Fecha_dt'])
        df['Fecha'] = df['Fecha_dt'].dt.date
        return df, unparsable, real_sistemas
    except Exception as e:
        return str(e), [], []

df, unparsable, real_sistemas = load_data('datos_pgn.xlsx')
print('Sistemas (antes de reemplazo):', real_sistemas)
print('Sistemas en DF final:', df['Sistema'].unique() if not isinstance(df, str) else 'N/A')
print('Alarma counts:', df['Alarma'].value_counts() if not isinstance(df, str) else 'N/A')
