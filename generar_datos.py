import pandas as pd
import random
from datetime import datetime, timedelta

def generar_datos_v2(ruta='c:\\DSB\\datos_test.xlsx'):
    print(f"Generando datos para el nuevo dashboard en: {ruta}")
    
    sistemas = ['Sistema de Gestión', 'Portal Web PGN', 'Base de Datos Central', 'Correo Institucional', 'Servidor de Aplicaciones']
    fechas = [datetime(2023, 1, 1) + timedelta(days=i) for i in range(100)]
    
    datos = []
    for _ in range(500):
        fecha = random.choice(fechas)
        hora = f"{random.randint(0,23):02d}:{random.randint(0,59):02d}"
        sistema = random.choice(sistemas)
        alarma = 'Sí' if random.random() < 0.2 else 'No'
        
        datos.append({
            'Fecha': fecha.strftime('%Y-%m-%d'),
            'Hora': hora,
            'Sistema': sistema,
            'Alarma': alarma
        })
        
    df = pd.DataFrame(datos)
    df.to_excel(ruta, index=False)
    print("¡Archivo generado!")

if __name__ == '__main__':
    generar_datos_v2()