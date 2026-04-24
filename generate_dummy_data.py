import pandas as pd
import numpy as np
from pathlib import Path

output_path = Path("datos.xlsx")

np.random.seed(42)

sistemas = [
    "ALFA", "APP MOVIL", "APPS EXTERNAS", "APPS INTRANET", "APPS PORTAL",
    "GESTOR DOKUS", "HOMINIS", "IGA", "INSAP", "ITA", "NUEVA SEDE ELECTRONICA",
    "PORTAL WEB E INTRANET", "SIAF", "SIGDEA PORTAL EMPLEADO", "SIGDEA SEDE ELECTRONICA",
    "SIM", "SIRI", "STRATEGOS", "X-ROAD"
]

fechas = pd.date_range(start="2025-01-01", end="2025-03-31", freq="D")
filas = []

for fecha in fechas:
    for sistema in np.random.choice(sistemas, size=4, replace=False):
        hora = f"{np.random.randint(6, 23):02d}:{np.random.choice(['00', '15', '30', '45'])}"
        alarma = np.random.choice(["Sí", "No"], p=[0.25, 0.75])
        filas.append({
            "Fecha": fecha.strftime("%Y-%m-%d"),
            "Hora": hora,
            "Sistema": sistema,
            "Alarma": alarma,
        })

if not filas:
    raise SystemExit("No se generaron filas de datos.")

pd.DataFrame(filas).to_excel(output_path, index=False)
print(f"Archivo de prueba generado: {output_path.resolve()}")
