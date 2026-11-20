"""Importaciones bajo el Régimen de Turismo según el Anexo Estadístico del BCP (Cuadro 52a), por año.

Uso (desde la raíz del repositorio):
    python python/leer_anexo_bcp.py "Anexo_Estadístico_del_Informe_Económico_13_11_2026_.xlsx"

El cuadro trae, por mes, tres columnas: importación registrada, importación bajo el Régimen de Turismo e
importación para consumo (miles de USD FOB). Escribe docs/datos/turismo_bcp_por_anio.csv.
"""
import datetime
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl
import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent


def main(ruta):
    wb = openpyxl.load_workbook(ruta, read_only=True)
    filas = list(wb["Cuadro 52a"].iter_rows(values_only=True))
    fechas = next(r for r in filas if sum(isinstance(c, datetime.datetime) for c in r) > 5)
    total = next(r for r in filas if r[0] and str(r[0]).strip() == "Total")
    acc = defaultdict(lambda: [0.0, 0.0, 0])
    for j, f in enumerate(fechas):
        if isinstance(f, datetime.datetime) and isinstance(total[j], (int, float)):
            acc[f.year][0] += total[j]
            acc[f.year][1] += total[j + 1] or 0
            acc[f.year][2] += 1
    d = pd.DataFrame([{"AÑO": a, "MESES": v[2], "IMPORTACION_REGISTRADA_USD": v[0] * 1e3, "REGIMEN_TURISMO_USD": v[1] * 1e3,
                       "PCT_REGIMEN_TURISMO": 100 * v[1] / v[0]} for a, v in sorted(acc.items())])
    d.to_csv(RAIZ / "docs" / "datos" / "turismo_bcp_por_anio.csv", index=False, encoding="utf-8-sig")
    print(d.tail(4).to_string(index=False))


if __name__ == "__main__":
    main(sys.argv[1])
