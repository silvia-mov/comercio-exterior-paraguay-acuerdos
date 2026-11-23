"""Completa la lista de cobros y regímenes especiales con lo que se ve en los datos de la DNIT.

Uso (desde la raíz del repositorio):
    python regimenes_tributos/completar_cobros.py

Lee regimenes_tributos/cobros_original.xlsx (la lista armada a partir de las normas, con la tasa y el artículo de
cada cobro) y salidas/dnit_anual/DNIT_2025_importacion.xlsx. Escribe regimenes_tributos/Cobros_y_regimenes_especiales.xlsx:
- celdas en verde: tasa leída en la norma (como en el original);
- celdas en celeste: tasa observada en lo que cobró la DNIT en 2025, sin norma leída todavía;
- hoja Control_DNIT_2025: la tasa efectiva que se ve en los datos para cada cobro.
"""
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

RAIZ = Path(__file__).resolve().parent.parent
CARPETA = RAIZ / "regimenes_tributos"
CELESTE = PatternFill("solid", fgColor="DCEBFA")
VERDE = PatternFill("solid", fgColor="B9EEE1")


def control(anio=2025):
    d = pd.read_excel(RAIZ / "salidas" / "dnit_anual" / f"DNIT_{anio}_importacion.xlsx", sheet_name="Datos", engine="calamine",
                      usecols=["SE_CUENTA", "IMPONIBLE GS", "IMPONIBLE DOLAR", "IVA", "RENTA", "SERVICIO", "OTROS",
                               "REGIMEN_TRIBUTARIO"])
    d = d[(d["SE_CUENTA"] == "SI") & (d["IMPONIBLE GS"] > 0)]
    w = d["IMPONIBLE DOLAR"]
    rtc = d["REGIMEN_TRIBUTARIO"].astype(str).str.startswith("Turismo de Compras")

    def pct(col, m=None):
        m = (d[col] > 0) if m is None else (m & (d[col] > 0))
        r = d.loc[m, col] / d.loc[m, "IMPONIBLE GS"] * 100
        return round(r.median(), 2), round(100 * w[m].sum() / w.sum(), 1)

    filas = []
    for nombre, col, m, norma in [
        ("Servicio de valoración", "SERVICIO", None, "Norma de creación no encontrada todavía"),
        ("Otros (en la importación)", "OTROS", None, "Puede incluir el canon informático y el arancel consular; sin desglose en los datos"),
        ("IVA, tasa general y reducida", "IVA", ~rtc & (d["IVA"] / d["IMPONIBLE GS"] > 0.03), "Ley 6380/2019, art. 90 (10 % y 5 %)"),
        ("IVA de Turismo de Compras", "IVA", rtc, "Decreto 2063/2024, art. 7 (10 % sobre el 12,5 % de la base = 1,25 %)"),
        ("Anticipo del IRE, fuera de Turismo de Compras", "RENTA", ~rtc, "Decreto 3182/2019, anexo, art. 87 (0,4 %)"),
        ("Anticipo del IRE, Turismo de Compras", "RENTA", rtc, "Decreto 2063/2024, art. 7 (1 %)"),
    ]:
        med, part = pct(col, m)
        filas.append([nombre, col, med, part, norma])
    return filas


def main():
    wb = openpyxl.load_workbook(CARPETA / "cobros_original.xlsx")
    ws = wb["Lista"]
    fila = {ws.cell(r, 3).value: r for r in range(2, ws.max_row + 1)}

    def poner(nombre, tasa=None, base=None, articulo=None, nota=None, relleno=CELESTE):
        r = fila[nombre]
        for col, v in zip((12, 13, 14, 15), (tasa, base, articulo, nota)):
            if v is not None:
                ws.cell(r, col).value = v
        if tasa is not None:
            ws.cell(r, 12).fill = relleno

    poner("Tasa por servicio de valoración aduanera", "0,5 %", "Valor en aduana (base gravable de las mercaderías importadas)",
          "Ley 2422/2004 (Código Aduanero), arts. 263 y 393 a)",
          "El servicio lo creó el Decreto-Ley 5/1974, aprobado por la Ley 489/1974. En 2025 el 99 % del valor importado "
          "pagó 0,5 % (hoja Control_DNIT_2025).", relleno=VERDE)
    ws.cell(fila["Tasa por servicio de valoración aduanera"], 7).value = "Ley 2422/2004, Código Aduanero"
    ws.cell(fila["Tasa por servicio de valoración aduanera"], 9).value = "Vigente (se cobró en 2025: hoja Control_DNIT_2025)"
    poner("Importación de vehículos eléctricos (2012)", "0 % de arancel y 0 % de IVA", "Vehículos eléctricos e híbridos importados",
          "Ley 4601/2012, art. 1", "Modificada por la Ley 5183/2014 (carpeta normas). Texto de la BACN.", relleno=VERDE)
    poner("Donaciones al Estado y a otras instituciones", "Exentas de todo tributo", "Bienes donados",
          "Ley 302/1993, art. 2", "No se pueden vender ni transferir por 5 años (art. 5). Texto de la BACN (escaneado, leído con OCR).",
          relleno=VERDE)
    poner("Canon informático",
          "Monto fijo por despacho, en jornales mínimos diarios: remesa expresa 0 % a 50 % de un jornal; fracciones complementarias "
          "25 %; exportación simplificada 25 %",
          "Cada despacho oficializado (no depende del valor)",
          "Res. DNA 198/2017, art. 2 (remesa expresa); Res. DNA 218/2020, art. 7 (IFC1); Res. DGA 25/1994, art. 1",
          "Remesa expresa: documentos 0 %, mercaderías de menos de USD 100 25 %, de USD 100 a 1.000 50 % (vigente desde el 01/04/2017). "
          "Para el despacho general, un resumen sin verificar habla de 2,5 jornales; el monto original está en la Res. MH 1144/1992, "
          "que falta. Lo percibe el Centro de Despachantes de Aduana (CDAP), no va al presupuesto de Aduanas. La exportación "
          "general no paga (Res. DGA 25/1994).", relleno=VERDE)
    # Fila nueva: el cobro de la carne exportada (ya documentado en este repositorio).
    r = ws.max_row + 1
    ws.cell(r, 1).value = r - 1
    for col, v in {2: "2. Cobros y reglas en la exportación", 3: "Cobro sobre la carne bovina exportada",
                   4: "Porcentaje sobre la exportación de carne (NCM 0201, 0202 y 0206); figura en OTROS",
                   5: "Exportación", 6: "Directa", 7: "Ley 808/1996, art. 14 c)", 8: "Resolución CIEFA 3/2021 (texto en la carpeta normas)",
                   9: "Vigente", 10: "Sí (Ley 808/1996)", 11: "Texto de la BACN; datos de la DNIT (OTROS en EC01 y EC03)", 12: "0,25 % por cada tonelada de carne",
                   13: "Valor de la carne importada y exportada", 14: "Ley 808/1996, art. 14 c) (tope); Res. CIEFA 3/2021, art. 1 (valor desde el 01/01/2022)",
                   15: "La Res. CIEFA 3/2021 (17/12/2021, Acta CIEFA 385) fija el valor en 0,25 % desde el 01/01/2022. En los datos de la DNIT se ve cerca de 0,25 %."}.items():
        ws.cell(r, col).value = v
        ws.cell(r, col).font = Font(name="Arial", size=10)
        ws.cell(r, col).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(r, 12).fill = VERDE
    # Las fórmulas del resumen cubrían hasta la fila 45.
    res = wb["Resumen"]
    for row in res.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = c.value.replace("$45", f"${r}")

    # Hoja de control.
    hc = wb.create_sheet("Control_DNIT_2025")
    hc.append(["Cobro", "Columna DNIT", "Tasa mediana (% de la base en guaraníes)", "% del valor importado que lo pagó", "Qué dice la norma"])
    for f in control():
        hc.append(f)
    hc.append([])
    hc.append(["Importaciones definitivas (SE_CUENTA = SI) de 2025. Tasa = cobro / IMPONIBLE GS, mediana de las filas que pagan. "
               "Programa: regimenes_tributos/completar_cobros.py."])
    for c in hc[1]:
        c.font = Font(name="Arial", size=10, bold=True)
    for w_, col in zip((42, 14, 22, 20, 70), "ABCDE"):
        hc.column_dimensions[col].width = w_

    # Leeme: colores, avisos y normas que faltan.
    le = wb["Leeme"]
    for t in ["", "Agregado en el repositorio (20/11/2026)",
              "Celdas en celeste: tasa observada en lo que cobró la DNIT en 2025, sin norma leída todavía.",
              "Ley 6380/2019, art. 153, num. 5: deroga todas las exoneraciones de impuestos salvo las de su lista (leído con OCR). Leídas en el texto el 20/11/2026: Leyes 808/1996, 302/1993, 4601/2012, 5542/2015 (BACN), 1064/1997, 110/1992, 523/1995, Decretos 11.771/2000 y 11.525/2007 Res. CIEFA 3/2021 y Leyes 4838/2012, 6925/2022, 4427/2012, 1295/1998, 3180/2007, 5501/2015, 921/1996, 7577/2025, 7467/2025, 5372/2014, 285/1993, 5819, 6178 y 6524/2020; tasa de valoración en el Código Aduanero.",
              "Hoja Control_DNIT_2025: la tasa que se ve en los datos para cada cobro, comparada con la norma.",
              "Las tasas del ISC de cada producto están en la hoja ISC_Decreto_3109 (texto de 2019 del decreto).",
              "El tributo único de maquila (1 %) se paga como impuesto interno: no figura en los despachos.",
              "El Código Aduanero (Ley 2422/2004) ya está en la carpeta normas del repositorio."]:
        le.append([t])
    for c in le["A"]:
        if c.value in ("Agregado en el repositorio (20/11/2026)",):
            c.font = Font(name="Arial", bold=True)
    wb.save(CARPETA / "Cobros_y_regimenes_especiales.xlsx")
    print("Listo:", CARPETA / "Cobros_y_regimenes_especiales.xlsx")


if __name__ == "__main__":
    main()
