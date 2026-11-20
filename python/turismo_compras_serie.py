"""Parte de las importaciones que entró con el IVA reducido de Turismo de Compras, por año, y el ejemplo de los celulares.

Uso (desde la raíz del repositorio):
    python python/turismo_compras_serie.py

Lee salidas/dnit_anual/DNIT_<año>_importacion.xlsx (columna REGIMEN_TRIBUTARIO, que se deduce del IVA pagado
sobre la base imponible y de la lista de productos del Decreto 2063/2024; no es un campo de la DNIT).
Escribe docs/datos/turismo_compras_por_anio.csv y docs/datos/celulares_doble_conteo.csv.
IVA reducido (otro): IVA de hasta 2,2 % de la base en productos que no están en la lista con los códigos de hoy
(antes de 2022 la NCM usaba otros códigos, así que parte del turismo de compras cae ahí).
Solo despachos definitivos (SE_CUENTA = SI). Valor: IMPONIBLE DOLAR (CIF).
"""
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
ANUAL = RAIZ / "salidas" / "dnit_anual"
DATOS = RAIZ / "docs" / "datos"
RTC = "Turismo de Compras"
CEL = "85171"   # 8517.12 (hasta 2022) y 8517.13 (desde 2022): teléfonos celulares


def main():
    # Meses de cada archivo anual (los mismos que usa importaciones_ventaja.py)
    meses = pd.read_csv(DATOS / "importaciones_acuerdo_ventaja.csv", encoding="utf-8-sig").set_index("AÑO")["MESES"]
    filas, cel = [], []
    for f in sorted(ANUAL.glob("DNIT_*_importacion.xlsx")):
        a = int(f.stem.split("_")[1])
        d = pd.read_excel(f, sheet_name="Datos", engine="calamine",
                          usecols=["SE_CUENTA", "POSICION", "IMPONIBLE DOLAR", "REGIMEN_TRIBUTARIO"])
        pos = d["POSICION"].astype(str).str.replace(".", "", regex=False)
        es_cel = pos.str.startswith("851712") | pos.str.startswith("851713")
        ok = d["SE_CUENTA"] == "SI"
        rtc = d["REGIMEN_TRIBUTARIO"].astype(str).str.startswith(RTC)
        tot = d.loc[ok, "IMPONIBLE DOLAR"].sum()
        v_rtc = d.loc[ok & rtc, "IMPONIBLE DOLAR"].sum()
        v_otro = d.loc[ok & (d["REGIMEN_TRIBUTARIO"] == "IVA reducido (otro)"), "IMPONIBLE DOLAR"].sum()
        c_tot = d.loc[ok & es_cel, "IMPONIBLE DOLAR"].sum()
        c_rtc = d.loc[ok & es_cel & rtc, "IMPONIBLE DOLAR"].sum()
        filas.append({"AÑO": a, "MESES": int(meses.get(a, 12)), "IMPORTACIONES_USD": tot, "TURISMO_COMPRAS_USD": v_rtc,
                      "PCT_TURISMO_COMPRAS": 100 * v_rtc / tot,
                      "IVA_REDUCIDO_OTRO_USD": v_otro, "PCT_IVA_REDUCIDO_OTRO": 100 * v_otro / tot, "CELULARES_USD": c_tot,
                      "PCT_CELULARES_TURISMO_COMPRAS": 100 * c_rtc / c_tot if c_tot else None})
        cel.append({"AÑO": a, "TODOS_LOS_REGIMENES_USD": d.loc[es_cel, "IMPONIBLE DOLAR"].sum(), "SE_CUENTA_SI_USD": c_tot})
        print(filas[-1])
    pd.DataFrame(filas).to_csv(DATOS / "turismo_compras_por_anio.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(cel).to_csv(DATOS / "celulares_doble_conteo.csv", index=False, encoding="utf-8-sig")


if __name__ == "__main__":
    main()
