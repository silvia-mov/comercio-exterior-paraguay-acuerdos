"""
DNIT "Nivel Item" de un año, en Excel, lo más parecido al dato original pero agrupado por año.

Sin filtrar regímenes ni agrupar acuerdos. Se suman los ítems que coinciden en todos los campos de la DNIT
salvo los que cambian en cada despacho (número de despacho, oficialización, cancelación, mes, ítem,
cotización, canal, mercadería y marca). Las descripciones son las del primer ítem de cada POSICION.

Uso: python python/dnit_raw_por_anio.py 2023 2024     (baja del respaldo en GitHub si hace falta)
Salida: salidas/raw/DNIT_raw_<año>.xlsx
"""
import subprocess
import sys
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
CLAVES = ["AÑO", "OPERACION", "DESTINACION", "REGIMEN", "ADUANA", "MEDIO TRANSPORTE", "PAIS ORIGEN",
          "PAIS PROCEDENCIA/DESTINO", "USO", "UNIDAD MEDIDA ESTADISTICA", "POSICION", "ACUERDO"]
DESCRIP = ["RUBRO", "DESC CAPITULO", "DESC PARTIDA", "DESC POSICION"]
NUM = ["CANTIDAD ESTADISTICA", "KILO NETO", "KILO BRUTO", "FOB DOLAR", "FLETE DOLAR", "SEGURO DOLAR",
       "IMPONIBLE DOLAR", "IMPONIBLE GS", "AJUSTE A INCLUIR", "AJUSTE A DEDUCIR",
       "DERECHO", "ISC", "SERVICIO", "RENTA", "IVA", "OTROS", "TOTAL"]
MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE",
         "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]


def anio(a, carpeta):
    if not carpeta.exists() or not list(carpeta.rglob("*.csv")):
        subprocess.run([sys.executable, str(RAIZ / "python" / "descargar_dnit.py"), "--anios", str(a),
                        "--fuente", "github", "--destino", str(carpeta)], check=True)
    rutas = sorted(carpeta.rglob("*_Nivel_Item.csv"), key=lambda p: MESES.index(p.stem.split("_")[1]))
    partes, items, desc, meses = [], 0, [], 0
    for r in rutas:
        d = pd.read_csv(r, dtype=str, keep_default_na=False, encoding="utf-8")
        d.columns = [c.strip() for c in d.columns]
        for c in NUM:
            d[c] = pd.to_numeric(d[c].str.replace(",", ".", regex=False), errors="coerce").fillna(0)
        items += len(d)
        meses += 1
        desc.append(d.drop_duplicates("POSICION")[["POSICION"] + DESCRIP])
        g = d.groupby(CLAVES, sort=False)
        x = g[NUM].sum()
        x["ITEMS"] = g.size()
        x["_DESP"] = g["DESPACHO CIFRADO"].agg(lambda s: set(s))
        partes.append(x)
    t = pd.concat(partes).groupby(level=list(range(len(CLAVES))), sort=False)
    out = t[NUM + ["ITEMS"]].sum()
    out["DESPACHOS"] = t["_DESP"].agg(lambda s: len(set().union(*s)))
    out = out.reset_index()
    descr = pd.concat(desc).drop_duplicates("POSICION")
    out = out.merge(descr, on="POSICION", how="left")
    cols = CLAVES[:11] + DESCRIP + ["ACUERDO", "ITEMS", "DESPACHOS"] + NUM
    return out[cols], items, meses


def main():
    for a in [int(x) for x in sys.argv[1:]]:
        carpeta = Path("/tmp/claude-0/dnit_raw") / str(a)
        out, items, meses = anio(a, carpeta)
        assert out["ITEMS"].sum() == items
        destino = RAIZ / "salidas" / "raw" / f"DNIT_raw_{a}.xlsx"
        destino.parent.mkdir(parents=True, exist_ok=True)
        notas = ["Notas", f"DNIT, datos abiertos 'Nivel Item', año {a} ({meses} meses).",
                 "Fuente: https://datosabiertos.aduana.gov.py",
                 f"{items:,} ítems de despacho originales. agrupados en {len(out):,} filas.".replace(",", "."),
                 "Se sumó sobre todos los campos de la DNIT salvo los que cambian en cada despacho: número de despacho, "
                 "oficialización, cancelación, mes, ítem, cotización, canal, mercadería y marca.",
                 "No se filtró ningún régimen ni se agrupó ningún acuerdo: están todas las operaciones (incluidos depósitos, "
                 "tránsitos, admisiones temporales, etc.).",
                 "ITEMS = cantidad de ítems de despacho sumados; DESPACHOS = despachos distintos.",
                 "Las descripciones (RUBRO, DESC ...) son las del primer ítem de cada POSICION.",
                 "Valores en dólares, salvo IMPONIBLE GS y los impuestos (DERECHO a TOTAL), en guaraníes."]
        with pd.ExcelWriter(destino, engine="xlsxwriter") as w:
            b = w.book
            b.formats[0].set_font_name("Calibri"); b.formats[0].set_font_size(10)
            out.to_excel(w, sheet_name="Datos", index=False)
            pd.DataFrame({"Notas": notas[1:]}).to_excel(w, sheet_name="Notas", index=False)
            w.sheets["Datos"].freeze_panes(1, 0)
            w.sheets["Datos"].autofilter(0, 0, len(out), len(out.columns) - 1)
        print(f"{destino.name}: {len(out):,} filas de {items:,} ítems; {meses} meses", flush=True)


if __name__ == "__main__":
    main()
