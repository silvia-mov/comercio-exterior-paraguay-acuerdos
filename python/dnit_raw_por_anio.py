"""
DNIT "Nivel Item" de un año, en Excel, lo más parecido al dato original pero agrupado por año.

Sin filtrar regímenes ni agrupar acuerdos. Se suman los ítems que coinciden en todos los campos de la DNIT
salvo los que cambian en cada despacho (número de despacho, oficialización, cancelación, mes, ítem,
cotización, canal, mercadería y marca). Las descripciones son las del primer ítem de cada POSICION.

Uso: python python/dnit_raw_por_anio.py 2023 2024     (baja del respaldo en GitHub si hace falta)
Salida: salidas/raw/DNIT_raw_<año>.xlsx
"""
import shutil
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


QUITAR = ["TOTAL", "AJUSTE A INCLUIR", "AJUSTE A DEDUCIR", "KILO BRUTO"]


def main():
    """Un Excel por año y operación, para la página: descripciones y regímenes en hojas aparte."""
    salida = RAIZ / "salidas" / "dnit_anual"
    salida.mkdir(parents=True, exist_ok=True)
    lista = []
    for a in [int(x) for x in sys.argv[1:]]:
        carpeta = Path("/tmp/claude-0/dnit_raw") / str(a)
        out, items, meses = anio(a, carpeta)
        assert out["ITEMS"].sum() == items
        for op in ("EXPORTACION", "IMPORTACION"):
            d = out[out["OPERACION"] == op]
            desc = d[["POSICION"] + DESCRIP].drop_duplicates("POSICION").sort_values("POSICION")
            reg = d[["DESTINACION", "REGIMEN"]].drop_duplicates("DESTINACION").sort_values("DESTINACION")
            datos = d.drop(columns=DESCRIP + ["REGIMEN"] + QUITAR)
            nombre = f"DNIT_{a}_{op.lower()}.xlsx"
            notas = [f"DNIT, datos abiertos 'Nivel Item', {op.lower()}es {a} ({meses} meses)." + (" Año incompleto." if meses < 12 else ""),
                     "Fuente: Dirección Nacional de Ingresos Tributarios (DNIT), https://datosabiertos.aduana.gov.py. Elaboración propia.",
                     f"{int(datos['ITEMS'].sum()):,} ítems de despacho agrupados en {len(datos):,} filas.".replace(",", "."),
                     "Se sumaron los ítems que coinciden en todos los campos, salvo los que cambian en cada despacho: número de despacho, "
                     "oficialización, cancelación, mes, ítem, cotización, canal, mercadería y marca.",
                     "No se filtró ningún régimen ni se agrupó ningún acuerdo: están todas las operaciones (incluidos depósitos, tránsitos, "
                     "zonas francas, admisiones temporales). Ojo: la mercadería que pasa por zona franca figura al entrar y al salir.",
                     "Hoja 'Descripciones': textos de cada POSICION (los del primer ítem). Hoja 'Regímenes': nombre de cada código DESTINACION.",
                     "Se quitaron TOTAL (= DERECHO + ISC + SERVICIO + RENTA + IVA + OTROS), los ajustes a incluir/deducir y el kilo bruto.",
                     "ITEMS = ítems de despacho sumados; DESPACHOS = despachos distintos.",
                     "Valores en dólares, salvo IMPONIBLE GS y los impuestos (DERECHO a OTROS), en guaraníes."]
            with pd.ExcelWriter(salida / nombre, engine="xlsxwriter") as w:
                b = w.book
                b.formats[0].set_font_name("Calibri"); b.formats[0].set_font_size(10)
                for hoja, df in (("Datos", datos), ("Descripciones", desc), ("Regímenes", reg)):
                    df.to_excel(w, sheet_name=hoja, index=False)
                    w.sheets[hoja].freeze_panes(1, 0)
                    w.sheets[hoja].autofilter(0, 0, len(df), len(df.columns) - 1)
                pd.DataFrame({"Notas": notas}).to_excel(w, sheet_name="Notas", index=False)
            lista.append({"archivo": str(salida / nombre), "hoja_datos": "Datos", "filas": ["ACUERDO"],
                          "columnas": [], "filtros": ["PAIS PROCEDENCIA/DESTINO", "DESTINACION"],
                          "valores": [["FOB DOLAR", "suma", "FOB (USD)"]]})
            print(f"{nombre}: {len(datos):,} filas", flush=True)
        shutil.rmtree(carpeta, ignore_errors=True)
    import json
    cfg = salida / "_tablas.json"
    cfg.write_text(json.dumps(lista, ensure_ascii=False), encoding="utf-8")
    subprocess.run(["/usr/bin/python3", str(RAIZ / "python" / "tabla_dinamica.py"), str(cfg)], check=True)
    cfg.unlink()
    sys.path.insert(0, str(RAIZ / "python"))
    from excel_descargas import etiquetas_en_espanol
    for t in lista:
        etiquetas_en_espanol(t["archivo"])


if __name__ == "__main__":
    main()
