"""
Impuestos por destinación (régimen) de la DNIT: qué se paga, sobre qué base y con qué norma,
más un control con lo que efectivamente se cobró en los datos de la DNIT.

Dos pasos (se pueden correr juntos o por separado):

  1. Control: suma, por destinación, la base imponible y cada impuesto de los archivos
     mensuales "Nivel Item" de la DNIT de un año. Guarda regimenes_tributos/control_<año>.csv
     (archivo chico, se sube al repositorio).
        python regimenes_tributos/tributos_por_regimen.py --control --anio 2025 --datos datos_dnit/2025
     Si la carpeta no tiene los CSV, bájelos antes con:
        python python/descargar_dnit.py --anios 2025 --destino datos_dnit

  2. Excel: une la tabla de normas (regimenes_tributos/normas_por_regimen.csv, editable a mano)
     con config/regimenes.csv y con el control, y arma
     regimenes_tributos/Tributos_por_regimen.xlsx (Calibri 10).
        python regimenes_tributos/tributos_por_regimen.py --excel --anio 2025

Los montos de la DNIT vienen en guaraníes: las tasas se calculan en guaraníes
(impuesto / imponible) y los montos en dólares se pasan con la cotización de cada despacho.
"""
import argparse
import sys
from pathlib import Path

import pandas as pd

AQUI = Path(__file__).resolve().parent
RAIZ = AQUI.parent
sys.path.insert(0, str(RAIZ / "python"))
import procesar_despachos as pdesp  # noqa: E402

IMPUESTOS = ["DERECHO", "IVA", "ISC", "SERVICIO", "RENTA", "OTROS"]
NOMBRE_IMP = {"DERECHO": "Arancel", "IVA": "IVA", "ISC": "ISC", "SERVICIO": "Servicio de valoración",
              "RENTA": "Anticipo de renta", "OTROS": "Otros"}


# ------------------------------------------------------------------------------
# 1. CONTROL CON LOS DATOS DE LA DNIT
# ------------------------------------------------------------------------------

def control(carpeta, anio):
    archivos = sorted(p for p in Path(carpeta).rglob("*") if p.suffix.lower() in (".csv", ".json", ".txt"))
    if not archivos:
        sys.exit(f"No hay archivos de la DNIT en {carpeta}. Bájelos con python/descargar_dnit.py")
    partes, despachos = [], []
    for ruta in archivos:
        print("leyendo", ruta.name, flush=True)
        df = pdesp.estandarizar_columnas(pdesp.leer_archivo(ruta))
        df = df.drop_duplicates(["DESPACHO_CIFRADO", "ITEM"])
        cot = df["COTIZACION"].where(df["COTIZACION"] > 0)
        df["IMPONIBLE_USD"] = df["IMPONIBLE_DOLAR"].where(df["IMPONIBLE_DOLAR"] > 0, df["IMPONIBLE_GS"] / cot).fillna(0)
        df["TOTAL_GS"] = df[IMPUESTOS].sum(axis=1)
        for c in IMPUESTOS + ["TOTAL_GS"]:
            df[c + "_USD"] = (df[c] / cot).fillna(0)
            df[c + "_ITEMS"] = (df[c] > 0).astype(int)
        df["ITEMS"] = 1
        sumar = (["ITEMS", "IMPONIBLE_GS", "IMPONIBLE_USD", "FOB_DOLAR"] + IMPUESTOS + ["TOTAL_GS"]
                 + [c + s for c in IMPUESTOS + ["TOTAL_GS"] for s in ("_USD", "_ITEMS")])
        partes.append(df.groupby(["COD_REGIMEN", "OPERACION"])[sumar].sum())
        despachos.append(df[["COD_REGIMEN", "OPERACION", "DESPACHO_CIFRADO"]].drop_duplicates())
    t = pd.concat(partes).groupby(level=[0, 1]).sum()
    d = pd.concat(despachos).drop_duplicates().groupby(["COD_REGIMEN", "OPERACION"]).size().rename("DESPACHOS")
    t = t.join(d).reset_index()
    t.insert(0, "AÑO", anio)
    salida = AQUI / f"control_{anio}.csv"
    t.round(2).to_csv(salida, index=False)
    print("Guardado", salida.relative_to(RAIZ), f"({len(t)} destinaciones con datos)")


# ------------------------------------------------------------------------------
# 2. EXCEL
# ------------------------------------------------------------------------------

def pct(a, b):
    return (a / b * 100).where(b > 0)


TRIBUTOS = ["DERECHO", "IVA", "ISC", "RENTA"]   # se comparan con la norma
TASAS = ["SERVICIO", "OTROS"]                    # se cobran también en algunos suspensivos


def veredicto(esperado, tasa_tributos):
    """Compara lo que dice la norma con los tributos cobrados (arancel+IVA+ISC+renta, % de la base)."""
    if pd.isna(tasa_tributos):
        return "Sin datos 2025"
    esperado = str(esperado).strip().lower()
    if esperado == "no paga":
        return "Coincide" if tasa_tributos < 0.5 else "Revisar: la norma dice que no paga y se cobró"
    if esperado == "paga":
        return "Coincide" if tasa_tributos >= 1 else "Revisar: la norma dice que paga y casi no se cobró"
    if esperado == "parcial":
        return "Coincide" if tasa_tributos < 5 else "Revisar: se cobró más de lo esperado"
    return "Sin norma para comparar"


def excel(anio):
    reg = pd.read_csv(RAIZ / "config" / "regimenes.csv", dtype=str, keep_default_na=False)
    normas = pd.read_csv(AQUI / "normas_por_regimen.csv", dtype=str, keep_default_na=False)
    faltan = pd.read_csv(AQUI / "normas_faltantes.csv", dtype=str, keep_default_na=False)
    ruta_ctl = AQUI / f"control_{anio}.csv"
    ctl = pd.read_csv(ruta_ctl) if ruta_ctl.exists() else pd.DataFrame(columns=["COD_REGIMEN"])
    if len(ctl):
        ctl = ctl.groupby("COD_REGIMEN").sum(numeric_only=True).reset_index()

    t = reg[["codigo", "descripcion", "tipo_dnit", "operacion_dnit"]].merge(
        normas.drop(columns=["descripcion"], errors="ignore"), on="codigo", how="left")
    t = t.merge(ctl, left_on="codigo", right_on="COD_REGIMEN", how="left")

    base = t["IMPONIBLE_GS"]
    t["Despachos 2025"] = t["DESPACHOS"]
    t["Base imponible 2025 (USD)"] = t["IMPONIBLE_USD"]
    t["Impuestos cobrados 2025 (USD)"] = t["TOTAL_GS_USD"]
    t["Impuesto promedio por despacho (USD)"] = t["TOTAL_GS_USD"] / t["DESPACHOS"]
    t["Total cobrado (% de la base)"] = pct(t["TOTAL_GS"], base)
    t["Tributos: arancel+IVA+ISC+renta (% de la base)"] = pct(t[TRIBUTOS].sum(axis=1, min_count=1), base)
    t["Tasas: servicio+otros (% de la base)"] = pct(t[TASAS].sum(axis=1, min_count=1), base)
    for c in IMPUESTOS:
        t[f"{NOMBRE_IMP[c]} (% de la base)"] = pct(t[c], base)
    t["Ítems que pagaron algo (%)"] = pct(t["TOTAL_GS_ITEMS"], t["ITEMS"])
    t["¿Coincide con la norma?"] = [veredicto(e, x) for e, x in
                                    zip(t["esperado"], t["Tributos: arancel+IVA+ISC+renta (% de la base)"])]

    # se quitan las columnas crudas del control (IVA, ISC...) para que no choquen con las de la norma
    hoja = t.drop(columns=[c for c in ctl.columns if c in t.columns]).rename(columns={
        "codigo": "Código", "descripcion": "Destinación", "tipo_dnit": "Tipo (DNIT)",
        "operacion_dnit": "Operación", "arancel": "Arancel", "iva": "IVA", "isc": "ISC",
        "servicio": "Servicio de valoración", "renta": "Anticipo de renta", "otros": "Otros",
        "base": "Base de cálculo", "norma": "Norma", "esperado": "Según la norma",
        "comentario": "Comentario"})
    columnas = ["Código", "Destinación", "Tipo (DNIT)", "Operación", "Arancel", "IVA", "ISC",
                "Servicio de valoración", "Anticipo de renta", "Otros", "Base de cálculo", "Norma",
                "Según la norma", "Despachos 2025", "Base imponible 2025 (USD)",
                "Impuestos cobrados 2025 (USD)", "Impuesto promedio por despacho (USD)",
                "Total cobrado (% de la base)", "Tributos: arancel+IVA+ISC+renta (% de la base)",
                "Tasas: servicio+otros (% de la base)"] + [f"{NOMBRE_IMP[c]} (% de la base)" for c in IMPUESTOS] + \
               ["Ítems que pagaron algo (%)", "¿Coincide con la norma?", "Comentario"]
    hoja = hoja[columnas]

    # destinaciones que aparecen en los datos y no están en la lista de 83
    extra = ctl[~ctl["COD_REGIMEN"].isin(reg["codigo"])] if len(ctl) else ctl

    notas = pd.DataFrame({"Nota": [
        "Una fila por destinación de config/regimenes.csv (lista de la DNIT).",
        "Columnas Arancel a Otros: qué dice la norma. 'Sí' = se paga; 'No' = no se paga; "
        "'Suspendido' = no se paga mientras dura el régimen y se paga si se nacionaliza; "
        "'Exento' = la norma lo libera.",
        f"Columnas 2025: suma de los archivos 'Nivel Item' de la DNIT del año {anio}, todas las "
        "destinaciones (también las suspensivas). Tasas = impuesto / base imponible, en guaraníes.",
        "'¿Coincide con la norma?' compara solo los tributos (arancel + IVA + ISC + anticipo de renta): "
        "'Coincide' si lo que no paga cobró menos de 0,5 % de la base, si lo que paga cobró 1 % o más, "
        "y si lo parcial cobró menos de 5 %. El servicio de valoración y 'Otros' se dejan aparte porque "
        "se cobran también en algunos regímenes suspensivos.",
        "La DNIT publica el impuesto por ítem; 'Otros' junta varios conceptos que no se pueden separar.",
        "La tasa promedio de arancel o IVA es menor que la tasa legal porque incluye bienes con 0 %, "
        "exoneraciones (Ley 7548/25 art. 4, ex Ley 60/90) y bienes con IVA del 5 %.",
        "Norma con asterisco (*): no está en la biblioteca subida; el artículo sale de los documentos "
        "anteriores y hay que confirmarlo (ver hoja 'Normas faltantes'). Sin asterisco: leída en la biblioteca.",
        "Se rehace con: python regimenes_tributos/tributos_por_regimen.py --excel --anio " + str(anio),
    ]})

    salida = AQUI / "Tributos_por_regimen.xlsx"
    with pd.ExcelWriter(salida, engine="xlsxwriter") as xw:
        wb = xw.book
        wb.formats[0].set_font_name("Calibri")
        wb.formats[0].set_font_size(10)
        f = {"font_name": "Calibri", "font_size": 10}
        txt = wb.add_format({**f, "text_wrap": True, "valign": "top"})
        num = wb.add_format({**f, "num_format": "#,##0", "valign": "top"})
        dec = wb.add_format({**f, "num_format": "0.00", "valign": "top"})
        cab = wb.add_format({**f, "bold": True, "text_wrap": True, "valign": "top", "bg_color": "#DDEBF7",
                             "border": 1})
        rojo = wb.add_format({**f, "font_color": "#9C0006", "bg_color": "#FFC7CE"})
        verde = wb.add_format({**f, "font_color": "#006100", "bg_color": "#C6EFCE"})

        def escribir(df, nombre, anchos):
            df.to_excel(xw, sheet_name=nombre, index=False)
            ws = xw.sheets[nombre]
            for j, c in enumerate(df.columns):
                ws.write(0, j, c, cab)
                es_num = pd.api.types.is_numeric_dtype(df[c])
                fmt = (dec if "%" in c else num) if es_num else txt
                ws.set_column(j, j, anchos.get(c, 14), fmt)
            ws.freeze_panes(1, 2 if nombre == "Por destinación" else 0)
            ws.autofilter(0, 0, len(df), len(df.columns) - 1)
            return ws

        anchos = {"Destinación": 34, "Base de cálculo": 40, "Norma": 48, "Comentario": 48,
                  "¿Coincide con la norma?": 30, "Código": 7, "Operación": 11}
        ws = escribir(hoja, "Por destinación", anchos)
        j = columnas.index("¿Coincide con la norma?")
        ws.conditional_format(1, j, len(hoja), j, {"type": "text", "criteria": "begins with",
                                                   "value": "Revisar", "format": rojo})
        ws.conditional_format(1, j, len(hoja), j, {"type": "text", "criteria": "begins with",
                                                   "value": "Coincide", "format": verde})
        if len(ctl):
            det = ctl.rename(columns={"COD_REGIMEN": "Código"})
            escribir(det, f"Control {anio} (detalle)", {})
        if len(extra):
            escribir(extra.rename(columns={"COD_REGIMEN": "Código"}), "Códigos fuera de la lista", {})
        escribir(faltan, "Normas faltantes", {"Norma": 40, "Para qué": 60, "Destinaciones": 30,
                                              "Dónde buscar": 40})
        escribir(notas, "Notas", {"Nota": 120})
    print("Guardado", salida.relative_to(RAIZ))
    return hoja


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--anio", type=int, default=2025)
    ap.add_argument("--datos", help="carpeta con los CSV mensuales de la DNIT del año")
    ap.add_argument("--control", action="store_true", help="recalcular control_<año>.csv")
    ap.add_argument("--excel", action="store_true", help="armar el Excel")
    a = ap.parse_args()
    if not (a.control or a.excel):
        a.control = a.excel = True
    if a.control:
        control(a.datos or RAIZ / "datos_dnit" / str(a.anio), a.anio)
    if a.excel:
        excel(a.anio)


if __name__ == "__main__":
    main()
