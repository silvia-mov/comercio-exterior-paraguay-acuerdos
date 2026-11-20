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


REG = pd.read_csv(RAIZ / "config" / "regimenes.csv", dtype=str, keep_default_na=False).drop_duplicates("codigo")
CUENTA = REG.set_index("codigo")["incluir"]
# Arancel Externo Común de la versión de la NCM vigente cada año (config/aec_por_anio.csv.gz, armado con
# las 63 versiones de SAREM del repositorio nomenclaturas-y-correlaciones-del-mercosur) y listas de Paraguay
# notificadas al MERCOSUR (config/listas_py.csv)
_A = pd.read_csv(RAIZ / "config" / "aec_por_anio.csv.gz", dtype={"NCM": str})
AEC_ANIO = {a: g.set_index("NCM")["AEC"] for a, g in _A.groupby("ANIO")}
LISTAS = pd.read_csv(RAIZ / "config" / "listas_py.csv", dtype=str)
NOMBRE_LISTA = {"LNE": "Lista Nacional de Excepciones", "BK": "Bienes de Capital",
                "BIT": "Informática y Telecomunicaciones", "RMP": "Materias Primas"}


def listas_del_anio(a):
    """Para cada lista, la última notificación hasta fin del año (o la primera disponible)."""
    out = {}
    for lista, g in LISTAS.groupby("LISTA"):
        fechas = sorted(g["FECHA_NOTIFICACION"].unique())
        previas = [f for f in fechas if f <= f"{a}-12-31"]
        f = previas[-1] if previas else fechas[0]
        out[lista] = (set(g.loc[g["FECHA_NOTIFICACION"] == f, "NCM"]), f)
    return out


def motivo_arancel(d, a):
    ls = listas_del_anio(a)
    cod = d["POSICION"].astype(str).str.replace(".", "", regex=False).str[:8]
    aec = cod.map(AEC_ANIO[min(max(a, min(AEC_ANIO)), max(AEC_ANIO))])
    m = pd.Series("Otra ley o régimen nacional (sin identificar)", index=d.index)
    for lista in ("RMP", "BIT", "BK", "LNE"):        # el último que se aplica tiene prioridad
        m[cod.isin(ls[lista][0])] = NOMBRE_LISTA[lista]
    m[aec.isna()] = "Sin dato de AEC"
    m[aec == 0] = "AEC 0 %"
    m[d["DERECHO"] > 0] = "Pagó arancel"
    con_acuerdo = ~d["ACUERDO"].isin(["SIN ACUERDO", "OMC"])
    m[con_acuerdo] = "Con acuerdo"
    # Si el AEC ya es 0 %, invocar el acuerdo no da ninguna ventaja: entraría igual con arancel 0
    m[con_acuerdo & (aec == 0)] = "Con acuerdo, pero AEC 0 % (el acuerdo no da ventaja)"
    return m, {k: v[1] for k, v in ls.items()}


def alerta_peso(d):
    """Compara el valor por kilo de cada fila con la mediana de su NCM en el año.
    Un error de tipeo en los kilos (5 en vez de 50) hace que el valor por kilo salga mucho más alto o más bajo;
    se marca solo desde 50 veces, porque dentro de un mismo NCM hay productos de valores muy distintos."""
    cod = d["POSICION"].astype(str).str.replace(".", "", regex=False).str[:8]
    uv = d["FOB DOLAR"] / d["KILO NETO"].where(d["KILO NETO"] > 0)
    grupo = pd.DataFrame({"cod": cod, "uv": uv})
    med = grupo.groupby("cod")["uv"].transform("median")
    n = grupo.groupby("cod")["uv"].transform("count")
    r = uv / med
    a = pd.Series("", index=d.index)
    a[(d["KILO NETO"] <= 0) & (d["FOB DOLAR"] > 0)] = "Sin kilos"
    a[(n >= 10) & (r >= 50)] = "Valor por kilo 50 veces mayor que lo normal: posibles kilos de menos"
    a[(n >= 10) & (r <= 0.02)] = "Valor por kilo 50 veces menor que lo normal: posibles kilos de más"
    return a


RTC = set(pd.read_csv(RAIZ / "config" / "lista_rtc.csv", dtype=str)["NCM"])


def regimen_tributario(d):
    """Clasifica por el IVA pagado sobre la base imponible (los dos en guaraníes)."""
    cod = d["POSICION"].astype(str).str.replace(".", "", regex=False).str[:8]
    tasa = (d["IVA"] / d["IMPONIBLE GS"].where(d["IMPONIBLE GS"] > 0)) * 100
    r = pd.Series("Sin base imponible", index=d.index)
    r[tasa == 0] = "Sin IVA (exento o suspendido)"
    r[(tasa > 0) & (tasa <= 2.2)] = "IVA reducido (otro)"
    r[(tasa > 0) & (tasa <= 2.2) & cod.isin(RTC)] = "Turismo de Compras (IVA ≈ 1,25 %)"
    r[(tasa > 2.2) & (tasa <= 7.5)] = "IVA 5 %"
    r[tasa > 7.5] = "IVA 10 %"
    return r


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
            # Para no contar dos veces: solo se suman los regímenes definitivos (config/regimenes.csv).
            # Ej.: la mercadería que entra a una zona franca (ZF01) se cuenta al salir al país (ZF2I).
            datos.insert(list(datos.columns).index("DESTINACION") + 1, "SE_CUENTA",
                         datos["DESTINACION"].map(CUENTA).fillna("NO"))
            datos["ALERTA_PESO"] = alerta_peso(datos)
            nota_listas = ""
            if op == "IMPORTACION":
                motivo, fechas_listas = motivo_arancel(datos, a)
                datos.insert(list(datos.columns).index("ACUERDO") + 1, "MOTIVO_ARANCEL", motivo)
                datos.insert(list(datos.columns).index("MOTIVO_ARANCEL") + 1, "REGIMEN_TRIBUTARIO", regimen_tributario(datos))
                nota_listas = ("MOTIVO_ARANCEL: por qué pagó o no arancel. 'Con acuerdo' = usó una preferencia; 'Con acuerdo, pero AEC 0 %' = declaró un acuerdo pero el producto ya tiene arancel 0 para cualquier país, así que el acuerdo no le dio ventaja; 'Pagó arancel' = DERECHO > 0; "
                               "'AEC 0 %' = el Arancel Externo Común ya es 0; luego, si el producto está en una lista de Paraguay notificada "
                               "al MERCOSUR (Bienes de Capital, Informática y Telecomunicaciones, Excepciones, Materias Primas); si no, "
                               "'Otra ley o régimen nacional'. AEC de la versión de la NCM vigente ese año (última versión publicada hasta fin de año que contiene el código). Listas usadas (fecha de notificación): "
                               + ", ".join(f"{NOMBRE_LISTA[k]} {v}" for k, v in fechas_listas.items()) + ". "
                               "REGIMEN_TRIBUTARIO: según el IVA pagado sobre la base imponible. 'Turismo de Compras' = producto de la lista "
                               "del Decreto 2063/2024 (Régimen de Turismo de Compras, antes Decreto 1931/2019) con IVA cercano al 1,25 % "
                               "(10 % sobre el 12,5 % de la base); no es un régimen aduanero, entra como importación común. "
                               "Para años anteriores a 2024 se usa la misma lista de productos.")
            reg = reg.merge(REG[["codigo", "incluir", "tipo_dnit", "nota"]].rename(
                columns={"codigo": "DESTINACION", "incluir": "SE_CUENTA", "tipo_dnit": "TIPO", "nota": "MOTIVO"}),
                on="DESTINACION", how="left").fillna({"SE_CUENTA": "NO", "MOTIVO": "Régimen no informado en la tabla oficial"})
            nombre = f"DNIT_{a}_{op.lower()}.xlsx"
            notas = [f"DNIT, datos abiertos 'Nivel Item', {op.lower()}es {a} ({meses} meses)." + (" Año incompleto." if meses < 12 else ""),
                     "Fuente: Dirección Nacional de Ingresos Tributarios (DNIT), https://datosabiertos.aduana.gov.py. Elaboración propia.",
                     f"{int(datos['ITEMS'].sum()):,} ítems de despacho agrupados en {len(datos):,} filas.".replace(",", "."),
                     "Se sumaron los ítems que coinciden en todos los campos, salvo los que cambian en cada despacho: número de despacho, "
                     "oficialización, cancelación, mes, ítem, cotización, canal, mercadería y marca.",
                     "No se filtró ningún régimen ni se agrupó ningún acuerdo: están todas las operaciones (incluidos depósitos, tránsitos, "
                     "zonas francas, admisiones temporales).",
                     "IMPORTANTE, columna SE_CUENTA: para sumar el comercio sin contar dos veces, filtrar SE_CUENTA = SI (solo regímenes "
                     "definitivos). La mercadería que pasa por una zona franca figura al entrar (ZF01, SE_CUENTA = NO) y al salir al país "
                     "(ZF2I, SE_CUENTA = SI). Ej.: celulares 8517.13 en 2023: USD 3.100 millones sumando todo, 1.722 con SE_CUENTA = SI "
                     "(cifra oficial de Paraguay en UN Comtrade: 1.729 CIF). La hoja 'Regímenes' explica cada código.",
                     "Hoja 'Descripciones': textos de cada POSICION (los del primer ítem). Hoja 'Regímenes': nombre de cada código DESTINACION.",
                     "Se quitaron TOTAL (= DERECHO + ISC + SERVICIO + RENTA + IVA + OTROS), los ajustes a incluir/deducir y el kilo bruto.",
                     "ITEMS = ítems de despacho sumados; DESPACHOS = despachos distintos.",
                     "ALERTA_PESO: compara el valor por kilo (FOB / KILO NETO) con la mediana del mismo NCM en el año (si hay al menos 10 filas). "
                     "Si es 50 veces mayor o menor, probablemente hay un error de carga en los kilos (por ejemplo 5 en vez de 50). Es una alerta, "
                     "no una corrección: los datos no se cambiaron.",
                     nota_listas,
                     "Valores en dólares, salvo IMPONIBLE GS y los impuestos (DERECHO a OTROS), en guaraníes."]
            with pd.ExcelWriter(salida / nombre, engine="xlsxwriter") as w:
                b = w.book
                b.formats[0].set_font_name("Calibri"); b.formats[0].set_font_size(10)
                for hoja, df in (("Datos", datos), ("Descripciones", desc), ("Regímenes", reg)):
                    df.to_excel(w, sheet_name=hoja, index=False)
                    w.sheets[hoja].freeze_panes(1, 0)
                    w.sheets[hoja].autofilter(0, 0, len(df), len(df.columns) - 1)
                pd.DataFrame({"Notas": [n for n in notas if n]}).to_excel(w, sheet_name="Notas", index=False)
            lista.append({"archivo": str(salida / nombre), "hoja_datos": "Datos", "filas": ["ACUERDO"],
                          "columnas": [], "filtros": ["SE_CUENTA", "PAIS PROCEDENCIA/DESTINO", "DESTINACION"],
                          "valores": [["FOB DOLAR", "suma", "FOB (USD)"]]})
            print(f"{nombre}: {len(datos):,} filas", flush=True)
        shutil.rmtree(carpeta, ignore_errors=True)
    # Sin tabla dinámica: cada usuario la arma si la necesita.



if __name__ == "__main__":
    main()
