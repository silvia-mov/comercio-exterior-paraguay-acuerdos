"""Gráficos interactivos de Datawrapper para docs/index.html y docs/aduana.html, hechos con los mismos datos que las imágenes.

Uso (desde la raíz del repositorio, con acceso a api.datawrapper.de y la clave de la API ya configurada en el entorno):
    python docs/datawrapper_graficos.py

Crea cada gráfico la primera vez y lo actualiza (datos, textos y estilo) las siguientes. Los números de los gráficos quedan
en docs/datawrapper_ids.json. Las imágenes PNG de siempre se siguen usando al imprimir.
"""
import json
import os
from pathlib import Path

import pandas as pd
import requests

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "docs" / "datos"
IDS = RAIZ / "docs" / "datawrapper_ids.json"
API = "https://api.datawrapper.de/v3"
H = {"Authorization": "Bearer " + os.environ.get("DATAWRAPPER_TOKEN", "proxy-injected")}
FUENTE = {"source-name": "DNIT, despachos aduaneros (cálculo propio)",
          "source-url": "https://silvia-mov.github.io/comercio-exterior-paraguay-acuerdos/", "byline": ""}
VERDE, VERDE2, MENTA, VIOLETA, LILA, GRIS = "#5e8f88", "#3f6b65", "#8db2ac", "#774d78", "#af8eb1", "#9a9097"
PUBLICAR = {"blocks": {"logo": {"enabled": False}, "embed": False, "download-pdf": False, "download-svg": False,
                       "get-the-data": True, "download-image": True}, "autoDarkMode": True}
MES = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre"]


def coma(x, d=0):
    return f"{x:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def grafico(clave, tipo, titulo, df, describe, visualize, notas="", datos=None):
    ids = json.loads(IDS.read_text()) if IDS.exists() else {}
    if clave not in ids:
        r = requests.post(f"{API}/charts", headers=H, json={"title": titulo, "type": tipo}, timeout=60)
        r.raise_for_status()
        ids[clave] = r.json()["id"]
        IDS.write_text(json.dumps(ids, indent=1, ensure_ascii=False))
    cid = ids[clave]
    requests.put(f"{API}/charts/{cid}/data", headers={**H, "Content-Type": "text/csv"},
                 data=df.to_csv(index=False).encode("utf-8"), timeout=60).raise_for_status()
    meta = {"describe": {**FUENTE, **describe}, "annotate": {"notes": notas},
            "visualize": {"dark-mode-invert": True, "show-tooltips": True, **visualize}, "publish": PUBLICAR}
    if datos:
        meta["data"] = datos
    requests.patch(f"{API}/charts/{cid}", headers=H, timeout=60,
                   json={"title": titulo, "type": tipo, "language": "es-AR", "metadata": meta}).raise_for_status()
    requests.post(f"{API}/charts/{cid}/publish", headers=H, timeout=120).raise_for_status()
    print(clave, f"https://datawrapper.dwcdn.net/{cid}/")
    return cid


def parcial(meses_por_anio):
    p = [(a, m) for a, m in meses_por_anio.items() if m < 12]
    return (f"* {p[0][0]}: enero a {MES[int(p[0][1])]}.", {a: f"{a}*" if m < 12 else str(a) for a, m in meses_por_anio.items()}) if p \
        else ("", {a: str(a) for a in meses_por_anio})


def destino_aduana_bcp():
    d = pd.read_csv(DATOS / "destino_aduana_bcp.csv", encoding="utf-8-sig")
    w = pd.DataFrame({"Año": sorted(d["AÑO"].unique())})
    for p, n in (("ARGENTINA", "Argentina"), ("URUGUAY", "Uruguay")):
        x = d[d["PAIS"] == p].set_index("AÑO")
        w[f"{n}: destino en la aduana"] = w["Año"].map(x["DNIT_EXPORTADO_USD"] / 1e6).round(1)
        w[f"{n}: destino final (BCP)"] = w["Año"].map(x["BCP_SIN_ELECTRICIDAD_USD"] / 1e6).round(1)
    ult = d[d["AÑO"] == d["AÑO"].max()].set_index("PAIS")["DIFERENCIA_USD"] / 1e6
    w["Año"] = w["Año"].astype(str)
    return grafico("destino_aduana_bcp", "d3-lines",
                   f"En {d['AÑO'].max()}, USD {coma(ult['ARGENTINA'])} millones declarados hacia Argentina y {coma(ult['URUGUAY'])} "
                   "hacia Uruguay tuvieron otro destino final", w,
                   {"intro": "Exportaciones de Paraguay sin electricidad, millones de USD FOB. Línea llena: destino declarado en el despacho "
                             "(DNIT). Línea punteada: destino según el certificado de origen (BCP). La distancia entre las dos es lo que pasa "
                             "por ese país pero termina en otro.", "number-format": "0,0", "number-append": ""},
                   {"custom-colors": {"Argentina: destino en la aduana": VERDE, "Argentina: destino final (BCP)": VERDE,
                                      "Uruguay: destino en la aduana": VIOLETA, "Uruguay: destino final (BCP)": VIOLETA},
                    "lines": {"Argentina: destino final (BCP)": {"dash": "style2"}, "Uruguay: destino final (BCP)": {"dash": "style2"}},
                    "labeling": "right", "label-colors": True, "label-margin": 190, "x-tick-format": "YYYY", "tooltip-x-format": "YYYY",
                    "y-grid-format": "0,0", "tooltip-number-format": "0,0", "interpolation": "linear"},
                   "Años completos.")


def importaciones_ventaja():
    d = pd.read_csv(DATOS / "importaciones_acuerdo_ventaja.csv", encoding="utf-8-sig")
    d = d[d["MESES"] == 12]
    ult = d.iloc[-1]
    w = pd.DataFrame({"Año": d["AÑO"].astype(str), "Usó un acuerdo": d["PCT_CON_ACUERDO"].round(1),
                      "Con ventaja arancelaria potencial": d["PCT_CON_VENTAJA_POTENCIAL"].round(1)})
    return grafico("importaciones_ventaja", "d3-lines",
                   f"En {int(ult['AÑO'])}, el {coma(ult['PCT_CON_ACUERDO'], 1)}% de lo importado usó un acuerdo; con ventaja "
                   f"arancelaria potencial, el {coma(ult['PCT_CON_VENTAJA_POTENCIAL'], 1)}%", w,
                   {"intro": "% de lo importado (CIF). Con ventaja potencial: sin los productos con Arancel Externo Común 0% ni los que, "
                             "importados sin acuerdo ese año, casi no pagaron arancel. La franja es lo que usó un acuerdo sin ventaja "
                             "(ya entraba con 0%).", "number-format": "0.[0]"},
                   {"custom-colors": {"Usó un acuerdo": MENTA, "Con ventaja arancelaria potencial": VERDE2},
                    "lines": {"Usó un acuerdo": {"width": "style2"}, "Con ventaja arancelaria potencial": {"width": "style2"}},
                    "labeling": "right", "label-colors": True, "label-margin": 150, "x-tick-format": "YYYY", "tooltip-x-format": "YYYY",
                    "custom-range-y": ["15", "35"], "y-grid-format": "0%", "tooltip-number-format": "0.[0]%", "interpolation": "linear"},
                   "Años completos. Despachos definitivos de la DNIT.")


def evolucion_acuerdos():
    corto = {"MERCOSUR intrazona (ACE 18)": "MERCOSUR (ACE 18)", "ALADI Bilateral Brasil (AAP.CE 74)": "Brasil (AAP.CE 74)",
             "ALADI Bilateral Argentina (AAP.CE 13)": "Argentina (AAP.CE 13)", "ALADI Bilateral México (AAP.R 38)": "México (AAP.R 38)",
             "ALADI Preferencia Arancelaria Regional (PAR 4)": "ALADI PAR 4"}
    a = pd.read_csv(DATOS / "comercio_por_acuerdo.csv", encoding="utf-8-sig")
    meses = pd.read_csv(DATOS / "utilizacion_por_anio.csv", encoding="utf-8-sig").set_index("AÑO")["MESES"]
    e = a[(a["OPERACION"] == "EXPORTACION") & (a["TIPO_ACUERDO"] != "SIN_ACUERDO") & a["AÑO"].map(lambda x: meses.get(x) == 12)]
    e = e.assign(N=e["ACUERDO_CONSOLIDADO"].replace(corto))
    piv = e.pivot_table(index="AÑO", columns="N", values="VALOR_USD", aggfunc="sum").fillna(0) / 1e6
    orden = list(piv.sum().sort_values(ascending=False).index[:9])
    w = piv[orden].round(1).reset_index().rename(columns={"AÑO": "Año"})
    w["Año"] = w["Año"].astype(str)
    ch = piv["Mercosur-Chile (ACE 35)"]
    return grafico("evolucion_acuerdos", "multiple-lines",
                   f"Fuera del MERCOSUR, el acuerdo con Chile es el que más sumó: de {coma(ch.iloc[0])} a {coma(ch.iloc[-1])} millones de USD",
                   w, {"intro": f"Exportaciones de Paraguay con cada acuerdo, millones de USD FOB, {w['Año'].iloc[0]}–{w['Año'].iloc[-1]} "
                                "(años completos). Los 9 acuerdos más usados. <b>Cada panel tiene su propia escala</b>: el MERCOSUR es "
                                "unas diez veces más grande que el que le sigue.", "number-format": "0,0.[0]"},
                   {"independentYScales": True, "gridLayout": "fixedCount", "gridColumnCount": 3, "gridColumnCountMobile": 2,
                    "base-color": VERDE, "x-tick-format": "YYYY", "tooltip-x-format": "YYYY", "tooltip-number-format": "0,0.[0]",
                    "interpolation": "linear"},
                   "Despachos definitivos de la DNIT. Un envío puede usar un acuerdo con un país distinto del destino declarado.")


def diversificacion_y_toneladas():
    c = pd.read_csv(DATOS / "concentracion.csv", encoding="utf-8-sig")
    nota, et = parcial(dict(zip(c["AÑO"], c["MESES"])))
    w = pd.DataFrame({"Año": c["AÑO"].map(et), "Productos exportados": c["PRODUCTOS_EXPORTADOS"],
                      "Productos exportados con acuerdo": c["PRODUCTOS_CON_ACUERDO"]})
    c0, u = c.iloc[0], c[c["MESES"] == 12].iloc[-1]
    grafico("diversificacion", "d3-lines",
            f"Cada vez más productos usan un acuerdo: {int(c0['PRODUCTOS_CON_ACUERDO'])} en {int(c0['AÑO'])}, "
            f"{int(u['PRODUCTOS_CON_ACUERDO'])} en {int(u['AÑO'])}", w,
            {"intro": "Cantidad de productos (partidas de 4 dígitos de la NCM) exportados por Paraguay, en total y con al menos un "
                      "envío con acuerdo.", "number-format": "0,0"},
            {"custom-colors": {"Productos exportados": MENTA, "Productos exportados con acuerdo": VERDE2},
             "lines": {k: {"width": "style2", "symbols": {"enabled": True, "style": "fill", "on": "every"}}
                       for k in ("Productos exportados", "Productos exportados con acuerdo")},
             "labeling": "right", "label-colors": True, "label-margin": 150, "custom-range-y": ["0", ""],
             "y-grid-format": "0,0", "tooltip-number-format": "0,0", "interpolation": "linear"},
            nota)
    t = pd.read_csv(DATOS / "toneladas_por_anio.csv", encoding="utf-8-sig")
    w = pd.DataFrame({"Año": t["AÑO"].map(et), "Con acuerdo": t["MILES_TON_CON_ACUERDO"].round(0),
                      "Sin acuerdo": t["MILES_TON_SIN_ACUERDO"].round(0)})
    tc = t[t["MESES"] == 12]
    pc = lambda r: r["MILES_TON_CON_ACUERDO"] / (r["MILES_TON_CON_ACUERDO"] + r["MILES_TON_SIN_ACUERDO"]) * 100
    grafico("toneladas", "stacked-column-chart",
            f"En toneladas, el {coma(pc(tc.iloc[-1]))}% de lo exportado en {int(tc.iloc[-1]['AÑO'])} usó un acuerdo "
            f"(en {int(tc.iloc[0]['AÑO'])}, el {coma(pc(tc.iloc[0]))}%)", w,
            {"intro": "Exportaciones de Paraguay en miles de toneladas (peso neto), con y sin acuerdo. Separa la cantidad del precio: "
                      "un año con soja cara sube los dólares, no las toneladas.", "number-format": "0,0"},
            {"custom-colors": {"Con acuerdo": VERDE, "Sin acuerdo": "#c9b3ca"}, "stack-percentages": False, "grid-lines": "show",
             "y-grid-format": "0,0", "value-label-format": "0,0", "show-color-key": True, "tooltip-number-format": "0,0"},
            nota, datos={"transpose": True})


def productos():
    p4 = pd.read_excel(RAIZ / "salidas" / "Productos_acuerdos.xlsx", sheet_name="Todos 4 dígitos", dtype={"PRODUCTO": str})
    cortos = pd.read_csv(RAIZ / "config" / "nombres_cortos_ncm.csv", dtype=str).set_index("partida")["nombre_corto"].to_dict()
    anio = int(p4.loc[p4["MESES"] == 12, "AÑO"].max())

    def corto(cod, desc):
        if cod in cortos:
            return cortos[cod]
        d = str(desc).split(",")[0].split(";")[0].strip().capitalize()
        return d if len(d) <= 34 else d[:32].rstrip() + "…"
    for op, nombre, valor in (("EXPORTACION", "exportacion", "FOB"), ("IMPORTACION", "importacion", "CIF")):
        d = p4[(p4["AÑO"] == anio) & (p4["OPERACION"] == op)]
        total_con = d["CON_ACUERDO_USD"].sum()
        d = d.sort_values("CON_ACUERDO_USD", ascending=False).head(12)
        n = [f"{corto(c, x)} ({coma(p)}%)" for c, x, p in zip(d["PRODUCTO"], d["DESCRIPCION"], d["CON_ACUERDO_PCT"])]
        w = pd.DataFrame({"Producto": n, "Usó el acuerdo": (d["CON_ACUERDO_USD"] / 1e6).round(1).values,
                          "Sin acuerdo": ((d["TOTAL_USD"] - d["CON_ACUERDO_USD"]) / 1e6).round(1).values})
        top3 = [corto(c, x).lower() for c, x in zip(d["PRODUCTO"].head(3), d["DESCRIPCION"].head(3))]
        parte = d["CON_ACUERDO_USD"].head(3).sum() / total_con * 100
        verbo = "exportado" if op == "EXPORTACION" else "importado"
        grafico(f"productos_{nombre}", "d3-bars-stacked",
                f"{(top3[0] + ', ' + top3[1] + ' y ' + top3[2]).capitalize()}: el {coma(parte)}% de lo {verbo} con acuerdo en {anio}", w,
                {"intro": f"Los 12 productos con más {'exportaciones' if op == 'EXPORTACION' else 'importaciones'} que usaron un acuerdo "
                          f"comercial, millones de USD {valor}, {anio}. Entre paréntesis, la parte de todo lo comerciado del producto que "
                          "usó un acuerdo. 4 dígitos de la NCM.", "number-format": "0,0"},
                {"custom-colors": {"Usó el acuerdo": VERDE, "Sin acuerdo": "#c9b3ca"}, "show-color-key": True,
                 "value-label-format": "0,0", "tooltip-number-format": "0,0.[0]", "sort-bars": False, "thick": True},
                "Despachos definitivos de la DNIT.")


def turismo_y_celulares():
    c = pd.read_csv(DATOS / "celulares_doble_conteo.csv", encoding="utf-8-sig").set_index("AÑO").loc[2023] / 1e6
    w = pd.DataFrame({"Forma de contar": ["Sumando todas las filas", "Contando una sola vez (SE_CUENTA = SI)", "Cifra oficial en UN Comtrade"],
                      "Millones de USD": [round(c["TODOS_LOS_REGIMENES_USD"]), round(c["SE_CUENTA_SI_USD"]), 1729]})
    grafico("celulares_doble_conteo", "d3-bars", "Celulares importados en 2023: sumar todo casi duplica el valor", w,
            {"intro": "Teléfonos celulares (NCM 8517.13 y la anterior 8517.12), millones de USD CIF. Lo que pasa por zona franca o "
                      "depósito figura dos veces en los datos. Contando una sola vez, la cifra queda cerca de la oficial.",
             "number-format": "0,0", "number-prepend": "USD ", "number-append": " millones"},
            {"base-color": VERDE, "custom-colors": {"Sumando todas las filas": VIOLETA, "Cifra oficial en UN Comtrade": GRIS},
             "value-label-format": "0,0", "show-values": True},
            "UN Comtrade: importaciones de Paraguay, partida 8517.13, 2023.")


if __name__ == "__main__":
    destino_aduana_bcp()
    importaciones_ventaja()
    evolucion_acuerdos()
    diversificacion_y_toneladas()
    productos()
    turismo_y_celulares()
