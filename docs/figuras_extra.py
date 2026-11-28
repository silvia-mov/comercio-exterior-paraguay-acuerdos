"""Tres gráficos nuevos de la página: evolución de cada acuerdo, concentración y toneladas.

Uso (desde la raíz del repositorio), después de datos_web.py:
    python docs/figuras_extra.py

Lee docs/datos/comercio_por_acuerdo.csv, docs/datos/exportaciones_por_pais.csv y salidas/productos/productos_<año>.csv.
Guarda en docs/img/ (versión clara y oscura) y los datos en docs/datos/concentracion.csv y toneladas_por_anio.csv.
Solo exportaciones de origen paraguayo (DNIT, regímenes definitivos). 2026 tiene menos meses: se marca con *.
"""
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from figuras_cruce import FONDO, TEMAS, estilo, guardar, leyenda, titulo  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "docs" / "datos"
PALETA = ["#4a7c75", "#5e8f88", "#8db2ac", "#774d78", "#af8eb1", "#9a7a9c", "#8a7f86"]
CORTO = {"MERCOSUR intrazona (ACE 18)": "MERCOSUR (ACE 18)", "ALADI Bilateral Brasil (AAP.CE 74)": "Brasil (AAP.CE 74)",
         "ALADI Bilateral Argentina (AAP.CE 13)": "Argentina (AAP.CE 13)", "ALADI Bilateral México (AAP.R 38)": "México (AAP.R 38)",
         "ALADI Preferencia Arancelaria Regional (PAR 4)": "ALADI PAR 4"}


def hhi(valores):
    s = valores / valores.sum()
    return float((s ** 2).sum() * 10000)


def coma(x, d=1):
    return f"{x:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def main():
    acu = pd.read_csv(DATOS / "comercio_por_acuerdo.csv")
    pais = pd.read_csv(DATOS / "exportaciones_por_pais.csv", keep_default_na=False)
    prods = pd.concat([pd.read_csv(f, dtype={"NCM": str}, keep_default_na=False)
                       for f in sorted((RAIZ / "salidas" / "productos").glob("productos_20*.csv"))])
    prods = prods[prods["OPERACION"] == "EXPORTACION"]
    prods["CON"] = prods["TIPO_ACUERDO"].isin(["PREFERENCIAL", "MERCOSUR_INTRAZONA"])
    meses = prods.groupby("AÑO")["MESES"].max().astype(int)
    anios = sorted(meses.index)
    completos = [a for a in anios if meses[a] == 12]

    # ---------------- datos: concentración (HHI de 0 a 10.000) y toneladas
    filas = []
    for a in anios:
        p = prods[prods["AÑO"] == a]
        p4 = p.assign(P4=p["NCM"].str[:4])
        con, sin = p4[p4["CON"]].groupby("P4")["VALOR_USD"].sum(), p4[~p4["CON"]].groupby("P4")["VALOR_USD"].sum()
        tot = p4.groupby("P4")["VALOR_USD"].sum()
        d = pais[pais["AÑO"] == a].groupby("CODIGO_PAIS")[["DNIT_CON_ACUERDO_USD", "DNIT_SIN_ACUERDO_USD", "DNIT_EXPORTADO_USD"]].sum()
        filas.append({"AÑO": a, "MESES": int(meses[a]),
                      "HHI_PRODUCTOS_CON_ACUERDO": hhi(con), "HHI_PRODUCTOS_SIN_ACUERDO": hhi(sin), "HHI_PRODUCTOS_TOTAL": hhi(tot),
                      "TOP5_PRODUCTOS_PCT": tot.nlargest(5).sum() / tot.sum() * 100,
                      "HHI_DESTINOS_CON_ACUERDO": hhi(d["DNIT_CON_ACUERDO_USD"]), "HHI_DESTINOS_SIN_ACUERDO": hhi(d["DNIT_SIN_ACUERDO_USD"]),
                      "TOP10_DESTINOS_PCT": d["DNIT_EXPORTADO_USD"].nlargest(10).sum() / d["DNIT_EXPORTADO_USD"].sum() * 100,
                      "PRODUCTOS_EXPORTADOS": int((tot > 0).sum()), "PRODUCTOS_CON_ACUERDO": int((con > 0).sum()),
                      "DESTINOS": int((d["DNIT_EXPORTADO_USD"] > 0).sum())})
    conc = pd.DataFrame(filas)
    conc.round(2).to_csv(DATOS / "concentracion.csv", index=False, encoding="utf-8-sig")
    ton = (prods.groupby(["AÑO", "CON"])["PESO_NETO_KG"].sum().unstack(fill_value=0) / 1e6)
    ton.columns = ["MILES_TON_SIN_ACUERDO" if not c else "MILES_TON_CON_ACUERDO" for c in ton.columns]
    ton = ton.reset_index()
    ton["MESES"] = ton["AÑO"].map(meses)
    ton["USD_POR_TON_CON_ACUERDO"] = prods[prods["CON"]].groupby("AÑO")["VALOR_USD"].sum().values / (ton["MILES_TON_CON_ACUERDO"] * 1000)
    ton.round(2).to_csv(DATOS / "toneladas_por_anio.csv", index=False, encoding="utf-8-sig")

    etiq = [f"{a}*" if meses[a] < 12 else str(a) for a in anios]
    for modo, t in TEMAS.items():
        # ---------------- 1. evolución de cada acuerdo (exportaciones, sin el MERCOSUR en un panel aparte)
        e = acu[(acu["OPERACION"] == "EXPORTACION") & acu["TIPO_ACUERDO"].isin(["PREFERENCIAL", "MERCOSUR_INTRAZONA"]) & acu["AÑO"].isin(completos)]
        e = e.assign(N=e["ACUERDO_CONSOLIDADO"].replace(CORTO))
        piv = e.pivot_table(index="AÑO", columns="N", values="VALOR_USD", aggfunc="sum").fillna(0) / 1e6
        otros = [c for c in piv.sum().sort_values(ascending=False).index if c != "MERCOSUR (ACE 18)"][:6]
        fig, (a1, a2) = plt.subplots(1, 2, figsize=(11, 5), gridspec_kw={"width_ratios": [1, 2.2]})
        fig.patch.set_facecolor(FONDO[modo])
        fig.subplots_adjust(left=0.07, right=0.83, top=0.74, bottom=0.1, wspace=0.28)
        a1.plot(piv.index, piv["MERCOSUR (ACE 18)"], color=PALETA[0], linewidth=2.6)
        a1.set_title("MERCOSUR (ACE 18)", loc="left", fontsize=11, color=t["texto"])
        a1.set_ylim(0, piv["MERCOSUR (ACE 18)"].max() * 1.1)
        ymax = piv[otros].max().max() * 1.1
        # etiquetas al final de cada línea, separadas para que no se pisen
        pos = sorted(((piv[c].iloc[-1], k, c) for k, c in enumerate(otros)))
        ult_y = -1e9
        for v, k, c in pos:
            y = max(v, ult_y + ymax * 0.065)
            ult_y = y
            a2.text(piv.index[-1] + 0.3, y, f"{c} ({coma(v, 0)})", va="center", fontsize=8.5, color=PALETA[(k + 1) % len(PALETA)])
        for k, c in enumerate(otros):
            a2.plot(piv.index, piv[c], color=PALETA[(k + 1) % len(PALETA)], linewidth=2.2)
        a2.set_title("Los demás acuerdos (otra escala)", loc="left", fontsize=11, color=t["texto"])
        a2.set_ylim(0, piv[otros].max().max() * 1.1)
        for ax in (a1, a2):
            ax.set_xticks([completos[0], completos[len(completos) // 2], completos[-1]])
            ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, _p: coma(x, 0)))
            estilo(ax, t)
        crece = (piv[otros].iloc[-3:].mean() / piv[otros].iloc[:3].mean().replace(0, float("nan"))).dropna()
        titulo(fig, t, f"Fuera del MERCOSUR, el acuerdo con Chile es el que más sumó: de "
               f"{coma(piv['Mercosur-Chile (ACE 35)'].iloc[0], 0)} a {coma(piv['Mercosur-Chile (ACE 35)'].iloc[-1], 0)} millones de USD",
               f"Exportaciones de Paraguay con cada acuerdo, millones de USD FOB, {completos[0]}–{completos[-1]} (años completos). "
               "Despachos definitivos de la DNIT. El MERCOSUR va aparte porque es unas nueve veces más grande que el que le sigue.")
        guardar(fig, "evolucion_acuerdos", modo)

        # ---------------- 2. diversificación: cuántos productos se exportan, en total y con acuerdo
        fig, ax = plt.subplots(figsize=(10, 4.6))
        fig.patch.set_facecolor(FONDO[modo])
        fig.subplots_adjust(left=0.08, right=0.95, top=0.7, bottom=0.12)
        ax.plot(etiq, conc["PRODUCTOS_EXPORTADOS"], color=t["menta"], linewidth=2.5, marker="o", markersize=4)
        ax.plot(etiq, conc["PRODUCTOS_CON_ACUERDO"], color=t["verde2"], linewidth=2.5, marker="o", markersize=4)
        ax.set_ylim(0, conc["PRODUCTOS_EXPORTADOS"].max() * 1.15)
        estilo(ax, t)
        c0, u = conc.iloc[0], conc[conc["MESES"] == 12].iloc[-1]
        titulo(fig, t, f"Cada vez más productos usan un acuerdo: {int(c0['PRODUCTOS_CON_ACUERDO'])} en {int(c0['AÑO'])}, "
               f"{int(u['PRODUCTOS_CON_ACUERDO'])} en {int(u['AÑO'])}",
               "Cantidad de productos (partidas de 4 dígitos de la NCM) exportados por Paraguay, en total y con al menos un envío "
               "con acuerdo. Despachos de la DNIT. * 2026: enero a septiembre.")
        leyenda(fig, t, 0.08, 0.77, [("Productos exportados", t["menta"], "●"), ("Productos exportados con acuerdo", t["verde2"], "●")])
        guardar(fig, "diversificacion", modo)

        # ---------------- 3. toneladas con y sin acuerdo
        fig, ax = plt.subplots(figsize=(10, 4.6))
        fig.patch.set_facecolor(FONDO[modo])
        fig.subplots_adjust(left=0.08, right=0.95, top=0.7, bottom=0.12)
        ax.bar(etiq, ton["MILES_TON_CON_ACUERDO"], color=t["verde"], width=0.72)
        ax.bar(etiq, ton["MILES_TON_SIN_ACUERDO"], bottom=ton["MILES_TON_CON_ACUERDO"], color=t["rosa_claro"], width=0.72)
        ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, _p: coma(x, 0)))
        estilo(ax, t)
        tc = ton[ton["MESES"] == 12]
        pc = lambda r: r["MILES_TON_CON_ACUERDO"] / (r["MILES_TON_CON_ACUERDO"] + r["MILES_TON_SIN_ACUERDO"]) * 100
        titulo(fig, t, f"En toneladas, el {coma(pc(tc.iloc[-1]), 0)}% de lo exportado en {int(tc.iloc[-1]['AÑO'])} usó un acuerdo "
               f"(en {int(tc.iloc[0]['AÑO'])}, el {coma(pc(tc.iloc[0]), 0)}%)",
               "Exportaciones de Paraguay en miles de toneladas (peso neto), con y sin acuerdo. Despachos de la DNIT. "
               "Separa la cantidad del precio: un año con soja cara sube los dólares, no las toneladas. * 2026: enero a septiembre.")
        leyenda(fig, t, 0.08, 0.77, [("Con acuerdo", t["verde"], "■"), ("Sin acuerdo", t["rosa_claro"], "■")])
        guardar(fig, "toneladas", modo)
    print(conc[["AÑO", "HHI_PRODUCTOS_CON_ACUERDO", "HHI_PRODUCTOS_SIN_ACUERDO", "TOP5_PRODUCTOS_PCT", "HHI_DESTINOS_CON_ACUERDO",
                "HHI_DESTINOS_SIN_ACUERDO", "TOP10_DESTINOS_PCT", "PRODUCTOS_EXPORTADOS", "PRODUCTOS_CON_ACUERDO"]].round(0).to_string())
    print(ton.round(0).to_string())


if __name__ == "__main__":
    main()
