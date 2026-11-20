"""Gráficos de la página de aduana: importaciones con IVA reducido (Turismo de Compras) y el doble conteo de los celulares.

Uso (desde la raíz del repositorio), después de python/turismo_compras_serie.py:
    python docs/figuras_turismo.py

Guarda docs/img/turismo_compras_{light,dark}.png y docs/img/celulares_doble_conteo_{light,dark}.png.
"""
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from figuras_cruce import FONDO, TEMAS, estilo, guardar, leyenda, titulo  # noqa: E402
from figuras_extra import coma  # noqa: E402

DATOS = Path(__file__).resolve().parent / "datos"
COMTRADE_CEL_2023 = 1729.0   # Paraguay, importaciones 8517.13 en 2023, UN Comtrade (millones de USD, CIF)


def turismo():
    d = pd.read_csv(DATOS / "turismo_compras_por_anio.csv", encoding="utf-8-sig")
    tot = d["PCT_TURISMO_COMPRAS"] + d["PCT_IVA_REDUCIDO_OTRO"]
    comp = d[d["MESES"] == 12]
    ult = comp.iloc[-1]
    for modo, t in TEMAS.items():
        fig, ax = plt.subplots(figsize=(10, 4.6))
        fig.patch.set_facecolor(FONDO[modo])
        fig.subplots_adjust(left=0.07, right=0.97, top=0.64, bottom=0.1)
        alfa = [1 if m == 12 else 0.45 for m in d["MESES"]]
        for i, r in d.iterrows():
            ax.bar(r["AÑO"], r["PCT_TURISMO_COMPRAS"], color=t["verde"], alpha=alfa[i], width=0.7)
            ax.bar(r["AÑO"], r["PCT_IVA_REDUCIDO_OTRO"], bottom=r["PCT_TURISMO_COMPRAS"], color=t["rosa"], alpha=alfa[i], width=0.7)
            if r["AÑO"] in (d["AÑO"].min(), comp["AÑO"].max()) or r["MESES"] < 12:
                ax.text(r["AÑO"], tot[i] + 0.6, coma(tot[i], 0) + " %", ha="center", fontsize=9.5, color=t["texto"])
        ax.set_ylim(0, 30)
        ax.set_yticks([0, 10, 20, 30], ["0 %", "10 %", "20 %", "30 %"])
        ax.set_xticks(list(d["AÑO"])[::3], [str(a) if m == 12 else f"{a}*" for a, m in zip(d["AÑO"][::3], d["MESES"][::3])])
        estilo(ax, t)
        titulo(fig, t, f"Cerca de una quinta parte de lo importado entra con IVA reducido ({coma(tot[ult.name], 0)} % en {int(ult['AÑO'])})",
               "% de lo importado (CIF, despachos definitivos de la DNIT) que pagó IVA de hasta 2,2 % de la base. "
               "Antes de 2022 la NCM usaba otros códigos: parte del turismo de compras queda en rosa. *2026: enero a septiembre.")
        leyenda(fig, t, 0.07, 0.72, [("Productos de la lista del Decreto 2063/2024 (Turismo de Compras)", t["verde"], "■"),
                                     ("Otros productos con IVA reducido", t["rosa"], "■")])
        guardar(fig, "turismo_compras", modo)


def celulares():
    c = pd.read_csv(DATOS / "celulares_doble_conteo.csv", encoding="utf-8-sig").set_index("AÑO").loc[2023] / 1e6
    vals = [c["TODOS_LOS_REGIMENES_USD"], c["SE_CUENTA_SI_USD"], COMTRADE_CEL_2023]
    nombres = ["Sumando todas las filas", "Contando una sola vez\n(SE_CUENTA = SI)", "Cifra oficial en\nUN Comtrade"]
    for modo, t in TEMAS.items():
        fig, ax = plt.subplots(figsize=(10, 4.2))
        fig.patch.set_facecolor(FONDO[modo])
        fig.subplots_adjust(left=0.24, right=0.93, top=0.7, bottom=0.08)
        y = [2, 1, 0]
        ax.barh(y, vals, color=[t["rosa"], t["verde"], t["texto2"]], height=0.6)
        for yy, v in zip(y, vals):
            ax.text(v + 40, yy, f"USD {coma(v, 0)} millones", va="center", fontsize=10.5, color=t["texto"])
        ax.set_yticks(y, nombres)
        ax.set_xlim(0, max(vals) * 1.3)
        ax.set_xticks([])
        estilo(ax, t, grilla="x")
        ax.grid(False)
        ax.spines["bottom"].set_visible(False)
        titulo(fig, t, "Celulares importados en 2023: sumar todo casi duplica el valor",
               "Teléfonos celulares (NCM 8517.13 y la anterior 8517.12), millones de USD CIF. Lo que pasa por zona franca o depósito figura dos veces "
               "en los datos. Contando una sola vez, la cifra queda cerca de la oficial.")
        guardar(fig, "celulares_doble_conteo", modo)


if __name__ == "__main__":
    turismo()
    celulares()
