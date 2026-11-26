"""Gráfico de importaciones: lo que usó un acuerdo y lo que además tuvo ventaja real (AEC mayor a 0).

Uso (desde la raíz del repositorio):
    python docs/figuras_importaciones.py

Lee docs/datos/importaciones_acuerdo_ventaja.csv y guarda docs/img/importaciones_ventaja_{light,dark}.png.
Solo años completos (el año en curso tiene menos meses).
"""
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from figuras_cruce import FONDO, IMG, TEMAS, estilo, guardar, leyenda, titulo  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent


def main():
    d = pd.read_csv(RAIZ / "docs" / "datos" / "importaciones_acuerdo_ventaja.csv", encoding="utf-8-sig")
    d = d[d["MESES"] == 12]
    ult = d.iloc[-1]
    for modo, t in TEMAS.items():
        fig, ax = plt.subplots(figsize=(10, 4.6))
        fig.patch.set_facecolor(FONDO[modo])
        fig.subplots_adjust(left=0.07, right=0.97, top=0.66, bottom=0.1)
        ax.fill_between(d["AÑO"], d["PCT_CON_VENTAJA_POTENCIAL"], d["PCT_CON_ACUERDO"], color="#af8eb1", alpha=0.45, linewidth=0)
        ax.plot(d["AÑO"], d["PCT_CON_ACUERDO"], color="#8db2ac", linewidth=2.4)
        ax.plot(d["AÑO"], d["PCT_CON_VENTAJA_POTENCIAL"], color="#4a7c75", linewidth=2.4)
        for col, color, dy in (("PCT_CON_ACUERDO", "#8db2ac", 0.6), ("PCT_CON_VENTAJA_POTENCIAL", "#4a7c75", -0.6)):
            ax.text(ult["AÑO"] + 0.15, ult[col] + dy, f"{ult[col]:.1f}%".replace(".", ","), va="center",
                    fontsize=10, color=color, fontweight="bold")
        ax.set_ylim(15, 35)   # línea, no barras: el eje no necesita empezar en 0
        ax.set_yticks([15, 20, 25, 30, 35], ["15%", "20%", "25%", "30%", "35%"])
        ax.set_xticks(list(d["AÑO"])[::2])
        ax.set_xlim(d["AÑO"].min() - 0.3, d["AÑO"].max() + 1.2)
        estilo(ax, t)
        titulo(fig, t, f"En {int(ult['AÑO'])}, el {ult['PCT_CON_ACUERDO']:.1f}% de lo importado usó un acuerdo; "
               f"con ventaja arancelaria potencial, el {ult['PCT_CON_VENTAJA_POTENCIAL']:.1f}%".replace(".", ","),
               "% de lo importado (CIF). Con ventaja potencial: sin los productos con Arancel Externo Común 0% ni los que, "
               "importados sin acuerdo ese año, casi no pagaron arancel (listas nacionales u otros regímenes). Despachos definitivos de la DNIT.")
        leyenda(fig, t, 0.07, 0.73, [("Usó un acuerdo", "#8db2ac", "━"), ("Con ventaja arancelaria potencial", "#4a7c75", "━"),
                                     ("Usó un acuerdo sin ventaja (ya entraba con 0%)", "#af8eb1", "■")])
        guardar(fig, "importaciones_ventaja", modo)
    print("Listo: importaciones_ventaja_{light,dark}.png en", IMG)


if __name__ == "__main__":
    main()
