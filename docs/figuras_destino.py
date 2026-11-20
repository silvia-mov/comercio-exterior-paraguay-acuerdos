"""Destino según la aduana (DNIT) y según el BCP (certificado de origen) para Argentina y Uruguay.

Uso (desde la raíz del repositorio), después de datos_web.py:
    python docs/figuras_destino.py

Lee docs/datos/exportaciones_por_pais.csv y guarda docs/img/destino_aduana_bcp_{light,dark}.png y
docs/datos/destino_aduana_bcp.csv. Sin electricidad. Solo años completos.
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
from figuras_extra import coma  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "docs" / "datos"


def main():
    d = pd.read_csv(DATOS / "exportaciones_por_pais.csv", keep_default_na=False)
    meses = pd.read_csv(DATOS / "utilizacion_por_anio.csv").set_index("AÑO")["MESES"]
    d = d[d["AÑO"].map(meses) == 12]
    sel = d[d["PAIS"].isin(["ARGENTINA", "URUGUAY"])][["AÑO", "PAIS", "DNIT_EXPORTADO_USD", "BCP_SIN_ELECTRICIDAD_USD"]].copy()
    sel["DIFERENCIA_USD"] = sel["DNIT_EXPORTADO_USD"] - sel["BCP_SIN_ELECTRICIDAD_USD"]
    sel.to_csv(DATOS / "destino_aduana_bcp.csv", index=False, encoding="utf-8-sig")
    for modo, t in TEMAS.items():
        fig, axes = plt.subplots(1, 2, figsize=(10, 4.6), sharey=False)
        fig.patch.set_facecolor(FONDO[modo])
        fig.subplots_adjust(left=0.08, right=0.97, top=0.64, bottom=0.12, wspace=0.22)
        for ax, p, nombre in zip(axes, ["ARGENTINA", "URUGUAY"], ["Argentina", "Uruguay"]):
            x = sel[sel["PAIS"] == p].drop(columns="PAIS").set_index("AÑO") / 1e6
            ax.fill_between(x.index, x["BCP_SIN_ELECTRICIDAD_USD"], x["DNIT_EXPORTADO_USD"], color=t["rosa_claro"], alpha=0.8, linewidth=0)
            ax.plot(x.index, x["DNIT_EXPORTADO_USD"], color=t["verde"], linewidth=2.5)
            ax.plot(x.index, x["BCP_SIN_ELECTRICIDAD_USD"], color=t["texto"], linewidth=2.5)
            ax.set_title(nombre, loc="left", fontsize=11, color=t["texto"])
            ax.set_ylim(0, x[["DNIT_EXPORTADO_USD", "BCP_SIN_ELECTRICIDAD_USD"]].max().max() * 1.12)
            ax.set_xticks([x.index.min(), x.index.min() + (x.index.max() - x.index.min()) // 2, x.index.max()])
            ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _p: coma(v, 0)))
            estilo(ax, t)
        dif = sel[sel["AÑO"] == sel["AÑO"].max()].set_index("PAIS")["DIFERENCIA_USD"] / 1e6
        titulo(fig, t, f"En {sel['AÑO'].max()}, USD {coma(dif['ARGENTINA'], 0)} millones declarados hacia Argentina y "
               f"{coma(dif['URUGUAY'], 0)} hacia Uruguay tuvieron otro destino final",
               "Exportaciones de Paraguay sin electricidad, millones de USD FOB. Verde: destino declarado en el despacho (DNIT). "
               "La otra línea: destino según el certificado de origen (BCP). La franja es lo que pasa por ese país pero termina en otro.")
        leyenda(fig, t, 0.08, 0.75, [("Destino en la aduana (DNIT)", t["verde"], "━"), ("Destino final (BCP)", t["texto"], "━"),
                                     ("Diferencia", t["rosa_claro"], "■")])
        guardar(fig, "destino_aduana_bcp", modo)
    print(sel.groupby("PAIS")["DIFERENCIA_USD"].agg(["first", "last"]) / 1e6)


if __name__ == "__main__":
    main()
