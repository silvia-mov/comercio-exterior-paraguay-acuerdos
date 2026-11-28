"""Aplica la regla «acuerdo que no corresponde al destino» a los resúmenes ya calculados (salidas/), sin volver a
leer los archivos de la DNIT. Las corridas nuevas de procesar_despachos.py, resumen_productos.py y cruzar_dnit_vue.py
ya aplican la regla solas (ver procesar_despachos.marcar_acuerdo_no_corresponde y config/acuerdos_paises.csv).

Uso: python python/corregir_acuerdo_destino.py
Corrige, en su lugar:
  salidas/productos/productos_<año>.csv        (exportaciones reclasificadas)
  salidas/resumen_dnit/dnit_exportaciones_<año>.csv (lo reclasificado pasa de con acuerdo a sin acuerdo)
  salidas/por_acuerdo/por_acuerdo_<año>.csv     (lo reclasificado sale de cada acuerdo y va a la categoría nueva)
Se puede correr más de una vez: lo ya corregido no se vuelve a restar.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import procesar_despachos as pd_dnit  # noqa: E402

SAL = pd_dnit.RAIZ_REPO / "salidas"
MARCA = "CORREGIDO_ACUERDO_DESTINO"


def main():
    total = 0.0
    for f in sorted((SAL / "productos").glob("productos_20*.csv")):
        p = pd.read_csv(f, dtype={"NCM": str, "CODIGO_PAIS": str}, keep_default_na=False, na_values=[""], encoding="utf-8-sig")
        anio = int(p["AÑO"].iloc[0])
        antes = p.copy()
        p = pd_dnit.marcar_acuerdo_no_corresponde(p, col_anio="AÑO")
        nuevo = (p["TIPO_ACUERDO"] == pd_dnit.ACUERDO_NO_CORRESPONDE) & (antes["TIPO_ACUERDO"] != pd_dnit.ACUERDO_NO_CORRESPONDE)
        if nuevo.any():
            claves = ["AÑO", "OPERACION", "NCM", "CODIGO_PAIS", "TIPO_ACUERDO", "ACUERDO_CONSOLIDADO"]
            otras = [c for c in p.columns if c not in claves + ["VALOR_USD", "PESO_NETO_KG"]]
            p = (p.groupby(claves, as_index=False, dropna=False)
                  .agg(VALOR_USD=("VALOR_USD", "sum"), PESO_NETO_KG=("PESO_NETO_KG", "sum"), **{c: (c, "first") for c in otras}))
            p.to_csv(f, index=False, encoding="utf-8-sig")
        mal = antes[nuevo]
        total += mal["VALOR_USD"].sum()

        # Resumen de exportaciones por país y partida (cruce de fuentes): de con acuerdo a sin acuerdo
        r_f = SAL / "resumen_dnit" / f"dnit_exportaciones_{anio}.csv"
        if r_f.exists() and len(mal):
            r = pd.read_csv(r_f, dtype={"PARTIDA": str, "CODIGO_PAIS": str}, keep_default_na=False, na_values=[""], encoding="utf-8-sig")
            if MARCA not in r.columns:
                m = (mal.assign(PARTIDA=mal["NCM"].str[:4]).groupby(["CODIGO_PAIS", "PARTIDA"])["VALOR_USD"].sum())
                k = list(zip(r["CODIGO_PAIS"], r["PARTIDA"]))
                resta = pd.Series([m.get(x, 0.0) for x in k], index=r.index).clip(upper=r["DNIT_FOB_CON_ACUERDO"])
                r["DNIT_FOB_CON_ACUERDO"] -= resta
                r["DNIT_FOB_SIN_ACUERDO"] += resta
                r[MARCA] = "SI"
                r.to_csv(r_f, index=False, encoding="utf-8-sig")

        # Resumen por acuerdo: el monto sale de cada acuerdo y va a la categoría nueva
        a_f = SAL / "por_acuerdo" / f"por_acuerdo_{anio}.csv"
        if a_f.exists() and len(mal):
            a = pd.read_csv(a_f, encoding="utf-8-sig")
            if not (a["ACUERDO_CONSOLIDADO"] == pd_dnit.NOMBRE_NO_CORRESPONDE).any():
                m = mal.groupby("ACUERDO_CONSOLIDADO")["VALOR_USD"].sum()
                exp = a["OPERACION"].str.startswith("EXP")
                for acu, v in m.items():
                    a.loc[exp & (a["ACUERDO_CONSOLIDADO"] == acu), "VALOR_DOLAR"] -= v
                a = pd.concat([a, pd.DataFrame([{"AÑO": anio, "OPERACION": "EXPORTACION", "TIPO_ACUERDO": pd_dnit.ACUERDO_NO_CORRESPONDE,
                                                 "ACUERDO_CONSOLIDADO": pd_dnit.NOMBRE_NO_CORRESPONDE, "VALOR_DOLAR": m.sum()}])],
                              ignore_index=True)
                a.to_csv(a_f, index=False, encoding="utf-8-sig")
        print(f"{anio}: {mal['VALOR_USD'].sum() / 1e6:,.1f} millones reclasificados")
    print(f"Total: {total / 1e6:,.1f} millones de USD")


if __name__ == "__main__":
    main()
