"""Ahorro arancelario estimado de las importaciones con acuerdo, y lo que se pagó por productos de origen MERCOSUR
importados sin usar el acuerdo. También: exportaciones al MERCOSUR con ventaja arancelaria potencial.

Uso (desde la raíz del repositorio):  python python/ahorro_arancelario.py
Lee salidas/dnit_anual/DNIT_<año>_{importacion,exportacion}.xlsx y config/aec_por_anio.csv.gz.
Salidas: docs/datos/ahorro_arancelario.csv y docs/datos/ahorro_productos_mercosur_sin_acuerdo.csv

Método (aproximación):
- Ahorro estimado = base imponible (USD) x tasa que pagó ese año el mismo producto (NCM 8 dígitos) importado SIN acuerdo
  (derecho / base imponible). Si ese año nadie lo importó sin acuerdo, se usa el AEC de la versión de la NCM vigente.
- Pagado sin usar el MERCOSUR = derecho aduanero (en USD) de lo importado SIN acuerdo con país de ORIGEN Argentina, Brasil
  o Uruguay. Una parte es legítima: el azúcar y el sector automotor están fuera del libre comercio del MERCOSUR, y hay
  mercadería que no cumple las reglas de origen.
- Exportaciones al MERCOSUR: Argentina, Brasil y Uruguay aplican el AEC a terceros (salvo sus excepciones nacionales);
  con ventaja potencial = el AEC del producto ese año es mayor a 0.
"""
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
MERCOSUR = {"ARGENTINA", "BRASIL", "URUGUAY"}
AEC = pd.read_csv(RAIZ / "config" / "aec_por_anio.csv.gz", dtype={"NCM": str})


def pais(s):
    """Algunos años la DNIT escribe "BR - BRASIL" en lugar de "BRASIL"."""
    return s.astype(str).str.upper().str.replace(r"^[A-Z]{2} - ", "", regex=True).str.strip()


def ncm(s):
    return s.astype(str).str.replace(".", "", regex=False).str[:8]


def main():
    filas, prods = [], []
    for a in range(2011, 2027):
        ruta = RAIZ / "salidas" / "dnit_anual" / f"DNIT_{a}_importacion.xlsx"
        if not ruta.exists():
            continue
        d = pd.read_excel(ruta, sheet_name="Datos", engine="calamine",
                          usecols=["SE_CUENTA", "PAIS ORIGEN", "POSICION", "MOTIVO_ARANCEL", "IMPONIBLE DOLAR", "IMPONIBLE GS", "DERECHO"])
        d = d[d["SE_CUENTA"] == "SI"].copy()
        d["NCM"] = ncm(d["POSICION"])
        cot = d["IMPONIBLE GS"].sum() / d["IMPONIBLE DOLAR"].sum()          # guaraníes por dólar del año
        d["DERECHO_USD"] = d["DERECHO"] / cot
        ca = d["MOTIVO_ARANCEL"].str.startswith("Con acuerdo")
        s = d[~ca].groupby("NCM")[["DERECHO", "IMPONIBLE GS"]].sum()
        tasa = (s["DERECHO"] / s["IMPONIBLE GS"]).where(s["IMPONIBLE GS"] > 0)
        aec = AEC[AEC["ANIO"] == min(a, AEC["ANIO"].max())].set_index("NCM")["AEC"] / 100
        d["TASA"] = d["NCM"].map(tasa).fillna(d["NCM"].map(aec)).fillna(0)
        ahorro = (d.loc[ca, "IMPONIBLE DOLAR"] * d.loc[ca, "TASA"]).sum()
        ms = ~ca & pais(d["PAIS ORIGEN"]).isin(MERCOSUR)
        pagado = d.loc[ms, "DERECHO_USD"].sum()
        g = d[ms].groupby("NCM").agg(IMPORTADO_USD=("IMPONIBLE DOLAR", "sum"), DERECHO_USD=("DERECHO_USD", "sum"))
        g = g[g["DERECHO_USD"] > 0].reset_index().assign(AÑO=a)
        prods.append(g)
        # exportaciones al MERCOSUR
        e = pd.read_excel(RAIZ / "salidas" / "dnit_anual" / f"DNIT_{a}_exportacion.xlsx", sheet_name="Datos", engine="calamine",
                          usecols=["SE_CUENTA", "PAIS ORIGEN", "PAIS PROCEDENCIA/DESTINO", "POSICION", "ACUERDO", "FOB DOLAR"])
        e = e[(e["SE_CUENTA"] == "SI") & pais(e["PAIS ORIGEN"]).eq("PARAGUAY")
              & pais(e["PAIS PROCEDENCIA/DESTINO"]).isin(MERCOSUR)].copy()
        e["AEC"] = ncm(e["POSICION"]).map(aec)
        eca = ~e["ACUERDO"].astype(str).str.upper().isin(["SIN ACUERDO", "OMC", "NAN", ""])
        filas.append({"AÑO": a, "IMPORTADO_CON_ACUERDO_USD": d.loc[ca, "IMPONIBLE DOLAR"].sum(),
                      "AHORRO_ESTIMADO_USD": ahorro, "DERECHO_PAGADO_TOTAL_USD": d["DERECHO_USD"].sum(),
                      "IMPORTADO_MERCOSUR_SIN_ACUERDO_USD": d.loc[ms, "IMPONIBLE DOLAR"].sum(),
                      "DERECHO_PAGADO_ORIGEN_MERCOSUR_SIN_ACUERDO_USD": pagado,
                      "EXPORTADO_AL_MERCOSUR_USD": e["FOB DOLAR"].sum(),
                      "EXPORT_MERCOSUR_CON_ACUERDO_USD": e.loc[eca, "FOB DOLAR"].sum(),
                      "EXPORT_MERCOSUR_CON_ACUERDO_AEC_MAYOR_0_USD": e.loc[eca & (e["AEC"] > 0), "FOB DOLAR"].sum(),
                      "EXPORT_MERCOSUR_SIN_ACUERDO_AEC_MAYOR_0_USD": e.loc[~eca & (e["AEC"] > 0), "FOB DOLAR"].sum()})
        print(a, "listo", flush=True)
    out = RAIZ / "docs" / "datos"
    r = pd.DataFrame(filas)
    r.round(2).to_csv(out / "ahorro_arancelario.csv", index=False, encoding="utf-8-sig")
    p = pd.concat(prods)
    p.sort_values(["AÑO", "DERECHO_USD"], ascending=[True, False]).round(2).to_csv(
        out / "ahorro_productos_mercosur_sin_acuerdo.csv", index=False, encoding="utf-8-sig")
    print((r.set_index("AÑO") / 1e6).round(1).to_string())


if __name__ == "__main__":
    main()
