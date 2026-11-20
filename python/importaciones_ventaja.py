"""Importaciones con acuerdo y con ventaja arancelaria potencial, por año.

Usa los archivos anuales de la DNIT (salidas/dnit_anual/DNIT_<año>_importacion.xlsx, de dnit_raw_por_anio.py).
Lo importado con acuerdo NO tiene ventaja potencial cuando:
  1. el Arancel Externo Común del producto ese año es 0 % (MOTIVO_ARANCEL "Con acuerdo, pero AEC 0 %"), o
  2. el mismo producto (NCM 8 dígitos), importado SIN acuerdo ese año, casi no pagó arancel
     (derecho aduanero menor al 0,5 % de la base imponible): listas nacionales, otras leyes o regímenes.
Es una aproximación: no ve exoneraciones propias de cada importador ni preferencias parciales.

Uso (desde la raíz del repositorio):  python python/importaciones_ventaja.py
Salida: docs/datos/importaciones_acuerdo_ventaja.csv
"""
import pandas as pd, sys
sys.path.insert(0,"python")
R=[]
for a in range(2011,2027):
    d=pd.read_excel(f"salidas/dnit_anual/DNIT_{a}_importacion.xlsx",sheet_name="Datos",engine="calamine",
        usecols=["SE_CUENTA","POSICION","ACUERDO","MOTIVO_ARANCEL","FOB DOLAR","FLETE DOLAR","SEGURO DOLAR","DERECHO","IMPONIBLE GS"])
    d=d[d.SE_CUENTA=="SI"].copy()
    d["NCM"]=d.POSICION.astype(str).str.replace(".","",regex=False).str[:8]
    d["V"]=d["FOB DOLAR"]+d["FLETE DOLAR"]+d["SEGURO DOLAR"]
    ca=d.MOTIVO_ARANCEL.str.startswith("Con acuerdo")
    s=d[~ca].groupby("NCM")[["DERECHO","IMPONIBLE GS"]].sum()
    tasa=(s.DERECHO/s["IMPONIBLE GS"]).where(s["IMPONIBLE GS"]>0)
    d["TASA_SIN"]=d.NCM.map(tasa)
    aec0=d.MOTIVO_ARANCEL.str.contains("AEC 0")
    cero_sin=d.TASA_SIN.notna()&(d.TASA_SIN<0.005)   # el mismo producto, sin acuerdo, casi no pagó arancel
    sin_dato=d.TASA_SIN.isna()
    v=d.V; tot=v.sum()
    R.append(dict(AÑO=a,TOT=tot,CA=v[ca].sum(),AEC0=v[ca&aec0].sum(),CERO_SIN=v[ca&~aec0&cero_sin].sum(),SIN_COMP=v[ca&~aec0&sin_dato].sum()))
r=pd.DataFrame(R); 
for c in ["CA","AEC0","CERO_SIN","SIN_COMP"]: r[c+"%"]=r[c]/r.TOT*100
r["VENT%"]=(r.CA-r.AEC0-r.CERO_SIN)/r.TOT*100
print(r[["AÑO","CA%","AEC0%","CERO_SIN%","SIN_COMP%","VENT%"]].round(2).to_string())
meses = {a: 12 for a in r["AÑO"]}
try:
    meses.update(pd.read_csv("docs/datos/importaciones_acuerdo_ventaja.csv", encoding="utf-8-sig").set_index("AÑO")["MESES"].to_dict())
except FileNotFoundError:
    pass
o = pd.DataFrame({"AÑO": r.AÑO, "MESES": r.AÑO.map(meses), "IMPORTACIONES_USD": r.TOT, "CON_ACUERDO_USD": r.CA,
                  "CON_ACUERDO_AEC0_USD": r.AEC0, "CON_ACUERDO_SIN_ACUERDO_PAGA_0_USD": r.CERO_SIN,
                  "CON_ACUERDO_SIN_COMPARACION_USD": r.SIN_COMP,
                  "CON_ACUERDO_CON_VENTAJA_POTENCIAL_USD": r.CA - r.AEC0 - r.CERO_SIN,
                  "PCT_CON_ACUERDO": r.CA / r.TOT * 100, "PCT_CON_VENTAJA_POTENCIAL": (r.CA - r.AEC0 - r.CERO_SIN) / r.TOT * 100})
o.to_csv("docs/datos/importaciones_acuerdo_ventaja.csv", index=False, encoding="utf-8-sig")
