"""Lee las notificaciones de Paraguay al MERCOSUR (BK, BIT, LNE): código NCM y arancel nacional vigente."""
import re, glob, json, sys, openpyxl
from collections import defaultdict
def read_notif(f):
    """devuelve {lista: [(ncm8, anv, desc)]}"""
    out={}
    wb=openpyxl.load_workbook(f,read_only=True,data_only=True)
    for ws in wb.worksheets:
        rows=list(ws.iter_rows(values_only=True)); hdr=None; items=[]
        for r in rows:
            r=list(r)
            if hdr is None and r and any(isinstance(c,str) and 'Arancel' in c for c in r if c):
                hdr=r; ianv=[i for i,c in enumerate(r) if isinstance(c,str) and 'Arancel' in c][0]
                idesc=next((i for i,c in enumerate(r) if isinstance(c,str) and 'Descrip' in c),None); continue
            if hdr is None or not r or r[0] is None: continue
            c=re.sub(r'\D','',str(r[0]))
            if len(c)==8:
                items.append((c, r[ianv], r[idesc] if idesc is not None else None))
        out[ws.title.strip()]=items
    return out
if __name__=='__main__':
    for f in sys.argv[1:]:
        o=read_notif(f); print(f.split('/')[-1], {k:len(v) for k,v in o.items()})
