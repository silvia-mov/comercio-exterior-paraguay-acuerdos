"""Arma historia de BK, BIT, LNE (notificaciones de Paraguay) y del Régimen de Materias Primas."""
import sys, glob, re, json, os, hashlib
import openpyxl
sys.path.insert(0, os.path.dirname(__file__))
from leer_notificaciones import read_notif

U = sys.argv[1]   # carpeta con los zip de las releases mercosur_listas ya descomprimidos (se buscan archivos en subcarpetas)
out = sys.argv[2]

# (etiqueta, fecha ISO, patrón de archivo). Se excluyen: la copia "al 31.07.2026" (idéntica a 31/01/2026)
# y la versión "copia" de 31/07/2022 (se usa el anexo original).
NOTIFS = [
    ('2018-S1', '2018-06-30', '2018-08-07 Nota MRE-Py 095*'),
    ('2020-01', '2020-01-31', 'Notificaci*31_01_2020-PY_rev.xlsx'),
    ('2020-07', '2020-07-31', '83612_*'),
    ('2021-01', '2021-01-31', '85815_*'),
    ('2021-07', '2021-07-31', 'MRE-PY_E267_2022*'),
    ('2022-01', '2022-01-31', 'MRE-PY_E975_2022*'),
    ('2022-07', '2022-07-31', 'MRE-PY_E1742_2022*_ANE.xlsx'),
    ('2023-07', '2023-07-31', '96835_*'),
    ('2023-12', '2023-12-31', '99502_*'),
    ('2025-01', '2025-01-31', '105305_*'),
    ('2025-07', '2025-07-31', '*31.07.2025*'),
    ('2026-01', '2026-01-31', '*31.01.2026*'),
]


def one(pat):
    # los zip repiten archivos: se aceptan varias copias si el contenido es idéntico
    m = sorted(glob.glob(os.path.join(U, '**', pat), recursive=True))
    distintos = {hashlib.md5(open(f, 'rb').read()).hexdigest() for f in m}
    assert len(distintos) == 1, (pat, m)
    return m[0]


def fname(f):
    # si el zip guardó el nombre en CP437 (ej. 'Notificación'), se corrige
    b = os.fsencode(os.path.basename(f))
    try:
        return b.decode('utf-8')
    except UnicodeDecodeError:
        return b.decode('cp437')


def clean(v):
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return None
    s = str(v).strip().replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return str(v).strip()


hist = {'notifs': [], 'BK': {}, 'BIT': {}, 'LNE': {}}
for lab, date, pat in NOTIFS:
    f = one(pat)
    hist['notifs'].append({'etiqueta': lab, 'fecha': date, 'archivo': fname(f)})
    data = read_notif(f)
    for L in ('BK', 'BIT', 'LNE'):
        key = next(k for k in data if k.strip().upper() == L)
        per = {}
        for c, anv, _ in data[key]:
            per.setdefault(c, []).append(clean(anv))
        for c, vals in per.items():
            uniq = []
            for v in vals:
                if v not in uniq:
                    uniq.append(v)
            hist[L].setdefault(c, {})[lab] = uniq[0] if len(uniq) == 1 else ' / '.join(
                ('%g' % v if isinstance(v, float) else str(v)) for v in uniq)

# Régimen de Materias Primas
RMP = [('2018', 'MRE-PY_E001_2021*(5).xlsx', 'AÑO 2018'), ('2019', 'MRE-PY_E001_2021*(5).xlsx', 'AÑO 2019'),
       ('2020-S1', 'MRE-PY_E001_2021*(5).xlsx', '1ER SEMESTRE 2020'),
       ('2023', 'RMP.LISTA_2023-2024(6).xlsx', 'AÑO 2023'), ('2024', 'RMP.LISTA_2023-2024(6).xlsx', 'AÑO 2024'),
       ('2025', 'RMP.LISTA_2025.xlsx', 'AÑO 2025')]
hist['rmp_years'] = [y for y, _, _ in RMP]
hist['RMP'] = {}
hist['rmp_files'] = {}
for y, pat, sheet in RMP:
    f = one(pat)
    hist['rmp_files'][y] = fname(f)
    ws = openpyxl.load_workbook(f, read_only=True, data_only=True)[sheet]
    for r in ws.iter_rows(values_only=True):
        if r and r[0] and len(re.sub(r'\D', '', str(r[0]))) == 8:
            hist['RMP'].setdefault(re.sub(r'\D', '', str(r[0])), []).append(y)

json.dump(hist, open(out, 'w'), ensure_ascii=False)
for L in ('BK', 'BIT', 'LNE'):
    print(L, len(hist[L]), 'códigos en alguna notificación')
print('RMP', len(hist['RMP']), {y: sum(y in v for v in hist['RMP'].values()) for y in hist['rmp_years']})
