"""Arma el Excel del Arancel Nacional Vigente frente al AEC.

Uso: python armar_excel.py arancel.json NCM_2026_11.xlsx salida.xlsx historia.json
"""
import json, re, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

rows_path, ncm_path, out, hist_path = sys.argv[1:5]
HIST = json.load(open(hist_path, encoding='utf-8'))
N_RMP = len(HIST['RMP'])
R = json.load(open(rows_path, encoding='utf-8'))

F = 'Arial'
HFILL = PatternFill('solid', fgColor='5E8F88')
HFONT = Font(name=F, bold=True, color='FFFFFF')
BASE = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
TITLE = Font(name=F, size=14, bold=True, color='5E8F88')
BLUE = Font(name=F, size=10, color='0000FF')
GREEN = Font(name=F, size=10, color='008000')
WRAP = Alignment(wrap_text=True, vertical='top')
THIN = Border(bottom=Side(style='thin', color='8DB2AC'))

wb = openpyxl.Workbook()

# ---------------- NCM noviembre 2026 (control) ----------------
src = openpyxl.load_workbook(ncm_path, read_only=True).active
ncm_rows = []
for r in src.iter_rows(min_row=4, values_only=True):
    if r[0] and len(str(r[0])) == 8:
        desc = re.sub(r'<[^>]+>', '', str(r[1] or '')).replace('\t', ' ')
        ncm_rows.append((str(r[0]), re.sub(r'\s+', ' ', desc).strip(), r[3]))
N_NCM = len(ncm_rows)

# ---------------- Arancel ----------------
ws = wb.active
ws.title = 'Arancel'
hdr = ['NCM', 'NCM (como en el PDF)', 'Capítulo', 'Descripción', 'Desglose nacional',
       'AEC %', 'ANV %', 'Lista', 'Diferencia ANV − AEC (puntos)', 'Diferencia relativa',
       'Sentido', 'Página PDF', 'AEC NCM nov-2026 %', 'Control con NCM nov-2026', 'Qué es la lista', 'Régimen de Materias Primas 2025']
ws.append(hdr)
lineas = [r for r in R if r['tipo'] in ('linea', 'desglose')]
for i, r in enumerate(lineas, start=2):
    ws.append([r['ncm'], r['codigo_pdf'], r['ncm'][:2], r['desc'], r['desglose'] or '',
               r['aec'], r['anv'], r['lista'] or '',
               f'=G{i}-F{i}',
               f'=IF(F{i}=0,"",I{i}/F{i})',
               f'=IF(I{i}<0,"Menor que el AEC",IF(I{i}>0,"Mayor que el AEC","Igual al AEC"))',
               r['pag'],
               f'=IFERROR(INDEX(NCM_nov2026!$C$2:$C${N_NCM+1},MATCH(A{i},NCM_nov2026!$A$2:$A${N_NCM+1},0)),"no existe")',
               f'=IF(M{i}="no existe","Código no existe en 2026",IF(M{i}=F{i},"AEC igual","AEC cambió"))',
               f'=IF(H{i}="","",IFERROR(INDEX(Listas!$B$2:$B$12,MATCH(SUBSTITUTE(H{i},"*","~*"),Listas!$A$2:$A$12,0)),"sin leyenda"))',
               f'=IF(COUNTIFS(Materias_primas!$A$2:$A${N_RMP+1},A{i},Materias_primas!$H$2:$H${N_RMP+1},"Sí")>0,"Sí","No")'])
NL = len(lineas) + 1
for c, cell in enumerate(ws[1], start=1):
    cell.font, cell.fill, cell.alignment = HFONT, HFILL, Alignment(wrap_text=True, vertical='center')
for row in ws.iter_rows(min_row=2, max_row=NL):
    for cell in row:
        cell.font = BASE
    for k in (5, 6):  # F, G: datos del PDF
        row[k].font = BLUE
    row[12].font = GREEN
    row[8].number_format = '+0.0;-0.0;0'
    row[9].number_format = '+0.0%;-0.0%;0%'
    row[5].number_format = row[6].number_format = row[12].number_format = '0.0'
widths = [11, 13, 9, 60, 32, 8, 8, 12, 14, 12, 18, 9, 12, 22, 34, 14]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.row_dimensions[1].height = 45
ws.freeze_panes = 'B2'
ws.auto_filter.ref = f'A1:P{NL}'
ws.conditional_formatting.add(f'K2:K{NL}', FormulaRule(formula=[f'$I2<0'], fill=PatternFill('solid', fgColor='B7D0CC')))
ws.conditional_formatting.add(f'K2:K{NL}', FormulaRule(formula=[f'$I2>0'], fill=PatternFill('solid', fgColor='E6DAE7')))
ws.conditional_formatting.add(f'N2:N{NL}', FormulaRule(formula=[f'$N2<>"AEC igual"'], fill=PatternFill('solid', fgColor='AF8EB1')))

# ---------------- Estructura completa ----------------
we = wb.create_sheet('Estructura')
we.append(['Orden', 'NCM', 'NCM (como en el PDF)', 'Nivel', 'Tipo de fila', 'Descripción',
           'Desglose nacional', 'AEC %', 'ANV %', 'Lista', 'Página PDF'])
TIPO = {'partida': 'Partida (4 dígitos)', 'subpartida': 'Subpartida (6 dígitos)', 'item': 'Ítem (7 dígitos)',
        'linea': 'Código (8 dígitos)', 'abierta': 'Código con desglose nacional', 'desglose': 'Desglose nacional'}
for i, r in enumerate(R, start=1):
    we.append([i, r['ncm'], r['codigo_pdf'], r['nivel'], TIPO[r['tipo']], r['desc'],
               r['desglose'] or '', r['aec'], r['anv'], r['lista'] or '', r['pag']])
for cell in we[1]:
    cell.font, cell.fill, cell.alignment = HFONT, HFILL, Alignment(wrap_text=True, vertical='center')
for row in we.iter_rows(min_row=2):
    lvl = row[3].value
    for cell in row:
        cell.font = BOLD if lvl in (4,) else BASE
    row[7].number_format = row[8].number_format = '0.0'
for i, w in enumerate([8, 11, 13, 7, 26, 70, 32, 8, 8, 12, 9], start=1):
    we.column_dimensions[get_column_letter(i)].width = w
we.freeze_panes = 'C2'
we.auto_filter.ref = f'A1:K{len(R)+1}'

# ---------------- NCM nov 2026 ----------------
wn = wb.create_sheet('NCM_nov2026')
wn.append(['NCM', 'Descripción (SAREM)', 'AEC %'])
for t in ncm_rows:
    wn.append(list(t))
for cell in wn[1]:
    cell.font, cell.fill = HFONT, HFILL
for row in wn.iter_rows(min_row=2):
    for cell in row:
        cell.font = BASE
    row[2].number_format = '0.0'
wn.column_dimensions['A'].width = 11
wn.column_dimensions['B'].width = 80
wn.column_dimensions['C'].width = 8
wn.freeze_panes = 'A2'

# ---------------- Resumen ----------------
wr = wb.create_sheet('Resumen', 0)
A = f'Arancel!$I$2:$I${NL}'
K = f'Arancel!$K$2:$K${NL}'
H = f'Arancel!$H$2:$H${NL}'
Nn = f'Arancel!$N$2:$N${NL}'
wr['A1'] = 'Arancel Nacional Vigente (octubre 2025) frente al Arancel Externo Común'
wr['A1'].font = TITLE
wr['A2'] = 'Diferencia = ANV − AEC, en puntos porcentuales. Negativo: Paraguay cobra menos que el AEC.'
wr['A2'].font = Font(name=F, size=10, italic=True)

wr['A4'] = 'Por sentido'
for c, h in zip('ABCD', ['Sentido', 'Líneas', '% de las líneas', 'Diferencia promedio (puntos)']):
    wr[f'{c}5'] = h
r0 = 6
for j, s in enumerate(['Menor que el AEC', 'Igual al AEC', 'Mayor que el AEC']):
    r = r0 + j
    wr[f'A{r}'] = s
    wr[f'B{r}'] = f'=COUNTIF({K},A{r})'
    wr[f'C{r}'] = f'=B{r}/$B${r0+3}'
    wr[f'D{r}'] = f'=IFERROR(AVERAGEIF({K},A{r},{A}),0)'
rt = r0 + 3
wr[f'A{rt}'] = 'Total'
wr[f'B{rt}'] = f'=SUM(B{r0}:B{rt-1})'
wr[f'C{rt}'] = f'=B{rt}/$B${rt}'
wr[f'D{rt}'] = f'=AVERAGE({A})'
for c in 'ABCD':
    wr[f'{c}{rt}'].font = BOLD

wr['A12'] = 'Por lista (columna "Lista" del PDF)'
for c, h in zip('ABCDE', ['Lista', 'Líneas', 'Menor que el AEC', 'Mayor que el AEC', 'Diferencia promedio (puntos)']):
    wr[f'{c}13'] = h
listas = ['(sin lista)'] + sorted({r['lista'] for r in lineas if r['lista']}, key=lambda x: -sum(1 for q in lineas if q['lista'] == x))
for j, l in enumerate(listas):
    r = 14 + j
    crit = '""' if l == '(sin lista)' else f'SUBSTITUTE(A{r},"*","~*")'
    wr[f'A{r}'] = l
    wr[f'B{r}'] = f'=COUNTIF({H},{crit})'
    wr[f'C{r}'] = f'=COUNTIFS({H},{crit},{K},"Menor que el AEC")'
    wr[f'D{r}'] = f'=COUNTIFS({H},{crit},{K},"Mayor que el AEC")'
    wr[f'E{r}'] = f'=IFERROR(AVERAGEIF({H},{crit},{A}),0)'
rl = 14 + len(listas)
wr[f'A{rl}'] = 'Total'
for c in 'BCD':
    wr[f'{c}{rl}'] = f'=SUM({c}14:{c}{rl-1})'
    wr[f'A{rl+1}'] = 'Control: el total debe ser igual al total de líneas'
    wr[f'B{rl+1}'] = f'=IF(B{rl}=B{rt},"OK","Revisar")'
    wr[f'{c}{rl}'].font = BOLD
wr[f'A{rl}'].font = BOLD

rc = rl + 3
wr[f'A{rc}'] = 'Control con la NCM de noviembre 2026 (SAREM)'
for c, h in zip('AB', ['Resultado', 'Líneas']):
    wr[f'{c}{rc+1}'] = h
for j, s in enumerate(['AEC igual', 'AEC cambió', 'Código no existe en 2026']):
    r = rc + 2 + j
    wr[f'A{r}'] = s
    wr[f'B{r}'] = f'=COUNTIF({Nn},A{r})'

for rr in (5, 13, rc + 1):
    for cell in wr[rr]:
        if cell.value is not None:
            cell.font, cell.fill = HFONT, HFILL
            cell.alignment = Alignment(wrap_text=True, vertical='center')
for rr in (4, 12, rc):
    wr[f'A{rr}'].font = Font(name=F, size=11, bold=True, color='774D78')
for row in wr.iter_rows(min_row=6):
    for cell in row:
        if cell.font is None or (cell.font.name != F):
            cell.font = BASE
        if cell.column_letter in 'D' and cell.row < rl and cell.row >= r0:
            pass
for r in range(r0, rt + 1):
    wr[f'C{r}'].number_format = '0.0%'
    wr[f'D{r}'].number_format = '+0.00;-0.00;0'
for r in range(14, rl):
    wr[f'E{r}'].number_format = '+0.00;-0.00;0'
for col, w in zip('ABCDE', [34, 12, 16, 16, 18]):
    wr.column_dimensions[col].width = w

# ---------------- Léeme ----------------
wl = wb.create_sheet('Leeme', 0)
notas = [
    ('Qué es este archivo', None),
    ('1.', 'El "Arancel Nacional Vigente (ANV) actualizado a octubre 2025" de Paraguay, pasado del PDF a Excel.'),
    ('2.', 'Fuente: PDF de la Dirección General de Política Comercial e Integración (Ministerio de Economía y Finanzas), 455 páginas, documento base Anexo del Decreto N° 6897/2022, ajustado por los Decretos 7031/2022, 8662/2022, 556/2023, 8767/2023, 2530/2024, 3379/2025 y 4754/2025. NCM 2022. El propio PDF aclara que es un documento de trabajo y no tiene fuerza legal.'),
    ('3.', 'Archivo: silvia-mov/varios, "Arancel Nacional Vigente_octubre 2025_1.pdf".'),
    ('Hojas', None),
    ('Resumen', 'Cuántas líneas tienen un ANV menor, igual o mayor que el AEC, y la diferencia promedio. También por lista.'),
    ('Arancel', 'Una fila por línea arancelaria (10.459 códigos + 103 desgloses nacionales). Es la hoja para filtrar y analizar.'),
    ('Listas', 'Qué significa cada valor de la columna Lista, con su fuente.'),
    ('Historia_BK, _BIT, _LNE', 'Qué códigos estuvieron en cada lista y con qué arancel, según las notificaciones semestrales de Paraguay al MERCOSUR (2018 a enero 2026).'),
    ('Materias_primas', 'Códigos del Régimen de Materias Primas de Paraguay por año (2018, 2019, 1er semestre 2020, 2023, 2024 y 2025). Dentro del régimen el arancel es 0 %.'),
    ('Estructura', 'Todas las filas del PDF en orden (partidas, subpartidas, ítems, códigos y desgloses), para leer la jerarquía.'),
    ('NCM_nov2026', 'La NCM de noviembre 2026 de SAREM (archivo NCM_2026_11.xlsx del repositorio comercio-exterior-paraguay-acuerdos), solo para control.'),
    ('Columnas nuevas (hoja Arancel)', None),
    ('Régimen de Materias Primas 2025', 'Sí, si el código está en la lista 2025 del régimen. Ojo: el PDF muestra el ANV normal para esos códigos, pero el importador que usa el régimen paga 0 %.'),
    ('Diferencia ANV − AEC', 'En puntos porcentuales. Ejemplo: AEC 16 y ANV 14,4 da −1,6. Negativo = Paraguay cobra menos que el AEC.'),
    ('Diferencia relativa', 'La diferencia dividida por el AEC. Ejemplo: −1,6 / 16 = −10 %. Queda vacía cuando el AEC es 0.'),
    ('Sentido', 'Menor, igual o mayor que el AEC.'),
    ('AEC NCM nov-2026 y Control', 'Busca el mismo código en la NCM de noviembre 2026 y avisa si el AEC cambió o si el código ya no existe.'),
    ('Colores', 'Azul: números copiados del PDF. Verde: dato traído de otra hoja. Negro: fórmulas.'),
    ('Cosas a tener en cuenta', None),
    ('Desgloses nacionales', '43 códigos no traen un único ANV: el PDF los abre en sub-líneas sin número (por ejemplo 1902.30.00 "Fideos de arroz" con ANV 2 y "Las demás" con ANV 16). Cada sub-línea es una fila, con el texto en la columna "Desglose nacional" y el mismo código NCM. Al contar productos por código, tenerlo en cuenta.'),
    ('8703.10.00', 'En el desglose "Propulsados por motor eléctrico..." el PDF pone AEC 0, aunque el código tiene AEC 20. Se copió tal cual.'),
    ('Columna Lista', 'Se copió tal como está en el PDF. La explicación de cada lista está en la hoja Listas y en la columna "Qué es la lista" de la hoja Arancel.'),
    ('Posibles errores del PDF', 'De las 1.095 líneas marcadas DEC. 08/22, 1.089 cumplen ANV = AEC × 0,9. Las otras 6 no: 4802.57.10 (AEC 15 en el PDF; la NCM nov-2026 dice 16, y 16 × 0,9 = 14,4), 8481.20.11 (ANV 14 mayor que el AEC 12,6) y 8539.52.00, 8539.90.10, 8539.90.20 y 8540.60.10 (ANV igual al AEC). Conviene revisarlas con el decreto.'),
    ('Encabezados de grupo', 'Las líneas de agrupación sin código (por ejemplo "- Peces ornamentales:") no se copiaron; sí están todas las partidas, subpartidas e ítems con código.'),
    ('Control con las notificaciones', 'Las listas BK (1.257), BIT (361) y LNE (715 códigos) notificadas al 31/07/2025 coinciden código por código y arancel por arancel con el PDF. LNE y LNE* del PDF son la misma lista notificada.'),
    ('Materias primas 2025 sin código en el PDF', 'De los 683 códigos de la lista 2025, 679 están en el PDF. Los otros 4 no existen en la NCM del PDF: 2712.10.00 (existe en la NCM nov-2026), 3906.90.49, 3907.29.90 y 8544.19.10 (en el PDF esas posiciones están abiertas con otros números). Hay que revisarlos con la correlación de la NCM.'),
    ('Notificaciones usadas', 'Archivos de las releases mercosur_listas a mercosur_listas_5 del repositorio silvia-mov/varios. Se dejaron afuera: el archivo "al 31.07.2026" (adentro dice 31/01/2026 y es idéntico al de enero 2026), la versión "copia" de julio 2022 (se usó el anexo original) y las listas de Uruguay del zip de regímenes especiales. Hasta julio 2022 las notificaciones usan la NCM 2017.'),
    ('Control de lectura', 'De 10.496 códigos que están en el PDF y en la NCM nov-2026, 10.486 tienen el mismo AEC. Los 10 distintos se revisaron en el PDF y son cambios reales entre los dos documentos. 6 códigos del PDF ya no existen en 2026 (se dividieron).'),
]
wl['A1'] = 'Léeme'
wl['A1'].font = TITLE
r = 3
for k, v in notas:
    if v is None:
        r += 1 if r > 3 else 0
        wl[f'A{r}'] = k
        wl[f'A{r}'].font = Font(name=F, size=11, bold=True, color='774D78')
    else:
        wl[f'A{r}'], wl[f'B{r}'] = k, v
        wl[f'A{r}'].font, wl[f'B{r}'].font = BOLD, BASE
        wl[f'A{r}'].alignment = wl[f'B{r}'].alignment = WRAP
    r += 1
wl.column_dimensions['A'].width = 28
wl.column_dimensions['B'].width = 110

wlst = wb.create_sheet('Listas', 2)
wlst.append(['Lista', 'Qué es', 'Fuente', 'Control con los datos del PDF'])
LEG = [
 ('BK', 'Bienes de capital', 'Leyenda del propio PDF (página 2).', ''),
 ('BIT', 'Bienes de informática y telecomunicaciones', 'Leyenda del propio PDF (página 2).', ''),
 ('LNE', 'Lista Nacional de Excepciones al AEC', 'Indicado por la autora. El PDF no trae leyenda.', ''),
 ('LNE*', 'Lista Nacional de Excepciones al AEC (con asterisco)', 'Indicado por la autora. El PDF no explica qué marca el asterisco.', ''),
 ('AUT', 'Régimen automotor: sector no incorporado al AEC', 'Indicado por la autora. La Dec. CMC 08/22, Art. 4, dice que el arancel del sector automotor se rige por los acuerdos bilaterales o, si no hay, por la legislación interna, hasta que el sector se incorpore al MERCOSUR.', ''),
 ('AZ*', 'Sector azucarero: no armonizado en el MERCOSUR', 'Indicado por la autora. El PDF no trae leyenda.', 'Las 5 líneas son de la partida 17.01, con AEC 16 y ANV 30.'),
 ('DEC. 08/22', 'Decisión CMC 08/22, Art. 2: permite bajar 10 % el arancel respecto del AEC vigente al 01/VII/2022 para los códigos del Anexo IV (Paraguay)', 'normativa-mercosur: decisiones/2022/DEC_008-2022_ES_fe-de-erratas_arancel-externo-comun.txt. Incorporada por el Decreto 8767/2023 (incorporacion_paraguay/2022/decreto-8767-2023.txt). El Anexo IV (lista de Paraguay) no está en el repositorio.', '1.089 de 1.095 líneas cumplen ANV = AEC × 0,9 (ver Leeme).'),
 ('DEC.37/07', 'Decisión CMC 37/07: fija el AEC de tejidos, confecciones y calzados; Art. 3 permite a Paraguay y Uruguay mantener su arancel nacional en tejidos y confecciones', 'normativa-mercosur: decisiones/2007/DEC_037-2007_ES_texto_modificacion-del-arancel-externo-comun.txt', 'Casos más comunes: AEC 35 y ANV 20; AEC 26 y ANV 18.'),
 ('DEC. 29/15', 'Decisión CMC 29/15: AEC de 35 % para 2008.70 (duraznos) hasta el 31/12/2023; Art. 2 permite a Paraguay mantener su arancel', 'normativa-mercosur: decisiones/2015/DEC_029-2015_ES_texto_arancel-externo-comun.txt', 'En el PDF las 3 líneas tienen AEC 14 y ANV 14 (la vigencia de la decisión terminó en 2023).'),
 ('DEC. 30/15', 'Decisión CMC 30/15: AEC de 28 % para lácteos (04.02, 04.04, 04.06) hasta el 31/12/2023; Art. 2 permite a Paraguay mantener su arancel', 'normativa-mercosur: decisiones/2015/DEC_030-2015_ES_texto_arancel-externo-comun.txt', 'En el PDF las 10 líneas tienen ANV igual al AEC (16 o 14; la vigencia de la decisión terminó en 2023).'),
]
for t in LEG:
    wlst.append(list(t))
wlst['A13'] = 'Nota: la Dec. CMC 08/22, Art. 3, excluye de la rebaja del 10 % los códigos de las Decisiones 37/07, 26/09, 27/09, 28/15, 29/15 y 30/15.'
wlst['A13'].font = Font(name=F, size=10, italic=True)
for cell in wlst[1]:
    cell.font, cell.fill = HFONT, HFILL
for row in wlst.iter_rows(min_row=2, max_row=11):
    for cell in row:
        cell.font, cell.alignment = BASE, WRAP
    row[0].font = BOLD
for col, w in zip('ABCD', [13, 55, 70, 45]):
    wlst.column_dimensions[col].width = w
wlst.freeze_panes = 'A2'

# ---------------- Historia de listas ----------------
desc_pdf = {}
for r in R:
    if r['nivel'] == 8 and r['ncm'] not in desc_pdf:
        desc_pdf[r['ncm']] = r['desc']
NOT = HIST['notifs']
def fmt8(c):
    return f'{c[:4]}.{c[4:6]}.{c[6:]}'
for L, crit in (('BK', ['"BK"']), ('BIT', ['"BIT"']), ('LNE', ['"LNE"', '"LNE~*"'])):
    wh = wb.create_sheet(f'Historia_{L}')
    head = ['NCM', 'NCM (formato)', 'Descripción en el PDF oct-2025'] + [f"{n['etiqueta']}" for n in NOT] + ['Notificaciones en que aparece', f'En {L} en el PDF oct-2025']
    wh.append(head)
    codes = sorted(HIST[L])
    first, last = 4, 3 + len(NOT)
    cf, cl = get_column_letter(first), get_column_letter(last)
    for i, c in enumerate(codes, start=2):
        vals = [HIST[L][c].get(n['etiqueta']) for n in NOT]
        cnt = '+'.join(f'COUNTIFS(Arancel!$A$2:$A${NL},A{i},Arancel!$H$2:$H${NL},{k})' for k in crit)
        wh.append([c, fmt8(c), desc_pdf.get(c, '(no está en el PDF: código de otra versión de la NCM)')] + vals +
                  [f'=COUNTA({cf}{i}:{cl}{i})', f'=IF({cnt}>0,"Sí","No")'])
    for cell in wh[1]:
        cell.font, cell.fill, cell.alignment = HFONT, HFILL, Alignment(wrap_text=True, vertical='center', horizontal='center')
    for row in wh.iter_rows(min_row=2):
        for cell in row:
            cell.font = BASE
        for cell in row[first-1:last]:
            cell.font = BLUE
            cell.number_format = '0.0'
            cell.alignment = Alignment(horizontal='center')
    wh.column_dimensions['A'].width = 11; wh.column_dimensions['B'].width = 12; wh.column_dimensions['C'].width = 50
    for k in range(first, last + 1):
        wh.column_dimensions[get_column_letter(k)].width = 8.5
    wh.column_dimensions[get_column_letter(last + 1)].width = 14
    wh.column_dimensions[get_column_letter(last + 2)].width = 14
    wh.row_dimensions[1].height = 42
    wh.freeze_panes = 'D2'
    wh.auto_filter.ref = f'A1:{get_column_letter(last+2)}{len(codes)+1}'
    n = len(codes) + 2
    wh[f'A{n+1}'] = 'Fila de totales: cantidad de códigos en cada notificación.'
    wh[f'A{n+1}'].font = Font(name=F, size=9, italic=True)
    wh[f'C{n}'] = 'Total de códigos'
    wh[f'C{n}'].font = BOLD
    for k in range(first, last + 1):
        col = get_column_letter(k)
        wh[f'{col}{n}'] = f'=COUNTA({col}2:{col}{n-1})'
        wh[f'{col}{n}'].font = BOLD
        wh[f'{col}{n}'].alignment = Alignment(horizontal='center')
    wh[f'A{n+2}'] = 'Archivos: ' + ' | '.join(f"{x['etiqueta']}: {x['archivo']}" for x in NOT)
    wh[f'A{n+2}'].font = Font(name=F, size=9, italic=True)

# ---------------- Materias primas ----------------
wm = wb.create_sheet('Materias_primas')
YRS = HIST['rmp_years']
wm.append(['NCM', 'NCM (formato)'] + YRS + ['Años en la lista', 'Descripción en el PDF oct-2025', 'ANV en el PDF oct-2025 %'])
rcodes = sorted(HIST['RMP'])
for i, c in enumerate(rcodes, start=2):
    wm.append([c, fmt8(c)] + ['Sí' if y in HIST['RMP'][c] else '' for y in YRS] +
              [f'=COUNTIF(C{i}:H{i},"Sí")', desc_pdf.get(c, '(no está en el PDF: código de otra versión de la NCM)'),
               f'=IFERROR(INDEX(Arancel!$G$2:$G${NL},MATCH(A{i},Arancel!$A$2:$A${NL},0)),"no está")'])
for cell in wm[1]:
    cell.font, cell.fill, cell.alignment = HFONT, HFILL, Alignment(wrap_text=True, vertical='center', horizontal='center')
for row in wm.iter_rows(min_row=2):
    for cell in row:
        cell.font = BASE
    for cell in row[2:8]:
        cell.font = BLUE; cell.alignment = Alignment(horizontal='center')
    row[10].number_format = '0.0'; row[10].font = GREEN
for col, w in zip('ABCDEFGHIJK', [11, 12, 7, 7, 8, 7, 7, 7, 10, 55, 12]):
    wm.column_dimensions[col].width = w
wm.row_dimensions[1].height = 32
wm.freeze_panes = 'C2'
wm.auto_filter.ref = f'A1:K{len(rcodes)+1}'
n = len(rcodes) + 2
wm[f'B{n}'] = 'Total'; wm[f'B{n}'].font = BOLD
for col in 'CDEFGH':
    wm[f'{col}{n}'] = f'=COUNTIF({col}2:{col}{n-1},"Sí")'; wm[f'{col}{n}'].font = BOLD
wm[f'A{n+1}'] = 'Fuente: ' + ' | '.join(f'{y}: {f}' for y, f in HIST['rmp_files'].items()) + '. Hasta 2020 los códigos son de la NCM 2017. Dentro del régimen el arancel es 0 %.'
wm[f'A{n+1}'].font = Font(name=F, size=9, italic=True)

# ---------------- Resumen: materias primas ----------------
rm = wr.max_row + 2
wr[f'A{rm}'] = 'Régimen de Materias Primas 2025 (arancel 0 % dentro del régimen)'
wr[f'A{rm}'].font = Font(name=F, size=11, bold=True, color='774D78')
for c, h in zip('ABC', ['Sentido (ANV del PDF frente al AEC)', 'Líneas en el régimen', 'ANV promedio en el PDF %']):
    wr[f'{c}{rm+1}'] = h
    wr[f'{c}{rm+1}'].font, wr[f'{c}{rm+1}'].fill = HFONT, HFILL
    wr[f'{c}{rm+1}'].alignment = Alignment(wrap_text=True, vertical='center')
P = f'Arancel!$P$2:$P${NL}'; G = f'Arancel!$G$2:$G${NL}'
for j, s_ in enumerate(['Menor que el AEC', 'Igual al AEC', 'Mayor que el AEC']):
    r = rm + 2 + j
    wr[f'A{r}'] = s_
    wr[f'B{r}'] = f'=COUNTIFS({K},A{r},{P},"Sí")'
    wr[f'C{r}'] = f'=IFERROR(AVERAGEIFS({G},{K},A{r},{P},"Sí"),0)'
    wr[f'C{r}'].number_format = '0.0'
    for c in 'ABC':
        wr[f'{c}{r}'].font = BASE
r = rm + 5
wr[f'A{r}'] = 'Total'; wr[f'B{r}'] = f'=SUM(B{rm+2}:B{rm+4})'; wr[f'C{r}'] = f'=IFERROR(AVERAGEIFS({G},{P},"Sí"),0)'
wr[f'C{r}'].number_format = '0.0'
for c in 'ABC':
    wr[f'{c}{r}'].font = BOLD

ORDER = ['Leeme', 'Resumen', 'Listas', 'Arancel', 'Historia_BK', 'Historia_BIT', 'Historia_LNE', 'Materias_primas', 'Estructura', 'NCM_nov2026']
wb._sheets = [wb[n] for n in ORDER]
wb.active = 0
wb.save(out)
print('ok', NL - 1, 'lineas', len(R), 'filas estructura', N_NCM, 'ncm')
