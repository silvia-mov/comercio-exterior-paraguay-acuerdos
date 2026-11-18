"""Lee el texto (pdftotext -layout) del Arancel Nacional Vigente y devuelve filas.

Tipos de fila:
  - 'partida' / 'subpartida' / 'item': encabezados sin arancel (4, 6, 7 dígitos)
  - 'linea': código de 8 dígitos con AEC y ANV
  - 'abierta': código de 8 dígitos que solo trae AEC; su ANV está en sub-líneas
  - 'desglose': sub-línea nacional sin número dentro de un código 'abierta'
"""
import re, sys, json

NUM = r'\d+(?:,\d+)?'
LISTA = r'[A-Z][A-Za-z .\/0-9*]*?'
CODE = re.compile(r'^\s{0,6}(\d{2}\.\d{2}|\d{4}(?:\.\d{2}(?:\.\d{1,2})?)?)\.?\s+(\S.*)$')
TWO = re.compile(rf'^(.*?)\s{{2,}}({NUM})\s+({NUM})(?:\s+({LISTA}))?\s*$')
ONE = re.compile(rf'^(.*?)\s{{2,}}({NUM})\s*$')
ONLYLIST = re.compile(rf'^(.*?)\s{{2,}}((?:DEC|LNE|AUT|AZ|BK|BIT)[A-Za-z .\/0-9*]*?)\s*$')
SKIP = re.compile(r'Documento Base|Ajustado por los Decretos|Arancel Externo Común \(NCM|^\s*AEC\s+ANV|^\s*NCM\s+DESCRIPCI|Página \d+ de 455|^\s*%\s+%|^\s*Capítulo \d+\s*$')
NIVEL = {4: 'partida', 6: 'subpartida', 7: 'item', 8: 'linea'}


def num(s):
    return None if s is None else float(s.replace(',', '.'))


def parse(path):
    pages = open(path, encoding='utf-8').read().split('\f')
    rows = []
    cur = None      # fila de código abierta
    sub = None      # sub-línea en construcción
    state = None    # 'desc' (faltan tasas), 'subs' (leyendo desgloses), 'done'
    started = False

    def close_sub():
        nonlocal sub
        sub = None

    for pno, page in enumerate(pages, start=1):
        for line in page.split('\n'):
            if not line.strip() or SKIP.search(line):
                continue
            m = CODE.match(line)
            if m and (started or m.group(1) == '01.01'):
                started = True
                close_sub()
                code, rest = m.group(1), m.group(2)
                d = code.replace('.', '')
                cur = {'pag': pno, 'codigo_pdf': code, 'ncm': d, 'nivel': len(d),
                       'tipo': NIVEL[len(d)], 'desc': rest.strip(), 'desglose': None,
                       'aec': None, 'anv': None, 'lista': None}
                rows.append(cur)
                state = 'done'
                if len(d) == 8:
                    t, o, l = TWO.match(rest), ONE.match(rest), ONLYLIST.match(rest)
                    if t:
                        cur.update(desc=t.group(1).strip(), aec=num(t.group(2)), anv=num(t.group(3)), lista=t.group(4))
                    elif o:
                        cur.update(desc=o.group(1).strip(), aec=num(o.group(2)), tipo='abierta')
                        state = 'subs'
                    elif l:
                        cur.update(desc=l.group(1).strip(), lista=l.group(2))
                        state = 'desc'
                    else:
                        state = 'desc'
                else:
                    state = 'head'
                continue
            if not started or cur is None:
                continue
            s = line.strip()
            if state == 'head':
                # continuación de la descripción de un encabezado (texto sangrado que no es un "- grupo")
                if not s.startswith('-') and (len(line) - len(line.lstrip())) >= 8:
                    cur['desc'] += ' ' + s
                else:
                    state = 'done'
            elif state == 'desc':
                t, o = TWO.match(line), ONE.match(line)
                if t:
                    cur['desc'] += ' ' + t.group(1).strip()
                    cur.update(aec=num(t.group(2)), anv=num(t.group(3)))
                    if t.group(4):
                        cur['lista'] = t.group(4)
                    state = 'done'
                elif o:
                    cur['desc'] += ' ' + o.group(1).strip()
                    cur.update(aec=num(o.group(2)), tipo='abierta')
                    state = 'subs'
                else:
                    cur['desc'] += ' ' + s
            elif state == 'subs':
                t, o = TWO.match(line), ONE.match(line)
                if o and not t and sub is None and not any(r.get('padre') is cur for r in rows[-3:]) and cur['tipo'] == 'abierta' and cur['anv'] is None and not _has_subs(rows, cur):
                    # caso raro: el ANV del código quedó en la línea siguiente (ej. 2517.20.00)
                    cur['desc'] += ' ' + o.group(1).strip()
                    cur.update(anv=num(o.group(2)), tipo='linea')
                    state = 'done'
                    continue
                if s.startswith('-'):
                    state = 'done'; close_sub(); continue
                if sub is None:
                    sub = {'pag': pno, 'codigo_pdf': cur['codigo_pdf'], 'ncm': cur['ncm'], 'nivel': 8,
                           'tipo': 'desglose', 'desc': cur['desc'], 'desglose': '',
                           'aec': None, 'anv': None, 'lista': None, 'padre': cur}
                    rows.append(sub)
                if t:
                    sub['desglose'] = (sub['desglose'] + ' ' + t.group(1).strip()).strip()
                    sub.update(aec=num(t.group(2)), anv=num(t.group(3)), lista=t.group(4))
                    close_sub()
                else:
                    sub['desglose'] = (sub['desglose'] + ' ' + s).strip()
    for r in rows:
        r.pop('padre', None)
        r['desc'] = re.sub(r'\s+', ' ', r['desc']).strip()
        if r['desglose']:
            r['desglose'] = re.sub(r'\s+', ' ', r['desglose']).strip()
    return rows


def _has_subs(rows, cur):
    return any(r.get('padre') is cur for r in rows)


if __name__ == '__main__':
    rows = parse(sys.argv[1])
    json.dump(rows, open(sys.argv[2], 'w'), ensure_ascii=False, indent=0)
    from collections import Counter
    print(Counter(r['tipo'] for r in rows))
