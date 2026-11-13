#!/usr/bin/env python3
"""Genera index.html a partir de master_captura.xlsx y app_template.html.

Uso:  python build_captura.py <plantilla.xlsx> [salida.html]

Reejecutar cuando cambie la plantilla. Limpia del libro lo que no debe publicarse
(vínculos externos con sus datos en caché, rutas locales, autor, complementos) y
falla si alguna celda de captura ya no está vacía o cambió de estructura.
"""
import colorsys, json, re, sys, zipfile
from pathlib import Path
import openpyxl

SHEET = 'xl/worksheets/sheet1.xml'
KEEP = ['[Content_Types].xml', '_rels/.rels', 'docProps/app.xml', 'docProps/core.xml',
        'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', SHEET, 'xl/theme/theme1.xml',
        'xl/styles.xml', 'xl/sharedStrings.xml', 'xl/metadata.xml', 'xl/calcChain.xml']
# Equipo nuevo sin pruebas: claves que la macro del Motor NO importa (función ClaveExcluidaNuevos).
# Mantener igual a la macro: si cambia allá, cambia aquí y se vuelve a generar la página.
NUEVO_EXCL_CLAVES = [
    'cap_Origen', 'cap_Estado', 'cap_Val_Modelo', 'cap_Val_FechaCal', 'cap_Val_Orificio',
    'cap_Man_Modelo', 'cap_Man_FechaCal', 'cap_Man_Nota',
    'cap_FechaInsp', 'cap_FechaEmision', 'cap_HoraIni', 'cap_HoraFin',
    'cap_CondServicio', 'cap_Acabado', 'cap_ObsEND',
    'cap_TipoEND', 'cap_ResultadoEND', 'cap_FechaPruebaVal']
NUEVO_EXCL_PATRONES = [r'cap_E_.*', r'cap_UT\d\d', r'cap_PM_Z.*']   # Like "cap_E_*", "cap_UT##", "cap_PM_Z*"
FREEZE = 5  # filas inmovilizadas en la plantilla (pane ySplit)
KINDS = {4: 'matrix', 7: 'ut', 11: 'checklist', 12: 'zones', 13: 'evidence'}
ELEM = list('DEFGHI')


def sanitize(z):
    p = {n: z.read(n).decode('utf-8') for n in KEEP}
    ct = p['[Content_Types].xml']
    ct = re.sub(r'<Override PartName="/xl/(externalLinks|webextensions)/[^>]*/>', '', ct)
    p['[Content_Types].xml'] = ct
    p['_rels/.rels'] = re.sub(r'<Relationship [^>]*webextension[^>]*/>', '', p['_rels/.rels'])
    p['xl/_rels/workbook.xml.rels'] = re.sub(r'<Relationship [^>]*externalLink[^>]*/>', '', p['xl/_rels/workbook.xml.rels'])
    w = p['xl/workbook.xml']
    w = re.sub(r'<mc:AlternateContent.*?</mc:AlternateContent>', '', w, flags=re.S)
    w = re.sub(r'<externalReferences>.*?</externalReferences>', '', w, flags=re.S)
    w = re.sub(r'<definedName [^>]*>[^<]*\[\d+\][^<]*</definedName>', '', w)
    w, n = re.subn(r'<calcPr ', '<calcPr fullCalcOnLoad="1" ', w)
    assert n == 1, 'calcPr no encontrado'
    p['xl/workbook.xml'] = w
    c = p['docProps/core.xml']
    c = re.sub(r'<dc:creator>.*?</dc:creator>', '<dc:creator></dc:creator>', c)
    c = re.sub(r'<cp:lastModifiedBy>.*?</cp:lastModifiedBy>', '<cp:lastModifiedBy></cp:lastModifiedBy>', c)
    p['docProps/core.xml'] = c
    s = p[SHEET]
    assert not re.search(r'\[\d+\]', s), 'La hoja usa vínculos externos'
    s = s.replace('topLeftCell="A142"', 'topLeftCell="A6"')
    s = re.sub(r'(<selection pane="bottomLeft") activeCell="[^"]*" sqref="[^"]*"', r'\1 activeCell="D7" sqref="D7:I7"', s)
    p[SHEET] = s
    for name, text in p.items():
        for bad in ('Users\\', 'externalLink', 'webextension', 'H:\\'):
            assert bad not in text, f'{bad} sigue en {name}'
    return p


def dv_map(ws):
    out = {}
    for d in ws.data_validations.dataValidation:
        for rng in d.sqref.ranges:
            for row in ws[rng.coord] if ':' in rng.coord else ((ws[rng.coord],),):
                for c in row:
                    out[c.coordinate] = d
    return out


def spec(d):
    r = spec_kind(d)
    if d is not None:
        for k, v in (('pt', d.promptTitle), ('p', d.prompt), ('et', d.errorTitle), ('e', d.error)):
            if v:
                r[k] = v
    return r


def spec_kind(d):
    if d is None or d.type == 'custom':
        return {'k': 'text'}
    if d.type == 'list':
        o = d.formula1.strip('"').split(',')
        return {'k': 'list', 'o': o, 'free': d.errorStyle == 'warning'}
    if d.type in ('decimal', 'whole'):
        r = {'k': 'num' if d.type == 'decimal' else 'int', 'min': float(d.formula1)}
        r['op'] = {'greaterThan': 'gt', 'greaterThanOrEqual': 'ge', None: 'ge'}[d.operator]
        if d.formula2:
            r['max'] = float(d.formula2)
        return r
    if d.type == 'date':
        return {'k': 'date', 'min': float(d.formula1), 'max': float(d.formula2)}
    if d.type == 'time':
        return {'k': 'time'}
    raise ValueError(d.type)


def metadata(ws):
    dvs = dv_map(ws)
    secs, cur = [], None
    txt = lambda c: ws[c].value if isinstance(ws[c].value, str) else None
    for r in range(6, ws.max_row + 1):
        a, b = txt(f'A{r}'), txt(f'B{r}')
        m = re.match(r'(\d+)\. ', b or '')
        if not a and m:
            n = int(m.group(1))
            cur = {'n': n, 'hr': r, 'title': b[m.end():], 'kind': KINDS.get(n, 'scalar'), 'rows': []}
            secs.append(cur)
            continue
        if not cur:
            continue
        unit = txt(f'C{r}')
        base = {'r': r, 'label': b, 'unit': None if unit in ('—', 'lista', None) else unit,
                'note': txt(f'J{r}'), 'rule': txt(f'K{r}')}
        if a:
            if cur['kind'] == 'zones':
                cols = list('DEFGH')
            elif cur['kind'] in ('matrix', 'ut', 'evidence'):
                cols = ELEM
            else:
                cols = ['D']
            cells = {c: spec(dvs.get(f'{c}{r}')) for c in cols}
            for c in cols:
                if ws[f'{c}{r}'].number_format == '0.00':
                    cells[c]['nf'] = '0.00'
            cur['rows'].append({**base, 'key': a, 'cells': cells})
        elif cur['kind'] == 'evidence' and b:
            cur['rows'].append({**base, 'calc': True})
        elif cur['kind'] == 'ut' and txt(f'K{r}'):
            cur['head'] = base['rule']
    zones = next(s for s in secs if s['kind'] == 'zones')
    zones['cols'] = {c: ws[f'{c}142'].value for c in 'DEFGH'}
    return secs


def theme_colors(xml):
    names = re.search(r'<a:clrScheme.*?</a:clrScheme>', xml, re.S).group(0)
    vals = re.findall(r'<a:(dk1|lt1|dk2|lt2|accent\d|hlink|folHlink)>.*?(?:lastClr|srgbClr val)="([0-9A-Fa-f]{6})"', names, re.S)
    d = dict(vals)
    return [d[k] for k in ('lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink')]


def color(c, theme):
    if c is None:
        return None
    if c.type == 'rgb' and isinstance(c.rgb, str):
        return None if c.rgb == '00000000' else c.rgb[-6:].upper()
    if c.type == 'theme':
        base = theme[c.theme]
        t = c.tint or 0
        if t:
            r, g, b = (int(base[i:i + 2], 16) / 255 for i in (0, 2, 4))
            h, l, s = colorsys.rgb_to_hls(r, g, b)
            l = l * (1 + t) if t < 0 else l * (1 - t) + t
            base = ''.join(f'{round(v * 255):02X}' for v in colorsys.hls_to_rgb(h, l, s))
        return base.upper()
    return None


def side(b, theme):
    if not b or not b.style:
        return None
    w = {'thin': '1px', 'hair': '1px', 'dotted': '1px', 'dashed': '1px', 'medium': '2px', 'thick': '3px', 'double': '3px'}.get(b.style, '1px')
    kind = {'double': 'double', 'dotted': 'dotted', 'dashed': 'dashed', 'hair': 'dotted'}.get(b.style, 'solid')
    return f'{w} {kind} #{color(b.color, theme) or "000000"}'


def grid(ws, theme):
    widths, styles, index, rows = {}, [], {}, []
    for dim in ws.column_dimensions.values():
        for i in range(dim.min, dim.max + 1):
            widths[i] = dim.width
    default = ws.sheet_format.defaultRowHeight or 15
    merged = {}
    covered = set()
    for m in ws.merged_cells.ranges:
        merged[(m.min_row, m.min_col)] = m
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                if (r, c) != (m.min_row, m.min_col):
                    covered.add((r, c))
    for r in range(1, ws.max_row + 1):
        cells = []
        for ci in range(1, 13):
            if (r, ci) in covered:
                continue
            c = ws.cell(r, ci)
            m = merged.get((r, ci))
            right = ws.cell(r, m.max_col) if m else c
            fill = color(c.fill.fgColor, theme) if c.fill.patternType == 'solid' else None
            font = color(c.font.color, theme)
            al = c.alignment
            bd = {k: side(getattr(b, k), theme) for k, b in (('top', c.border), ('bottom', c.border), ('left', c.border), ('right', right.border))}
            css = [f'font-family:{c.font.name},Carlito,Arial,sans-serif', f'font-size:{c.font.sz}pt']
            if fill: css.append(f'background:#{fill}')
            css.append(f'color:#{font or "000000"}')
            if c.font.b: css.append('font-weight:700')
            if c.font.i: css.append('font-style:italic')
            if al.horizontal in ('left', 'center', 'right'): css.append(f'text-align:{al.horizontal}')
            css.append('vertical-align:' + {'top': 'top', 'center': 'middle', 'bottom': 'bottom'}.get(al.vertical, 'bottom'))
            css.append('white-space:' + ('normal' if al.wrap_text else 'nowrap'))
            css += [f'border-{k}:{v}' for k, v in bd.items() if v]
            key = ';'.join(css)
            if key not in index:
                index[key] = len(styles)
                styles.append({'css': key, 'fill': fill, 'font': font, 'b': bool(c.font.b)})
            x = {'c': c.column_letter, 's': index[key]}
            if m:
                x['n'] = m.max_col - m.min_col + 1
            if c.data_type == 'f' or (c.value is not None and not isinstance(c.value, (str, int, float))):
                x['f'] = 1
            elif c.value is not None:
                x['v'] = c.value
            cells.append(x)
        dim = ws.row_dimensions[r] if r in ws.row_dimensions else None
        rows.append({'r': r, 'h': (dim.height if dim and dim.height else default), 'c': cells})
    cols = [round((widths.get(i, ws.sheet_format.baseColWidth or 8.43)) * 7) for i in range(1, 13)]
    return {'cols': cols, 'styles': styles, 'rows': rows, 'freeze': FREEZE}


def col_num(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def cf_js(ws, theme):
    tok = re.compile(r'"[^"]*"|\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|[A-Z]+(?=\()|<>|<=|>=|=|<|>|\*|[(),]|\d+(?:\.\d+)?')
    ref = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')
    out = []
    for rng, rules in ws.conditional_formatting._cf_rules.items():
        sq = next(iter(rng.sqref.ranges))
        assert len(rng.sqref.ranges) == 1
        ax, ay = sq.min_col, sq.min_row
        def one(m):
            dc, c, dr, r = m.groups()
            x = str(col_num(c)) if dc else f'x+{col_num(c) - ax}'
            y = r if dr else f'y+{int(r) - ay}'
            return x, y
        for rule in rules:
            f = rule.formula[0]
            parts = tok.findall(f)
            assert ''.join(parts) == f, f
            js = []
            for t in parts:
                if t == '=': js.append('===')
                elif t == '<>': js.append('!==')
                elif t[0] == '"': js.append(json.dumps(t[1:-1], ensure_ascii=False))
                elif ref.fullmatch(t):
                    x, y = one(ref.fullmatch(t)); js.append(f'C({x},{y})')
                elif ':' in t:
                    a, b = t.split(':')
                    x1, y1 = one(ref.fullmatch(a)); x2, y2 = one(ref.fullmatch(b)); js.append(f'Rng({x1},{y1},{x2},{y2})')
                else: js.append(t)
            d = rule.dxf
            fill = color(d.fill.bgColor, theme) if d and d.fill else None
            font = color(d.font.color, theme) if d and d.font else None
            out.append((rule.priority, '{a:[%d,%d,%d,%d],p:%d,stop:%s,fill:%s,font:%s,f:function(x,y){return %s;}}' % (
                sq.min_col, sq.min_row, sq.max_col, sq.max_row, rule.priority, 'true' if rule.stopIfTrue else 'false',
                json.dumps(fill), json.dumps(font), ''.join(js))))
    return '[' + ',\n'.join(t for _, t in sorted(out, key=lambda v: v[0])) + ']'


def check_empty(sheet, secs):
    miss = []
    for s in secs:
        for row in s['rows']:
            for c in row.get('cells', {}):
                if not re.search(rf'<c r="{c}{row["r"]}" s="\d+"/>', sheet):
                    miss.append(f'{c}{row["r"]}')
    assert not miss, f'Celdas de captura no vacías o inexistentes: {miss}'


def main():
    src = Path(sys.argv[1])
    here = Path(__file__).parent
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else here.parent / 'index.html'
    wb = openpyxl.load_workbook(src)
    ws = wb.worksheets[0]
    secs = metadata(ws)
    with zipfile.ZipFile(src) as z:
        parts = sanitize(z)
    check_empty(parts[SHEET], secs)
    theme = theme_colors(parts['xl/theme/theme1.xml'])
    g = grid(ws, theme)
    css = '\n'.join(f'.s{i}{{{st["css"]}}}' for i, st in enumerate(g['styles']))
    meta = {'grid': g, 'version': ws['D3'].value, 'sheet': ws.title, 'title': ws['A1'].value,
            'rules': ws['K2'].value, 'sections': secs}
    keys = [r['key'] for sec in secs for r in sec['rows'] if 'key' in r]
    falta = [k for k in NUEVO_EXCL_CLAVES if k not in keys]
    assert not falta, f'Claves de equipo nuevo que ya no existen en la plantilla: {falta}'
    meta['nuevoExcl'] = [k for k in keys if k in NUEVO_EXCL_CLAVES or any(re.fullmatch(p, k) for p in NUEVO_EXCL_PATRONES)]
    print(f'Equipo nuevo: {len(keys) - len(meta["nuevoExcl"])} de {len(keys)} claves se capturan')
    dump = lambda o: json.dumps(o, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    html = (here / 'app_template.html').read_text(encoding='utf-8')
    html = html.replace('__SHEET_CSS__', css).replace('__CF__', cf_js(ws, theme))
    html = html.replace('__META__', dump(meta)).replace('__PARTS__', dump([[k, parts[k]] for k in KEEP]))
    out.write_text(html, encoding='utf-8', newline='\n')
    n = sum(len(r.get('cells', {})) for s in secs for r in s['rows'])
    print(f'{out}: {out.stat().st_size // 1024} KB - {len(secs)} secciones - {n} celdas de captura - hoja {ws.title}')


if __name__ == '__main__':
    main()
