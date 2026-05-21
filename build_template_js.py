#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Converts template_data_raw.json → template_data.js
Deduplicates styles and resolves theme colors to RGB.
"""
import json, re
from pathlib import Path

RAW  = Path(__file__).parent / 'template_data_raw.json'
OUT  = Path(__file__).parent / 'template_data.js'

# Office default theme index → ARGB
THEME = {
    0: 'FFFFFFFF', 1: 'FF000000', 2: 'FFE7E6E6', 3: 'FF44546A',
    4: 'FF4472C4', 5: 'FFED7D31', 6: 'FFA5A5A5', 7: 'FFFFC000',
    8: 'FF5B9BD5', 9: 'FF70AD47',
}

def tint_rgb(hex6, tint):
    r,g,b = int(hex6[2:4],16), int(hex6[4:6],16), int(hex6[6:8],16)
    if tint > 0:
        r,g,b = (int(x+(255-x)*tint) for x in (r,g,b))
    else:
        r,g,b = (int(x*(1+tint)) for x in (r,g,b))
    return 'FF%02X%02X%02X' % (r,g,b)

def resolve_color(c):
    if not c: return None
    if 'auto' in c: return None
    if 'indexed' in c: return None
    if 'rgb' in c: rgb = c['rgb']; return rgb if len(rgb)==8 else 'FF'+rgb
    if 'theme' in c:
        base = THEME.get(c['theme'], 'FF000000')
        t = c.get('tint', 0)
        return tint_rgb(base, t) if abs(t) > 0.001 else base
    return None

def conv_font(f):
    if not f: return None
    o = {}
    if f.get('name'): o['n'] = f['name']
    if f.get('size'): o['s'] = f['size']
    if f.get('bold'): o['b'] = 1
    if f.get('italic'): o['i'] = 1
    c = resolve_color(f.get('color'))
    if c and c not in ('FF000000',): o['c'] = c  # skip default black
    return o or None

def conv_fill(f):
    if not f or f.get('type') != 'solid': return None
    c = resolve_color(f.get('fgColor'))
    if not c or c in ('FF000000', 'FFFFFFFF', '00000000'): return None
    return c  # just the ARGB string

def conv_border_side(s):
    if not s or not s.get('style'): return None
    st = s['style']
    c  = resolve_color(s.get('color'))
    if c and c not in ('FF000000',):
        return [st, c]
    return st  # just style string if default color

def conv_border(b):
    if not b: return None
    t = conv_border_side(b.get('top'))
    bt = conv_border_side(b.get('bottom'))
    l = conv_border_side(b.get('left'))
    r = conv_border_side(b.get('right'))
    if not any([t, bt, l, r]): return None
    return [t, bt, l, r]

def conv_align(a):
    if not a: return None
    h = a.get('horizontal')
    v = a.get('vertical')
    w = a.get('wrapText')
    if not h and not v and not w: return None
    o = {}
    if h: o['h'] = h
    if v: o['v'] = v
    if w: o['w'] = 1
    return o

def make_style(cell):
    f  = conv_font(cell.get('font'))
    fl = conv_fill(cell.get('fill'))
    bo = conv_border(cell.get('border'))
    al = conv_align(cell.get('align'))
    if not any([f, fl, bo, al]): return None
    return (
        json.dumps(f,  separators=(',',':'), ensure_ascii=False) if f  else 'null',
        json.dumps(fl, separators=(',',':'), ensure_ascii=False) if fl else 'null',
        json.dumps(bo, separators=(',',':'), ensure_ascii=False) if bo else 'null',
        json.dumps(al, separators=(',',':'), ensure_ascii=False) if al else 'null',
    )

def process_sheet(data):
    style_map = {}   # tuple → index
    styles    = []
    cells_out = []

    for cell in data.get('cells', []):
        v  = cell.get('v')       # None = cleared numeric
        sk = make_style(cell)
        if sk is None:
            si = -1
        elif sk in style_map:
            si = style_map[sk]
        else:
            si = len(styles)
            style_map[sk] = si
            styles.append(sk)

        if v is None and si == -1:
            continue  # fully empty, skip
        cells_out.append([cell['r'], cell['c'], v, si])

    # Convert styles from tuples → compact list [font,fill,border,align]
    styles_js = [[json.loads(s[0]), json.loads(s[1]),
                  json.loads(s[2]), json.loads(s[3])] for s in styles]

    # rowHeights: only keep rows TALLER than default (≈15); don't shrink rows
    rh = {k: v for k, v in data.get('rowHeights', {}).items()
          if v and v > 15.5}

    # colWidths: round to 2 decimals
    cw = {k: round(v, 2) for k, v in data.get('colWidths', {}).items() if v}

    return {
        'mc':  data.get('mergedCells', []),
        'cw':  cw,
        'rh':  rh,
        'fp':  data.get('freezePanes'),
        'tc':  data.get('tabColor'),
        'cells': cells_out,
        'styles': styles_js,
    }

def main():
    raw = json.loads(RAW.read_text('utf-8'))
    result = {}
    for key, sheet in raw.items():
        result[key] = process_sheet(sheet)
        c = len(result[key]['cells'])
        s = len(result[key]['styles'])
        print(f'  {key}: {c} cells, {s} styles')

    js = 'window.GL_TMPL=' + json.dumps(result, separators=(',', ':'),
                                         ensure_ascii=False) + ';'
    OUT.write_text(js, encoding='utf-8')
    kb = OUT.stat().st_size // 1024
    print(f'\n  → {OUT.name}  {kb} KB')

if __name__ == '__main__':
    main()
