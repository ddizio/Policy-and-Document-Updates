#!/usr/bin/env python3
"""Convert Vantage policy/procedure .docx files into the Procedure and Policy
Template Example look: logo header table, green section banners, Century Gothic
body, page-numbered footer."""
import os, re, sys, glob, zipfile, datetime, shutil
from xml.etree import ElementTree as ET

W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
w  = '{%s}' % W
SRC_DIR  = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TPL_DOCX = os.path.join(SRC_DIR, 'Procedure and Policy Template Example.docx')
TPL_DIR  = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.template_unpacked')
CONTENT_W = 10800
GREEN  = '098164'
BODY_FONT = 'Century Gothic'
HEAD_FONT = 'Aptos'
TODAY = datetime.date(2026, 8, 18).strftime('%m/%d/%Y')
COMPANY_RE = re.compile(r'^vantage\s+specialty\s+chemicals[\s,.]*$', re.I)

def esc(s):
    s = ''.join(ch for ch in s if ch >= ' ' or ch in '\t')
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

# ---------------------------------------------------------------- source parse
def run_items(p):
    """Flatten a paragraph into [(text, bold, italic)] preserving order."""
    out = []
    for r in p.iter(w+'r'):
        rPr = r.find(w+'rPr')
        b = rPr is not None and rPr.find(w+'b') is not None
        i = rPr is not None and rPr.find(w+'i') is not None
        t = ''
        for n in r:
            if n.tag == w+'t':   t += n.text or ''
            elif n.tag == w+'tab': t += ' '
            elif n.tag == w+'br':  t += ' '
        if t:
            out.append([t, b, i])
    # merge adjacent runs with identical formatting
    merged = []
    for it in out:
        if merged and merged[-1][1] == it[1] and merged[-1][2] == it[2]:
            merged[-1][0] += it[0]
        else:
            merged.append(list(it))
    return merged

def p_info(p):
    pPr = p.find(w+'pPr')
    style, numid, ilvl = '', None, 0
    if pPr is not None:
        s = pPr.find(w+'pStyle')
        if s is not None: style = s.get(w+'val') or ''
        np = pPr.find(w+'numPr')
        if np is not None:
            n = np.find(w+'numId'); l = np.find(w+'ilvl')
            if n is not None: numid = n.get(w+'val')
            if l is not None: ilvl = int(l.get(w+'val') or 0)
    sz, color, bold = None, '', False
    r = p.find(w+'r')
    if r is not None:
        rPr = r.find(w+'rPr')
        if rPr is not None:
            e = rPr.find(w+'sz')
            if e is not None: sz = int(e.get(w+'val'))
            c = rPr.find(w+'color')
            if c is not None: color = (c.get(w+'val') or '').upper()
            bold = rPr.find(w+'b') is not None
    if sz is None and pPr is not None:
        rPr = pPr.find(w+'rPr')
        if rPr is not None:
            e = rPr.find(w+'sz')
            if e is not None: sz = int(e.get(w+'val'))
    runs = run_items(p)
    text = ''.join(x[0] for x in runs).strip()
    all_bold = bool(runs) and all(x[1] for x in runs if x[0].strip())
    return dict(kind='p', style=style, numid=numid, ilvl=ilvl, sz=sz or 22,
                color=color, bold=bold or all_bold, all_bold=all_bold,
                runs=runs, text=text)

def tbl_info(tbl):
    grid = [int(g.get(w+'w') or 0) for g in tbl.findall(w+'tblGrid/'+w+'gridCol')]
    rows = []
    for tr in tbl.findall(w+'tr'):
        cells = []
        for tc in tr.findall(w+'tc'):
            tcPr = tc.find(w+'tcPr')
            span = 1; shd = ''
            if tcPr is not None:
                gs = tcPr.find(w+'gridSpan')
                if gs is not None: span = int(gs.get(w+'val') or 1)
                sh = tcPr.find(w+'shd')
                if sh is not None: shd = (sh.get(w+'fill') or '').upper()
            paras = [p_info(p) for p in tc.findall(w+'p')]
            cells.append(dict(span=span, shd=shd, paras=paras))
        if cells: rows.append(cells)
    return dict(kind='tbl', grid=grid, rows=rows)

def load_numfmt(z):
    try: nd = z.read('word/numbering.xml').decode('utf8', 'ignore')
    except KeyError: return {}
    absfmt = {}
    for m in re.finditer(r'<w:abstractNum w:abstractNumId="(\d+)"(.*?)</w:abstractNum>', nd, re.S):
        l0 = re.search(r'<w:lvl w:ilvl="0".*?</w:lvl>', m.group(2), re.S)
        fmt = 'bullet'
        if l0:
            f = re.search(r'<w:numFmt w:val="([^"]+)"', l0.group(0))
            if f: fmt = f.group(1)
        absfmt[m.group(1)] = fmt
    out = {}
    for m in re.finditer(r'<w:num w:numId="(\d+)"[^>]*>\s*<w:abstractNumId w:val="(\d+)"/>', nd):
        out[m.group(1)] = absfmt.get(m.group(2), 'bullet')
    return out

def parse(path):
    z = zipfile.ZipFile(path)
    root = ET.fromstring(z.read('word/document.xml'))
    body = root.find(w+'body')
    blocks = []
    def walk(el):
        for ch in el:
            if ch.tag == w+'p':   blocks.append(p_info(ch))
            elif ch.tag == w+'tbl':
                t = tbl_info(ch)
                if len(t['rows']) == 1 and len(t['rows'][0]) == 1:
                    c = t['rows'][0][0]
                    for p in c['paras']:
                        p['cell_shd'] = c['shd']
                        blocks.append(p)
                else:
                    blocks.append(t)
            elif ch.tag == w+'sdt':
                c = ch.find(w+'sdtContent')
                if c is not None: walk(c)
    walk(body)
    return blocks, load_numfmt(z)

# ------------------------------------------------------------- classification
KNOWN_SECTIONS = {'executive summary','introduction','purpose','scope','background',
                  'definitions','responsibilities','references','appendix','conclusion',
                  'document control','revision history'}

def is_section(b):
    if b['kind'] != 'p' or not b['text']: return False
    t = b['text']
    if b['style'].lower().startswith('heading1'): return True
    if b['style'].lower() in ('title',): return False
    if COMPANY_RE.match(t): return False
    if re.match(r'^\d+\.\s', t) and b['bold'] and len(t) < 130: return True
    if b['bold'] and t.lower().rstrip(':').strip() in KNOWN_SECTIONS: return True
    if b['bold'] and t.isupper() and 3 <= len(t) <= 60 and not t.endswith('.'): return True
    if b.get('cell_shd') and b['cell_shd'] not in ('', 'AUTO', 'FFFFFF') \
       and b['bold'] and len(t) < 130: return True
    return False

def is_subhead(b):
    if b['kind'] != 'p' or not b['text']: return False
    t = b['text']
    if b['style'].lower().startswith(('heading2','heading3','heading4')): return True
    if re.match(r'^\d+\.\d+', t) and b['bold']: return True
    if re.match(r'^[a-z]\)\s', t) and b['bold']: return True
    if b['all_bold'] and 2 < len(t) <= 85 and not t.endswith(('.', ':')) and len(b['runs']) <= 2:
        return True
    return False

# ---------------------------------------------------------------- XML emitters
def rpr(bold=False, italic=False, font=BODY_FONT, sz=20, color=None):
    s = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    if bold:   s += '<w:b/><w:bCs/>'
    if italic: s += '<w:i/><w:iCs/>'
    if color:  s += f'<w:color w:val="{color}"/>'
    s += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    return '<w:rPr>' + s + '</w:rPr>'

def run(text, bold=False, italic=False, font=BODY_FONT, sz=20, color=None):
    return ('<w:r>' + rpr(bold, italic, font, sz, color) +
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r>')

def runs_xml(items, font=BODY_FONT, sz=20, color=None, force_bold=None):
    if not items: return ''
    return ''.join(run(t, (force_bold if force_bold is not None else b), i, font, sz, color)
                   for t, b, i in items)

def para(inner, after=80, before=0, numid=None, ilvl=0, jc=None, ind=None,
         keep=False, line=None, font=BODY_FONT, sz=20):
    pPr = ''
    if keep: pPr += '<w:keepNext/><w:keepLines/>'
    if numid is not None:
        pPr += f'<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{numid}"/></w:numPr>'
    sp = f'<w:spacing w:before="{before}" w:after="{after}"'
    sp += f' w:line="{line}" w:lineRule="auto"' if line else ''
    pPr += sp + '/>'
    if ind: pPr += ind
    if jc:  pPr += f'<w:jc w:val="{jc}"/>'
    pPr += rpr(font=font, sz=sz)
    return f'<w:p><w:pPr>{pPr}</w:pPr>{inner}</w:p>'

def banner(text):
    """Full-width green section heading bar (single-cell table = robust + faithful)."""
    body = para(run(text, bold=True, font=HEAD_FONT, sz=24, color='FFFFFF'),
                after=0, before=0, jc='center', font=HEAD_FONT, sz=24)
    return (
      '<w:tbl><w:tblPr>'
      f'<w:tblW w:w="{CONTENT_W}" w:type="dxa"/>'
      '<w:tblBorders>'
      '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
      '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
      '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
      '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
      '</w:tblBorders><w:tblLayout w:type="fixed"/>'
      '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>'
      '<w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>'
      '</w:tblPr>'
      f'<w:tblGrid><w:gridCol w:w="{CONTENT_W}"/></w:tblGrid>'
      '<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr>'
      f'<w:tcW w:w="{CONTENT_W}" w:type="dxa"/>'
      f'<w:shd w:val="clear" w:color="auto" w:fill="{GREEN}"/>'
      '<w:vAlign w:val="center"/></w:tcPr>'
      f'{body}</w:tc></w:tr></w:tbl>'
    )

def cell(width, inner, shd=None, valign='center'):
    tcPr = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if shd: tcPr += f'<w:shd w:val="clear" w:color="auto" w:fill="{shd}"/>'
    tcPr += f'<w:vAlign w:val="{valign}"/>'
    return f'<w:tc><w:tcPr>{tcPr}</w:tcPr>{inner}</w:tc>'

TBL_BORDERS = ('<w:tblBorders>'
  '<w:top w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>'
  '<w:left w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>'
  '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>'
  '<w:right w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>'
  '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>'
  '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="A6A6A6"/>'
  '</w:tblBorders>')

def scale_grid(grid, ncols):
    if not grid or sum(grid) <= 0:
        base = CONTENT_W // max(ncols, 1)
        grid = [base] * max(ncols, 1)
    tot = sum(grid)
    out = [max(int(round(CONTENT_W * g / tot)), 300) for g in grid]
    out[-1] += CONTENT_W - sum(out)
    return out

def table_xml(t):
    rows = t['rows']
    if not rows: return ''
    ncols = max(sum(c['span'] for c in r) for r in rows)
    grid = t['grid'] if len(t['grid']) == ncols else []
    cols = scale_grid(grid, ncols)
    # header-row detection
    r0 = rows[0]
    r0_bold = all(any(p['all_bold'] for p in c['paras'] if p['text']) or
                  not any(p['text'] for p in c['paras']) for c in r0)
    r0_white = any(p['color'] == 'FFFFFF' for c in r0 for p in c['paras'])
    r0_shd = any(c['shd'] and c['shd'] not in ('AUTO', 'FFFFFF') for c in r0)
    r1_plain = len(rows) > 1 and any(
        p['text'] and not p['all_bold'] for c in rows[1] for p in c['paras'])
    header = bool(r0_bold and (r0_white or r0_shd or r1_plain))

    out = ['<w:tbl><w:tblPr>'
           f'<w:tblW w:w="{CONTENT_W}" w:type="dxa"/>'
           + TBL_BORDERS + '<w:tblLayout w:type="fixed"/>'
           '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="90" w:type="dxa"/>'
           '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar>'
           '</w:tblPr><w:tblGrid>' +
           ''.join(f'<w:gridCol w:w="{c}"/>' for c in cols) + '</w:tblGrid>']
    for ri, r in enumerate(rows):
        hdr = header and ri == 0
        trPr = '<w:trPr><w:cantSplit/>' + ('<w:tblHeader/>' if hdr else '') + '</w:trPr>'
        tcs, ci = [], 0
        for c in r:
            wd = sum(cols[ci:ci+c['span']]) or cols[min(ci, len(cols)-1)]
            ci += c['span']
            ps = [p for p in c['paras']]
            if not any(p['text'] for p in ps):
                inner = para('', after=0, sz=18)
            else:
                bits = []
                for p in ps:
                    if not p['text']:
                        continue
                    bits.append(para(
                        runs_xml(p['runs'], sz=18,
                                 color='FFFFFF' if hdr else None,
                                 force_bold=True if hdr else None),
                        after=0, sz=18))
                inner = ''.join(bits) or para('', after=0, sz=18)
            tcs.append(cell(wd, inner, shd=GREEN if hdr else None,
                            valign='center' if hdr else 'top'))
        out.append(f'<w:tr>{trPr}{"".join(tcs)}</w:tr>')
    out.append('</w:tbl>')
    return ''.join(out)

def header_table(title, owner, date, logo_rid='rId11'):
    def hcell(width, inner, span=None, borders_right=True):
        tcPr = f'<w:tcW w:w="{width}" w:type="dxa"/>'
        if span: tcPr += f'<w:gridSpan w:val="{span}"/>'
        tcPr += ('<w:tcBorders>'
                 '<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                 '<w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                 '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                 + ('<w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
                    if borders_right else '<w:right w:val="nil"/>') +
                 '</w:tcBorders><w:vAlign w:val="center"/>')
        return f'<w:tc><w:tcPr>{tcPr}</w:tcPr>{inner}</w:tc>'
    logo = (
      '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
      '<wp:inline distT="0" distB="0" distL="0" distR="0">'
      '<wp:extent cx="2343150" cy="666750"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
      '<wp:docPr id="1001" name="Picture 1" descr="Vantage Specialty Chemicals logo"/>'
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
      '</wp:cNvGraphicFramePr>'
      '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
      '<pic:nvPicPr><pic:cNvPr id="0" name="Picture 1"/><pic:cNvPicPr>'
      '<a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr></pic:nvPicPr>'
      f'<pic:blipFill><a:blip r:embed="{logo_rid}"/><a:srcRect/>'
      '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
      '<pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/>'
      '<a:ext cx="2343150" cy="666750"/></a:xfrm>'
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
      '<a:ln><a:noFill/></a:ln></pic:spPr></pic:pic>'
      '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
    c1, c2, c3 = 4275, 4695, 1830
    r1 = ('<w:tr><w:trPr><w:trHeight w:val="1335"/></w:trPr>'
          + hcell(c1, para(logo, after=0))
          + hcell(c2, para(run(title, bold=True), after=0))
          + hcell(c3, para('', after=0)) + '</w:tr>')
    r2 = ('<w:tr><w:trPr><w:trHeight w:val="240"/></w:trPr>'
          + hcell(c1 + c2, para(run('Document Owner: ' + owner, bold=True), after=0),
                  span=2, borders_right=False)
          + hcell(c3, para(run('Date: ' + date, bold=True), after=0)) + '</w:tr>')
    return ('<w:tbl><w:tblPr>'
            f'<w:tblW w:w="{CONTENT_W}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:left w:w="80" w:type="dxa"/>'
            '<w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{c1}"/><w:gridCol w:w="{c2}"/>'
            f'<w:gridCol w:w="{c3}"/></w:tblGrid>{r1}{r2}</w:tbl>')

def kv_table(pairs):
    c1, c2 = 3200, CONTENT_W - 3200
    rows = []
    for k, v in pairs:
        rows.append('<w:tr><w:trPr><w:cantSplit/></w:trPr>'
                    + cell(c1, para(run(k, bold=True, sz=18, color=GREEN), after=0, sz=18),
                           valign='top')
                    + cell(c2, para(run(v, sz=18), after=0, sz=18), valign='top')
                    + '</w:tr>')
    return ('<w:tbl><w:tblPr>'
            f'<w:tblW w:w="{CONTENT_W}" w:type="dxa"/>'
            + TBL_BORDERS + '<w:tblLayout w:type="fixed"/>'
            '<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="90" w:type="dxa"/>'
            '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar>'
            '</w:tblPr>'
            f'<w:tblGrid><w:gridCol w:w="{c1}"/><w:gridCol w:w="{c2}"/></w:tblGrid>'
            + ''.join(rows) + '</w:tbl>')

# --------------------------------------------------------------- front matter
OWNER_KEYS = ('document owner', 'owner', 'policy owner')
DATE_KEYS  = ('date', 'effective date', 'date prepared', 'effective')
DROP_KEYS  = ('reporting period',)

def norm_key(k):
    return k.strip().rstrip(':').strip().lower()

def split_meta_line(text):
    """Pull 'Label: value' pairs out of a front-matter line (may be |-separated)."""
    pairs, leftovers = [], []
    for part in re.split(r'\s*[|·]\s*', text):
        part = part.strip()
        if not part: continue
        m = re.match(r'^([A-Za-z][A-Za-z &/\'-]{2,28}):\s*(.+)$', part)
        if m and len(m.group(2)) < 120:
            pairs.append((m.group(1).strip(), m.group(2).strip()))
        else:
            leftovers.append(part)
    return pairs, leftovers

def cell_pairs(c):
    """Pull bold-label / plain-value pairs out of a single table cell."""
    ps = [p for p in c['paras'] if p['text']]
    pairs, i = [], 0
    while i < len(ps) - 1:
        if ps[i]['all_bold'] and not ps[i+1]['all_bold'] and len(ps[i]['text']) < 40:
            pairs.append((ps[i]['text'], ps[i+1]['text'])); i += 2
        else:
            i += 1
    return pairs


def extract_front(blocks):
    """Split blocks into (front, rest) at the first real section heading.

    Title text is the largest type in the document, so a heading that is
    strictly smaller than the maximum is the first true section."""
    sizes = [b['sz'] for b in blocks if b['kind'] == 'p' and b['text']]
    maxsz = max(sizes) if sizes else 22
    for limit in (maxsz, maxsz + 1):   # strict, then inclusive fallback
        for i, b in enumerate(blocks):
            if b['kind'] == 'p' and is_section(b) and b['sz'] < limit:
                return blocks[:i], blocks[i:]
    return [], blocks

def build_meta(front, fallback_title):
    title_parts, pairs, intro = [], [], []
    paras = []
    for b in front:
        if b['kind'] == 'p':
            paras.append(b)
        else:  # table in front matter
            rows = b['rows']
            single = len(rows) == 1 and len(rows[0]) == 1
            if single:
                paras.extend(p for p in rows[0][0]['paras'])
                continue
            for r in rows:
                inner = []
                for c in r:
                    inner += cell_pairs(c)
                if inner:
                    pairs += inner
                    continue
                cells = [(' '.join(p['text'] for p in c['paras'] if p['text'])).strip()
                         for c in r]
                cells = [c for c in cells if c]
                if len(cells) >= 2:
                    for j in range(0, len(cells) - 1, 2):
                        pairs.append((cells[j], cells[j+1]))
                elif len(cells) == 1:
                    pr, lo = split_meta_line(cells[0])
                    pairs += pr; intro += lo
    big = [p for p in paras if p['sz'] >= 32 and p['text'] and not COMPANY_RE.match(p['text'])]
    if big:
        title_parts = [' '.join(p['text'] for p in big)]
        idx = paras.index(big[-1])
        nxt = paras[idx+1] if idx + 1 < len(paras) else None
        if nxt and nxt['text'] and 24 <= nxt['sz'] < 32 and len(nxt['text']) < 45 \
           and '|' not in nxt['text'] and not COMPANY_RE.match(nxt['text']):
            title_parts.append(nxt['text']); big.append(nxt)
    used = set(id(p) for p in big)
    for p in paras:
        if id(p) in used or not p['text'] or COMPANY_RE.match(p['text']):
            continue
        pr, lo = split_meta_line(p['text'])
        pairs += pr
        intro += lo
    title = ' — '.join(title_parts) if len(title_parts) > 1 else (
            title_parts[0] if title_parts else fallback_title)
    title = re.sub(r'\s+', ' ', title).strip()
    # dedupe + classify pairs
    seen, clean = set(), []
    owner = date = None
    for k, v in pairs:
        nk = norm_key(k)
        if not nk or not v or (nk, v) in seen: continue
        seen.add((nk, v))
        if nk in OWNER_KEYS:
            if owner is None: owner = v
            continue
        if nk in DATE_KEYS:
            if date is None: date = v
            continue
        clean.append((k.strip().rstrip(':'), v))
    intro = [t for t in intro
             if len(t) > 3 and t.strip().rstrip(':').lower() not in KNOWN_SECTIONS]
    return title, owner, date, clean, intro

# ----------------------------------------------------------------- assembly
def build_body(blocks, numfmt, state):
    """Emit XML for the post-front-matter content."""
    out = []
    i = 0
    while i < len(blocks):
        b = blocks[i]
        if b['kind'] == 'tbl':
            out.append(table_xml(b))
            out.append(para('', after=100))
            i += 1; continue
        if not b['text']:
            i += 1; continue
        if b['numid'] is not None:
            fmt = numfmt.get(b['numid'], 'bullet')
            group = []
            while i < len(blocks) and blocks[i]['kind'] == 'p' \
                  and blocks[i]['numid'] == b['numid'] and blocks[i]['text']:
                group.append(blocks[i]); i += 1
            nid = state['next_num'](fmt)
            for g in group:
                lvl = min(g['ilvl'], 2)
                ind = (f'<w:ind w:left="{720 + 360*lvl}" w:hanging="360"/>')
                out.append(para(runs_xml(g['runs']), after=40, numid=nid,
                                ilvl=lvl, ind=ind))
            out.append(para('', after=60))
            continue
        if is_section(b):
            out.append(para('', after=60))
            out.append(banner(b['text']))
            out.append(para('', after=80))
            i += 1; continue
        if is_subhead(b):
            out.append(para(runs_xml(b['runs'], color=GREEN, force_bold=True),
                            before=120, after=60, keep=True))
            i += 1; continue
        out.append(para(runs_xml(b['runs']), after=120))
        i += 1
    return ''.join(out)

BULLET_ABS = 900
DECIMAL_ABS = 901

def numbering_xml(src, n_lists):
    def lvl_bullet(i):
        return (f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
                f'<w:lvlText w:val="{["","o",""][i%3] or ""}"/>'
                f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{720+360*i}" w:hanging="360"/></w:pPr>'
                '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>')
    def lvl_bullet_real(i):
        ch = ['', 'o', ''][i % 3]
        font = ['Symbol', 'Courier New', 'Wingdings'][i % 3]
        return (f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
                f'<w:lvlText w:val="{esc(ch)}"/><w:lvlJc w:val="left"/>'
                f'<w:pPr><w:ind w:left="{720+360*i}" w:hanging="360"/></w:pPr>'
                f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr></w:lvl>')
    def lvl_dec(i):
        fmt = ['decimal', 'lowerLetter', 'lowerRoman'][i % 3]
        return (f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                f'<w:lvlText w:val="%{i+1}."/><w:lvlJc w:val="left"/>'
                f'<w:pPr><w:ind w:left="{720+360*i}" w:hanging="360"/></w:pPr></w:lvl>')
    b_abs = (f'<w:abstractNum w:abstractNumId="{BULLET_ABS}"><w:multiLevelType w:val="hybridMultilevel"/>'
             + ''.join(lvl_bullet_real(i) for i in range(3)) + '</w:abstractNum>')
    d_abs = (f'<w:abstractNum w:abstractNumId="{DECIMAL_ABS}"><w:multiLevelType w:val="hybridMultilevel"/>'
             + ''.join(lvl_dec(i) for i in range(3)) + '</w:abstractNum>')
    nums = ''.join(
        f'<w:num w:numId="{900+i}"><w:abstractNumId w:val="{aid}"/></w:num>'
        for i, aid in n_lists)
    j = src.find('<w:num ')
    if j < 0: j = src.rfind('</w:numbering>')
    src = src[:j] + b_abs + d_abs + src[j:]
    k = src.rfind('</w:num>')
    k = k + len('</w:num>') if k >= 0 else src.rfind('</w:numbering>')
    return src[:k] + nums + src[k:]

def convert(path, outdir):
    base = os.path.basename(path)
    blocks, numfmt = parse(path)
    fallback = re.sub(r'\.docx$', '', base)
    fallback = re.sub(r'^[\d\s&-]+', '', fallback).replace('_', ' ').strip()
    front, rest = extract_front(blocks)
    title, owner, date, pairs, intro = build_meta(front, fallback)
    owner = owner or 'Sustainability'
    if not date or re.fullmatch(r'\[.*\]', date.strip()):
        date = TODAY

    lists = []
    def next_num(fmt):
        aid = DECIMAL_ABS if fmt == 'decimal' else BULLET_ABS
        lists.append((len(lists), aid))
        return 900 + len(lists) - 1
    state = {'next_num': next_num}

    parts = [header_table(title, owner, date), para('', after=120)]
    for t in intro:
        parts.append(para(run(t), after=120))
    if pairs:
        parts.append(para('', after=40))
        parts.append(banner('Document Control'))
        parts.append(para('', after=80))
        parts.append(kv_table(pairs))
        parts.append(para('', after=140))
    parts.append(build_body(rest, numfmt, state))

    tpl_doc = open(os.path.join(TPL_DIR, 'word/document.xml'), encoding='utf8').read()
    open_tag = tpl_doc[:tpl_doc.index('<w:body>') + len('<w:body>')]
    sect = tpl_doc[tpl_doc.rfind('<w:sectPr'):]
    doc = open_tag + ''.join(parts) + sect

    footer = open(os.path.join(TPL_DIR, 'word/footer1.xml'), encoding='utf8').read()
    footer = footer.replace('In-House Anti-Corruption Questionnaires', esc(title))
    numb = open(os.path.join(TPL_DIR, 'word/numbering.xml'), encoding='utf8').read()
    numb = numbering_xml(numb, lists)

    out = os.path.join(outdir, base)
    override = {'word/document.xml': doc, 'word/footer1.xml': footer,
                'word/numbering.xml': numb}
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(TPL_DIR):
            for f in files:
                full = os.path.join(root, f)
                rel = os.path.relpath(full, TPL_DIR).replace(os.sep, '/')
                if rel in override:
                    z.writestr(rel, override[rel].encode('utf8'))
                else:
                    z.write(full, rel)
    return out, title, owner, date, len(lists)

def ensure_template():
    if os.path.isdir(TPL_DIR):
        return
    with zipfile.ZipFile(TPL_DOCX) as z:
        z.extractall(TPL_DIR)
    for root, _, files in os.walk(TPL_DIR):
        for f in files:
            p = os.path.join(root, f)
            if os.path.islink(p):
                os.unlink(p)


if __name__ == '__main__':
    ensure_template()
    outdir = sys.argv[1] if len(sys.argv) > 1 else os.path.join(SRC_DIR, 'Converted')
    os.makedirs(outdir, exist_ok=True)
    targets = sys.argv[2:] or sorted(
        f for f in glob.glob(os.path.join(SRC_DIR, '*.docx'))
        if 'Template Example' not in f)
    for f in targets:
        o, t, ow, d, n = convert(f, outdir)
        print(f'OK  {os.path.basename(o)}\n    title={t!r} owner={ow!r} date={d!r} lists={n}')
