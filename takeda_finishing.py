"""
takeda_finishing.py — Generalized Mode-B FINISHING PASSES (brand-agnostic).

These are the "make it look right" polish passes you run AFTER cloning/recoloring/
retheming a deck, before the QA gate. They encode fixes that recur on every
conversion regardless of source or brand:

  1. align_blade_to_title   — header blade top aligns with the title's first text line
  2. apply_footer_strip_locked / make_strip_bleed
                            — footer as LOCKED layout furniture, behind content,
                              bled edge-to-edge (no gap)
  3. (covered by #2)        — footer behind all slide content => text always in front
  4. fix_text_contrast      — background-aware contrast (white on dark, dark on light),
                              checking slide/layout/master backgrounds AND underlying panels
  5. tidy_title_text        — collapse awkward whitespace, grow the box so a title wraps
                              neatly instead of reproducing a broken source layout

Import alongside takeda:
    import sys; sys.path.insert(0, '/mnt/project')
    import takeda_finishing as fin
    from takeda import _chrome_png            # strip asset source

Then, after your per-slide clone/recolor/retheme and before saving:
    fin.finish_deck(deck.prs, _chrome_png('ENTYVIO_BANNER'),
                    blade_layouts={'Title and Content','Title Only',
                                   'One Third LHS Dark','One Third LHS Light'})

Depends only on python-pptx, lxml, Pillow, numpy (no takeda internals).
"""
import re, io
from pptx.oxml.ns import qn
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
EMU = 914400

# ----------------------------- colour helpers -----------------------------
def _theme_map(prs):
    """Resolve the (possibly rethemed) theme scheme to a {slot: hex} dict, including
    bg1/tx1/bg2/tx2 aliases, so schemeClr references can be resolved to real hexes."""
    m = prs.slide_masters[0]
    for rel in m.part.rels.values():
        if 'theme' in rel.reltype:
            root = etree.fromstring(rel.target_part.blob)
            cs = root.find(f'.//{{{A}}}clrScheme'); d = {}
            for slot in ('dk1','lt1','dk2','lt2','accent1','accent2','accent3',
                         'accent4','accent5','accent6'):
                node = cs.find(f'{{{A}}}{slot}')
                if node is None: continue
                s = node.find(f'{{{A}}}srgbClr'); sy = node.find(f'{{{A}}}sysClr')
                d[slot] = s.get('val') if s is not None else (sy.get('lastClr','000000') if sy is not None else None)
            d['bg1']=d.get('lt1','FFFFFF'); d['tx1']=d.get('dk1','000000')
            d['bg2']=d.get('lt2','FFFFFF'); d['tx2']=d.get('dk2','000000')
            return d
    return {}

def _resolve(el, TH):
    if el is None: return None
    s = el.find(qn('a:srgbClr'))
    if s is not None: return s.get('val')
    sc = el.find(qn('a:schemeClr'))
    if sc is not None: return TH.get(sc.get('val'))
    return None

def _lum(h):
    h = h.lstrip('#'); r,g,b = [int(h[i:i+2],16)/255 for i in (0,2,4)]
    f = lambda c: c/12.92 if c<=0.03928 else ((c+0.055)/1.055)**2.4
    return 0.2126*f(r)+0.7152*f(g)+0.0722*f(b)

def _ratio(a,b):
    la,lb = _lum(a),_lum(b); hi,lo = max(la,lb),min(la,lb)
    return (hi+0.05)/(lo+0.05)

def _shape_fill(sh, TH):
    spPr = sh._element.find(qn('p:spPr'))
    return _resolve(spPr.find(qn('a:solidFill')), TH) if spPr is not None else None

def _base_bg(slide, TH):
    for obj in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        bg = obj.element.find('.//'+qn('p:bg'))
        if bg is not None:
            c = _resolve(bg.find('.//'+qn('a:solidFill')), TH)
            if c: return c
    return 'FFFFFF'

def _box(sh):
    try: return (sh.left/EMU, sh.top/EMU, (sh.left+sh.width)/EMU, (sh.top+sh.height)/EMU)
    except Exception: return None

def _overlap(a,b):
    return not (a[2]<=b[0] or b[2]<=a[0] or a[3]<=b[1] or b[3]<=a[1])

def _set_run_color(rpr, hexv):
    """Set an explicit run/endPara colour, inserting <a:solidFill> in its SCHEMA-CORRECT
    position. In CT_TextCharacterProperties the fill group must come right after <a:ln>
    and BEFORE <a:latin>/<a:ea>/<a:cs>/<a:sym>/<a:hlink*>. Appending it at the end (after
    latin/cs) is tolerated by LibreOffice but IGNORED by PowerPoint — which then falls back
    to the inherited/lstStyle colour (e.g. bg2=purple). This ordering is the fix."""
    for tag in ('a:noFill','a:solidFill','a:gradFill','a:blipFill','a:pattFill','a:grpFill'):
        for f in rpr.findall(qn(tag)): rpr.remove(f)
    sf = rpr.makeelement(qn('a:solidFill'), {})
    etree.SubElement(sf, qn('a:srgbClr')).set('val', hexv)
    ln = rpr.find(qn('a:ln'))
    if ln is not None:
        ln.addnext(sf)                 # fill goes immediately after a:ln
    else:
        rpr.insert(0, sf)              # else first child

# ------------------- PASS 4/3: background-aware contrast -------------------
def fix_text_contrast(slide, TH, strip_band=(7.13, 7.5), dark_hex='1A1628'):
    """White text on a dark effective background, dark text on a light one.
    Effective bg = the shape's own fill -> the nearest underlying filled shape
    (slide, then layout, then master) -> the footer-strip band -> the slide/layout/
    master background. Colour is written explicitly on runs AND endParaRPr so it
    can't revert to an inherited theme colour. This is the correct fix for
    'headings the same colour as the background' — never blindly force one colour."""
    base = _base_bg(slide, TH)
    shapes = list(slide.shapes)
    for zi, sh in enumerate(shapes):
        if not sh.has_text_frame: continue
        tb = _box(sh)
        eff = _shape_fill(sh, TH)
        if not eff and tb:                                   # underlying slide shapes
            for zj in range(zi-1, -1, -1):
                u = shapes[zj]; ub = _box(u); uf = _shape_fill(u, TH)
                if uf and ub and _overlap(tb, ub): eff = uf; break
        if not eff and tb:                                   # layout, then master shapes
            for container in (slide.slide_layout, slide.slide_layout.slide_master):
                for u in container.shapes:
                    ub = _box(u); uf = _shape_fill(u, TH)
                    if uf and ub and _overlap(tb, ub): eff = uf; break
                if eff: break
        if not eff and tb and strip_band[0] <= tb[1] < strip_band[1]:
            eff = '1A1628'                                   # sits on the dark footer strip
        if not eff: eff = base
        bg_dark = _lum(eff) < 0.5
        target = 'FFFFFF' if bg_dark else dark_hex
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                if not run.text.strip(): continue
                rpr = run._r.find(qn('a:rPr'))
                if rpr is None:
                    rpr = run._r.makeelement(qn('a:rPr'), {}); run._r.insert(0, rpr)
                cur = _resolve(rpr, TH)
                if bg_dark:
                    if cur != 'FFFFFF': _set_run_color(rpr, 'FFFFFF')
                else:
                    if cur is None or _ratio(cur, eff) < 3.0: _set_run_color(rpr, target)
            epr = para._p.find(qn('a:endParaRPr'))
            if epr is not None: _set_run_color(epr, target)

# --------------------- PASS 5: tidy title whitespace/fit ---------------------
def tidy_title_text(slide, max_h_in=2.2):
    """Collapse runs of spaces, strip spaces around soft line-breaks / line ends, and
    grow the title box (height only) so it never clips — so an awkward source title
    wraps neatly instead of being reproduced verbatim. NEVER writes a partial xfrm
    (a bare <a:off> without y collapses the box to zero width in PowerPoint), so it
    only resizes when a complete off+ext already exists."""
    for sh in slide.shapes:
        if not (sh.is_placeholder and sh.has_text_frame): continue
        try:
            if sh.placeholder_format.type is not None and sh.placeholder_format.idx != 0:
                continue
        except Exception:
            continue
        tb = sh.text_frame
        for para in tb.paragraphs:
            runs = para.runs
            for k, run in enumerate(runs):
                t = re.sub(r' {2,}', ' ', run.text)
                nxt = run._r.getnext()
                if k == len(runs)-1 or (nxt is not None and nxt.tag == qn('a:br')):
                    t = t.rstrip()
                prv = run._r.getprevious()
                if prv is not None and prv.tag == qn('a:br'):
                    t = t.lstrip()
                if t != run.text: run.text = t
        n_lines = 0; max_sz = 24
        for para in tb.paragraphs:
            n_lines += 1 + len(para._p.findall(qn('a:br')))
            for run in para.runs:
                if run.font.size: max_sz = max(max_sz, run.font.size.pt)
        need = n_lines * (max_sz * 1.2 / 72.0) + 0.15
        spPr = sh._element.find(qn('p:spPr'))
        if spPr is None: continue
        xfrm = spPr.find(qn('a:xfrm'))
        if xfrm is None: continue
        off = xfrm.find(qn('a:off')); ext = xfrm.find(qn('a:ext'))
        if off is None or ext is None: continue
        if None in (off.get('x'), off.get('y'), ext.get('cx'), ext.get('cy')): continue
        if need > int(ext.get('cy'))/EMU and need <= max_h_in:
            ext.set('cy', str(int(need*EMU)))                # width & position preserved

# ---------------------- PASS 1: align blade to title ----------------------
def align_blade_to_title(slide, blade_name='TakedaBlade'):
    """Move the header blade so its TOP edge meets the title's first text-line top
    (title box top + top inset), instead of floating at a fixed Y."""
    blade = title = None
    for sh in slide.shapes:
        if getattr(sh, 'name', '') == blade_name: blade = sh
        if sh.is_placeholder and sh.has_text_frame:
            try:
                if sh.placeholder_format.idx == 0: title = sh
            except Exception: pass
    if blade is None or title is None or title.top is None: return
    bodyPr = title._element.find('.//'+qn('a:bodyPr'))
    t_ins = int(bodyPr.get('tIns')) if (bodyPr is not None and bodyPr.get('tIns')) else 45720
    off = blade._element.find('.//'+qn('a:xfrm')).find(qn('a:off'))
    off.set('y', str(int(title.top) + t_ins))

# ---------------- PASS 2/3: locked, edge-to-edge footer strip ----------------
def make_strip_bleed(src_png, out_png='/tmp/strip_bleed.png', sat_thresh=0.30):
    """Crop the near-white margins baked into the footer asset's edges so the
    saturated gradient bleeds edge-to-edge (fixes the perimeter 'gap')."""
    from PIL import Image
    import numpy as np
    im = Image.open(src_png).convert('RGB')
    a = np.asarray(im).astype(float)
    mx = a.max(axis=2); mn = a.min(axis=2)
    sat = np.where(mx > 0, (mx - mn) / np.clip(mx, 1, None), 0)
    content = sat > sat_thresh
    cols = np.where(content.mean(axis=0) > 0.4)[0]
    rows = np.where(content.mean(axis=1) > 0.4)[0]
    if len(cols) == 0 or len(rows) == 0:
        im.save(out_png); return out_png
    im.crop((int(cols.min()), int(rows.min()),
             int(cols.max())+1, int(rows.max())+1)).save(out_png)
    return out_png

def apply_footer_strip_locked(prs, strip_png, band=(0.0, 7.15, 13.333, 0.35),
                              bleed=True, name='entyvio_footer_strip'):
    """Place the footer strip once on every USED layout as LOCKED, non-selectable
    furniture: it renders behind all slide content (so text is always in front) and
    above each layout's own background (so it shows even on dark dividers), and it
    cannot be clicked/moved/resized on a slide. Removes any per-slide strip pics first.
    bleed=True crops the asset margins so the gradient reaches both edges."""
    if bleed: strip_png = make_strip_bleed(strip_png)
    L, T, W, H = band
    for s in prs.slides:
        for sh in list(s.shapes):
            if sh.shape_type == 13 and sh.top is not None and abs(sh.top/EMU - T) < 0.06:
                sh._element.getparent().remove(sh._element)
    seen, used = set(), []
    for s in prs.slides:
        lay = s.slide_layout
        if id(lay) not in seen: seen.add(id(lay)); used.append(lay)
    with open(strip_png, 'rb') as f: blob = f.read()
    for lay in used:
        if any(getattr(sh,'name','')==name for sh in lay.shapes): continue
        _img, rId = lay.part.get_or_add_image_part(io.BytesIO(blob))
        pic = f'''<p:pic xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">
          <p:nvPicPr>
            <p:cNvPr id="990" name="{name}"/>
            <p:cNvPicPr><a:picLocks noSelect="1" noMove="1" noResize="1" noChangeAspect="1"/></p:cNvPicPr>
            <p:nvPr userDrawn="1"/>
          </p:nvPicPr>
          <p:blipFill><a:blip r:embed="{rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
          <p:spPr><a:xfrm><a:off x="{int(L*EMU)}" y="{int(T*EMU)}"/>
            <a:ext cx="{int(W*EMU)}" cy="{int(H*EMU)}"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        </p:pic>'''
        lay.shapes._spTree.insert(2, etree.fromstring(pic))   # behind content, above layout bg

# ------------------------------ orchestrator ------------------------------
def finish_deck(prs, strip_png, blade_layouts=None, add_blade=None):
    """Run every finishing pass in the correct order. `add_blade` is an optional
    callable add_blade(slide) (e.g. takeda.add_blade_accent) used on slides whose
    layout name is in `blade_layouts`. Contrast runs LAST so titles/headings are
    coloured against their real, final backgrounds."""
    TH = _theme_map(prs)
    for slide in prs.slides:
        tidy_title_text(slide)
    if add_blade and blade_layouts:
        for slide in prs.slides:
            if slide.slide_layout.name in blade_layouts:
                add_blade(slide)
                align_blade_to_title(slide)
    apply_footer_strip_locked(prs, strip_png)
    for slide in prs.slides:
        fix_text_contrast(slide, TH)
