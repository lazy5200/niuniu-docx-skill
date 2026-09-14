# -*- coding: utf-8 -*-
"""docxlib.py —— 按《规范.md》以基底文档重建整份 .docx 的通用构建库

设计原则
--------
1. **基底提供"看不见的部件"**：numbering.xml / styles.xml / header*.xml / footer*.xml / sectPr。
   目标文档若本身缺这些（python-docx 从零生成的常见病），不要就地补，直接换基底重建。
2. **样式 id 不硬编码**：一律用 styles.xml 的 `w:name → styleId` 反查。
   Word COM 保存后会重编号 styleId（heading1 2→1、Table Grid 11→a6），硬编码必炸。
3. **body 子元素索引不硬编码**：自动探测封面段 / 修订表 / '目录' 标题 / TOC 原型 / 外层域 end。
   （`kids[i]` ≠ 段索引：body 里除 `w:p` 还有 `w:tbl` 和 `w:sectPr`。）
4. **所有 rPr / pPr 每次新建**：lxml 元素只能有一个父节点，复用会被"搬走"。
5. **段前段后必须带四伴随属性**：`beforeLines/afterLines/beforeAutospacing/afterAutospacing`，
   否则 Word 保存时会把 before/after 当冗余属性清掉（实测标题顶距页眉 46.4pt → 28.2pt）。
"""
import os
import re
import copy
import math
import shutil
import zipfile
import datetime

from lxml import etree

WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = '{%s}' % WNS
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
R = '{%s}' % RNS
XS = '{http://www.w3.org/XML/1998/namespace}space'

# ---- 绘图（插图）相关命名空间 ----
WPNS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
ANS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PICNS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
PKGNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CTNS = 'http://schemas.openxmlformats.org/package/2006/content-types'
REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

EMU_PT = 12700                      # 1 磅 = 12700 EMU
IMG_DEFAULT_DPI = 96.0              # 像素 → 自然尺寸的换算基准


def q(t):
    return W + t


# ============================== schema 子元素顺序表 ==============================
PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
             'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
             'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
             'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
             'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle',
             'rPr', 'sectPr', 'pPrChange']
RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
             'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color',
             'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd',
             'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']
TBL_ORDER = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize',
             'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar',
             'tblLook', 'tblCaption', 'tblDescription']
TRPR_ORDER = ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight',
              'tblHeader', 'tblCellSpacing', 'jc', 'hidden']
TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar',
              'textDirection', 'tcFitText', 'vAlign', 'hideMark']


def ins(parent, child, order):
    """按 schema 顺序把 child 插入 parent（顺序错 → Word 报"文件损坏"）"""
    tag = child.tag.split('}')[1]
    idx = order.index(tag)
    for i, ch in enumerate(parent):
        t = ch.tag.split('}')[1]
        if t in order and order.index(t) > idx:
            parent.insert(i, child)
            return child
    parent.append(child)
    return child


def mk(tag, **kw):
    e = etree.Element(q(tag))
    for k, v in kw.items():
        e.set(q(k), str(v))
    return e


# ============================== 规范参数（与 references/规范.md 对齐） ==============================
SPEC = {
    'avail': 8306,                # 版心宽 = A4(11906) − 左右页边距(1800×2)，单位 twip
    'h1': dict(style='heading 1', size=30, before=340, after=330, line=576, rule='auto', jc='both'),
    'h2': dict(style='heading 2', size=28, before=260, after=260, line=413, rule='auto', jc='both'),
    'h3': dict(style='heading 3', size=24, before=240, after=120, line=312, rule='auto', jc='both'),
    'body': dict(size=24, line=400, rule='exact', first=480, first_chars=200),
    'cmd': dict(size=24, line=400, rule='exact', left=440),
    'cap': dict(size=21, before=240, after=120),      # 表题：表格上方，段前12磅/段后6磅
    'fig': dict(size=21, before=120, after=240),      # 图题：图片下方，段前6磅/段后12磅
    'figimg': dict(before=120, after=120),            # 图片段：段前段后各 6 磅，单倍行距
    'table': dict(text=21, hdr=21, border_sz=4, cellmar=108),
    'revision': dict(hdr=18, body=21, widths=[2076, 2076, 2077, 2077]),
    'cover': dict(title=36, sub=28),
}


# ============================== 基础元素 ==============================
def build_rpr(east='宋体', ascii_font='Times New Roman', size=None, bold=None, italic=None):
    rpr = etree.Element(q('rPr'))
    rf = ins(rpr, mk('rFonts'), RPR_ORDER)
    rf.set(q('ascii'), ascii_font)
    rf.set(q('hAnsi'), ascii_font)
    rf.set(q('eastAsia'), east)
    rf.set(q('cs'), ascii_font)
    if bold:
        ins(rpr, mk('b'), RPR_ORDER)
        ins(rpr, mk('bCs'), RPR_ORDER)
    if italic:
        ins(rpr, mk('i'), RPR_ORDER)
        ins(rpr, mk('iCs'), RPR_ORDER)
    if size:
        ins(rpr, mk('sz', val=size), RPR_ORDER)
        ins(rpr, mk('szCs', val=size), RPR_ORDER)
    return rpr


def build_ppr(style=None, keep_next=False, num=None, before=None, after=None, line=None, line_rule=None,
              ind_first=None, ind_chars=None, ind_left=None, ind_hang=None, jc=None, rpr=None, sect=None,
              sp_lines=True, bidi0=False, snap=False):
    ppr = etree.Element(q('pPr'))
    if style:
        ins(ppr, mk('pStyle', val=style), PPR_ORDER)
    if keep_next:
        ins(ppr, mk('keepNext'), PPR_ORDER)
    if num:
        npr = ins(ppr, mk('numPr'), PPR_ORDER)
        ins(npr, mk('ilvl', val=num[1]), ['ilvl', 'numId'])
        ins(npr, mk('numId', val=num[0]), ['ilvl', 'numId'])
    if bidi0:
        ins(ppr, mk('bidi', val='0'), PPR_ORDER)
    if snap:
        ins(ppr, mk('snapToGrid'), PPR_ORDER)
    if any(x is not None for x in (before, after, line, line_rule)):
        sp = mk('spacing')
        # ★★ before/after 必须配 beforeLines/afterLines/beforeAutospacing/afterAutospacing，
        #    否则 Word COM 保存时把 before/after 当"冗余属性"整组删掉。
        if before is not None:
            sp.set(q('before'), str(before))
            if sp_lines:
                sp.set(q('beforeLines'), '0')
                sp.set(q('beforeAutospacing'), '0')
        if after is not None:
            sp.set(q('after'), str(after))
            if sp_lines:
                sp.set(q('afterLines'), '0')
                sp.set(q('afterAutospacing'), '0')
        if line is not None:
            sp.set(q('line'), str(line))
        if line_rule:
            sp.set(q('lineRule'), line_rule)
        ins(ppr, sp, PPR_ORDER)
    if any(x is not None for x in (ind_first, ind_chars, ind_left, ind_hang)):
        ind = mk('ind')
        ind.set(q('left'), str(ind_left or 0))
        ind.set(q('leftChars'), '0')
        ind.set(q('right'), '0')
        ind.set(q('rightChars'), '0')
        ind.set(q('firstLine'), str(ind_first or 0))
        ind.set(q('firstLineChars'), str(ind_chars or 0))
        if ind_hang is not None:
            ind.set(q('hanging'), str(ind_hang))
            ind.set(q('hangingChars'), '0')
        ins(ppr, ind, PPR_ORDER)
    if jc:
        ins(ppr, mk('jc', val=jc), PPR_ORDER)
    if rpr is not None:
        ins(ppr, rpr, PPR_ORDER)
    if sect is not None:
        ins(ppr, sect, PPR_ORDER)
    return ppr


def make_run(text, rpr=None):
    r = etree.Element(q('r'))
    if rpr is not None:
        r.append(rpr)
    t = etree.SubElement(r, q('t'))
    t.set(XS, 'preserve')
    t.text = text
    return r


def para(ppr, runs):
    p = etree.Element(q('p'))
    if ppr is not None:
        p.append(ppr)
    for r in runs:
        p.append(r)
    return p


# ============================== 探测工具 ==============================
def tag(el):
    return etree.QName(el).localname


def ptext(el):
    return ''.join(t.text or '' for t in el.iter(q('t'))).strip()


def resolve_style_ids(styles_bytes):
    """w:name → styleId 反查（COM 保存后 styleId 会重编号，绝不能硬编码）"""
    root = etree.fromstring(styles_bytes)
    out = {}
    for st in root.findall(q('style')):
        nm = st.find(q('name'))
        if nm is not None:
            out[(nm.get(q('val')) or '').strip().lower()] = st.get(q('styleId'))
    return out


def detect_layout(doc, sid):
    """自动探测基底结构，返回素材索引与原型。"""
    body = doc.find(q('body'))
    kids = list(body)
    L = {'kids': kids, 'body': body}

    first_tbl = next((i for i, e in enumerate(kids) if tag(e) == 'tbl'), None)
    if first_tbl is None:
        raise RuntimeError('基底里找不到任何表格（修订记录表）——请换一个带封面的基底')
    L['cover'] = list(range(0, first_tbl))
    L['rev_tbl'] = first_tbl

    dir_idx = next((i for i, e in enumerate(kids)
                    if i > first_tbl and tag(e) == 'p' and ptext(e) == '目录'), None)
    L['dir_para'] = dir_idx

    toc1_id, toc2_id = sid.get('toc 1'), sid.get('toc 2')
    L['toc_paras'] = [i for i, e in enumerate(kids)
                      if tag(e) == 'p' and i > (dir_idx or first_tbl)
                      and (e.find(q('pPr')) is not None
                           and e.find(q('pPr')).find(q('pStyle')) is not None
                           and e.find(q('pPr')).find(q('pStyle')).get(q('val')) in (toc1_id, toc2_id))]
    L['toc1_idx'] = next((i for i in L['toc_paras']
                          if kids[i].find(q('pPr')).find(q('pStyle')).get(q('val')) == toc1_id), None)
    L['toc2_idx'] = next((i for i in L['toc_paras']
                          if kids[i].find(q('pPr')).find(q('pStyle')).get(q('val')) == toc2_id), None)

    # 外层 TOC 域 end 所在段（最后一条 TOC 条目之后，含 fldChar end）
    last_toc = max(L['toc_paras']) if L['toc_paras'] else (dir_idx or first_tbl)
    L['toc_tail'] = next((i for i, e in enumerate(kids)
                          if i > last_toc and tag(e) == 'p'
                          and e.find('.//' + q('fldChar')) is not None), None)

    # 修订表与目录之间的分节符载体（含段内 sectPr 的空段）
    L['sect_mid'] = next((i for i, e in enumerate(kids)
                          if (dir_idx is not None and first_tbl < i < dir_idx) and tag(e) == 'p'
                          and e.find(q('pPr')) is not None
                          and e.find(q('pPr')).find(q('sectPr')) is not None), None)

    # body 末尾的 body 级 sectPr
    L['sect'] = kids[-1] if tag(kids[-1]) == 'sectPr' else None
    if L['sect'] is None:
        raise RuntimeError('基底 body 末尾没有 body 级 sectPr')
    return L


def detect_h1_numid(doc, h1_style_id):
    """从基底的章标题段落里读出 numId（保证 %1、 这一套编号定义被沿用）"""
    for el in doc.iter(q('p')):
        pp = el.find(q('pPr'))
        if pp is None:
            continue
        st = pp.find(q('pStyle'))
        if st is None or st.get(q('val')) != h1_style_id:
            continue
        npr = pp.find(q('numPr'))
        if npr is None:
            continue
        nid = npr.find(q('numId'))
        if nid is not None:
            return nid.get(q('val'))
    return None


# ============================== 各类段落 ==============================
def H1(text, sid, num_id):
    cfg = SPEC['h1']
    pr = build_rpr(east='黑体', ascii_font='黑体', size=cfg['size'], bold=True)
    pp = build_ppr(style=sid.get('heading 1'), jc=cfg['jc'],
                   num=(num_id, 0) if num_id else None,
                   before=cfg['before'], after=cfg['after'],
                   line=cfg['line'], line_rule=cfg['rule'],
                   rpr=build_rpr(east='黑体', ascii_font='黑体', size=cfg['size'], bold=True),
                   bidi0=True, snap=True)
    return para(pp, [make_run(text, pr)])


def H2(text, sid):
    cfg = SPEC['h2']
    pr = build_rpr(east='黑体', ascii_font='黑体', size=cfg['size'], bold=True)
    pp = build_ppr(style=sid.get('heading 2'), keep_next=True, jc=cfg['jc'],
                   before=cfg['before'], after=cfg['after'],
                   line=cfg['line'], line_rule=cfg['rule'],
                   rpr=build_rpr(east='黑体', ascii_font='黑体', size=cfg['size'], bold=True),
                   bidi0=True)
    return para(pp, [make_run(text, pr)])


def H3(text, sid):
    cfg = SPEC['h3']
    pr = build_rpr(east='黑体', ascii_font='黑体', size=cfg['size'], bold=True)
    pp = build_ppr(style=sid.get('heading 3'), keep_next=True, jc=cfg['jc'],
                   before=cfg['before'], after=cfg['after'],
                   line=cfg['line'], line_rule=cfg['rule'],
                   rpr=build_rpr(east='黑体', ascii_font='黑体', size=cfg['size'], bold=True))
    return para(pp, [make_run(text, pr)])


def BODY(text):
    cfg = SPEC['body']
    pr = build_rpr(east='宋体', ascii_font='Times New Roman', size=cfg['size'])
    pp = build_ppr(line=cfg['line'], line_rule=cfg['rule'],
                   ind_first=cfg['first'], ind_chars=cfg['first_chars'],
                   rpr=build_rpr(east='宋体', ascii_font='Times New Roman', size=cfg['size']))
    return para(pp, [make_run(text, pr)])


def CMD(text):
    cfg = SPEC['cmd']
    pr = build_rpr(east='宋体', ascii_font='Times New Roman', size=cfg['size'])
    pp = build_ppr(line=cfg['line'], line_rule=cfg['rule'],
                   ind_first=0, ind_chars=0, ind_left=cfg['left'],
                   rpr=build_rpr(east='宋体', ascii_font='Times New Roman', size=cfg['size']))
    return para(pp, [make_run(text, pr)])


def CAP(text, kind='T'):
    """表题（置表上，段前12/段后6）或图题（置图下，段前6/段后12）"""
    cfg = SPEC['cap'] if kind == 'T' else SPEC['fig']
    pr = build_rpr(east='宋体', ascii_font='Times New Roman', size=cfg['size'])
    pp = build_ppr(keep_next=(kind == 'T'), before=cfg['before'], after=cfg['after'],
                   ind_first=0, ind_chars=0, ind_left=0, ind_hang=0, jc='center',
                   rpr=build_rpr(east='宋体', ascii_font='Times New Roman', size=cfg['size']))
    return para(pp, [make_run(text, pr)])


# ============================== 插图（OOXML drawing） ==============================
IMG_CT = {'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'gif': 'image/gif',
          'bmp': 'image/bmp', 'tif': 'image/tiff', 'tiff': 'image/tiff',
          'emf': 'image/x-emf', 'wmf': 'image/x-wmf'}
IMG_NO_SIZE = ('emf', 'wmf')        # 矢量图读不出像素尺寸 → 必须显式给宽度


def image_size(path):
    """读图片像素尺寸（纯标准库实现，不依赖 Pillow）。"""
    with open(path, 'rb') as f:
        b = f.read()
    if b[:8] == b'\x89PNG\r\n\x1a\n':
        return (int.from_bytes(b[16:20], 'big'), int.from_bytes(b[20:24], 'big'))
    if b[:6] in (b'GIF87a', b'GIF89a'):
        return (int.from_bytes(b[6:8], 'little'), int.from_bytes(b[8:10], 'little'))
    if b[:2] == b'BM':
        return (abs(int.from_bytes(b[18:22], 'little', signed=True)),
                abs(int.from_bytes(b[22:26], 'little', signed=True)))
    if b[:2] == b'\xff\xd8':                       # JPEG：扫 SOF 段
        i, n = 2, len(b)
        while i + 9 < n:
            if b[i] != 0xFF:
                i += 1
                continue
            m = b[i + 1]
            if m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                     0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                return (int.from_bytes(b[i + 7:i + 9], 'big'),
                        int.from_bytes(b[i + 5:i + 7], 'big'))
            if m in (0xD8, 0xD9) or 0xD0 <= m <= 0xD7:
                i += 2
                continue
            i += 2 + int.from_bytes(b[i + 2:i + 4], 'big')
        raise ValueError('JPEG 里找不到 SOF 段：%s' % path)
    raise ValueError('不支持的图片格式（png/jpg/gif/bmp/emf/wmf）：%s' % path)


def fit_image(path, max_pt, width_pt=None):
    """算显示尺寸（磅），返回 (w_pt, h_pt)；矢量图返回 (w, None)。

    默认等比缩放到版心宽，但**不放大**：原图自然尺寸（96 DPI 换算）小于版心宽时保持原大。
    显式给 width_pt 时按给定宽度等比缩放。
    """
    ext = os.path.splitext(path)[1].lower().lstrip('.')
    if ext in IMG_NO_SIZE:
        if not width_pt:
            raise ValueError('.%s 是矢量图，读不出像素尺寸，请在内容表里显式给宽度' % ext)
        return float(width_pt), None
    pw, ph = image_size(path)
    nw, nh = pw * 72.0 / IMG_DEFAULT_DPI, ph * 72.0 / IMG_DEFAULT_DPI
    w = float(width_pt) if width_pt else min(nw, max_pt)
    return w, nh * w / nw


def add_image_part(order, parts, img_path):
    """把图片写进 zip：word/media/ + document.xml.rels 关系 + [Content_Types] Default。返回 (rId, name)。"""
    ext = os.path.splitext(img_path)[1].lower().lstrip('.')
    if ext not in IMG_CT:
        raise ValueError('不支持的图片扩展名 .%s（支持 %s）' % (ext, '/'.join(sorted(IMG_CT))))
    if not os.path.exists(img_path):
        raise FileNotFoundError('图片不存在：%s' % img_path)

    used = [int(m.group(1)) for n in parts
            for m in [re.match(r'word/media/image(\d+)\.', n)] if m]
    name = 'word/media/image%d.%s' % ((max(used) if used else 0) + 1, ext)
    data = open(img_path, 'rb').read()
    parts[name] = data
    zi = zipfile.ZipInfo(name, date_time=datetime.datetime.now().timetuple()[:6])
    zi.compress_type = zipfile.ZIP_DEFLATED
    zi.external_attr = 0o600 << 16
    order.append((zi, data))

    # ---- 关系部件 ----
    rn = 'word/_rels/document.xml.rels'
    rels = etree.fromstring(parts[rn])
    ids = {e.get('Id') for e in rels}
    i = 1
    while ('rIdImg%d' % i) in ids:
        i += 1
    rid = 'rIdImg%d' % i
    e = etree.SubElement(rels, '{%s}Relationship' % PKGNS)
    e.set('Id', rid)
    e.set('Type', REL_IMAGE)
    e.set('Target', 'media/' + os.path.basename(name))
    parts[rn] = etree.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)

    # ---- Content_Types：缺该扩展名的 Default 就补，否则 Word 报"内容有问题" ----
    cn = '[Content_Types].xml'
    ct = etree.fromstring(parts[cn])
    if not any((d.get('Extension') or '').lower() == ext
               for d in ct.findall('{%s}Default' % CTNS)):
        d = etree.Element('{%s}Default' % CTNS)
        d.set('Extension', ext)
        d.set('ContentType', IMG_CT[ext])
        ct.insert(0, d)
        parts[cn] = etree.tostring(ct, xml_declaration=True, encoding='UTF-8', standalone=True)

    return rid, name


_DRAW_TPL = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:p xmlns:w="{w}" xmlns:wp="{wp}" xmlns:a="{a}" xmlns:pic="{pic}" xmlns:r="{r}">'
    '<w:pPr><w:keepNext/>'
    '<w:spacing w:before="{before}" w:after="{after}" w:line="240" w:lineRule="auto"'
    ' w:beforeLines="0" w:beforeAutospacing="0" w:afterLines="0" w:afterAutospacing="0"/>'
    '<w:ind w:left="0" w:leftChars="0" w:right="0" w:rightChars="0"'
    ' w:firstLine="0" w:firstLineChars="0"/>'
    '<w:jc w:val="center"/></w:pPr>'
    '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
    '<wp:extent cx="{cx}" cy="{cy}"/>'
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
    '<wp:docPr id="{id}" name="{name}"/>'
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
    '<a:graphic><a:graphicData uri="{pic}"><pic:pic>'
    '<pic:nvPicPr><pic:cNvPr id="{id}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
    '<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
)


def make_figure_para(rid, w_pt, h_pt, img_id, img_name):
    """生成图片段落（居中 + 单倍行距 + keepNext）。

    ★ 行距必须 auto，**绝不能继承正文的 `lineRule=exact` 固定 20 磅** ——
      否则图片被裁成 20 磅高：`w:drawing` / `r:embed` / `word/media/*` 全部齐备，
      查 XML 查不出任何问题，但导出后完全看不见。
    """
    cfg = SPEC['figimg']
    cx = int(w_pt * EMU_PT)
    cy = int((h_pt if h_pt else w_pt * 0.6) * EMU_PT)
    xml = _DRAW_TPL.format(w=WNS, wp=WPNS, a=ANS, pic=PICNS, r=RNS,
                           before=cfg['before'], after=cfg['after'],
                           cx=cx, cy=cy, id=img_id, rid=rid, name=img_name)
    return etree.fromstring(xml.encode('utf-8'))


def build_figure(path, avail, order, parts, width_pt=None, img_id=1):
    """插图（自动等比缩放）并返回段落元素。"""
    w_pt, h_pt = fit_image(path, avail / 20.0, width_pt)
    rid, name = add_image_part(order, parts, path)
    print('[插图] %-28s → %s  %.1f×%.1f pt'
          % (os.path.basename(path), rid, w_pt, h_pt if h_pt else w_pt * 0.75))
    return make_figure_para(rid, w_pt, h_pt, img_id, name)


def make_table(header, rows, widths, avail=None):
    avail = avail or SPEC['avail']
    tcfg = SPEC['table']
    assert sum(widths) == avail, '列宽合计 %d != 版心宽 %d' % (sum(widths), avail)
    assert len(header) == len(widths), '表头列数 %d != 列宽数 %d' % (len(header), len(widths))
    for r in rows:
        assert len(r) == len(widths), '数据行列数不匹配：%r' % (r[:2],)
    tbl = etree.Element(q('tbl'))
    tpr = etree.SubElement(tbl, q('tblPr'))
    ins(tpr, mk('tblStyle', val='Table Grid'), TBL_ORDER)
    ins(tpr, mk('tblW', w=avail, type='dxa'), TBL_ORDER)
    ins(tpr, mk('jc', val='center'), TBL_ORDER)
    bd = ins(tpr, mk('tblBorders'), TBL_ORDER)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = mk(edge)
        e.set(q('val'), 'single')
        e.set(q('sz'), str(tcfg['border_sz']))   # 4 half-point = 0.5 磅
        e.set(q('space'), '0')
        e.set(q('color'), 'auto')
        bd.append(e)
    ins(tpr, mk('tblLayout', type='fixed'), TBL_ORDER)
    cm = ins(tpr, mk('tblCellMar'), TBL_ORDER)
    for edge, w in (('top', '0'), ('left', str(tcfg['cellmar'])),
                    ('bottom', '0'), ('right', str(tcfg['cellmar']))):
        e = mk(edge)
        e.set(q('w'), w)
        e.set(q('type'), 'dxa')
        cm.append(e)
    ins(tpr, mk('tblLook', val='04A0', firstRow='1', lastRow='0', firstColumn='1',
                lastColumn='0', noHBand='0', noVBand='1'), TBL_ORDER)
    grid = etree.SubElement(tbl, q('tblGrid'))
    for w in widths:
        etree.SubElement(grid, q('gridCol')).set(q('w'), str(w))

    for ri, cells in enumerate([header] + rows):
        tr = etree.SubElement(tbl, q('tr'))
        trpr = etree.SubElement(tr, q('trPr'))
        ins(trpr, mk('cantSplit'), TRPR_ORDER)      # 行不跨页
        if ri == 0:
            ins(trpr, mk('tblHeader'), TRPR_ORDER)  # 表头重复
        for ci, txt in enumerate(cells):
            tc = etree.SubElement(tr, q('tc'))
            cpr = etree.SubElement(tc, q('tcPr'))
            ins(cpr, mk('tcW', w=widths[ci], type='dxa'), TCPR_ORDER)
            ins(cpr, mk('vAlign', val='center'), TCPR_ORDER)
            sz = tcfg['hdr'] if ri == 0 else tcfg['text']
            bold = True if ri == 0 else None
            ppr = build_ppr(before=20, after=20, line=260, line_rule='auto',
                            ind_first=0, ind_chars=0, ind_left=0, ind_hang=0,
                            jc='center' if ri == 0 else 'left',
                            rpr=build_rpr(east='宋体', ascii_font='Times New Roman', size=sz, bold=bold))
            p = etree.SubElement(tc, q('p'))
            p.append(ppr)
            p.append(make_run(txt, build_rpr(east='宋体', ascii_font='Times New Roman',
                                             size=sz, bold=bold)))
    return tbl


def build_revision_table(rows):
    """修订记录表：表头小五 9pt / 表体五号 10.5pt，列宽合计 = 版心宽"""
    rcfg = SPEC['revision']
    while len(rows) < 6:
        rows.append(['', '', '', ''])
    return make_table(['日期', '版本', '说明', '作者'], rows, rcfg['widths'])


# ============================== 目录 ==============================
def derive_toc_items(content):
    items = []
    for it in content:
        if it[0] == 'H1':
            items.append((1, it[1]))
        elif it[0] == 'H2':
            items.append((2, it[1]))
    return items


def cjk_w(s, pt):
    return sum(pt if ord(c) > 0x2E80 else pt * 0.5 for c in s)


def est_body_pages(content, lines_per_page=34, text_w=415.3):
    lines = 0.0
    pages = {}
    for it in content:
        k = it[0]
        if k == 'H1':
            lines += 1.5 + 17 + 16.5
        elif k == 'H2':
            lines += 1.2 + 13 + 13
        elif k == 'H3':
            lines += 1.1 + 12 + 6
        elif k in ('P', 'C'):
            lines += max(1.0, math.ceil(cjk_w(it[1], 12.0) / text_w))
        elif k in ('TBLCAP', 'FIGCAP'):
            lines += 1.0 + 12 + 6
        elif k == 'FIG':
            w = it[2] if len(it) > 2 else None
            try:
                _, h = fit_image(it[1], SPEC['avail'] / 20.0, w)
            except Exception:
                h = None
            lines += ((h or (w or SPEC['avail'] / 20.0) * 0.75) / 20.0) + 0.6
        elif k == 'T':
            h = 1.0
            for r in [it[1]] + list(it[2]):
                mx = 1
                for ci, c in enumerate(r):
                    wpt = it[3][ci] / 20.0
                    mx = max(mx, math.ceil(cjk_w(c, 10.5) / max(20.0, wpt - 6)))
                h += mx * 1.55 + 0.15
            lines += h
        if k in ('H1', 'H2'):
            pages[it[1]] = int(lines // lines_per_page) + 1
    return pages


def _strip_outer_field(toc1_proto):
    """从 TOC1 原型上剥离外层 TOC 域 begin 组（fldChar begin / instrText TOC… / fldChar separate）

    ⚠️ 只能剥「TOC 域自己的」这三段。首个 TOC 条目段落紧跟其后的还有该条目自己的
    HYPERLINK begin —— 若按"扫到域就收"的写法会把 HYPERLINK begin 一起剥掉，
    结果是每条条目少一个 begin、多一个 end，全文域字符不配平（实测 begin=15 end=17）。
    """
    runs = [c for c in toc1_proto if c.tag == q('r')]
    outer = []
    for r in runs:
        fc = r.find(q('fldChar'))
        it = r.find(q('instrText'))
        if len(outer) == 0:
            if fc is not None and fc.get(q('fldCharType')) == 'begin':
                outer.append(r)
                continue
            return [], toc1_proto
        if len(outer) == 1:
            if it is not None and (it.text or '').strip().startswith('TOC'):
                outer.append(r)
                continue
            return [], toc1_proto
        if len(outer) == 2:
            if fc is not None and fc.get(q('fldCharType')) == 'separate':
                outer.append(r)
                break
            return [], toc1_proto
    if len(outer) != 3:
        return [], toc1_proto
    keep = [copy.deepcopy(r) for r in outer]
    for r in outer:
        toc1_proto.remove(r)
    return keep, toc1_proto


def build_toc_proto_tail(proto, text, bookmark, page):
    """改写目录条目原型：w:t[0]=条目文本，w:t[1]=页码，instrText 书签改名"""
    ts = list(proto.iter(q('t')))
    if len(ts) >= 2:
        ts[0].text = text
        ts[1].text = str(page)
    elif len(ts) == 1:
        ts[0].text = text
    for it in proto.iter(q('instrText')):
        it.text = re.sub(r'_Toc[0-9A-Za-z]+', bookmark, it.text or '')
    return proto


def synth_toc_field_parts():
    """基底没有目录时，合成一个可被 Word 重建的 TOC 域"""
    begin = etree.Element(q('r'))
    etree.SubElement(begin, q('fldChar')).set(q('fldCharType'), 'begin')
    instr = etree.Element(q('r'))
    it = etree.SubElement(instr, q('instrText'))
    it.set(XS, 'preserve')
    it.text = ' TOC \\o "1-2" \\h \\z \\u '
    sep = etree.Element(q('r'))
    etree.SubElement(sep, q('fldChar')).set(q('fldCharType'), 'separate')
    ph = etree.Element(q('r'))
    t = etree.SubElement(ph, q('t'))
    t.set(XS, 'preserve')
    t.text = '（在 Word 中按 F9 更新目录）'
    end = etree.Element(q('r'))
    etree.SubElement(end, q('fldChar')).set(q('fldCharType'), 'end')
    return begin, instr, sep, ph, end


# ============================== zip 读写 ==============================
def load_parts(path):
    z = zipfile.ZipFile(path)
    order = [(i, z.read(i.filename)) for i in z.infolist()]
    z.close()
    return order, {i.filename: d for i, d in order}


def save_parts(path, order, parts):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it, d in order:
            zi = zipfile.ZipInfo(it.filename, date_time=it.date_time)
            zi.compress_type = it.compress_type
            zi.external_attr = it.external_attr
            zo.writestr(zi, parts.get(it.filename, d))


# ============================== 主流程 ==============================
def rebuild(base, out, content, cover_texts=None, revision_rows=None,
            header_text=None, header_old=None, doc_title=None,
            avail=None, toc_items=None, banner=True, backup=True):
    """以 base 为基底，按《规范》重建 out。

    参数
    ----
    base        : 基底 docx（提供 numbering/styles/header/footer/sectPr/封面与目录原型）
    out         : 目标 docx
    content     : 内容表 [(类型, ...), ...]，类型见 CONTENT_TYPES
    cover_texts : 封面 4 个文案 [主标题, 文档类型, 下部模块名, 年月]，按顺序替换基底封面里的非空段落
    revision_rows: 修订记录行 [['日期','版本','说明','作者'], ...]
    header_text : 页眉文案（如 "XXX - 运维手册"）
    header_old  : 页眉里要被替换的原文案（不传则自动替换最长的那段文本）
    doc_title   : 文档属性标题
    toc_items   : 目录条目 [(层级, 文本)]，不传则由 H1/H2 推导
    """
    avail = avail or SPEC['avail']
    revision_rows = revision_rows or []
    cover_texts = cover_texts or []
    toc_items = toc_items or derive_toc_items(content)

    # 提醒：章标题的编号由基底 numbering 自动生成（%1、），内容表里 H1 文本不要再写编号
    for it in content:
        if it[0] == 'H1' and re.match(r'^\s*\d+\s*[、.．]', it[1]):
            print('[提醒] H1 文本「%s」自带编号，基底的自动编号会再叠一层 → 建议把编号去掉'
                  % it[1])
            break

    if backup and os.path.exists(out):
        bak = os.path.join(os.path.dirname(out), '.workbuddy', 'WPS_Backup')
        os.makedirs(bak, exist_ok=True)
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M')
        stem = os.path.splitext(os.path.basename(out))[0]
        dst = os.path.join(bak, '%s_重建前_%s.docx' % (stem, ts))
        shutil.copy2(out, dst)
        print('[备份] 原稿 →', dst, os.path.getsize(dst), 'B')

    order, parts = load_parts(base)
    doc = etree.fromstring(parts['word/document.xml'])
    sid = resolve_style_ids(parts['word/styles.xml'])
    L = detect_layout(doc, sid)
    kids = L['kids']

    if not sid.get('heading 1') or not sid.get('heading 2'):
        raise RuntimeError('基底 styles.xml 里找不到 heading 1 / heading 2 样式')
    h1_num = detect_h1_numid(doc, sid.get('heading 1'))
    print('[基底] heading1=%s heading2=%s toc1=%s toc2=%s | H1 numId=%s'
          % (sid.get('heading 1'), sid.get('heading 2'), sid.get('toc 1'), sid.get('toc 2'), h1_num))

    # ---- 封面 ----
    COVER = [copy.deepcopy(kids[i]) for i in L['cover']]
    slots = [p for p in COVER if ptext(p)]
    if len(cover_texts) > len(slots):
        raise RuntimeError('封面文案 %d 条 > 基底封面文本段 %d 个' % (len(cover_texts), len(slots)))
    for txt, el in zip(cover_texts, slots):
        ts = list(el.iter(q('t')))
        if ts:
            ts[0].text = txt
            for extra in ts[1:]:
                extra.text = ''
    print('[封面] 替换 %d 个文本段：%s' % (len(cover_texts), ' | '.join(cover_texts)))

    # ---- 目录原型 ----
    if L['toc1_idx'] is not None and L['toc2_idx'] is not None:
        T1 = copy.deepcopy(kids[L['toc1_idx']])
        T2 = copy.deepcopy(kids[L['toc2_idx']])
        OPEN_RUNS, T1 = _strip_outer_field(T1)
        TAIL = copy.deepcopy(kids[L['toc_tail']]) if L['toc_tail'] is not None else None
        have_toc = True
    else:
        T1 = T2 = TAIL = None
        OPEN_RUNS, have_toc = [], False
        print('[目录] 基底无 TOC 原型 → 将合成 TOC 域（Word 打开后 F9 重建）')

    SECT_MID = copy.deepcopy(kids[L['sect_mid']]) if L['sect_mid'] is not None else None
    DIRP = copy.deepcopy(kids[L['dir_para']]) if L['dir_para'] is not None else None
    EMPTY = [copy.deepcopy(kids[i]) for i in range(L['rev_tbl'] + 1, L['sect_mid'])] \
        if (L['sect_mid'] is not None and L['sect_mid'] > L['rev_tbl'] + 1) else []
    SECT = L['sect']

    # ---- 清空正文 ----
    for ch in kids:
        if ch is not SECT:
            L['body'].remove(ch)

    outEls = list(COVER)
    outEls.append(build_revision_table(revision_rows))
    outEls += EMPTY
    if SECT_MID is not None:
        outEls.append(SECT_MID)
    if DIRP is not None:
        outEls.append(DIRP)

    # ---- 目录 ----
    pages = est_body_pages(content)
    if have_toc:
        for i, (lv, txt) in enumerate(toc_items):
            p = copy.deepcopy(T1 if lv == 1 else T2)
            bm = '_Toc%05d' % (90001 + i)
            build_toc_proto_tail(p, txt, bm, pages.get(txt, 1))
            if i == 0 and OPEN_RUNS:
                cur = p.find(q('pPr'))
                for r in OPEN_RUNS:
                    nxt = copy.deepcopy(r)
                    cur.addnext(nxt)
                    cur = nxt
            outEls.append(p)
        if TAIL is not None:
            outEls.append(TAIL)
    else:
        begin, instr, sep, ph, end = synth_toc_field_parts()
        p = etree.Element(q('p'))
        ppr = build_ppr(style=sid.get('toc 1'))
        p.append(ppr)
        for r in (begin, instr, sep, ph, end):
            p.append(r)
        outEls.append(p)

    # ---- 正文 ----
    img_seq = 0
    for it in content:
        k = it[0]
        if k == 'H1':
            outEls.append(H1(it[1], sid, h1_num))
        elif k == 'H2':
            outEls.append(H2(it[1], sid))
        elif k == 'H3':
            outEls.append(H3(it[1], sid))
        elif k == 'P':
            outEls.append(BODY(it[1]))
        elif k == 'C':
            outEls.append(CMD(it[1]))
        elif k == 'TBLCAP':
            outEls.append(CAP(it[1], 'T'))
        elif k == 'FIGCAP':
            outEls.append(CAP(it[1], 'F'))
        elif k == 'FIG':
            # ('FIG', 路径) / ('FIG', 路径, 宽度pt) / ('FIG', 路径, 宽度pt, 图题)
            img_seq += 1
            w = it[2] if len(it) > 2 else None
            outEls.append(build_figure(it[1], avail, order, parts, w, img_seq))
            if len(it) > 3 and it[3]:
                outEls.append(CAP(it[3], 'F'))
        elif k == 'T':
            outEls.append(make_table(it[1], it[2], it[3], avail))
        else:
            raise ValueError('未知内容类型 %r' % (k,))

    for el in outEls:
        L['body'].append(el)
    L['body'].append(SECT)

    parts['word/document.xml'] = etree.tostring(doc, xml_declaration=True,
                                                encoding='UTF-8', standalone=True)

    # ---- Phase B：页眉 / settings / core ----
    if header_text:
        hname = next((n for n in parts if re.match(r'word/header\d*\.xml$', n)), None)
        if hname:
            h = parts[hname].decode('utf-8')
            if header_old and header_old in h:
                h = h.replace(header_old, header_text)
            else:
                # 自动：替换页眉里最长的一段 w:t
                cand = re.findall(r'<w:t[^>]*>([^<]+)</w:t>', h)
                cand = [c for c in cand if c.strip()]
                if cand:
                    old = max(cand, key=len)
                    if old != header_text:
                        h = h.replace(old, header_text)
                        print('[页眉] "%s" → "%s"' % (old, header_text))
                else:
                    print('[页眉] 警告：页眉里没有可替换的文本')
            parts[hname] = h.encode('utf-8')

    s = parts['word/settings.xml'].decode('utf-8')
    m = re.search(r'<w:compatSetting[^>]*w:name="compatibilityMode"[^>]*/>', s)
    if m:
        s = s[:m.start()] + re.sub(r'w:val="\d+"', 'w:val="15"', m.group(0)) + s[m.end():]
    if '<w:updateFields' not in s:
        if '<w:footnotePr' in s:
            s = s.replace('<w:footnotePr', '<w:updateFields w:val="true"/><w:footnotePr', 1)
        else:
            s = s.replace('</w:settings>', '<w:updateFields w:val="true"/></w:settings>', 1)
    parts['word/settings.xml'] = s.encode('utf-8')

    if doc_title and 'docProps/core.xml' in parts:
        c = parts['docProps/core.xml'].decode('utf-8')
        if '<dc:title>' in c:
            c = re.sub(r'<dc:title>.*?</dc:title>',
                       '<dc:title>%s</dc:title>' % doc_title, c, flags=re.S)
            parts['docProps/core.xml'] = c.encode('utf-8')

    # ---- 写出 ----
    save_parts(out, order, parts)

    # ---- 断言 ----
    _verify(out, toc_items, have_toc, sid)
    print('[产出] %s  %d B' % (out, os.path.getsize(out)))
    return out


def _verify(path, toc_items, have_toc, sid):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    for nm in names:
        if nm.endswith('.xml') or nm.endswith('.rels'):
            etree.fromstring(z.read(nm))
    x = z.read('word/document.xml').decode('utf-8')
    doc = etree.fromstring(z.read('word/document.xml'))
    media = sorted(n for n in names if n.startswith('word/media/'))
    n_beg = x.count('w:fldCharType="begin"')
    n_end = x.count('w:fldCharType="end"')
    n_h1 = x.count('<w:pStyle w:val="%s"/>' % sid.get('heading 1'))
    n_h2 = x.count('<w:pStyle w:val="%s"/>' % sid.get('heading 2'))

    # ---- 插图自检（三件套齐全 + 行距未被裁切）----
    n_draw = x.count('<w:drawing>')
    if n_draw or media:
        # ⚠️ 不能要求 n_draw == len(media)：Word COM 保存时会把**重复图片去重**
        #    （同一文件插入多次 → 合并成一个 media + 一条关系，多个 drawing 共用同一 rId），
        #    所以只要求"没有未被引用的 media"。
        assert n_draw >= len(media), 'drawing 数 %d < media 文件数 %d（有图未被引用）' % (n_draw, len(media))
        rels = etree.fromstring(z.read('word/_rels/document.xml.rels'))
        tgts = {e.get('Target') for e in rels if (e.get('Type') or '') == REL_IMAGE}
        for m in media:
            assert m.split('word/')[-1] in tgts, 'media %s 没有对应的关系条目' % m
        ct = z.read('[Content_Types].xml').decode('utf-8')
        for ext in {os.path.splitext(m)[1].lstrip('.').lower() for m in media}:
            assert 'Extension="%s"' % ext in ct, '[Content_Types] 缺 .%s 的 Default（Word 会报内容有问题）' % ext
        bad = []
        for el in doc.iter(q('p')):
            if el.find('.//' + q('drawing')) is None:
                continue
            pp = el.find(q('pPr'))
            sp = pp.find(q('spacing')) if pp is not None else None
            if sp is None or sp.get(q('lineRule')) == 'exact':
                bad.append(ptext(el)[:20])
        assert not bad, '有 %d 个图片段是固定行距 → 图片会被裁掉：%r' % (len(bad), bad[:3])
    z.close()
    assert n_beg == n_end, '域字符不配平 begin=%d end=%d' % (n_beg, n_end)
    if have_toc:
        n_toc = x.count('TOC \\o')
        assert n_toc == 1, '外层 TOC 域指令出现 %d 次（期望 1）' % n_toc
    print('[自检] XML 全部可解析 ｜ 域 begin=%d end=%d ｜ H1 %d ｜ H2 %d ｜ 表 %d ｜ 图 %d ｜ 目录条目 %d'
          % (n_beg, n_end, n_h1, n_h2, x.count('<w:tbl>'), len(media), len(toc_items)))
