# -*- coding: utf-8 -*-
"""insert_toc.py —— 给已有 .docx 插入 / 重建「目录」页

用法
----
python insert_toc.py 目标.docx                       # 默认 1-2 级，标题「目录」
python insert_toc.py 目标.docx --levels 3            # 收录到三级
python insert_toc.py 目标.docx --title 目  录 --avail 8306

幂等
----
重跑时先定位并剔除上一轮的目录块，再重新插入，所以反复执行不会出现双目录。
识别旧目录用的是 **Word 一定会保留的特征**：TOC 域的 `instrText` + 紧邻其前的「目录」标题段。
⚠️ **不要**用 `w:rsidR` 标记判旧目录 —— Word COM 打开并保存时会把 rsid 重写掉，
   标记随之失效，重跑就插出第二份目录（实测 Word 报 `TablesOfContents 数 = 2`）。
⚠️ 也不要用"文本等于条目名"判断 —— 条目段把**页码也拼进文本**了
   （`''.join(el.itertext())` 得到 `'1、 项目概况12'`），永远匹配不上。

页码
----
条目先写**占位页码 1**，插入后必须跑一次 Word COM 更新域拿真实页码：

    python refresh_toc.py 目标.docx --pdf out.pdf

（Word 的 `TablesOfContents(1).Update()` 会重排页码与制表位，比每次手工两遍编译可靠。）
"""
import os
import sys
import argparse

from lxml import etree

sys.stdout.reconfigure(encoding='utf-8')
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import docxlib as D  # noqa: E402

MARK = '00A0C001'                       # 合法 rsid，仅作视觉留痕；★ 幂等**不**依赖它（Word 会重写 rsid）


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


# ---------------------------------- 段落构造 ----------------------------------
def title_xml(title):
    return (
        '<w:p xmlns:w="%s" w:rsidR="%s"><w:pPr>'
        '<w:pageBreakBefore/>'
        '<w:spacing w:before="240" w:after="240" w:line="240" w:lineRule="auto"'
        ' w:beforeLines="0" w:beforeAutospacing="0" w:afterLines="0" w:afterAutospacing="0"/>'
        '<w:ind w:left="0" w:leftChars="0" w:right="0" w:rightChars="0"'
        ' w:firstLine="0" w:firstLineChars="0"/>'
        '<w:jc w:val="center"/>'
        '<w:rPr><w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体"/><w:b/><w:bCs/>'
        '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体"/><w:b/><w:bCs/>'
        '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>'
        '<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
        % (D.WNS, MARK, esc(title))
    )


def field_run(kind, text=None):
    if kind == 'begin':
        return '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    if kind == 'instr':
        return '<w:r><w:instrText xml:space="preserve">%s</w:instrText></w:r>' % esc(text)
    if kind == 'sep':
        return '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    return '<w:r><w:fldChar w:fldCharType="end"/></w:r>'


def entry_xml(text, level, sid, avail, pre='', post=''):
    """一条目录条目：文本 + 右对齐点线制表位 + 页码占位。"""
    st = sid.get('toc %d' % level)
    pstyle = '<w:pStyle w:val="%s"/>' % st if st else ''
    ind = 0 if level == 1 else 420 * (level - 1)
    return (
        '<w:p xmlns:w="%s" w:rsidR="%s"><w:pPr>%s'
        '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="%d"/></w:tabs>'
        '<w:spacing w:line="240" w:lineRule="auto"/>'
        '<w:ind w:left="%d" w:leftChars="%d" w:right="0" w:rightChars="0"'
        ' w:firstLine="0" w:firstLineChars="0"/>'
        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"'
        ' w:eastAsia="宋体"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:pPr>'
        '%s<w:r><w:t xml:space="preserve">%s</w:t></w:r>'
        '<w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r>%s</w:p>'
        % (D.WNS, MARK, pstyle, avail, ind, ind, pre, esc(text), post)
    )


# ---------------------------------- 探测 ----------------------------------
def pstyle_of(el, sid):
    pp = el.find(D.q('pPr'))
    if pp is None:
        return None
    ps = pp.find(D.q('pStyle'))
    return ps.get(D.q('val')) if ps is not None else None


def collect_headings(body, sid, max_level):
    """按样式名扫正文直接子元素，收集 (层级, 文本)。"""
    want = {sid.get('heading %d' % lv): lv for lv in range(1, max_level + 1)}
    out = []
    for el in body:
        if el.tag != D.q('p'):
            continue
        lv = want.get(pstyle_of(el, sid))
        if lv:
            out.append((lv, D.ptext(el)))
    return out


def find_insert_anchor(body, sid, kids):
    """返回 (anchor元素, first_h元素)。目录插在 anchor 之前。"""
    h1 = sid.get('heading 1')
    first_h = next((el for el in kids if el.tag == D.q('p') and pstyle_of(el, sid) == h1), None)
    if first_h is None:
        return (kids[-1] if kids else None), None
    # 正文之前若有承载 sectPr 的分节符段，目录要插在它**之前**（归入上一节）
    idx = kids.index(first_h)
    carrier = None
    for i in range(idx - 1, -1, -1):
        pp = kids[i].find(D.q('pPr'))
        if pp is not None and pp.find(D.q('sectPr')) is not None:
            carrier = kids[i]
            break
    return (carrier if carrier is not None else first_h), first_h


def find_old_block(body, sid, title):
    """定位上一轮插入的目录块：从正文起点**向前回溯**，连续收集属于目录的段落。

    收集判据（三者之一，遇到第一个不属于目录的段落就停）：
      ① 文本 == 「目录」标题；② 样式是 toc 1/2/3；③ 含域字符且文本为空。

    ⚠️ **不要靠 `w:rsidR` 标记做幂等** —— Word COM 打开并保存时会把 rsid 重写掉，
    标记随之失效，重跑就插出第二份目录（实测 Word 报 `TablesOfContents 数 = 2`）。
    ⚠️ **也不要「从 TOC begin 向后找第一个 `fldChar end`」** —— Word 更新目录后，
    首条条目段内部就带了条目自己（HYPERLINK/PAGEREF）的 `end`，会被误判成外层域的结尾，
    结果只删掉 2 段、其余目录段全部残留（实测域字符 begin=8 end=9）。
    ⚠️ 同样不能用"文本等于条目名"判断 —— 条目段的文本把**页码也拼进去**了。
    """
    kids = list(body)
    anchor, _ = find_insert_anchor(body, sid, kids)
    if anchor is None:
        return []
    idx = kids.index(anchor)
    toc_ids = {sid.get('toc 1'), sid.get('toc 2'), sid.get('toc 3')} - {None}
    norm = lambda s: (s or '').replace(' ', '').replace('\u3000', '')   # noqa: E731

    out = []
    for i in range(idx - 1, max(-1, idx - 40), -1):
        el = kids[i]
        if el.tag != D.q('p'):
            break
        txt = D.ptext(el)
        sty = pstyle_of(el, sid)
        has_fld = el.find('.//' + D.q('fldChar')) is not None
        if norm(txt) == norm(title) or (sty and sty in toc_ids):
            out.append(el)
        elif has_fld and not txt:
            out.append(el)              # 外层 TOC 域 end 的载体段（空文本）
        else:
            break
    out.reverse()
    return out


# ---------------------------------- 主流程 ----------------------------------
def run(path, title='目录', max_level=2, avail=None, instr=None):
    avail = avail or D.SPEC['avail']
    if instr is None:                      # ★ 必须 raw 字符串：'\u' 会被 Python 当转义
        instr = r' TOC \o "1-%d" \h \z \u ' % max_level

    order, parts = D.load_parts(path)
    doc = etree.fromstring(parts['word/document.xml'])
    sid = D.resolve_style_ids(parts['word/styles.xml'])
    body = doc.find(D.q('body'))
    if body is None:
        raise SystemExit('document.xml 里没有 body')

    if not sid.get('heading 1'):
        raise SystemExit('styles.xml 里找不到 heading 1 —— 请先在 Word 里把章标题套上标题样式')

    # ---- 幂等：先剔除上一轮插入的目录块 ----
    old = find_old_block(body, sid, title)
    for el in old:
        body.remove(el)
    if old:
        print('[幂等] 剔除上一轮的目录块（%d 个节点）' % len(old))

    kids = list(body)                      # ★ 剔除后**必须重算**索引
    items = collect_headings(body, sid, max_level)
    if not items:
        print('[警告] 没扫到任何标题（heading 1..%d），将只插入一个空的 TOC 域' % max_level)

    anchor, first_h = find_insert_anchor(body, sid, kids)
    if anchor is None:
        raise SystemExit('找不到插入锚点')

    # ---- 组块 ----
    blocks = [etree.fromstring(title_xml(title).encode('utf-8'))]
    if items:
        for i, (lv, txt) in enumerate(items):
            pre = ''
            post = ''
            if i == 0:
                pre = field_run('begin') + field_run('instr', instr) + field_run('sep')
            if i == len(items) - 1:
                post = field_run('end')
            blocks.append(etree.fromstring(entry_xml(txt, lv, sid, avail, pre, post).encode('utf-8')))
    else:
        p = etree.fromstring(entry_xml('', 1, sid, avail,
                                       field_run('begin') + field_run('instr', instr) + field_run('sep'),
                                       field_run('end')).encode('utf-8'))
        blocks.append(p)

    for b in blocks:                       # 顺序 addprevious，保持次序
        anchor.addprevious(b)

    parts['word/document.xml'] = etree.tostring(doc, xml_declaration=True,
                                                encoding='UTF-8', standalone=True)

    # ★ 校验放在写盘**之前**：否则断言失败时半成品已经落盘（实测会把文档写成双目录）
    x = parts['word/document.xml'].decode('utf-8')
    n_beg = x.count('w:fldCharType="begin"')
    n_end = x.count('w:fldCharType="end"')
    n_toc = x.count('TOC \\o')
    assert n_beg == n_end, '域字符不配平 begin=%d end=%d' % (n_beg, n_end)
    assert n_toc == 1, 'TOC 指令出现 %d 处（期望 1）—— 旧目录没被清理干净' % n_toc

    D.save_parts(path, order, parts)

    n1 = sum(1 for lv, _ in items if lv == 1)
    n2 = len(items) - n1
    print('[目录] 已插入「%s」+ %d 条（一级 %d / 二级及以下 %d）｜ 锚点=%s'
          % (title, len(items), n1, n2, '分节符前' if first_h is not None and anchor is not first_h
             else '首个章标题前'))
    print('[提示] 页码现为占位值 1，请执行：python refresh_toc.py "%s" --pdf out.pdf' % path)
    return len(items)


def main():
    ap = argparse.ArgumentParser(description='给已有 docx 插入/重建目录页')
    ap.add_argument('docx', help='目标 docx（就地修改）')
    ap.add_argument('--title', default='目录', help='目录页标题，默认「目录」')
    ap.add_argument('--levels', type=int, default=2, help='收录层级 1/2/3，默认 2')
    ap.add_argument('--avail', type=int, default=None, help='版心宽 twip，默认 8306')
    a = ap.parse_args()
    if not os.path.exists(a.docx):
        raise SystemExit('文件不存在：%s' % a.docx)
    run(a.docx, a.title, a.levels, a.avail)


if __name__ == '__main__':
    main()
