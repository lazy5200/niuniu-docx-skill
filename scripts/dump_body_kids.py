#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""dump_body_kids.py —— dump docx 的 body 直接子元素（索引 / 标签 / 样式 / 是否含域 / 是否含段内 sectPr）

用途（重要）：
    拿模板当基底重建文档时，大家都会写 `kids = list(body)` 再按 `kids[i]` 摘封面 / 修订表 /
    目录原型。但 body 的直接子元素除 `w:p` 外还有 `w:tbl` 和 `w:sectPr`，
    所以从第一个表之后 `kids[i]` 就 **不等于** 第 i 个段落 —— 按"段号"猜索引必然取错素材
    （典型后果：'目录' 标题段丢失、TOC1/TOC2 原型互换、外层 TOC 域被复制到每条条目上、
    目录节 sectPr 丢失，最终 Word 报 TablesOfContents.Count = 0、目录全是"错误!未定义书签。"）。

用法：
    python dump_body_kids.py <a.docx> [b.docx ...]
    # 或按索引区间细看某个段的 run 序列：
    python dump_body_kids.py <a.docx> --runs 18 19 20 21 37
"""
import sys
import zipfile

try:
    from lxml import etree
except ImportError:  # 兜底：标准库 xml.etree
    import xml.etree.ElementTree as etree

WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = '{%s}' % WNS
FC = {'begin': 'B', 'separate': 'S', 'end': 'E'}


def para_summary(el):
    """段落摘要：styleId / numPr / 段内 sectPr / 是否含域 / 文本"""
    st = ''
    ppr = el.find(W + 'pPr')
    if ppr is not None:
        pst = ppr.find(W + 'pStyle')
        st = ' style=%s' % (pst.get(W + 'val') if pst is not None else '-')
        npr = ppr.find(W + 'numPr')
        if npr is not None:
            nid = npr.find(W + 'numId')
            st += ' numPr=%s' % (nid.get(W + 'val') if nid is not None else '?')
        if ppr.find(W + 'sectPr') is not None:
            st += ' [含段内sectPr]'
    if el.find('.//' + W + 'fldChar') is not None:
        st += ' [有域]'
    txt = ''.join(t.text or '' for t in el.iter(W + 't'))[:44]
    return st, txt


def run_detail(el):
    """把一个段落的子元素展开成可读串：B/S/E + {域指令} + "文本" + <tab>"""
    out = []
    for c in el:
        tag = etree.QName(c).localname
        if tag == 'pPr':
            ppr = c
            pst = ppr.find(W + 'pStyle')
            extra = ', 有sectPr' if ppr.find(W + 'sectPr') is not None else ''
            out.append('pPr(style=%s%s)' % (pst.get(W + 'val') if pst is not None else '-', extra))
        elif tag == 'r':
            inner = []
            for cc in c:
                t = etree.QName(cc).localname
                if t == 'fldChar':
                    inner.append(FC.get(cc.get(W + 'fldCharType'), '?'))
                elif t == 'instrText':
                    inner.append('{%s}' % (cc.text or '').strip())
                elif t == 't':
                    inner.append('"%s"' % (cc.text or ''))
                elif t == 'tab':
                    inner.append('<tab>')
            if inner:
                out.append('r[%s]' % ' '.join(inner))
        else:
            out.append('<%s>' % tag)
    return ' '.join(out)


def dump(path, runs=None):
    print('=' * 78)
    print(path)
    print('=' * 78)
    z = zipfile.ZipFile(path)
    root = etree.fromstring(z.read('word/document.xml'))
    z.close()
    body = root.find(W + 'body')
    kids = list(body)
    print('body 直接子元素 %d 个 ｜ 段 %d ｜ 表 %d ｜ body 级 sectPr %d' % (
        len(kids),
        sum(1 for e in kids if etree.QName(e).localname == 'p'),
        sum(1 for e in kids if etree.QName(e).localname == 'tbl'),
        sum(1 for e in kids if etree.QName(e).localname == 'sectPr')))
    for i, el in enumerate(kids):
        tag = etree.QName(el).localname
        if tag == 'p':
            st, txt = para_summary(el)
        elif tag == 'tbl':
            trs = el.findall(W + 'tr')
            st, txt = ' 行数=%d 首行列数=%d' % (
                len(trs), len(trs[0].findall(W + 'tc')) if trs else 0), ''
        elif tag == 'sectPr':
            st, txt = ' <body级节>', ''
        else:
            st, txt = '', ''
        print('%3d  %-8s%s  %r' % (i, tag, st, txt))
    if runs:
        for i in runs:
            print('--- kids[%d] run 序列 ---' % i)
            print('   ', run_detail(kids[i]))


if __name__ == '__main__':
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(1)
    run_ids = []
    if '--runs' in args:
        k = args.index('--runs')
        run_ids = [int(x) for x in args[k + 1:]]
        args = args[:k]
    for p in args:
        dump(p, run_ids)
        print()
