# -*- coding: utf-8 -*-
"""dump_para_spacing.py —— 查"标题段前/段后有没有被 Word 吃掉"

用法：
    python dump_para_spacing.py <docx> [--keys 运维概述,1.1]

输出每个命中的段落：pStyle / 完整 spacing 属性 / 是否带 before/after /
beforeLines·afterLines·beforeAutospacing·afterAutospacing（这四个缺失 → Word 保存后 before/after 会被清掉）。

判读：
  spacing 里**没有 before/after** → 段距已被 Word 删除（成品缺陷），要回构建脚本补 sp_lines=True。
  有 before/after 但缺 beforeLines/*Autospacing → 下次再被 Word 打开保存仍可能丢。
"""
import sys, io, os, re, zipfile
import argparse
from lxml import etree

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NEED = ('beforeLines', 'afterLines', 'beforeAutospacing', 'afterAutospacing')


def q(t):
    return W + t


def doc_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read('word/document.xml')


def styles_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read('word/styles.xml') if 'word/styles.xml' in z.namelist() else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--keys', default=None, help='逗号分隔的段落文本（默认：命中标题样式的前若干段）')
    ap.add_argument('--all', action='store_true', help='列出所有使用标题样式的段落')
    a = ap.parse_args()

    root = etree.fromstring(doc_xml(a.docx))
    sroot = etree.fromstring(styles_xml(a.docx)) if styles_xml(a.docx) else None

    # 找出哪些 pStyle 的名字像标题
    head_ids = set()
    if sroot is not None:
        for st in sroot.findall(q('style')):
            nm = st.find(q('name'))
            nm = (nm.get(q('val')) or '').lower() if nm is not None else ''
            if nm.startswith('heading') or nm.startswith('标题'):
                head_ids.add(st.get(q('styleId')))
    print('标题样式 styleId =', sorted(head_ids) or '(未找到，按--keys匹配)')

    keys = [k.strip() for k in a.keys.split(',')] if a.keys else None
    n = 0
    for el in root.iter(q('p')):
        t = ''.join(x.text or '' for x in el.iter(q('t'))).strip()
        if not t:
            continue
        pp = el.find(q('pPr'))
        sid = None
        if pp is not None:
            s = pp.find(q('pStyle'))
            sid = s.get(q('val')) if s is not None else None
        hit = (t in keys) if keys else (sid in head_ids)
        if not hit:
            continue
        n += 1
        sp = pp.find(q('spacing')) if pp is not None else None
        attrs = {etree.QName(k).localname: v for k, v in sp.attrib.items()} if sp is not None else {}
        miss = [k for k in NEED if k not in attrs]
        flag = 'OK' if ('before' in attrs and 'after' in attrs) else '!! 缺 before/after'
        print('--- 「%s」  pStyle=%s  %s' % (t[:40], sid, flag))
        print('    spacing:', attrs if attrs else '(无 spacing 元素)')
        if miss:
            print('    缺配套属性:', miss)
        if a.all is False and keys is None and n >= 8:
            break
    print('命中 %d 段' % n)
    if n == 0:
        print('提示：用 --keys "某标题文本" 指定段落。')


if __name__ == '__main__':
    main()
