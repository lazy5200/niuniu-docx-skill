# -*- coding: utf-8 -*-
"""导出 docx 的段落与表格清单，用于改写前的结构比对和改写后的复核。

用法:
    python dump_docx.py <文件路径> [> out.txt]

输出格式:
    P[idx] style=... |font=...|size=...|b=...| align=... :: 段落文本
    === TABLE n rows=.. cols=.. style=.. ===
      R0: 单元格 | 单元格
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

path = sys.argv[1]
doc = Document(path)
body = doc.element.body


def iter_block(parent):
    for child in parent.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, doc)
        elif child.tag == qn('w:tbl'):
            yield Table(child, doc)


idx_p = idx_t = 0
for block in iter_block(body):
    if isinstance(block, Paragraph):
        txt = block.text.strip()
        info = ""
        if block.runs:
            f = block.runs[0].font
            info = f"|font={f.name}|size={f.size.pt if f.size else None}|b={f.bold}|"
        print(f"P[{idx_p}] style={block.style.name} {info} align={block.alignment} :: {txt}")
        idx_p += 1
    else:
        print(f"=== TABLE {idx_t} rows={len(block.rows)} cols={len(block.columns)} "
              f"style={block.style.name if block.style else None} ===")
        for ri, row in enumerate(block.rows):
            cells = [c.text.replace('\n', '\\n').strip() for c in row.cells]
            print(f"  R{ri}: " + " | ".join(cells))
        print("=== END TABLE ===")
        idx_t += 1
