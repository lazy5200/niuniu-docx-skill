# -*- coding: utf-8 -*-
"""build_docx.py —— 按《规范.md》从内容表生成整份 .docx 的 CLI

用法
----
python build_docx.py --base 模板.docx --out 目标.docx --content content_x.py \
    --cover "系统名|文档类型|系统名|2026年9月" \
    --revision "2026-09-14|1.0|初版|项目组" \
    --header "系统名 - 文档类型" \
    --title "系统名-文档类型"

内容表（--content 指向的 .py）需导出：
    AVAIL   = 8306                      # 版心宽（twip），A4 左右 3.17cm 时 = 8306
    CONTENT = [(类型, ...), ...]        # 见下

内容类型：H1 / H2 / H3 / P / C / TBLCAP / FIGCAP / FIG / T
    ('H1', '运维概述')                       章标题（★ 不要写编号，基底 numbering 自动输出「1、」）
    ('H2', '1.1 编写目的与适用范围')          一级节标题（编号写进文本）
    ('H3', '1.1.1 xxx')                      二级节标题
    ('P',  '正文……')                         正文（首行缩进 2 汉字）
    ('C',  '$ systemctl restart xxx')        命令（左缩进，不首行缩进）
    ('TBLCAP', '表1-1 手册使用约定')          表题（置于表格上方）
    ('T', ['列1','列2'], [['a','b']], [1600, 6706])   表格（列宽合计须 == AVAIL）
    ('FIG',  r'D:\\img\\arch.png')                       插图（默认等比缩放到版心宽，不放大）
    ('FIG',  r'D:\\img\\arch.png', 380)                  插图（指定显示宽度 380 磅）
    ('FIG',  r'D:\\img\\arch.png', 380, '图2-1 系统架构')  插图 + 图题（一步到位）
    ('FIGCAP', '图2-1 系统架构')              图题（单独放在图片下方时用）
"""
import os
import sys
import argparse
import importlib.util

sys.stdout.reconfigure(encoding='utf-8')
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import docxlib  # noqa: E402


def load_content(path):
    spec = importlib.util.spec_from_file_location('content_mod', path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    content = getattr(mod, 'CONTENT', None)
    if content is None:
        raise SystemExit('内容表缺少 CONTENT')
    return content, getattr(mod, 'AVAIL', None)


def main():
    ap = argparse.ArgumentParser(description='按《规范.md》从内容表生成整份 docx')
    ap.add_argument('--base', required=True, help='基底 docx（提供编号/样式/页眉页脚/节/封面与目录原型）')
    ap.add_argument('--out', required=True, help='目标 docx')
    ap.add_argument('--content', required=True, help='内容表 .py')
    ap.add_argument('--cover', default='', help='封面文案，用 | 分隔，按顺序替换封面文本段')
    ap.add_argument('--revision', action='append', default=[],
                    help='修订行 "日期|版本|说明|作者"，可重复')
    ap.add_argument('--header', default=None, help='页眉文案')
    ap.add_argument('--header-old', default=None, help='页眉里要替换掉的原文案（不传则自动取最长文本）')
    ap.add_argument('--title', default=None, help='文档属性标题（dc:title）')
    ap.add_argument('--avail', type=int, default=None, help='版心宽（twip），默认 8306')
    ap.add_argument('--no-backup', action='store_true', help='不备份已存在的目标文件')
    a = ap.parse_args()

    content, avail_mod = load_content(a.content)
    avail = a.avail or avail_mod or docxlib.SPEC['avail']
    cover = [x for x in a.cover.split('|')] if a.cover else []
    rows = [r.split('|') for r in a.revision]

    docxlib.rebuild(
        base=a.base, out=a.out, content=content,
        cover_texts=cover, revision_rows=rows,
        header_text=a.header, header_old=a.header_old, doc_title=a.title,
        avail=avail, backup=not a.no_backup,
    )


if __name__ == '__main__':
    main()
