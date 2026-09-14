# -*- coding: utf-8 -*-
"""refresh_toc.py —— Word COM：刷新目录域为真实页码 + 导出 PDF

用法
----
python refresh_toc.py 目标.docx [--pdf 目标.pdf] [--snapshot]

要点
----
- `Fields.Update()` 刷新所有域；`TablesOfContents(i).Update()` 重建目录。
- 页眉页脚里的域（如页码 PAGE）也要单独刷新。
- 必须在保存后再导 PDF，否则 PDF 里的目录页码是占位值。
- 用前先 `--snapshot` 存一份 COM 前快照，便于排查"Word 保存改了什么"。
- 断言 `TablesOfContents.Count`：若为 0，说明目录域结构不对（外层 TOC 域未识别），
  回到 docxlib 检查 fldChar begin/separate/end 配平与 `TOC \\o` 唯一性。
"""
import os
import sys
import shutil
import argparse

sys.stdout.reconfigure(encoding='utf-8')
import win32com.client as wc  # noqa: E402


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('docx')
    ap.add_argument('--pdf', default=None, help='导出 PDF 路径')
    ap.add_argument('--snapshot', action='store_true', help='COM 前把原文件快照到 .workbuddy/_precom.docx')
    a = ap.parse_args()

    src = os.path.abspath(a.docx)
    if not os.path.exists(src):
        raise SystemExit('文件不存在：%s' % src)
    pdf = os.path.abspath(a.pdf) if a.pdf else None
    if pdf:
        os.makedirs(os.path.dirname(pdf), exist_ok=True)

    if a.snapshot:
        pre = os.path.join(os.path.dirname(src), '.workbuddy', '_precom.docx')
        os.makedirs(os.path.dirname(pre), exist_ok=True)
        shutil.copy2(src, pre)
        print('[快照] COM 前 →', pre, os.path.getsize(pre), 'B')

    try:
        app = wc.gencache.EnsureDispatch('Word.Application')
    except Exception:
        app = wc.Dispatch('Word.Application')
    app.Visible = False
    app.DisplayAlerts = 0
    print('[Word] 版本', app.Version)

    doc = app.Documents.Open(src, ReadOnly=False, AddToRecentFiles=False)
    try:
        doc.Fields.Update()
        n = doc.TablesOfContents.Count
        print('[目录] TablesOfContents 数 =', n)
        if n == 0:
            print('  !! 目录域未被识别 → 检查 TOC 域 fldChar 配平与 TOC \\o 唯一性')
        for i in range(1, n + 1):
            doc.TablesOfContents(i).Update()
        for si in range(1, doc.Sections.Count + 1):
            for hf in doc.Sections(si).Headers:
                hf.Range.Fields.Update()
            for hf in doc.Sections(si).Footers:
                hf.Range.Fields.Update()
        doc.Repaginate()
        print('[页数]', doc.ComputeStatistics(2))

        entries = []
        if n:
            for para in doc.TablesOfContents(1).Range.Paragraphs:
                t = para.Range.Text.replace('\r', '').strip()
                if t:
                    entries.append(t)
        print('[目录] %d 条' % len(entries))
        # 打印前 6 条；条目多时省略中段再补最后 3 条。
        # ⚠️ 只有真正省略了中段才打印尾部，否则 [:6] 与 [-3:] 会重叠，
        #    在 7~9 条时把同一条打印两遍（看着像"目录里有重复条目"）。
        for e in entries[:6]:
            print('    ', e)
        if len(entries) > 9:
            print('     ...')
            for e in entries[-3:]:
                print('    ', e)
        elif len(entries) > 6:
            for e in entries[6:]:
                print('    ', e)

        doc.Save()
        if pdf:
            doc.ExportAsFixedFormat(pdf, 17)
            print('[PDF] →', pdf, os.path.getsize(pdf), 'B')
    finally:
        doc.Close(SaveChanges=0)
        app.Quit()
    print('[完成] 目录已刷新')


if __name__ == '__main__':
    main()
