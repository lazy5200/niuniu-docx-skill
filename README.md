# docx-writer

> 中文正式文档（.docx）全流程技能 —— 按规范撰写、就地改写保格式、按模板重排结构、补/刷目录，并支持插图与 OMML 原生公式。

一个面向 **中文技术文档 / 交付文档**（运维手册、设计说明书、验收文档、报告）的文档工程技能。
它不靠"凭感觉排版"，而是以一份《规范.md》作为**格式基准**、以一个真实 docx 作为**外形基底**，
用脚本把排版规则固化下来，再用 PDF 实测做终检。

## 它能做什么

| 模式 | 场景 | 入口 |
| --- | --- | --- |
| **A 新建** | "写一份 XX 手册/报告""按规范生成 docx" | `scripts/build_docx.py` |
| **B 就地改写** | "改这份文档的内容，格式不要变" | `scripts/rewrite_helpers.py` |
| **C 套模板重排** | "按 XX 模板改结构""跟模板对齐" | `scripts/build_docx.py --base 模板` |
| **D 补/刷目录** | "文档缺目录""目录页码不对" | `scripts/insert_toc.py` + `scripts/refresh_toc.py` |

## 亮点

- **以模板为基底重建**：从同族模板继承 `numbering.xml` / `styles.xml` / 页眉页脚 / 节设置，
  不必从零重造看不见的部件，外形自动对齐。
- **格式由规范约束**：字号、行距、段前段后、编号、缩进、表格边框、页眉页脚逐项落成参数。
- **PDF 实测终检**：`check_spec.py` 把排版规则变成可执行断言（含文字越界检测），
  而不是"看着差不多"。
- **踩坑沉淀**：`references/踩坑与排障.md` 收录 34 条真实坑位与解法
  （Word 吞段距、COM 保存后 styleId 重编号、图片被固定行距裁切、目录插重复等）。
- **支持插图与公式**：本地图片自动等比缩放并写全 OOXML 关系；
  公式要求以可编译的 OMML 原生对象录入（行间公式独立成行右端编号，行内公式随文不编号）。

## 安装

把本目录整个放到 WorkBuddy 的技能目录下即可：

```
~/.workbuddy/skills/docx-writer/          # macOS / Linux
%USERPROFILE%\.workbuddy\skills\docx-writer\   # Windows
```

## 依赖

```bash
python -m pip install lxml pywin32 pymupdf python-docx
```

- `lxml` —— OOXML 直接操作（核心）
- `pywin32` —— Word COM，用于刷目录域与导出 PDF（仅 Windows 需要）
- `pymupdf` —— PDF 实测终检
- `python-docx` —— 仅 B 模式（就地改写）需要

## 目录结构

```
docx-writer/
├── SKILL.md                     # 技能主文档：模式路由 + 规范速查 + 故障表
├── references/
│   ├── 规范.md                   # 《文档格式规范》：格式基准（12 章）
│   └── 踩坑与排障.md             # 34 条实战坑位与解法
├── scripts/
│   ├── docxlib.py               # 核心库：基底解析、样式反查、构建、插图
│   ├── build_docx.py            # A/C 模式 CLI
│   ├── insert_toc.py            # D 模式：插入/重建目录（幂等）
│   ├── refresh_toc.py           # 刷目录域 + 导出 PDF
│   ├── check_spec.py            # 规范终检（退出码 = FAIL 数）
│   ├── rewrite_helpers.py       # B 模式辅助
│   └── dump_*.py                # 勘察脚本：先量再动手
└── examples/
    └── content_example.py       # 内容表示例，复制即用
```

## 快速开始

```bash
python scripts/build_docx.py \
  --base 模板.docx \
  --out 目标文档.docx \
  --content examples/content_example.py \
  --cover "系统名|文档类型|系统名|2026年9月" \
  --revision "2026-09-14|1.0|初版|项目组" \
  --header "系统名 - 文档类型" \
  --title "系统名-文档类型"

python scripts/refresh_toc.py --docx 目标文档.docx --pdf 目标文档.pdf
python scripts/check_spec.py --pdf 目标文档.pdf
```

## 内容表格式

```python
CONTENT = [
    ('H1',     '编写说明'),            # ★ H1 不写编号，由基底 numbering 自动生成
    ('H2',     '1.1 目的与范围'),       # ★ 节标题编号写在文本里
    ('P',      '正文……'),
    ('C',      '$ systemctl restart demo-service'),
    ('TBLCAP', '表1-1 术语与说明'),
    ('T',      ['术语', '说明'], [['内容表', '描述文档正文的元组列表']], [2000, 6306]),
    ('FIG',    r'D:\img\arch.png', 380, '图3-1 系统总体架构'),
]
```

## 许可

MIT License，详见 [LICENSE](LICENSE)。
