---
name: docx-writer
description: 中文正式文档（.docx）全流程助手——按《规范.md》撰写整份交付文档、就地改写已有文档并 100% 保留原格式、按模板骨架重排结构、插入或重建目录，并支持插图与 OMML 原生公式。当用户要求"写一份 XX 文档/手册/报告""按规范生成一份 docx""新建一份符合格式规范的 word""把这份内容排成文档""改这份文档的内容但格式不要变""按 XX 模板修改结构""跟模板对齐""加一页目录""目录页码不对""给文档配图/插图"时使用。
agent_created: true
---

# 中文 DOCX 撰写与改写（统一技能）

覆盖「写文档」的全生命周期：按规范从零产出、改已有文档保格式、按模板重排、补目录，
以及插图与公式。**格式基准**是一份《规范.md》，**外形基准**是一个「基底 docx」。

---

## 0. 先选模式

| 用户说什么 | 模式 | 入口 |
| --- | --- | --- |
| "写一份 XX 手册/报告"「按规范生成 docx」「新建一份 word」 | **A 新建** | `build_docx.py` |
| "改这份文档的内容，格式不要变"「更新一下 XX.docx」 | **B 就地改写** | `rewrite_helpers.py` + §3 |
| "按 XX 模板改结构"「跟模板对齐」「套模板」 | **C 套模板重排** | `build_docx.py --base 模板`（优先） |
| "文档缺目录"「目录页码不对」「按模板补目录」 | **D 补目录** | `insert_toc.py` + `refresh_toc.py` |

**怎么判断走 A/C 还是 B**：

- 目标文档**本身格式完备**（有自己的 numbering / styles / 页眉页脚 / 多节）→ 走 **B**，就地改，零风险。
- 目标文档是 **python-docx 从零生成的**（单节、无 `numbering.xml`、无 header/footer、标题手写编号）
  → **不要就地补部件**，换基底重建（走 **A** 或 **C**）。在残缺文档上打补丁的风险远大于重建。
- **有模板或兄弟成品** → 一律拿它当 `--base`：外形（封面几行字、页眉文案、目录形态）
  取自基底，**规范只管字号/间距/编号/表格/页面**。

四种模式共用三件设施：**基底**（提供看不见的部件）、**勘察脚本**（先量再动手）、**终检**（PDF 实测）。

---

## 1. 环境

```
C:\Users\mrlazy\.workbuddy\binaries\python\envs\default\Scripts\python.exe
```

依赖：`lxml`、`pywin32`、`pymupdf`、`python-docx`（B 模式用）。
缺则用该 venv 的 `-m pip install`，**不要**污染系统 Python。

> ⚠️ 托管 python 本体（`...\versions\3.13.12\python.exe`）**没有 lxml**，一律用上面的 venv。
> ⚠️ Windows 下 bash shim 可能半坏（`head`/`tail`/`ls`/`dirname` 报 command not found）。
> 一律 **Write 落盘脚本 + 绝对路径 python + `> out.txt 2>&1` + Read 读回**，不要用 shell 管道。

---

## 2. 模式 A：按规范新建（S0→S6）

```
S0 定基底 → S1 勘察 → S2 写内容表 → S3 构建 → S4 刷目录+导PDF → S5 终检 → S6 交付
```

**S0 定基底** — 找一个**同族模板 / 兄弟文档**当 `--base`。它提供你不想重造的部件：
`numbering.xml`（章标题 `1、` 编号）、`styles.xml`、`header*.xml`、`footer*.xml`、
`w:sectPr`（A4 + 页边距）、封面与目录的**原型段落**。没有模板就从任一已有成品拿。

**S1 勘察（不能省）**

```bash
python scripts/dump_body_kids.py <base.docx> --runs 18 19 20 21
python scripts/dump_para_spacing.py <base.docx>
```

**S2 写内容表** — 照 `examples/content_example.py` 写 `content_<doc>.py`，
导出 `CONTENT` 与 `AVAIL`（A4 左右 3.17cm → **8306**）。目录条目由 H1/H2 自动推导。

**S3 构建**

```bash
python scripts/build_docx.py \
  --base 模板.docx --out 目标.docx --content content_x.py \
  --cover "系统名|文档类型|系统名|2026年9月" \
  --revision "2026-09-14|1.0|初版|项目组" \
  --header "系统名 - 文档类型" \
  --title "系统名-文档类型"
```

脚本末尾自带断言：XML 全部可解析、域字符配平、`TOC \o` 唯一、**插图三件套齐全且图片段非固定行距**。

**S4 刷目录 + 导 PDF**

```bash
python scripts/refresh_toc.py 目标.docx --pdf .workbuddy/preview/目标.pdf
```

**必须看 Word COM 报的 `TablesOfContents 数`**：若为 0，说明目录域结构不对，
回去查 `fldChar` begin/separate/end 配平与 `TOC \o` 唯一性。

**S5 终检**

```bash
python scripts/check_spec.py 目标.docx --pdf .workbuddy/preview/目标.pdf
```

退出码 = FAIL 数。PDF 侧查越界与**实绘字号**（样式定义会骗人）。

**S6 交付** — `present_files` 交付 docx + PDF。

---

## 3. 模式 B：就地改写（保原格式）

铁律：**复用已有的 run 与 XML 元素，绝不重建段落/单元格**。

```python
from rewrite_helpers import set_para_text, set_term, insert_para_after, set_cell_text
```

| 函数 | 用途 |
| --- | --- |
| `set_para_text(p, text)` | 改段落文字，保留首个 run 格式（**不要** `p.text = text`，那会清空所有 run 级格式） |
| `set_term(p, name, desc)` | 改「词条名：释义」式段落（run0 加粗、run1 正常） |
| `insert_para_after(ref_p, text)` | 在 ref_p 后插同格式段落 |
| `set_cell_text(cell, text)` | 写单元格（定位到**最后一个文字非空**的 run，避免丢字体） |
| `set_row(t, ri, values)` | 按列序重写整行 |
| `add_column(t, values)` | 追加一列（追加后**必须**用 `set_row` 重排列序） |

**执行顺序**：先删后增、索引在使用前重算，避免错位。

**必须复核**：改完 dump 一遍段落/表格结构 + 渲染 PDF 抽查。

---

## 4. 模式 C：套模板重排

**优先用 A 的流水线**（`--base 模板`），它已经把「清空正文 → 按内容表重建」做成了机制。
只有当**内容本身要以原文档为底、只换骨架**时，才走「搬家式重构」：

1. dump 模板与目标的骨架，比对章节目录差异；
2. 定策略：**大章严格与模板一致** → 模板里没有的节标题要删，内容降为章前导语；
3. 用目标文档当基底（保住它的素材），按模板的形状重建；
4. 标题体系对齐见 `references/踩坑与排障.md` §1.5（样式级 vs 段落级 numPr 冲突）。

**同模板的兄弟文档：先判断能否「派生」子脚本** —— 只替换项目标识串、逻辑一行不动，
并打印 diff 逐行确认无误伤。实测 8 处标识变更、零调试一次通过。

---

## 5. 模式 D：补 / 重建目录

```bash
python scripts/insert_toc.py 目标.docx --title 目录 --levels 2
python scripts/refresh_toc.py 目标.docx --pdf out.pdf      # 拿真实页码
```

- **幂等**：重跑时先从正文起点**向前回溯**定位上一轮的目录块并剔除。
  判据是 Word 一定会保留的特征：① 文本 ==「目录」；② 样式是 `toc 1/2/3`；③ 含域字符且文本为空。
  ⚠️ **不要**用 `w:rsidR` 标记判旧目录 —— Word COM 保存会**重写 rsid**，标记失效 → 插出第二份目录
  （实测 Word 报 `TablesOfContents 数 = 2`）。
  ⚠️ **不要**用"从 TOC begin 向后找第一个 `fldChar end`"—— Word 更新目录后，
  首条条目段内部就带了自己的 `end`，会被误判成外层域结尾，只删掉 2 段、其余全残留（域字符 `begin=8 end=9`）。
  ⚠️ **不要**用"文本等于条目名"判断 —— 条目段的文本把**页码也拼进去**了
  （`''.join(el.itertext())` 得到 `'1、 项目概况12'`），永远匹配不上 → 双目录。
- **校验在写盘之前**：先把 XML 组装到内存、跑完全部断言，**再** `save_parts`。
  否则断言失败时半成品已经落盘（实测会把文档写成双目录）。
- **剔除之后必须重算索引**，否则会插到错误位置（症状：目录里套目录、页码整体错乱）。
- 插入位置：正文第一个章标题**之前**；若它前面有承载 `sectPr` 的分节符段，
  则插在**分节符之前**（目录归入封面那一节，不打乱正文页码）。
- 目录标题段用 `pageBreakBefore` 独占一页（**不要**用「空段 + 分页符」，会多一张空白页）。
- 页码交给 Word COM 更新，**不要手填**。

---

## 6. 内容表格式

```python
AVAIL = 8306            # 版心宽 twip = 11906 − 1800×2
CONTENT = [
    ('H1',  '运维概述'),                      # ★ 不要写编号，基底 numbering 自动输出「1、」
    ('H2',  '1.1 编写目的与适用范围'),         # ★ 节标题编号写进文本（H2 不用自动编号）
    ('H3',  '1.1.1 xxx'),
    ('P',   '正文段落，首行自动缩进 2 汉字……'),
    ('C',   '$ systemctl restart xxx'),       # 命令：左缩进、不首行缩进
    ('TBLCAP', '表2-1 部署主机与登录信息'),     # 表题（自动置于表格上方、居中）
    ('T',   ['主机', '账号'], [['192.0.2.10', 'admin']], [4153, 4153]),
    ('FIG', r'D:\img\arch.png', 380, '图2-1 系统总体架构'),   # 插图 + 图题（一步到位）
    ('FIGCAP', '图2-2 数据流'),                # 单独放图题时用（置于图片下方）
]
```

约束（构建脚本会断言）：

- **H1 文本不含编号**；H2/H3 文本**自带**编号（`1.1`、`1.1.1`）。
- `('T', 表头, 数据行列表, 列宽列表)`：列宽合计必须 `== AVAIL`，列数必须一致。
- `TBLCAP` 在表格**之前**；`FIGCAP` 在图片**之后**。
- `('FIG', 路径)` / `('FIG', 路径, 宽度pt)` / `('FIG', 路径, 宽度pt, 图题)` 三种写法。

---

## 7. 插图

### 支持的写法

```python
('FIG', r'D:\img\arch.png')                  # 等比缩放到版心宽（不放大超过原图自然尺寸）
('FIG', r'D:\img\arch.png', 380)             # 指定显示宽度 380 磅
('FIG', r'D:\img\arch.png', 380, '图2-1 xx')  # 插图 + 图题一步到位（推荐）
```

### 内部做了什么

一次调用完成 OOXML 插图的**三件套**（缺一不可）：

| 部件 | 位置 | 缺了会怎样 |
| --- | --- | --- |
| 图片二进制 | `word/media/imageN.png` | 图片位置空白 |
| 关系条目 | `word/_rels/document.xml.rels` | Word 报「内容有问题」 |
| Content Type | `[Content_Types].xml` 的 `<Default Extension="png" .../>` | Word 报「内容有问题」 |

尺寸规则：默认**等比缩放到版心宽，但不放大**（原图自然尺寸按 96 DPI 换算，小于版心宽时保持原大）；
显式给宽度时按给定宽度等比缩放。像素尺寸由纯标准库解析（PNG / JPEG / GIF / BMP，不依赖 Pillow）。
`.emf` / `.wmf` 矢量图读不出像素尺寸，**必须显式给宽度**。

### 🚨 图片段的行距必须是 auto

**绝不能继承正文的 `lineRule="exact"`（固定 20 磅）** —— 否则图片被裁成 20 磅高：
`w:drawing` / `r:embed` / `word/media/*` 全部齐备，查 XML 查不出任何问题，但导出后**完全看不见**。
`build_figure()` 已强制 `line="240" lineRule="auto"`，`_verify()` 会断言所有含 `w:drawing` 的段落都不是 `exact`。

**判据**：导出 PDF 前后**总页数**对比。页数没变就要怀疑图没占位。

### 图片从哪来

- **界面截图 / 已有图** → 直接引用文件路径。
- **流程、架构、关系图** → 自绘（截图表达不了），见 `references/踩坑与排障.md` §三。
- ⚠️ **别为了凑版面造图**：界面类文档用户只认真实截图；
  「要尽量多写」时**先想表格，再想自绘图** —— 表格是零风险的高密度载体。
- ⚠️ 自绘图**必做文字溢出/重叠几何自检**（模型看不了图的唯一替代方案），
  且中文标注**自绘优于 AI 生图**（生图会糊字）。

---

## 8. 公式（要求可编译的 OMML）

**目标**：公式是 Word 里双击可编辑的**原生公式对象**（`m:oMath`），不是图片、纯文本或字符拼凑。

- **最佳路径**：`pandoc formulas.tex -o formulas.docx`（一条命令，无需 TeX 发行版）。
  注入时**必须包一层带命名空间声明的根节点**，否则 lxml 报 `Namespace prefix m on oMath is not defined`。
- **行间公式**：居中 + 右端编号 `(X-Y)` + 段前段后各 6 磅。
  用**居中制表位 + 右制表位**（`4153` / `8306`），不要用表格。
- **行内公式**：正文里用 `$...$` 标记，直接 `append` `m:oMath`（不需要制表位）。
  只放**矮结构**（`y_i`、`\hat{y}`、`\mu_x`），含 `\frac`/`\sqrt`/带限 `\sum` 的提到行间 ——
  否则会被固定行高裁切。
- 🚨 **`compatibilityMode=11` 会让 Word 把 OMML 降级成 VML 图片**（`m:oMath` 变 0）。
  构建时必须改成 **15**，且验收要在 **COM 保存之后**再数一遍 `m:oMath`。
- ⚠️ 含 `\t`/`\b`/`\n` 开头的 LaTeX **一律用原始字符串** `r'\theta'`，否则 Python 先转义成控制字符。

---

## 9. 规范速查（详见 `references/规范.md`）

| 要素 | 取值 |
| --- | --- |
| 页面 | A4；上下 2.54cm；左右 3.17cm；版心宽 8306 twip |
| 章标题 H1 | 黑体 15pt(30) · 段前 17pt(340) / 段后 16.5pt(330) · 行距 2.4 倍(576 auto) · 顶格 |
| 节标题 H2 | 黑体 14pt(28) · 段前段后各 13pt(260) · 行距 ≈1.72 倍(413 auto) · 顶格 |
| 正文 | 宋体 12pt(24) · 西文 Times New Roman · 固定行距 20pt(400 exact) · 首行缩进 2 汉字(480/200) |
| 表格 | Table Grid · 六边单线 0.5 磅(sz=4) · 总宽 8306 · 行不跨页 · 表内文字 10.5pt(21) |
| 修订表 | 表头 9pt(18) · 表体 10.5pt(21) · 列宽 2076/2076/2077/2077 |
| 表题 | 表上方 · 居中 · 段前 12pt(240) / 段后 6pt(120) |
| 图题 | 图下方 · 居中 · 段前 6pt(120) / 段后 12pt(240) |
| 图片 | 居中 · 段前段后各 6pt(120) · **单倍行距** |
| 页眉 | 9pt(18) 斜体 `项目名 - 文档类型` |
| 页脚 | 9pt(18) `第 X 页`（PAGE 域） |
| 封面 | 主标题 18pt(36) 加粗 ×2 行 · 副标题/年月 14pt(28) |
| 公式 | 行间：居中 + 右端编号 + 段前段后 6pt；行内：随文不编号；**两者都必须可编译（OMML）** |

> 括号里是 OOXML 的 half-point / twip 值，直接用于脚本参数。

---

## 10. 硬规则（逐条都踩过，详见 `references/踩坑与排障.md`）

1. **`kids[i]` ≠ 段索引** —— body 直接子元素含 `w:tbl` 与 `w:sectPr`。摘任何原型前先跑 `dump_body_kids.py`。
2. **lxml 元素只能有一个父节点** —— 要挂到多个父节点的 `rPr`/`pPr` 必须**每次新建**，复用会被"搬走"。
3. **`w:spacing` 的 `before/after` 必须带四伴随属性**（`beforeLines`/`afterLines`/`beforeAutospacing`/`afterAutospacing`），
   否则 Word 保存时会整组删掉（实测标题顶距页眉 46.4pt → 28.2pt）。
4. **样式 id 绝不硬编码** —— 用 `w:name` 反查。COM 保存后会重编号（`heading1 2→1`、`TableGrid 11→a6`）。
5. **表格宽度 = 版心宽**，列宽合计必须相等；COM 后边框可能移入表样式，**效果等价**，别断言 `tblPr`。
6. **目录域三要素**：`TOC` 域的 `begin`/`instrText`/`separate` 只挂**首条条目**，`end` 挂末条之后；
   `TOC \o` 全文只允许 1 次；页码由域生成，**禁止手填**。
7. **PDF 实测是唯一可信口径** —— 样式定义会骗人（表体单元格常不写显式 `sz`）。
8. **插图三件套**：media + rels + Content_Types，缺一 Word 就报"内容有问题"。
9. **图片段行距必须是 auto** —— 固定行距会把图裁没，XML 却查不出来。
10. **公式必须可编译**，且 `compatibilityMode` 必须是 15。
11. **章标题编号只出现一次** —— H1 文本里不要再写编号。
12. **PDF 越界要区分「标点悬挂」** —— 右边界容差放宽 1 个 em（模板 PDF 亦如此），左/上/下仍严格。

---

## 11. 故障排查

| 现象 | 根因 | 处置 |
| --- | --- | --- |
| 章标题变成「1、1、运维概述」 | 内容表 H1 文本自带编号，与基底自动编号叠加 | 去掉 H1 文本里的 `1、` |
| `TablesOfContents.Count = 0`，目录全"错误!未定义书签" | 目录域结构坏（取错原型 / 多剥了 `HYPERLINK begin`） | 核对 `dump_body_kids.py` 索引；确认 `TOC \o` 唯一、域字符配平 |
| 章标题段前段后变 0，标题贴上页眉 | `spacing` 缺四伴随属性被 Word 清除 | `build_ppr(sp_lines=True)`，用 PDF 实测标题距页眉 |
| **图片导出后看不见** | 图片段继承了固定行距被裁 | 图片段必须 `lineRule="auto"`；比 PDF 页数 |
| Word 报「内容有问题」 | 插图三件套缺一；或段落/表格子元素顺序违反 schema | 查 media/rels/Content_Types；用 `ins()` 按顺序插入 |
| Word 报「可能已经损坏」 | 自建部件的 `Content_Types` Override 被清理循环删掉 | 二分定位：逐个部件换回基底字节 |
| 表框看不见了 | COM 把边框移入 Table Grid 样式 | 查样式定义或看 PDF 实绘线 |
| 断言按数字 styleId 全红 | Word 重编号了 styleId | 按 `w:name` 找样式 |
| 终检报「段距四伴随属性缺失」 | COM 归一化了等于默认值的属性 | 正常，只看 `before/after` 在不在（已降为 INFO） |
| 终检报「PDF 有 1 处越界」 | 中文行末标点悬挂 | 右边界容差放宽 1 个 em |
| 单元格写完字体丢失 | 写进了空 run，删掉了带格式的 run | 定位到**最后一个文字非空**的 run |
| 重跑后出现双目录 / `TablesOfContents 数 = 2` | 用 `w:rsidR` 标记判旧目录，但 Word 保存会**重写 rsid** | 改为「从正文起点向前回溯 + 目录特征」识别旧目录 |
| 删目录后域字符不配平（`begin=8 end=9`） | 把首条条目段内条目自己的 `end` 误判成外层 TOC 域的 `end` | 按段落特征回溯，或按域嵌套深度计数 |
| 断言失败但文件已被改坏 | 校验写在 `save_parts` 之后 | 先组装 XML → 断言 → 再写盘 |
| `WinError 5 拒绝访问` | WPS/Word 开着目标文件 | 别用临时文件 + `os.replace`；直接覆盖写 |

---

## 12. 脚本清单

| 脚本 | 用途 |
| --- | --- |
| `scripts/docxlib.py` | **核心库**：以基底重建整份 docx（结构自动探测 + 样式名反查 + 规范参数 + 插图 + 自检） |
| `scripts/build_docx.py` | 模式 A/C 入口：内容表 + 基底 → 目标 docx |
| `scripts/rewrite_helpers.py` | 模式 B 工具：就地改写保格式的函数集 |
| `scripts/insert_toc.py` | 模式 D：给已有文档插 / 重建目录页（幂等） |
| `scripts/refresh_toc.py` | Word COM：刷新目录域为真实页码 + 导 PDF |
| `scripts/check_spec.py` | 终检：XML 断言 + PDF 实测（退出码 = FAIL 数） |
| `scripts/dump_body_kids.py` | 勘察：body 直接子元素索引/标签/样式/域/sectPr |
| `scripts/dump_para_spacing.py` | 勘察：标题段前段后是否被 Word 吃掉 |
| `scripts/dump_docx.py` | 勘察：段落/表格/节总览 |
| `examples/content_example.py` | 内容表示例（含插图写法，可直接复制改） |
| `references/规范.md` | 默认格式基准（项目内有 `规范.md` 时**以项目内为准**） |
| `references/踩坑与排障.md` | 34 条实战坑的根因/判据/修法 |

---

## 13. 交付前自检清单

**结构**
- [ ] 封面（主标题 18pt / 副标题 14pt）→ 修订记录 7 行 4 列 → 目录 → 正文各章
- [ ] 目录条目数 = 章标题数 + 节标题数；`TOC \o` 唯一；无"未定义书签"

**格式**
- [ ] H1 15pt 黑体、H2 14pt 黑体、正文 12pt 宋体（**PDF 实测**复核）
- [ ] 标题段前段后齐全且带四伴随属性
- [ ] 全部表 `tblW` = 列宽合计 = 版心宽；边框 0.5 磅单线；行不跨页
- [ ] 表题在表上（12/6pt）、图题在图下（6/12pt），均居中

**插图**（有图时）
- [ ] `drawing` 数 = `word/media/*` 文件数；rels 有对应条目；`[Content_Types]` 有 Default
- [ ] 所有图片段**不是**固定行距
- [ ] 导出 PDF **页数**因插图而变化；渲染 PNG 目视确认图片可见、比例正常

**公式**（有公式时）
- [ ] COM 保存后 `m:oMath` 计数 == 期望值；`w:pict` == 0
- [ ] PDF 里有 CambriaMath 字体 span

**交付**
- [ ] `TablesOfContents.Count == 1` 且页码为真实值
- [ ] PDF 无越界文字（右边界按标点悬挂放宽）
- [ ] `compatibilityMode = 15`；`dc:title` 与文件名一致
