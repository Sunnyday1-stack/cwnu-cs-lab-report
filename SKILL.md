---
name: write-lab-report
description: 读取大学课程实验报告模板、实验要求与源代码，在保留原表格、页眉和字体格式的前提下按栏目填写实验目的、原理、环境、内容与代码、测试用例、运行结果、问题及解决办法、心得体会；没有现成代码时先实现并编译运行，截图保留占位。Use when 用户给出 .doc/.docx 实验报告模板、实验指导书或源程序，要求写/补/改实验报告、实验心得、实验总结，或提到实验报告、上机报告、课程实验、填写报告、报告模板。
---

# 写实验报告

## 概述

把实验要求、源码和真实运行结果整理成符合模板的课程实验报告。默认只填空白栏目，保留模板原有格式与已有内容，经不伪造运行结果或截图。

## 前置条件

- Windows + Microsoft Word（`.doc` 必须经 Word COM 读写）。
- Python 依赖：`pywin32`；`.docx` 快速读取需要 `python-docx`；预览 PNG 需要 `pymupdf`。缺失时先安装：
  `pip install pywin32 python-docx pymupdf`
- `.wps` 无法直接处理，先让用户另存为 `.docx`。
- 模板必须是“左侧栏目名 + 右侧填写区”的表格形式；纯段落式模板 `inspect` 会返回空栏目，此时说明限制并请用户提供表格式模板。

本文命令中的 `<skill-dir>` 指本 skill 的安装目录，例如
`%USERPROFILE%\.qoder-cn\skills\write-lab-report`。

## 效率约定（重要）

Word 每次启动/保存/导出都要十几秒，慢的根源是反复开关 Word，不是内容多少。

1. **默认用 `build` 一条命令**完成 填充 + 排版修补 + 另存 + PDF/PNG，全程只启动一次 Word。不要再写 `fix*.py` 后处理脚本，排版需求写进 `content.json` 的 `layout`。
2. `.docx` 的 `inspect`/`extract` 直接读 XML，约 0.2 秒，不会启动 Word；只有 `.doc` 才走 Word。
3. `preview` 发现 PDF 比源文件新就直接复用，0.2 秒出结果；需要强制重导时加 `--force`。
4. 一次只看改动页：`--pages 12-16`，不要每次渲染全部页面。
5. 不在命令里加 `--stats`，`ComputeStatistics` 会强制重新分页，多花数秒。
6. 改内容后从**原始模板**重新 `build` 到新文件名，不要在同一成品上反复 fill。

## 工作流程

### 1. 收集输入

必需：报告模板路径。
尽量获取：实验指导书/要求文档、源代码、运行截图、课程与实验名称。

只缺模板时暂停询问；其余缺失按下面规则继续，不要连续追问。

### 2. 读取模板与要求

```bash
python "<skill-dir>/scripts/report_tool.py" inspect "<模板路径>" --out "<工作目录>/report_structure.json"
python "<skill-dir>/scripts/report_tool.py" extract "<要求文档>" --out "<工作目录>/requirements.txt"
```

读 `report_structure.json`：

- `sections`：识别出的栏目，含 `label`、`row`、`existing`、`chars`、`preview`。同名栏目出现多次时用 `row` 指定。
- `meta_paragraphs`：页眉里的姓名、学号、班级、指导教师、日期等。
- `labels`：表格中所有短标签，可用于确认字段位置。

**不要填写**“成绩”“教师评语”等由教师填写的栏目，也不要擅自改姓名、学号。

### 3. 准备代码与数据

按实验类型处理：

- **编程类**（操作系统、数据结构、C/C++、Java 等）：先读现有源码并编译运行；没有源码时按题目要求实现，保存到模板同目录，编译运行，保留真实输出。
- **网络仿真类**（eNSP、Packet Tracer 等）：整理拓扑、设备型号、IP 规划、关键配置命令和验证结果。
- **数据库类**：整理建表 SQL、测试数据、查询语句和实际结果。
- **硬件/组成原理类**：整理器件、连线、关键波形或引脚状态、故障排查过程。
- **无法实际运行的内容**：只写设计与推演，明确标注“预期结果”，不得伪造成“实际运行结果”。

### 4. 设计测试用例

至少 3 组，包含正常情况和至少 1 个边界情况（空输入、同时到达、极值、异常参数等）。

每组先手工推演预期结果，再与程序输出逐项对照。不一致就查原因，不得把“实际”改成“预期”来凑一致。

### 5. 写 content.json

结构、块类型和示例见 `references/content-schema.md`。要点：

- 每个栏目一个 section，`label` 用模板中的栏目名。
- 默认 `mode: "fill"`，只填空白栏目；用户明确要求覆盖时才用 `"replace"`。
- 已有代码/正文的栏目不要重复贴一遍；把新增的测试用例、分析、心得写到对应空白栏目。
- 代码用 `code` 块，表格用 `table` 块，截图用 `placeholder` 或 `image` 块。
- 页面已有内容且需要补充时用 `"append"`。

各栏目写什么、心得怎么写，见 `references/writing-guide.md`。

### 6. 构建

一条命令完成填充 → 排版修补 → 另存 → 导出预览：

```bash
python "<skill-dir>/scripts/report_tool.py" build "<模板路径>" \
  --content "<工作目录>/content.json" \
  --out "<成品路径>" \
  --meta-paragraphs \
  --png-dir "<工作目录>/preview"
```

- 默认输出 `<模板名>_filled.<原扩展名>`，推荐显式传 `--out`。
- 成品里剩余的“【”数量见返回的 `placeholders_left`，必须为 0（除非确实等用户截图）。
- 返回 JSON 先看 `summary`、`warnings`、`layout`、`preview.pages`。
- 只填某个栏目或补截图时，才单独用 `fill`；它不会应用 `layout`，除非加 `--layout`。
- 原地保存会自动生成 `*.bak-时间戳` 备份。文件被 Word 占用时脚本报错码 2，提示用户关闭文档后重试。

排版修补写在 `content.json` 的 `layout` 段（表格断行、清行高、图注居中），字段见
`references/content-schema.md`。**不要**再手写 `fix*.py` 去逐段遍历 Word 对象——那是最慢且最容易卡死的做法。

### 7. 预览与自检

`build` 已经生成预览；单独补看时（PDF 未变化会直接复用，不会启动 Word）：

```bash
python "<skill-dir>/scripts/report_tool.py" preview "<成品路径>" --png-dir "<工作目录>/preview" --pages 12-16
```

查看新内容所在页的 PNG，逐项检查：

- 每个目标栏目都有内容，且没有残留“【…】”占位符（除非确实等用户截图）。
- 代码没有超出单元格右边界，长行已折行且字号可读。
- 嵌套表格列宽合理、边框完整、表头加粗；跨页表格续页后有表头（Word 不支持嵌套表头重复，超过 3 行要拆表）。
- 图片居中、有图注，未撑破表格。
- 分页合理，避免标题或表头单独留在页尾；没有整页空白。
- 原有栏目和页眉未被意外改动。

需要修改时，**从原模板重新构建到新文件**，不要在已填充文件上重复 fill。

### 8. 截图与收尾

- 没有截图时插入 `placeholder` 块，写清楚要截哪几张图，例如“菜单界面 / 输入进程 / 调度输出”。
- 用户提供截图后，只对“运行结果截图”栏目用 `mode: "replace"` 加 `image` 块重填，不要连带覆盖其他栏目。
- 不得声称已插入不存在的截图。

最后向用户报告：填了哪些栏目、成品与备份路径、预览图位置、还需要用户补的截图或信息。

## 写作红线

- 不编造实验数据、运行输出、报错信息和“遇到的问题”。
- 心得必须来自本次实验的具体代码、参数或调试过程，禁止“通过本次实验我学到了很多”式空话。
- 不复制整段教材或网络范文；术语与课程教材保持一致。
- 不破坏模板结构：页眉页脚、表格、已有内容、字体字号保持原样，除非用户明确要求替换。
- 中文正文用宋体/小四、1.5 倍行距、首行缩进 2 字符；代码用等宽字体、单倍行距。脚本已按此设置，可在 `content.json` 的 `format` 中覆盖。

## 资源

- `scripts/report_tool.py`：
  - `inspect` 读模板栏目（`.docx` 走 python-docx，`.doc` 走 Word）
  - `extract` 提取要求文档纯文本（`.docx` 走 python-docx）
  - `build` 一次完成填充、排版修补、另存、预览（推荐）
  - `fill` 只填充；`preview` 单独导出 PDF/PNG（带缓存）
- `references/content-schema.md`：`content.json` 完整字段与示例。
- `references/writing-guide.md`：各栏目写法、不同实验类型的差异、心得写法与常见扣分点。
