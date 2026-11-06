# 西华师范大学计算机学院 · 实验报告写作 Skill

> **write-lab-report** — 在 Word 原模板上填写课程实验报告，保留原有表格、页眉和字体格式。
> 面向西华师范大学计算机学院本科课程实验，也兼容其他院校的同类表格式模板。

> 本项目是个人开源工具，**不是西华师范大学官方项目**；实验报告的具体栏目与格式以任课教师下发的模板为准。

---

## 这个技能解决什么问题

学院下发的实验报告模板是「左侧栏目名 + 右侧填写区」的表格式 Word 文档，手动往里贴代码、表格和截图时，很容易出现：

- 表格被撑破、跨页错位，页眉和字体被改乱；
- 代码直接粘贴后缩进和字号失控；
- 截图大小不一、图注不居中；
- 改完一个栏目，把别的栏目或页眉也带坏了。

`write-lab-report` 通过 Word COM 在**原模板**上按栏目写入，只动目标单元格，不重排整份文档。它既可以作为 Agent Skill 被 Qoder / Claude Code 调用，也可以直接当命令行工具使用。

## 适用对象与模板

- **适用课程**：操作系统、数据结构、计算机网络、数据库原理、计算机组成原理等计算机学院本科课程实验。
- **模板形式**：Word 表格式实验报告（`.doc` / `.docx`），左侧是栏目名，右侧是填写区。
- **已适配的常见栏目**：

| 模板栏目 | 说明 |
| --- | --- |
| 课程名称 / 实验名称 | 表头字段，可用 `meta` 替换 |
| 实验目的 | 支持段落、列表 |
| 实验仪器和器材 | 自动识别“实验环境 / 实验设备”等别名 |
| 实验内容或上机调试程序 | 支持代码块、嵌套表格、小标题 |
| 问题及解决办法 | 写真实调试过程，不编造 |
| 运行结果截图 | 插入真实截图或占位符 |
| 实验心得体会 | 必须结合本次实验的代码与参数 |
| 页眉（姓名 / 学号 / 班级 / 指导教师 / 日期） | 加 `--meta-paragraphs` 才会替换，默认不动姓名学号 |

> 不同教师下发的模板栏目名可能略有差异。工具内置了常见别名（如“实验内容及步骤”“遇到的问题及解决方法”），同名栏目出现两次时可用 `row` 指定行号。纯段落式模板不支持，`inspect` 会返回空栏目。

## 它能做什么

- **读取模板结构**：识别栏目位置、现有内容、页眉字段，`.docx` 直接读 XML，约 0.2 秒。
- **按栏目填充**：默认只写空白栏目，不覆盖已有内容；支持段落、小标题、列表、代码块、表格、图片和截图占位符。
- **一条命令构建**：填充 → 排版修补 → 另存 → 导出 PDF/PNG 预览，全程只启动一次 Word。
- **排版自检**：统计残留的「【…】」占位符，提示嵌套表格跨页、图注是否居中、行高是否异常。
- **保留原格式**：页眉页脚、合并单元格、字体字号、表格边框保持模板原样。

它**不会**替你编造实验数据、运行输出或“遇到的问题”。没有截图时只插入占位符，并明确告诉你该截哪几张图。

## 环境要求

| 项目 | 要求 |
| --- | --- |
| 系统 | Windows 10 / 11 |
| Office | Microsoft Word 2016 及以上（含 Microsoft 365） |
| Python | 3.9+ |
| 依赖 | `pywin32`、`python-docx`、`pymupdf` |

```powershell
pip install -r requirements.txt
```

## 安装

仓库名是 `cwnu-cs-lab-report`，但技能目录名必须保持 **`write-lab-report`**（Qoder / Claude Code 用目录名作为技能名）：

```powershell
git clone https://github.com/Sunnyday1-stack/cwnu-cs-lab-report.git

# Qoder CN
xcopy /E /I cwnu-cs-lab-report "%USERPROFILE%\.qoder-cn\skills\write-lab-report"

# Claude Code
xcopy /E /I cwnu-cs-lab-report "%USERPROFILE%\.claude\skills\write-lab-report"
```

没有 Git 时，直接在 GitHub 页面下载 ZIP，解压后把文件夹重命名为 `write-lab-report`，再放进上面的 skills 目录。

安装后在对话里说「用学院模板帮我写操作系统实验三」即可触发；也可以完全不用 Agent，直接跑下面的命令。

## 快速开始

下文用 `<skill-dir>` 表示技能所在目录，例如 `C:\Users\you\.qoder-cn\skills\write-lab-report`。

### 1. 读取模板

```powershell
python "<skill-dir>\scripts\report_tool.py" inspect "D:\课程\实验报告模板.doc" --out structure.json
```

输出包含：

- `sections`：栏目名、所在行、是否已有内容、字数、预览；
- `meta_paragraphs`：页眉里的姓名、学号、班级、指导教师、日期；
- `labels`：表格中的全部短标签。

### 2. 提取实验要求

```powershell
python "<skill-dir>\scripts\report_tool.py" extract "D:\课程\实验指导书.docx" --out requirements.txt
```

支持 `.doc` / `.docx` / `.pdf` / 纯文本。

### 3. 写 `content.json`

最小示例（完整字段见 [`references/content-schema.md`](references/content-schema.md)）：

```json
{
  "meta": {
    "课程名称": "计算机操作系统",
    "实验名称": "实验二 银行家算法避免死锁模拟实验"
  },
  "sections": [
    {
      "label": "实验目的",
      "mode": "fill",
      "blocks": [
        {"type": "paragraph", "text": "理解死锁产生的条件，掌握银行家算法的安全性检查过程。"}
      ]
    },
    {
      "label": "实验内容",
      "mode": "fill",
      "blocks": [
        {"type": "heading", "text": "一、核心数据结构"},
        {"type": "code", "text": "int available[N];\nint max[N][M];\nint allocation[N][M];\nint need[N][M];"},
        {"type": "table",
         "headers": ["输入", "预期结果", "实际结果", "结论"],
         "rows": [["T0 时刻", "安全", "安全", "一致"]]}
      ]
    },
    {
      "label": "运行结果截图",
      "mode": "fill",
      "blocks": [
        {"type": "placeholder", "text": "【此处粘贴运行结果截图】", "note": "菜单界面、进程输入、安全性检查输出"}
      ]
    }
  ]
}
```

栏目名支持常见别名，例如「实验内容或上机调试程序」自动匹配到「实验内容」。同名栏目出现两次时，用 `"row": 6` 指定行号。

### 4. 构建

```powershell
python "<skill-dir>\scripts\report_tool.py" build "D:\课程\实验报告模板.doc" `
  --content content.json `
  --out "D:\课程\实验二报告.docx" `
  --meta-paragraphs `
  --png-dir preview
```

- `--meta-paragraphs`：同时替换页眉段落中的日期、指导教师等字段（不传则只改表格里的字段，姓名学号默认不动）。
- `--png-dir`：导出 PDF 并逐页渲染 PNG，方便检查分页和溢出。
- 返回 JSON 中的 `summary`、`warnings`、`layout`、`placeholders_left` 需要逐项看，`warnings` 不能忽略。
- 原地保存会自动生成 `*.bak-时间戳` 备份；文件被 Word 占用时脚本返回码为 2。

### 5. 只补一个栏目

截图拿到后，只替换「运行结果截图」，其他栏目不动：

```powershell
python "<skill-dir>\scripts\report_tool.py" fill "D:\课程\实验二报告.docx" `
  --content patch.json `
  --layout
```

`patch.json` 里该栏目使用 `"mode": "replace"` + `image` 块。注意：**改内容应从原模板重新 `build`**，不要在同一成品上反复填充。

### 6. 重新预览

```powershell
python "<skill-dir>\scripts\report_tool.py" preview "D:\课程\实验二报告.docx" --png-dir preview --pages 12-16
```

PDF 比源文件新时直接复用，不启动 Word；需要重新导出时加 `--force`。

## content.json 要点

- `sections[].mode`：`fill`（默认，仅填空白）、`append`（追加）、`replace`（清空重写）。
- 块类型：`paragraph`、`heading`、`caption`、`list`、`code`、`table`、`image`、`placeholder`。
- 顶层 `format` 可覆盖默认排版（宋体小四、1.5 倍行距、首行缩进 2 字符）。
- 顶层 `layout` 控制 Word 排版修补，字段和 Word 限制见 [`references/content-schema.md`](references/content-schema.md)。

## 学院模板适配说明

1. **页眉字段**：姓名、学号、班级、指导教师、日期在页眉段落里，只有加 `--meta-paragraphs` 才会替换；不传时保持模板原值，避免误改。
2. **不填教师栏目**：脚本不会填写“成绩”“教师评语”，也不会擅自修改姓名学号。
3. **运行结果截图**：没有真实截图时插入灰色占位符，写明需要截哪几张画面；拿到截图后用 `replace` 模式只更新这一栏。
4. **心得体会**：写作指南要求引用本次实验的具体代码、参数或调试过程，禁止“通过本次实验我学到了很多”式空话。
5. **模板差异**：如果学院换了新版模板，先用 `inspect` 看栏目名，再在 `content.json` 里用模板原文或别名匹配。

## 已测试环境

- Windows 11 + Microsoft Word（Microsoft 365）
- Python 3.11
- 使用真实课程模板做过完整构建：16 页报告、7 个栏目、2 张运行截图，填充后页眉、表格和字体未被破坏。

仓库不包含任何真实模板、学生姓名、学号或运行截图。

## 已知限制

这些限制来自 Word 本身，不是脚本能绕过的：

1. **模板必须是表格式**（左栏目名、右填写区）。纯段落式模板 `inspect` 会返回空栏目。
2. **嵌套表格无法重复表头**：Word 的“标题行重复”对单元格内的嵌套表格无效，数据超过 3 行建议拆成多张表。
3. **单元格内分页不可控**：在表格行上强制“不允许跨页断行”且该行高于一页时，Word 会直接裁掉内容。所以长行不要设 `no_split`。
4. **`.doc` 较慢**：老格式必须经过 Word 转换，`inspect` 大约 5 秒；能转 `.docx` 就转。
5. 图片插入后是 VML 还是 DrawingML 取决于源文档格式，脚本对两种都做了计数兼容。

## 目录结构

```
cwnu-cs-lab-report/
├── SKILL.md                     # Agent 使用说明（触发条件与工作流）
├── README.md
├── requirements.txt
├── examples/
│   └── content.example.json
├── references/
│   ├── content-schema.md        # content.json 完整字段
│   └── writing-guide.md         # 各栏目写法与常见扣分点
└── scripts/
    └── report_tool.py           # inspect / extract / fill / build / preview
```

## 隐私与学术规范

- 仓库不包含任何真实课程模板、学生姓名、学号或运行截图。
- 工具只负责排版与信息整理。实验数据、程序输出和调试过程必须真实，请遵守西华师范大学及所在课程的学术规范。
- 生成的报告会包含你的姓名、学号等信息，公开分享前请自行脱敏。

## License

[MIT](LICENSE)
