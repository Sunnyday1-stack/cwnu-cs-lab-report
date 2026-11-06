# write-lab-report

> 把课程实验要求、源代码和真实运行结果，整理成符合学校模板的 Word 实验报告。
> An Agent Skill + CLI that fills Chinese university lab-report templates in Microsoft Word without breaking the original layout.

这是一个面向 AI Agent 的 Skill（适用于 Qoder、Claude Code 等支持 `SKILL.md` 的客户端），核心是一个可独立使用的命令行工具 `report_tool.py`：它通过 Word COM 在**原模板**上填表，保留原有页眉、表格、合并单元格和字体格式，不会把老师的模板重排成另一份文档。

> 只适合 Windows + Microsoft Word。`.wps` 请先另存为 `.docx`。

---

## 它能做什么

- **读取模板结构**：识别「实验目的 / 实验原理 / 实验环境 / 实验内容 / 运行结果 / 问题及解决 / 心得体会」等栏目，支持常见别名与同名栏目多次出现。
- **按栏目填充**：默认只写空白栏目，不覆盖已有内容；支持段落、小标题、列表、代码块、表格、图片和截图占位符。
- **一条命令构建**：填充 → 排版修补 → 另存 → 导出 PDF/PNG 预览，全程只启动一次 Word。
- **排版自检**：统计残留的「【…】」占位符、提示嵌套表格跨页、图注是否居中、行高是否异常。
- **快速读取**：`.docx` 的 `inspect` / `extract` 直接读 XML，约 0.2 秒，不启动 Word。

它**不会**替你编造实验数据、运行输出或「遇到的问题」。没有截图时只会插入占位符，并明确告诉你该截哪几张图。

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

把整个仓库放到客户端的 skills 目录，文件夹名保持 `write-lab-report`：

```powershell
# Qoder
xcopy /E /I write-lab-report "%USERPROFILE%\.qoder-cn\skills\write-lab-report"

# Claude Code
xcopy /E /I write-lab-report "%USERPROFILE%\.claude\skills\write-lab-report"
```

之后在对话里说「用实验报告模板帮我写实验三」即可触发；也可以完全不用 Agent，直接跑下面的命令。

## 快速开始

下文用 `<skill-dir>` 表示仓库所在目录，例如 `C:\Users\you\.qoder-cn\skills\write-lab-report`。

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
    "实验名称": "实验二 银行家算法"
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

- `--meta-paragraphs`：同时替换页眉段落中的日期、指导教师等字段（不传则只改表格里的字段）。
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

## 已知限制

这些限制来自 Word 本身，不是脚本能绕过的：

1. **模板必须是表格式**（左栏目名、右填写区）。纯段落式模板 `inspect` 会返回空栏目。
2. **嵌套表格无法重复表头**：Word 的「标题行重复」对单元格内的嵌套表格无效，数据超过 3 行建议拆成多张表。
3. **单元格内分页不可控**：在表格行上强制「不允许跨页断行」且该行高于一页时，Word 会直接裁掉内容。所以长行不要设 `no_split`。
4. **`.doc` 较慢**：老格式必须经过 Word 转换，`inspect` 大约 5 秒；能转 `.docx` 就转。
5. 图片插入后是 VML 还是 DrawingML 取决于源文档格式，脚本对两种都做了计数兼容。

## 目录结构

```
write-lab-report/
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
- 工具只负责排版与信息整理。实验数据、程序输出和调试过程必须真实，请遵守所在课程的学术规范。
- 生成的报告会包含你的姓名、学号等信息，公开分享前请自行脱敏。

## License

[MIT](LICENSE)
