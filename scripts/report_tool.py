#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Word 实验报告模板工具（Windows + Microsoft Word）。

子命令：
  inspect  读取模板结构，输出栏目位置、现有内容、可用 meta 字段（JSON）
  extract  提取 doc/docx/pdf/txt 的纯文本，便于阅读实验要求或源码说明
  fill     按 content.json 填充栏目，保留模板原有表格与页眉格式
  build    一次性完成 填充 + 排版修补 + 另存 + 导出 PDF/PNG（推荐）
  preview  单独导出 PDF 并渲染为 PNG，用于检查排版

性能约定：
  * build 只启动一次 Word，替代 fill + 多个 fix 脚本 + preview 的多轮开关；
  * .docx 的 inspect/extract 直接读 XML，不启动 Word；
  * preview 默认复用比源文件新的 PDF，加 --force 强制重新导出；
  * fill/build 默认不调用 ComputeStatistics（会触发重新分页），需要页数时加 --stats。

依赖：pywin32（Word COM）、python-docx（.docx 快速读取）、
pymupdf（可选，用于预览 PNG）。
"""

import argparse
import json
import os
import re
import shutil
import sys
import time
import traceback

try:
    import pythoncom
    import win32com.client as win32
except ImportError:  # pragma: no cover
    pythoncom = None
    win32 = None

WORD_EXTS = {".doc", ".docx", ".docm", ".dot", ".dotx"}
TEXT_EXTS = {".txt", ".md", ".markdown", ".csv", ".c", ".h", ".cpp", ".java", ".py", ".sql", ".ini", ".log"}

WD_FORMAT_DOC = 0
WD_FORMAT_DOCX = 16
WD_FORMAT_PDF = 17
WD_ALIGN_LEFT = 0
WD_ALIGN_CENTER = 1
WD_ALIGN_JUSTIFY = 3
WD_LINE_SINGLE = 0
WD_LINE_1_5 = 1
WD_PREFERRED_WIDTH_PERCENT = 2
WD_BORDER_LEFT = -2
WD_STAT_WORDS = 0
WD_STAT_PAGES = 2
WD_STAT_CHARS = 3
WD_IN_TABLE = 12

CM_TO_PT = 28.3465

SECTION_ALIASES = [
    ("实验目的", ["实验目的", "实验目的和要求", "实验目的与要求", "目的与要求", "实验目标"]),
    ("实验原理", ["实验原理", "实验原理与方法", "基本原理", "实验原理及方法", "算法原理"]),
    ("实验环境", ["实验仪器和器材", "实验仪器与器材", "实验设备", "实验环境", "仪器设备", "软硬件环境", "实验器材"]),
    ("实验内容", ["实验内容或上机调试程序", "实验内容", "实验内容及步骤", "实验步骤", "上机调试程序", "实验内容与步骤", "实验过程"]),
    ("问题及解决", ["问题及解决办法", "问题及解决", "问题及解决方法", "实验过程中出现的问题及解决方法", "问题与解决", "遇到的问题及解决办法", "遇到的问题及解决方法", "实验问题"]),
    ("运行结果", ["运行结果截图", "运行结果", "实验结果", "实验结果与分析", "运行结果与分析", "测试结果", "实验结果及分析"]),
    ("心得体会", ["实验心得体会", "实验心得", "心得体会", "实验体会", "实验总结", "总结与体会", "心得与体会"]),
]

META_ALIASES = {
    "课程名称": ["课程名称", "课程"],
    "实验名称": ["实验名称", "实验项目", "实验题目"],
    "报告人姓名": ["报告人姓名", "报告人", "姓名", "学生姓名"],
    "学号": ["学号"],
    "班级": ["班级"],
    "年级": ["年级"],
    "指导教师": ["指导教师", "教师", "指导老师"],
    "日期": ["日期", "实验日期", "报告日期"],
    "成绩": ["成绩", "评分"],
}

META_PATTERNS = {
    "报告人姓名": (r"(报告人姓名[ \t\u3000]*)([^\s\x07]+)", 2),
    "学号": (r"(学号[ \t\u3000]*)([A-Za-z0-9]+)", 2),
    "指导教师": (r"(指导教师[ \t\u3000]*)([^\s\x07]+)", 2),
    "日期": (r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日", 0),
    "年级": (r"(\d{4})([ \t\u3000]*级)", 1),
    "班级": (r"(\d{1,3})([ \t\u3000]*班)", 1),
}

DEFAULT_FORMAT = {
    "body_font": "宋体",
    "latin_font": "Times New Roman",
    "heading_font": "黑体",
    "body_size": 12.0,
    "heading_size": 12.0,
    "code_font": "Consolas",
    "code_size": 9.0,
    "table_size": 10.5,
    "caption_size": 10.5,
    "line_spacing": 1.5,
}


class ToolError(Exception):
    def __init__(self, message, code=1, **extra):
        super().__init__(message)
        self.code = code
        for key, value in extra.items():
            setattr(self, key, value)


def emit(payload, out_path=None, code=0):
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if out_path:
        folder = os.path.dirname(os.path.abspath(out_path))
        if folder:
            os.makedirs(folder, exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(text)
    print(text)
    sys.exit(code)


def fail(message, code=1, **extra):
    payload = {"ok": False, "error": message}
    payload.update(extra)
    emit(payload, code=code)


def norm(text):
    text = (text or "").replace("\x07", "")
    text = re.sub(r"[\s\u3000]+", "", text)
    return text.strip("：:。.")


def clean_text(text):
    text = text or ""
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.strip() for line in text.split("\n")]
    while lines and not lines[0]:
        lines.pop(0)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def label_key(normalized):
    return SECTION_LOOKUP.get(normalized)


def meta_key(normalized):
    return META_LOOKUP.get(normalized)


def _build_lookup(aliases):
    table = {}
    for canonical, names in aliases:
        for name in list(names) + [canonical]:
            table[norm(name)] = canonical
    return table


SECTION_LOOKUP = _build_lookup(SECTION_ALIASES)
META_LOOKUP = _build_lookup(list(META_ALIASES.items()))


def find_lock_file(path):
    folder = os.path.dirname(os.path.abspath(path))
    base = os.path.basename(path)
    exact = os.path.join(folder, "~$" + base)
    if os.path.exists(exact):
        return exact
    if not os.path.isdir(folder):
        return None
    ext = os.path.splitext(base)[1].lower()
    for name in os.listdir(folder):
        if name.startswith("~$") and os.path.splitext(name)[1].lower() == ext:
            tail = name[2:]
            if base.startswith(tail) or tail.startswith(base):
                return os.path.join(folder, name)
    return None


class WordSession:
    def __init__(self):
        self.app = None

    def __enter__(self):
        if pythoncom is None or win32 is None:
            raise ToolError("缺少 pywin32，无法调用 Word。请先安装：pip install pywin32")
        pythoncom.CoInitialize()
        try:
            self.app = win32.DispatchEx("Word.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = 0
        except Exception as exc:
            pythoncom.CoUninitialize()
            raise ToolError(f"无法启动 Word：{exc}")
        return self.app

    def __exit__(self, exc_type, exc, tb):
        if self.app is not None:
            try:
                self.app.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        return False


def open_document(app, path, readonly):
    if not os.path.exists(path):
        raise ToolError(f"文件不存在：{path}")
    lock = find_lock_file(path)
    if lock and not readonly:
        raise ToolError(
            "文件正在 Word 中打开（检测到锁文件），请先关闭该文档再填充。",
            code=2,
            lock_file=lock,
        )
    doc = app.Documents.Open(os.path.abspath(path), ReadOnly=readonly, AddToRecentFiles=False, Visible=False)
    if not readonly:
        if doc.ReadOnly:
            doc.Close(0)
            raise ToolError("文档以只读方式打开，可能被占用或受保护。请关闭后重试。", code=2)
        try:
            if int(doc.ProtectionType) != -1:
                doc.Close(0)
                raise ToolError("文档处于保护状态，无法写入。请先在 Word 中取消保护。", code=2)
        except ToolError:
            raise
        except Exception:
            pass
    return doc


def doc_info(doc):
    info = {"pages": None, "words": None, "characters": None, "tables": None}
    try:
        info["pages"] = int(doc.ComputeStatistics(WD_STAT_PAGES))
    except Exception:
        pass
    try:
        info["words"] = int(doc.ComputeStatistics(WD_STAT_WORDS))
    except Exception:
        pass
    try:
        info["characters"] = int(doc.ComputeStatistics(WD_STAT_CHARS))
    except Exception:
        pass
    try:
        info["tables"] = int(doc.Tables.Count)
    except Exception:
        pass
    return info


def cell_font_info(cell):
    info = {"font": None, "size": None, "bold": None, "alignment": None, "paragraphs": None}
    try:
        first = cell.Range.Paragraphs(1).Range
        info["font"] = str(first.Font.NameFarEast or first.Font.Name or "") or None
        info["size"] = float(first.Font.Size) if first.Font.Size else None
        info["bold"] = bool(first.Font.Bold)
        info["alignment"] = int(first.ParagraphFormat.Alignment)
        info["paragraphs"] = int(cell.Range.Paragraphs.Count)
    except Exception:
        pass
    return info


def row_cells(table, row_index):
    try:
        return list(table.Rows(row_index).Cells)
    except Exception:
        return []


def locate_section_cell(doc, table_index, row_index, canonical):
    """按表格、行、栏目名重新定位，避免前面栏目写入后字符位置变化。"""
    try:
        table = doc.Tables(table_index)
    except Exception:
        return None
    for cell in row_cells(table, row_index):
        try:
            if label_key(norm(cell.Range.Text)) == canonical:
                return next_cell_in_row(cell)
        except Exception:
            continue
    return None


def next_cell_in_row(label_cell):
    try:
        table = label_cell.Range.Tables(1)
        cells = row_cells(table, label_cell.RowIndex)
    except Exception:
        return None
    start = label_cell.Range.Start
    for idx, cell in enumerate(cells):
        if cell.Range.Start == start:
            if idx + 1 < len(cells):
                return cells[idx + 1]
            return None
    return None


def collect_sections(doc):
    """返回模板中所有可识别的报告栏目及其填写单元格。"""
    found = []
    seen = set()
    try:
        table_count = int(doc.Tables.Count)
    except Exception:
        table_count = 0
    for ti in range(1, table_count + 1):
        table = doc.Tables(ti)
        try:
            cells = list(table.Range.Cells)
        except Exception:
            continue
        for cell in cells:
            try:
                key = label_key(norm(cell.Range.Text))
            except Exception:
                continue
            if not key:
                continue
            content_cell = next_cell_in_row(cell)
            if content_cell is None:
                continue
            signature = (ti, int(cell.RowIndex), content_cell.Range.Start)
            if signature in seen:
                continue
            seen.add(signature)
            raw = content_cell.Range.Text
            text = clean_text(raw)
            image_count = 0
            try:
                image_count = int(content_cell.Range.InlineShapes.Count)
            except Exception:
                pass
            found.append(
                {
                    "label": key,
                    "label_text": clean_text(cell.Range.Text),
                    "table": ti,
                    "row": int(cell.RowIndex),
                    "cell_start": int(content_cell.Range.Start),
                    "existing": bool(text) or image_count > 0,
                    "chars": len(text),
                    "images": image_count,
                    "preview": text[:160],
                    **cell_font_info(content_cell),
                }
            )
    found.sort(key=lambda item: (item["table"], item["row"]))
    return found


def collect_label_cells(doc):
    """返回所有短标签单元格（含 meta 字段），用于课程名称/实验名称等替换。"""
    rows = []
    try:
        table_count = int(doc.Tables.Count)
    except Exception:
        table_count = 0
    for ti in range(1, table_count + 1):
        table = doc.Tables(ti)
        try:
            cells = list(table.Range.Cells)
        except Exception:
            continue
        for cell in cells:
            try:
                text = clean_text(cell.Range.Text)
            except Exception:
                continue
            key = norm(text)
            if not key or len(key) > 12:
                continue
            content_cell = next_cell_in_row(cell)
            if content_cell is None:
                continue
            rows.append(
                {
                    "table": ti,
                    "row": int(cell.RowIndex),
                    "label": text,
                    "key": key,
                    "content_start": int(content_cell.Range.Start),
                    "content": clean_text(content_cell.Range.Text),
                    "cell": content_cell,
                }
            )
    return rows


def meta_paragraphs(doc, limit=80):
    """页眉信息总在正文前部，扫描前若干段即可，避免逐段 COM 调用拖慢整篇文档。"""
    out = []
    keys = list(META_PATTERNS.keys()) + ["课程名称", "实验名称", "成绩"]
    total = min(int(doc.Paragraphs.Count), int(limit))
    for i in range(1, total + 1):
        try:
            paragraph = doc.Paragraphs(i)
            if paragraph.Range.Information(WD_IN_TABLE):
                continue
            text = paragraph.Range.Text.replace("\r", "").replace("\x07", "")
        except Exception:
            continue
        if not text.strip() or len(text) > 220:
            continue
        hits = [k for k in keys if k in text]
        if hits:
            out.append({"paragraph": i, "start": int(paragraph.Range.Start), "hits": hits, "text": text})
    return out


def docx_text(path):
    """直接读 .docx XML 取纯文本，不启动 Word。"""
    from docx import Document

    document = Document(path)
    parts = [p.text for p in document.paragraphs]
    for table in document.tables:
        for row in table.rows:
            seen = []
            for cell in row.cells:
                if cell._tc in seen:
                    continue
                seen.append(cell._tc)
                parts.append(clean_text(cell.text))
    return "\n".join(part for part in parts if part)


def iter_docx_tables(tables, seen=None, prefix=""):
    """按文档顺序遍历顶层与嵌套表格，返回 (编号, 表格, 是否嵌套)。"""
    seen = seen if seen is not None else set()
    for table in tables:
        if id(table._tbl) in seen:
            continue
        seen.add(id(table._tbl))
        yield prefix + "top", table, False
        for row in table.rows:
            cells = []
            for cell in row.cells:
                if cell._tc in cells:
                    continue
                cells.append(cell._tc)
                for sub in cell.tables:
                    for item in iter_docx_tables([sub], seen):
                        yield item


def _docx_row_cells(row):
    cells = []
    seen = []
    for cell in row.cells:
        if cell._tc in seen:
            continue
        seen.append(cell._tc)
        cells.append(cell)
    return cells


def inspect_docx(path):
    """inspect 的快速路径：python-docx 读表，毫秒级。"""
    from docx import Document

    document = Document(path)
    sections = []
    labels = []
    table_no = 0
    for _, table, nested in iter_docx_tables(document.tables):
        table_no += 1
        for row_index, row in enumerate(table.rows, start=1):
            cells = _docx_row_cells(row)
            for cell_index, cell in enumerate(cells[:-1]):
                text = clean_text(cell.text)
                key = norm(text)
                if not key or len(key) > 12:
                    continue
                content = cells[cell_index + 1]
                content_text = clean_text(content.text)
                # .doc 转存后 Word 常用 VML(<w:pict>/<v:imagedata>) 存图片，两种都要数
                images = len(content._tc.xpath(".//*[local-name()='blip' or local-name()='imagedata']"))
                font = size = bold = None
                if content.paragraphs and content.paragraphs[0].runs:
                    first = content.paragraphs[0].runs[0]
                    font = first.font.name
                    size = float(first.font.size.pt) if first.font.size else None
                    bold = bool(first.font.bold)
                labels.append(
                    {
                        "table": table_no,
                        "row": row_index,
                        "label": text,
                        "key": key,
                        "content_start": None,
                        "content": content_text,
                    }
                )
                canonical = label_key(key)
                if not canonical:
                    continue
                sections.append(
                    {
                        "label": canonical,
                        "label_text": text,
                        "table": table_no,
                        "row": row_index,
                        "cell_start": None,
                        "existing": bool(content_text) or images > 0,
                        "chars": len(content_text),
                        "images": images,
                        "preview": content_text[:160],
                        "nested": nested,
                        "font": font,
                        "size": size,
                        "bold": bold,
                        "alignment": None,
                        "paragraphs": len(content.paragraphs),
                    }
                )

    metas = []
    keys = list(META_PATTERNS.keys()) + ["课程名称", "实验名称", "成绩"]
    for index, paragraph in enumerate(document.paragraphs, start=1):
        text = paragraph.text.replace("\x07", "")
        if not text.strip() or len(text) > 220:
            continue
        hits = [k for k in keys if k in text]
        if hits:
            metas.append({"paragraph": index, "start": None, "hits": hits, "text": text})

    filesize = os.path.getsize(path)
    info = {"pages": None, "words": None, "characters": None, "tables": table_no, "bytes": filesize}
    return {
        "ok": True,
        "file": path,
        "engine": "python-docx（未启动 Word）",
        "document": info,
        "sections": sections,
        "labels": labels,
        "meta_paragraphs": metas,
    }


def cell_content_end(cell):
    rng = cell.Range
    rng.MoveEnd(1, -1)
    return int(rng.End)
def clear_cell(cell):
    rng = cell.Range
    rng.MoveEnd(1, -1)
    rng.Delete()


def set_font(rng, east, latin, size, bold=False):
    rng.Font.NameFarEast = east
    rng.Font.NameAscii = latin
    rng.Font.NameOther = latin
    rng.Font.Size = float(size)
    rng.Font.Bold = -1 if bold else 0


def clear_paragraph_effects(rng):
    """清除从上一段继承的边框和底纹，避免占位框/代码块样式污染后续内容。"""
    try:
        rng.ParagraphFormat.Borders.Enable = False
    except Exception:
        pass
    try:
        rng.Shading.BackgroundPatternColorIndex = 0
    except Exception:
        pass
    try:
        rng.Shading.Texture = 0
        rng.Shading.BackgroundPatternColor = -16777216
    except Exception:
        pass


def format_body(rng, cfg, indent=True, align=WD_ALIGN_JUSTIFY, space_before=0, space_after=0, keep_with_next=False):
    clear_paragraph_effects(rng)
    set_font(rng, cfg["body_font"], cfg["latin_font"], cfg["body_size"], bold=False)
    pf = rng.ParagraphFormat
    pf.Alignment = align
    pf.LineSpacingRule = WD_LINE_1_5 if float(cfg["line_spacing"]) >= 1.5 else WD_LINE_SINGLE
    pf.SpaceBefore = float(space_before)
    pf.SpaceAfter = float(space_after)
    pf.LeftIndent = 0
    pf.RightIndent = 0
    if indent:
        pf.FirstLineIndent = 0
        pf.CharacterUnitFirstLineIndent = 2
    else:
        pf.FirstLineIndent = 0
        pf.CharacterUnitFirstLineIndent = 0
    try:
        pf.KeepWithNext = bool(keep_with_next)
    except Exception:
        pass


def format_heading(rng, cfg, keep_with_next=True):
    clear_paragraph_effects(rng)
    set_font(rng, cfg["heading_font"], cfg["latin_font"], cfg["heading_size"], bold=True)
    pf = rng.ParagraphFormat
    pf.Alignment = WD_ALIGN_LEFT
    pf.LineSpacingRule = WD_LINE_SINGLE
    pf.SpaceBefore = 6
    pf.SpaceAfter = 3
    pf.LeftIndent = 0
    pf.RightIndent = 0
    pf.FirstLineIndent = 0
    pf.CharacterUnitFirstLineIndent = 0
    try:
        pf.KeepWithNext = bool(keep_with_next)
    except Exception:
        pass


def choose_code_size(text, cell_width_pt, preferred):
    max_len = max([len(line) for line in text.split("\n")] + [0])
    available = max(160.0, float(cell_width_pt) - 12.0)
    sizes = [float(preferred), 8.5, 8.0, 7.5, 7.0]
    for size in sizes:
        if max_len * size * 0.56 <= available:
            return size
    return min(sizes)


def format_code(rng, cfg, text, cell_width_pt):
    size = choose_code_size(text, cell_width_pt, cfg["code_size"])
    clear_paragraph_effects(rng)
    rng.Font.NameFarEast = cfg["body_font"]
    rng.Font.NameAscii = cfg["code_font"]
    rng.Font.NameOther = cfg["code_font"]
    rng.Font.Size = size
    rng.Font.Bold = 0
    pf = rng.ParagraphFormat
    pf.Alignment = WD_ALIGN_LEFT
    pf.LineSpacingRule = WD_LINE_SINGLE
    pf.SpaceBefore = 0
    pf.SpaceAfter = 0
    pf.LeftIndent = 0
    pf.RightIndent = 0
    pf.FirstLineIndent = 0
    pf.CharacterUnitFirstLineIndent = 0
    try:
        rng.Shading.BackgroundPatternColor = 0xF5F5F5
        border = pf.Borders(WD_BORDER_LEFT)
        border.LineStyle = 1
        border.LineWidth = 12
        border.Color = 0xBFBFBF
    except Exception:
        pass
    return size


def insert_block_text(doc, cell, text, first):
    end = cell_content_end(cell)
    if not first:
        doc.Range(end, end).InsertAfter("\r")
        end = cell_content_end(cell)
    rng = doc.Range(end, end)
    rng.InsertAfter(text)
    return rng


def normalize_block_text(text):
    text = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return lines


def insert_table(doc, cell, block, cfg, first):
    headers = [str(h) for h in block.get("headers", [])]
    rows = block.get("rows", [])
    warnings = []
    if not headers and rows:
        headers = [""] * len(rows[0])
    if not headers:
        raise ToolError("table 块缺少 headers")
    ncols = len(headers)
    data = [headers]
    for row in rows:
        values = [str(v) for v in row][:ncols]
        values += [""] * (ncols - len(values))
        data.append(values)
    end = cell_content_end(cell)
    if not first:
        doc.Range(end, end).InsertAfter("\r")
        end = cell_content_end(cell)
    table = doc.Tables.Add(doc.Range(end, end), len(data), ncols)
    for r, values in enumerate(data, start=1):
        for c, value in enumerate(values, start=1):
            table.Cell(r, c).Range.Text = value
    try:
        table.Borders.Enable = True
        table.PreferredWidthType = WD_PREFERRED_WIDTH_PERCENT
        table.PreferredWidth = 100
    except Exception:
        pass
    if block.get("header_repeat", True):
        try:
            table.Rows(1).HeadingFormat = True
        except Exception:
            pass
        # Word 不支持嵌套表格跨页重复表头，行数多时建议拆表
        if len(data) - 1 > 3:
            warnings.append(
                f"表格有 {len(data) - 1} 行数据且嵌在单元格内，Word 无法跨页重复表头，"
                "建议拆成多个 3 行以内的表"
            )
    if block.get("no_split", True):
        try:
            table.Rows.AllowBreakAcrossPages = False
        except Exception:
            pass
    if str(block.get("row_height", "")).lower() in ("auto", "minimum"):
        try:
            for row in table.Rows:
                row.HeightRule = 0  # wdRowHeightAuto
        except Exception:
            pass
    widths = block.get("widths")
    try:
        cell_width = float(cell.Width)
    except Exception:
        cell_width = 420.0
    try:
        if widths and len(widths) == ncols:
            for idx, width_cm in enumerate(widths, start=1):
                table.Columns(idx).Width = float(width_cm) * CM_TO_PT
        else:
            each = max(60.0, (cell_width - 8.0) / ncols)
            for idx in range(1, ncols + 1):
                table.Columns(idx).Width = each
    except Exception:
        pass
    size = float(block.get("font_size") or cfg["table_size"])
    header_bold = bool(block.get("header_bold", True))
    for r in range(1, len(data) + 1):
        for c in range(1, ncols + 1):
            rng = table.Cell(r, c).Range
            set_font(rng, cfg["body_font"], cfg["latin_font"], size, bold=(r == 1 and header_bold))
            pf = rng.ParagraphFormat
            pf.Alignment = WD_ALIGN_CENTER
            pf.LineSpacingRule = WD_LINE_SINGLE
            pf.SpaceBefore = 0
            pf.SpaceAfter = 0
            pf.FirstLineIndent = 0
            pf.CharacterUnitFirstLineIndent = 0
    return table, warnings


def insert_image(doc, cell, block, cfg, first):
    path = block.get("path")
    if not path or not os.path.exists(path):
        raise ToolError(f"图片不存在：{path}")
    end = cell_content_end(cell)
    if not first:
        doc.Range(end, end).InsertAfter("\r")
        end = cell_content_end(cell)
    rng = doc.Range(end, end)
    shape = doc.InlineShapes.AddPicture(os.path.abspath(path), False, True, rng)
    try:
        shape.LockAspectRatio = -1
        shape.Width = float(block.get("width_cm", 12.0)) * CM_TO_PT
        img_para = shape.Range.Paragraphs(1).Range
        clear_paragraph_effects(img_para)
        pf = img_para.ParagraphFormat
        pf.Alignment = WD_ALIGN_CENTER
        pf.LeftIndent = 0
        pf.RightIndent = 0
        pf.CharacterUnitFirstLineIndent = 0
        pf.FirstLineIndent = 0
        pf.SpaceBefore = 3
        pf.SpaceAfter = 3
    except Exception:
        pass
    caption = block.get("caption")
    if caption:
        pos = int(shape.Range.End)
        cap = doc.Range(pos, pos)
        cap.InsertAfter("\r" + str(caption))
        clear_paragraph_effects(cap)
        set_font(cap, cfg["body_font"], cfg["latin_font"], cfg["caption_size"], bold=False)
        pf = cap.ParagraphFormat
        pf.Alignment = WD_ALIGN_CENTER
        pf.LineSpacingRule = WD_LINE_SINGLE
        pf.SpaceBefore = 0
        pf.SpaceAfter = 6
        pf.FirstLineIndent = 0
        pf.CharacterUnitFirstLineIndent = 0


def insert_placeholder(doc, cell, block, cfg, first):
    text = str(block.get("text", "【此处粘贴运行结果截图】")).replace("\r", "\v").replace("\n", "\v")
    note = block.get("note")
    if note:
        text = text + "\v" + str(note)
    insert_block_text(doc, cell, text, first=first)
    # 说明行用软换行放在同一段内，避免新段落继承边框/底纹
    ph_first = int(cell.Range.Paragraphs.Count)
    rng = cell.Range.Paragraphs(ph_first).Range
    clear_paragraph_effects(rng)
    set_font(rng, cfg["body_font"], cfg["latin_font"], cfg["caption_size"], bold=False)
    rng.Font.Color = 0x808080
    pf = rng.ParagraphFormat
    pf.Alignment = WD_ALIGN_CENTER
    pf.LineSpacingRule = WD_LINE_SINGLE
    pf.SpaceBefore = 6
    pf.SpaceAfter = 6
    pf.FirstLineIndent = 0
    pf.CharacterUnitFirstLineIndent = 0
    try:
        borders = pf.Borders
        borders.Enable = True
        for idx in (-1, -2, -3, -4):
            borders(idx).LineStyle = 1
            borders(idx).LineWidth = 6
            borders(idx).Color = 0xBFBFBF
    except Exception:
        pass


def insert_caption(doc, cell, block, cfg, first):
    text = " ".join(normalize_block_text(block.get("text", "")))
    if not text:
        return
    rng = insert_block_text(doc, cell, text, first)
    clear_paragraph_effects(rng)
    set_font(rng, cfg["body_font"], cfg["latin_font"], cfg["caption_size"], bold=False)
    pf = rng.ParagraphFormat
    pf.Alignment = WD_ALIGN_CENTER
    pf.LineSpacingRule = WD_LINE_SINGLE
    pf.SpaceBefore = float(block.get("space_before", 0))
    pf.SpaceAfter = float(block.get("space_after", 6))
    pf.FirstLineIndent = 0
    pf.CharacterUnitFirstLineIndent = 0
    try:
        pf.KeepWithNext = bool(block.get("keep_with_next", False))
    except Exception:
        pass


def write_blocks(doc, cell, blocks, cfg, starts_empty):
    first = bool(starts_empty)
    stats = {"paragraphs": 0, "code_lines": 0, "tables": 0, "images": 0, "placeholders": 0}
    warnings = []
    align_map = {"left": WD_ALIGN_LEFT, "center": WD_ALIGN_CENTER, "justify": WD_ALIGN_JUSTIFY}
    try:
        cell_width = float(cell.Width)
    except Exception:
        cell_width = 420.0
    for block in blocks:
        kind = str(block.get("type", "paragraph")).lower()
        if kind in ("paragraph", "body", "text"):
            lines = normalize_block_text(block.get("text", ""))
            if not lines:
                continue
            text = "\r".join(lines)
            rng = insert_block_text(doc, cell, text, first)
            align = align_map.get(str(block.get("align", "justify")).lower(), WD_ALIGN_JUSTIFY)
            format_body(
                rng,
                cfg,
                indent=bool(block.get("indent", True)),
                align=align,
                keep_with_next=bool(block.get("keep_with_next", False)),
            )
            stats["paragraphs"] += len(lines)
        elif kind in ("heading", "subheading", "title"):
            lines = normalize_block_text(block.get("text", ""))
            if not lines:
                continue
            rng = insert_block_text(doc, cell, "\r".join(lines), first)
            format_heading(rng, cfg, keep_with_next=bool(block.get("keep_with_next", True)))
            stats["paragraphs"] += len(lines)
        elif kind in ("caption", "figcaption"):
            insert_caption(doc, cell, block, cfg, first)
            stats["paragraphs"] += 1
        elif kind in ("list", "bullets", "steps"):
            items = block.get("items") or []
            if not items:
                continue
            style = str(block.get("style", "paren")).lower()
            marks = []
            for idx, _ in enumerate(items, start=1):
                if style in ("dot", "number", "arabic"):
                    marks.append(f"{idx}. ")
                elif style in ("circle", "cn"):
                    marks.append(f"{'①②③④⑤⑥⑦⑧⑨⑩'[idx - 1] if idx <= 10 else str(idx)} ")
                elif style in ("dash", "hyphen"):
                    marks.append("- ")
                else:
                    marks.append(f"（{idx}）")
            lines = []
            for mark, item in zip(marks, items):
                item_lines = normalize_block_text(item)
                if not item_lines:
                    continue
                lines.append(mark + item_lines[0])
                lines.extend(item_lines[1:])
            rng = insert_block_text(doc, cell, "\r".join(lines), first)
            format_body(rng, cfg, indent=False)
            stats["paragraphs"] += len(lines)
        elif kind == "code":
            lines = normalize_block_text(block.get("text", ""))
            if not lines:
                continue
            lines = [line.replace("\t", "    ").rstrip() for line in lines]
            text = "\r".join(lines)
            rng = insert_block_text(doc, cell, text, first)
            format_code(rng, cfg, "\n".join(lines), cell_width)
            stats["code_lines"] += len(lines)
        elif kind == "table":
            _, table_warnings = insert_table(doc, cell, block, cfg, first)
            warnings.extend(table_warnings)
            stats["tables"] += 1
        elif kind in ("image", "picture"):
            insert_image(doc, cell, block, cfg, first)
            stats["images"] += 1
        elif kind in ("placeholder", "screenshot"):
            insert_placeholder(doc, cell, block, cfg, first)
            stats["placeholders"] += 1
        else:
            raise ToolError(f"未知的内容块类型：{kind}")
        first = False
    if warnings:
        stats["warnings"] = warnings
    return stats


def set_cell_value(cell, value, cfg, center=True):
    old = cell_font_info(cell)
    clear_cell(cell)
    rng = cell.Range
    rng.MoveEnd(1, -1)
    rng.Text = str(value)
    size = old.get("size") or cfg["body_size"]
    set_font(cell.Range, old.get("font") or cfg["body_font"], cfg["latin_font"], size, bold=bool(old.get("bold")))
    pf = cell.Range.ParagraphFormat
    pf.Alignment = old.get("alignment") if old.get("alignment") is not None else (WD_ALIGN_CENTER if center else WD_ALIGN_LEFT)
    pf.FirstLineIndent = 0
    pf.CharacterUnitFirstLineIndent = 0
    pf.LineSpacingRule = WD_LINE_SINGLE
    pf.SpaceBefore = 0
    pf.SpaceAfter = 0


def replace_meta_in_paragraphs(doc, meta, cfg, limit=80):
    changes = []
    used = set()
    total = min(int(doc.Paragraphs.Count), int(limit))
    for i in range(1, total + 1):
        if len(used) >= len(meta):
            break
        try:
            paragraph = doc.Paragraphs(i)
            if paragraph.Range.Information(WD_IN_TABLE):
                continue
            text = paragraph.Range.Text.replace("\r", "")
        except Exception:
            continue
        for key, value in meta.items():
            if key in used or key not in META_PATTERNS:
                continue
            pattern, group = META_PATTERNS[key]
            match = re.search(pattern, text)
            if not match:
                continue
            start = match.start(group) if group else match.start(0)
            end = match.end(group) if group else match.end(0)
            abs_start = int(paragraph.Range.Start) + start
            abs_end = int(paragraph.Range.Start) + end
            rng = doc.Range(abs_start, abs_end)
            rng.Text = str(value)
            changes.append({"field": key, "old": match.group(group) if group else match.group(0), "new": str(value), "paragraph": i})
            used.add(key)
            break
    return changes


def run_inspect(path, need_meta=True, stats=False):
    with WordSession() as app:
        doc = open_document(app, path, readonly=True)
        try:
            sections = collect_sections(doc)
            labels = [{k: v for k, v in item.items() if k != "cell"} for item in collect_label_cells(doc)]
            metas = meta_paragraphs(doc) if need_meta else []
            if stats:
                info = doc_info(doc)
            else:
                try:
                    info = {"pages": None, "words": None, "characters": None, "tables": int(doc.Tables.Count)}
                except Exception:
                    info = {}
        finally:
            doc.Close(0)
    return {
        "ok": True,
        "file": path,
        "engine": "Word COM",
        "document": info,
        "sections": sections,
        "labels": labels,
        "meta_paragraphs": metas,
    }


def cmd_inspect(args):
    path = os.path.abspath(args.file)
    if not os.path.exists(path):
        fail(f"文件不存在：{path}")
    payload = None
    if path.lower().endswith(".docx") and not args.word:
        try:
            payload = inspect_docx(path)
        except Exception:
            payload = None
    if payload is None:
        payload = run_inspect(path, need_meta=not args.no_meta, stats=args.stats)
    emit(payload, args.out)


def cmd_extract(args):
    path = os.path.abspath(args.file)
    if not os.path.exists(path):
        fail(f"文件不存在：{path}")
    ext = os.path.splitext(path)[1].lower()
    engine = "plain"
    if ext == ".docx" and not args.word:
        try:
            text = docx_text(path)
            engine = "python-docx（未启动 Word）"
        except Exception:
            text = None
    else:
        text = None
    if text is None:
        if ext in WORD_EXTS:
            with WordSession() as app:
                doc = open_document(app, path, readonly=True)
                try:
                    text = doc.Content.Text.replace("\x07", "")
                finally:
                    doc.Close(0)
            engine = "Word COM"
        elif ext == ".pdf":
            try:
                import pymupdf
            except ImportError:
                fail("读取 PDF 需要 pymupdf：pip install pymupdf")
            document = pymupdf.open(path)
            text = "\n\n".join(page.get_text("text") for page in document)
            document.close()
        else:
            raw = open(path, "rb").read()
            for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
                try:
                    text = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                fail("无法识别文本编码，请手动转为 UTF-8 后再试")
    out = args.out or (os.path.splitext(path)[0] + ".extracted.txt")
    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(text)
    emit({"ok": True, "file": path, "out": out, "engine": engine, "chars": len(text), "preview": text[:300]})


def load_content(path):
    try:
        with open(path, "r", encoding="utf-8") as f:
            content = json.load(f)
    except Exception as exc:
        fail(f"读取 content.json 失败：{exc}")
    if not isinstance(content, dict):
        fail("content.json 顶层必须是对象")
    cfg = dict(DEFAULT_FORMAT)
    cfg.update(content.get("format") or {})
    return content, cfg


def validate_sections(content):
    """在启动 Word 之前校验栏目名，避免跑完一遍才发现标签写错。"""
    errors = []
    allowed = "、".join(canonical for canonical, _ in SECTION_ALIASES)
    for index, section in enumerate(content.get("sections") or [], start=1):
        label = str(section.get("label", "")).strip()
        if not label:
            errors.append(f"第 {index} 个 section 缺少 label")
            continue
        if label_key(norm(label)) is None:
            errors.append(f"第 {index} 个栏目「{label}」无法识别，可用栏目名：{allowed}")
        mode = str(section.get("mode", "fill")).lower()
        if mode not in ("fill", "replace", "append"):
            errors.append(f"栏目「{label}」的 mode={mode} 无效，应为 fill/replace/append")
    return errors


def fill_document(doc, content, cfg, meta_paragraphs=False):
    report = {"ok": True, "sections": [], "meta": [], "warnings": []}
    sections = collect_sections(doc)
    by_label = {}
    for item in sections:
        by_label.setdefault(item["label"], []).append(item)

    meta = {k: v for k, v in (content.get("meta") or {}).items() if v not in (None, "")}
    if meta:
        labels = collect_label_cells(doc)
        unresolved = {}
        for key, value in meta.items():
            normalized = norm(key)
            canonical = meta_key(normalized) or key
            target = None
            for item in labels:
                if item["key"] == normalized or (canonical in META_ALIASES and meta_key(item["key"]) == canonical):
                    target = item
                    break
            if target is None:
                unresolved[canonical] = value
                continue
            try:
                set_cell_value(target["cell"], value, cfg)
                report["meta"].append({"field": key, "action": "filled", "old": target["content"][:80]})
            except Exception as exc:
                report["warnings"].append(f"meta {key} 写入失败：{exc}")
        if unresolved:
            if meta_paragraphs:
                report["meta"].extend(replace_meta_in_paragraphs(doc, unresolved, cfg))
            else:
                report["warnings"].append(
                    "以下字段未在表格中找到，可在页眉段落中手动修改，或加 --meta-paragraphs 自动替换："
                    + "、".join(unresolved.keys())
                )

    for section in content.get("sections") or []:
        label = str(section.get("label", "")).strip()
        key = label_key(norm(label))
        if not key:
            report["sections"].append({"label": label, "action": "skipped", "reason": "未识别的栏目名"})
            report["warnings"].append(f"栏目「{label}」未识别，已跳过")
            continue
        candidates = by_label.get(key, [])
        if section.get("row") is not None:
            candidates = [c for c in candidates if c["row"] == int(section["row"])]
        if not candidates:
            report["sections"].append({"label": label, "action": "skipped", "reason": "模板中未找到该栏目"})
            report["warnings"].append(f"模板中未找到栏目「{label}」（row={section.get('row')}）")
            continue
        mode = str(section.get("mode", "fill")).lower()
        target = None
        if mode == "fill":
            target = next((c for c in candidates if not c["existing"]), None)
            if target is None:
                report["sections"].append({"label": label, "action": "skipped", "reason": "栏目已有内容（mode=fill）"})
                continue
        else:
            target = candidates[0]
        if mode not in ("fill", "replace", "append"):
            report["sections"].append({"label": label, "action": "skipped", "reason": f"未知 mode：{mode}"})
            continue

        # 前面的栏目写入后字符位置会整体后移，必须按行+栏目名重新定位
        cell = locate_section_cell(doc, target["table"], target["row"], key)
        if cell is None:
            report["sections"].append({"label": label, "action": "failed", "reason": "无法定位填写单元格"})
            report["warnings"].append(f"栏目「{label}」无法定位填写单元格")
            continue

        existing = bool(clean_text(cell.Range.Text))
        if mode == "replace" or not existing:
            clear_cell(cell)
            starts_empty = True
        else:
            starts_empty = False
        stats = write_blocks(doc, cell, section.get("blocks") or [], cfg, starts_empty)
        report["warnings"].extend(stats.pop("warnings", []))
        report["sections"].append(
            {
                "label": key,
                "row": target["row"],
                "action": "filled",
                "mode": mode,
                "chars": len(clean_text(cell.Range.Text)),
                **stats,
            }
        )

    filled = sum(1 for item in report["sections"] if item["action"] == "filled")
    skipped = sum(1 for item in report["sections"] if item["action"] != "filled")
    report["summary"] = f"填充 {filled} 个栏目，跳过 {skipped} 个"
    return report


def save_document(doc, out_path, source_path):
    out_path = os.path.abspath(out_path)
    folder = os.path.dirname(out_path)
    if folder:
        os.makedirs(folder, exist_ok=True)
    if out_path == os.path.abspath(source_path):
        doc.Save()
        return
    ext = os.path.splitext(out_path)[1].lower()
    if ext == ".pdf":
        doc.ExportAsFixedFormat(out_path, WD_FORMAT_PDF)
    elif ext == ".doc":
        doc.SaveAs2(out_path, FileFormat=WD_FORMAT_DOC)
    elif ext in (".docx", ".docm"):
        doc.SaveAs2(out_path, FileFormat=WD_FORMAT_DOCX)
    else:
        doc.SaveAs2(out_path)


def export_pdf(doc, pdf):
    pdf = os.path.abspath(pdf)
    folder = os.path.dirname(pdf)
    if folder:
        os.makedirs(folder, exist_ok=True)
    doc.ExportAsFixedFormat(pdf, WD_FORMAT_PDF)
    return pdf


def iter_word_tables(doc):
    """遍历顶层与嵌套表格，按 Range.Start 去重，返回 (table, nested)。"""
    seen = set()
    try:
        count = int(doc.Tables.Count)
    except Exception:
        return
    for i in range(1, count + 1):
        table = doc.Tables(i)
        try:
            start = int(table.Range.Start)
        except Exception:
            continue
        if start in seen:
            continue
        seen.add(start)
        yield table, False
        try:
            cells = list(table.Range.Cells)
        except Exception:
            continue
        for cell in cells:
            try:
                subs = list(cell.Tables)
            except Exception:
                continue
            for sub in subs:
                try:
                    sub_start = int(sub.Range.Start)
                except Exception:
                    continue
                if sub_start in seen:
                    continue
                seen.add(sub_start)
                yield sub, True


def _layout_targets(mode, nested):
    mode = str(mode or "off").strip().lower()
    if mode in ("all", "both", "true", "yes"):
        return True
    if mode in ("nested", "inner"):
        return nested
    if mode in ("top", "outer"):
        return not nested
    return False


def apply_layout(doc, layout):
    """content.json 的 layout 段：表格断行控制、清行高、居中图注。

    注意：顶层容器表格的单元格常常高过一页，给这种行设置“禁止跨页断行”
    会让 Word 直接裁掉放不下的正文，所以 no_split 默认只作用于嵌套小表。
    """
    stats = {
        "tables": 0,
        "rows_no_split": 0,
        "header_repeat": 0,
        "heights_cleared": 0,
        "captions_centered": 0,
        "warnings": [],
    }
    layout = layout or {}
    no_split = layout.get("no_split", "nested")
    header_repeat = layout.get("header_repeat", "off")
    for table, nested in iter_word_tables(doc):
        stats["tables"] += 1
        if _layout_targets(no_split, nested):
            try:
                table.Rows.AllowBreakAcrossPages = False
                stats["rows_no_split"] += int(table.Rows.Count)
            except Exception:
                pass
        if _layout_targets(header_repeat, nested):
            try:
                table.Rows(1).HeadingFormat = True
                stats["header_repeat"] += 1
            except Exception:
                pass

    for index in layout.get("clear_row_heights") or []:
        try:
            table = doc.Tables(int(index))
        except Exception:
            stats["warnings"].append(f"clear_row_heights：找不到第 {index} 个表格")
            continue
        try:
            for row in table.Rows:
                try:
                    row.HeightRule = 0  # wdRowHeightAuto
                    stats["heights_cleared"] += 1
                except Exception:
                    pass
        except Exception:
            stats["warnings"].append(f"clear_row_heights：第 {index} 个表格无法清行高")

    for prefix in layout.get("center_captions") or []:
        rng = doc.Content
        find = rng.Find
        try:
            find.ClearFormatting()
            find.Text = str(prefix)
            find.Forward = True
            find.Wrap = 0
            if find.Execute():
                rng.Paragraphs(1).Range.ParagraphFormat.Alignment = WD_ALIGN_CENTER
                stats["captions_centered"] += 1
        except Exception:
            stats["warnings"].append(f"center_captions：未能定位「{prefix}」")
    return stats


def render_preview(pdf, folder, pages, dpi, make_png=True):
    pdf = os.path.abspath(pdf)
    result = {"pdf": pdf, "pngs": [], "pages": None, "png_dir": None}
    if not make_png:
        try:
            import pymupdf
        except ImportError:
            return result
        document = pymupdf.open(pdf)
        result["pages"] = document.page_count
        document.close()
        return result
    try:
        import pymupdf
    except ImportError:
        result["warning"] = "未安装 pymupdf，只导出 PDF：pip install pymupdf"
        return result
    folder = os.path.abspath(folder)
    os.makedirs(folder, exist_ok=True)
    wanted = None
    if pages:
        wanted = set()
        for part in str(pages).split(","):
            part = part.strip()
            if "-" in part:
                a, b = part.split("-", 1)
                wanted.update(range(int(a), int(b) + 1))
            elif part:
                wanted.add(int(part))
    document = pymupdf.open(pdf)
    for index in range(document.page_count):
        page_no = index + 1
        if wanted and page_no not in wanted:
            continue
        pix = document[index].get_pixmap(dpi=dpi)
        name = os.path.join(folder, f"page-{page_no:02d}.png")
        pix.save(name)
        result["pngs"].append(name)
    result["pages"] = document.page_count
    result["png_dir"] = folder
    document.close()
    return result


def cmd_fill(args):
    path = os.path.abspath(args.file)
    if not os.path.exists(path):
        fail(f"文件不存在：{path}")
    content, cfg = load_content(args.content)
    errors = validate_sections(content)
    if errors:
        fail("content.json 校验失败：" + "；".join(errors))

    out_path = os.path.abspath(args.out) if args.out else path
    backup = None
    if out_path == path and not args.no_backup:
        stem, ext = os.path.splitext(path)
        backup = f"{stem}.bak-{time.strftime('%Y%m%d-%H%M%S')}{ext}"
        shutil.copy2(path, backup)

    report = {"ok": True, "file": path, "output": out_path, "backup": backup}
    with WordSession() as app:
        doc = open_document(app, path, readonly=False)
        try:
            report.update(fill_document(doc, content, cfg, meta_paragraphs=args.meta_paragraphs))
            if args.layout:
                report["layout"] = apply_layout(doc, content.get("layout") or {})
            save_document(doc, out_path, path)
            if args.stats:
                report["document"] = doc_info(doc)
        finally:
            doc.Close(0)

    emit(report, args.out_json)


def cmd_build(args):
    """一次 Word 会话完成：填充 → 排版修补 → 另存 → 导出 PDF/PNG。"""
    path = os.path.abspath(args.file)
    if not os.path.exists(path):
        fail(f"文件不存在：{path}")
    content, cfg = load_content(args.content)
    errors = validate_sections(content)
    if errors:
        fail("content.json 校验失败：" + "；".join(errors))

    if args.out:
        out_path = os.path.abspath(args.out)
    else:
        stem, ext = os.path.splitext(path)
        out_path = stem + "_filled" + ext
    backup = None
    if out_path == path and not args.no_backup:
        backup = f"{os.path.splitext(path)[0]}.bak-{time.strftime('%Y%m%d-%H%M%S')}{os.path.splitext(path)[1]}"
        shutil.copy2(path, backup)

    preview = not args.no_preview
    pdf = os.path.abspath(args.pdf) if args.pdf else os.path.splitext(out_path)[0] + ".preview.pdf"
    png_dir = os.path.abspath(args.png_dir) if args.png_dir else os.path.splitext(out_path)[0] + "_preview"

    report = {"ok": True, "file": path, "output": out_path, "backup": backup}
    with WordSession() as app:
        doc = open_document(app, path, readonly=False)
        try:
            report.update(fill_document(doc, content, cfg, meta_paragraphs=args.meta_paragraphs))
            report["layout"] = apply_layout(doc, content.get("layout") or {})
            try:
                report["placeholders_left"] = doc.Content.Text.count("【")
            except Exception:
                pass
            save_document(doc, out_path, path)
            if args.stats:
                report["document"] = doc_info(doc)
            if preview:
                export_pdf(doc, pdf)
        finally:
            doc.Close(0)

    if preview:
        report["preview"] = render_preview(pdf, png_dir, args.pages, args.dpi, not args.no_png)
    emit(report, args.out_json)


def cmd_preview(args):
    path = os.path.abspath(args.file)
    if not os.path.exists(path):
        fail(f"文件不存在：{path}")
    pdf = os.path.abspath(args.pdf) if args.pdf else os.path.splitext(path)[0] + ".preview.pdf"
    cached = os.path.exists(pdf) and os.path.getmtime(pdf) >= os.path.getmtime(path)
    if args.force or not cached:
        with WordSession() as app:
            doc = open_document(app, path, readonly=True)
            try:
                export_pdf(doc, pdf)
            finally:
                doc.Close(0)
    folder = args.png_dir or os.path.splitext(path)[0] + "_preview"
    result = render_preview(pdf, folder, args.pages, args.dpi, not args.no_png)
    result["cached"] = bool(cached and not args.force)
    result["ok"] = True
    emit(result)


def build_parser():
    parser = argparse.ArgumentParser(description="Word 实验报告模板检查与填充工具")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("inspect", help="读取模板结构")
    p.add_argument("file")
    p.add_argument("--out", help="把 JSON 结果写入文件")
    p.add_argument("--no-meta", action="store_true", help="跳过页眉 meta 段落扫描（更快）")
    p.add_argument("--stats", action="store_true", help="统计页数/字数（会触发重新分页，较慢）")
    p.add_argument("--word", action="store_true", help=".docx 也强制用 Word 读取（默认走 python-docx 快速路径）")
    p.set_defaults(func=cmd_inspect)

    p = sub.add_parser("extract", help="提取文档纯文本")
    p.add_argument("file")
    p.add_argument("--out", help="输出 txt 路径")
    p.add_argument("--word", action="store_true", help=".docx 也强制用 Word 读取")
    p.set_defaults(func=cmd_extract)

    p = sub.add_parser("fill", help="按 content.json 填充模板")
    p.add_argument("file")
    p.add_argument("--content", required=True, help="内容 JSON 文件")
    p.add_argument("--out", help="另存为新文件（不改原文件）")
    p.add_argument("--out-json", help="把填充摘要写入 JSON 文件")
    p.add_argument("--meta-paragraphs", action="store_true", help="同时替换页眉段落中的 meta 字段")
    p.add_argument("--layout", action="store_true", help="同时应用 content.json 的 layout 排版修补")
    p.add_argument("--stats", action="store_true", help="统计页数/字数（会触发重新分页，较慢）")
    p.add_argument("--no-backup", action="store_true", help="原地保存时不生成备份")
    p.set_defaults(func=cmd_fill)

    p = sub.add_parser("build", help="一次完成填充 + 排版修补 + 另存 + 预览（推荐）")
    p.add_argument("file")
    p.add_argument("--content", required=True, help="内容 JSON 文件")
    p.add_argument("--out", help="成品路径（默认模板名 + _filled）")
    p.add_argument("--out-json", help="把构建摘要写入 JSON 文件")
    p.add_argument("--meta-paragraphs", action="store_true", help="同时替换页眉段落中的 meta 字段")
    p.add_argument("--pdf", help="PDF 输出路径")
    p.add_argument("--png-dir", help="PNG 输出目录")
    p.add_argument("--pages", help="只渲染指定页，如 1,3-5")
    p.add_argument("--dpi", type=int, default=110)
    p.add_argument("--no-preview", action="store_true", help="不导出 PDF/PNG")
    p.add_argument("--no-png", action="store_true", help="只导出 PDF，不渲染 PNG")
    p.add_argument("--stats", action="store_true", help="统计页数/字数（会触发重新分页，较慢）")
    p.add_argument("--no-backup", action="store_true", help="原地保存时不生成备份")
    p.set_defaults(func=cmd_build)

    p = sub.add_parser("preview", help="导出 PDF/PNG 预览")
    p.add_argument("file")
    p.add_argument("--pdf", help="PDF 输出路径")
    p.add_argument("--png-dir", help="PNG 输出目录")
    p.add_argument("--pages", help="只渲染指定页，如 1,3-5")
    p.add_argument("--dpi", type=int, default=110)
    p.add_argument("--force", action="store_true", help="忽略已存在的较新 PDF，强制重新导出")
    p.add_argument("--no-png", action="store_true", help="只导出 PDF")
    p.set_defaults(func=cmd_preview)

    return parser


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    parser = build_parser()
    args = parser.parse_args()
    try:
        args.func(args)
    except ToolError as exc:
        payload = {"ok": False, "error": str(exc)}
        for key, value in vars(exc).items():
            if key not in ("args",):
                payload[key] = value
        emit(payload, code=exc.code)
    except Exception as exc:
        emit({"ok": False, "error": str(exc), "trace": traceback.format_exc(limit=3)}, code=1)


if __name__ == "__main__":
    main()
