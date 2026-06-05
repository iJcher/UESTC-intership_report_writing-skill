# -*- coding: utf-8 -*-
"""
将 Markdown 章节合并生成一份套用 Word 模板样式的 .docx。

用法:
    python build_docx.py --template <模板.docx> --input <md目录或文件...> \
                         --output <输出.docx> [--title "文档标题"]

特性:
- 标题 #/##/###/#### -> Heading 1~4（继承模板内置样式）
- 表格 -> 带边框 Word 表格
- ```mermaid 代码块 -> 联网渲染为 PNG 图片插入（Kroki 优先，mermaid.ink 兜底）
- 其他 ``` 代码块 -> 等宽字体代码块
- 图题（图X-Y ...）-> 居中小字
- **加粗** / `行内代码` -> 对应 run 格式

依赖: python-docx（pip install python-docx）。Mermaid 渲染需联网。
"""
import sys
import io
import re
import os
import glob
import base64
import argparse
import tempfile
import urllib.request

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CAPTION_RE = re.compile(r'^(图|表)\s*\d+[-－]\d+')  # 图题（图X-Y）与表题（表X-Y）统一居中小字
REF_RE = re.compile(r'^\[\d+\]')
INLINE_RE = re.compile(r'(\*\*.+?\*\*|`[^`]+`|\[\d+\])')

# 字体规格（依据已验收的初期报告标准）
EN_FONT = 'Times New Roman'
CN_BODY = '宋体'
CN_HEAD = '黑体'
BODY_SIZE = 12.0     # 正文 小四
REF_SIZE = 10.5      # 参考文献 五号
TABLE_SIZE = 10.5    # 表格 五号
FIG_SIZE = 9.0       # 图题
TITLE_SIZE = 18.0    # 文档标题
HEADING_SIZE = {1: 15.0, 2: 14.0, 3: 14.0, 4: 14.0}  # 一级小三，其余四号

STYLE_BY_ID = {}
H1_SEEN = [0]  # 一级标题（章）计数：从第二个起段前分页，第一个紧跟文档标题


def get_style(style_id):
    return STYLE_BY_ID.get(style_id)


def clear_body(doc):
    body = doc.element.body
    for child in list(body):
        if child.tag == qn('w:sectPr'):
            continue
        body.remove(child)


def set_font(run, cn=CN_BODY, size=BODY_SIZE, bold=False, superscript=False, en=EN_FONT):
    """显式设置中英文字体、字号，确保与标准报告一致。"""
    run.bold = bold
    if size is not None:
        run.font.size = Pt(size)
    if superscript:
        run.font.superscript = True
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:ascii'), en)
    rfonts.set(qn('w:hAnsi'), en)
    rfonts.set(qn('w:eastAsia'), cn)


def set_code_font(run):
    run.font.size = Pt(9)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:ascii'), 'Consolas')
    rfonts.set(qn('w:hAnsi'), 'Consolas')


def add_inline(paragraph, text, size=BODY_SIZE, cn=CN_BODY, sup_refs=True):
    """处理 **加粗**、`行内代码`、[n]上标引用，其余为普通 run。"""
    for part in INLINE_RE.split(text):
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            set_font(paragraph.add_run(part[2:-2]), cn, size, bold=True)
        elif part.startswith('`') and part.endswith('`'):
            set_code_font(paragraph.add_run(part[1:-1]))
        elif sup_refs and re.fullmatch(r'\[\d+\]', part):
            set_font(paragraph.add_run(part), cn, size, superscript=True)
        else:
            set_font(paragraph.add_run(part), cn, size)


def png_size(data):
    if data[:8] == b'\x89PNG\r\n\x1a\n':
        return int.from_bytes(data[16:20], 'big'), int.from_bytes(data[20:24], 'big')
    return None, None


def _kroki(code):
    req = urllib.request.Request(
        'https://kroki.io/mermaid/png', data=code.encode('utf-8'),
        headers={'Content-Type': 'text/plain', 'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=45) as r:
        return r.read()


def _mermaid_ink(code):
    b64 = base64.urlsafe_b64encode(code.encode('utf-8')).decode('ascii')
    url = 'https://mermaid.ink/img/' + b64 + '?type=png&bgColor=ffffff'
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=45) as r:
        return r.read()


def render_mermaid(code):
    # 优先 Kroki（POST 无 URL 长度限制），失败再退回 mermaid.ink
    try:
        return _kroki(code)
    except Exception:
        return _mermaid_ink(code)


def add_image(doc, png_bytes, max_cm=15.0):
    fd, tmp = tempfile.mkstemp(suffix='.png')
    os.close(fd)
    try:
        with open(tmp, 'wb') as f:
            f.write(png_bytes)
        w, _ = png_size(png_bytes)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        if w:
            cm_w = w / 96 * 2.54
            run.add_picture(tmp, width=Cm(min(cm_w, max_cm)))
        else:
            run.add_picture(tmp, width=Cm(max_cm))
    finally:
        os.remove(tmp)


def set_code_box(paragraph):
    """给段落加四周边框（白底无填充），形成代码框。"""
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:space'), '6')
        el.set(qn('w:color'), '999999')
        pbdr.append(el)
    pPr.append(pbdr)


def add_code_block(doc, lines):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    set_code_box(p)
    for i, ln in enumerate(lines):
        if i > 0:
            p.add_run().add_break()
        set_code_font(p.add_run(ln))


def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '000000')
        borders.append(el)
    tblPr.append(borders)


def add_table(doc, rows):
    ncol = len(rows[0])
    table = doc.add_table(rows=0, cols=ncol)
    set_table_borders(table)
    for ri, row in enumerate(rows):
        cells = table.add_row().cells
        for ci in range(ncol):
            txt = (row[ci] if ci < len(row) else '').replace('**', '').replace('`', '')
            cells[ci].text = ''
            run = cells[ci].paragraphs[0].add_run(txt)
            set_font(run, CN_BODY, TABLE_SIZE, bold=(ri == 0))
    return table


def split_table_row(line):
    s = line.strip().strip('|')
    return [c.strip() for c in s.split('|')]


def is_separator_row(line):
    return bool(re.match(r'^\s*\|?\s*:?-{2,}', line)) and set(line.strip()) <= set('|-: ')


def collect_md(inputs):
    """inputs 可为目录（取其中 *.md 排序）或文件路径，按给定顺序合并。"""
    files = []
    for p in inputs:
        if os.path.isdir(p):
            files.extend(sorted(glob.glob(os.path.join(p, '*.md'))))
        else:
            files.append(p)
    return files


def render_md(doc, lines, fig_counter):
    i, n = 0, len(lines)
    while i < n:
        line = lines[i]
        stripped = line.strip()

        if stripped.startswith('```'):
            lang = stripped[3:].strip().lower()
            buf = []
            i += 1
            while i < n and not lines[i].strip().startswith('```'):
                buf.append(lines[i])
                i += 1
            i += 1
            if lang == 'mermaid':
                fig_counter[0] += 1
                idx = fig_counter[0]
                try:
                    add_image(doc, render_mermaid('\n'.join(buf)))
                    print(f'  [图片] 第{idx}张 mermaid 渲染成功')
                except Exception as e:
                    print(f'  [警告] 第{idx}张 mermaid 渲染失败: {e}，改插文本')
                    add_code_block(doc, buf)
            else:
                add_code_block(doc, buf)
            continue

        if not stripped:
            i += 1
            continue

        m = re.match(r'^(#{1,4})\s+(.*)$', stripped)
        if m:
            level = len(m.group(1))
            p = doc.add_paragraph()
            hs = get_style(f'Heading{level}')
            if hs is not None:
                p.style = hs
            if level == 1:
                H1_SEEN[0] += 1
                if H1_SEEN[0] > 1:
                    pPr = p._p.get_or_add_pPr()
                    if pPr.find(qn('w:pageBreakBefore')) is None:
                        pPr.insert(0, OxmlElement('w:pageBreakBefore'))
                print(f'  [章] 第{H1_SEEN[0]}个一级标题 "{m.group(2)[:16]}" 段前分页={H1_SEEN[0] > 1}')
            set_font(p.add_run(m.group(2)), CN_HEAD, HEADING_SIZE.get(level, 14.0))
            i += 1
            continue

        if stripped.startswith('|'):
            tbl = [split_table_row(line)]
            i += 1
            while i < n and lines[i].strip().startswith('|'):
                if not is_separator_row(lines[i]):
                    tbl.append(split_table_row(lines[i]))
                i += 1
            add_table(doc, tbl)
            continue

        if CAPTION_RE.match(stripped):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font(p.add_run(stripped), CN_BODY, FIG_SIZE)
            i += 1
            continue

        # 参考文献条目：五号字，开头 [n] 不作上标
        if REF_RE.match(stripped):
            p = doc.add_paragraph()
            add_inline(p, stripped, size=REF_SIZE, sup_refs=False)
            i += 1
            continue

        if stripped.startswith('- '):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.74)
            set_font(p.add_run('• '), CN_BODY, BODY_SIZE)
            add_inline(p, stripped[2:])
            i += 1
            continue

        p = doc.add_paragraph()
        add_inline(p, stripped)
        i += 1


def build(template, inputs, output, title=None):
    doc = Document(template)
    STYLE_BY_ID.clear()
    H1_SEEN[0] = 0
    for s in doc.styles:
        STYLE_BY_ID[s.style_id] = s
    clear_body(doc)

    if title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(p.add_run(title), CN_HEAD, TITLE_SIZE, bold=True)

    files = collect_md(inputs)
    if not files:
        raise SystemExit('未找到任何输入 md 文件')
    fig_counter = [0]
    for path in files:
        print(f'处理: {path}')
        with open(path, encoding='utf-8') as f:
            render_md(doc, f.read().split('\n'), fig_counter)

    doc.save(output)
    print(f'\n已生成: {output}（共 {fig_counter[0]} 张图）')


def main():
    parser = argparse.ArgumentParser(description='Markdown 章节 + Word 模板 -> 排版好的 .docx')
    parser.add_argument('--template', '-t', required=True, help='Word 模板 .docx 路径（提供标题/正文样式）')
    parser.add_argument('--input', '-i', nargs='+', required=True, help='章节 md 目录或文件列表（按顺序合并）')
    parser.add_argument('--output', '-o', required=True, help='输出 .docx 路径')
    parser.add_argument('--title', default=None, help='可选：文档首行居中标题')
    args = parser.parse_args()
    build(args.template, args.input, args.output, args.title)


if __name__ == '__main__':
    main()
