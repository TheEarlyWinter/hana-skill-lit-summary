#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_docx.py - 化学/化工实验工作总结 Word 文档自动排版生成器
严格遵循：
- 正文字体：中文宋体 (SimSun) + 英文/数字 Times New Roman，小四 (12pt)，1.25倍行距
- 标题层级：
  * 大标题：小二 (18pt) 加粗 居中
  * 一级标题 (一、)：小三 (15pt) 加粗
  * 二级标题 ((一)、1.)：四号 (14pt) 加粗
  * 三级标题 ((1))：小四 (12pt) 加粗
- 表格：标准学术三线表 (顶底线 1.5pt，栏目线 0.75pt)，表头加粗居中，五号字 (10.5pt)
- 图表占位/图片：居中插入，配五号字居中图题
"""

import sys
import os
import json
import argparse
from typing import Dict, Any, List

import docx
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

FONT_CN = "宋体"
FONT_EN = "Times New Roman"

def set_run_font(run, font_name_cn=FONT_CN, font_name_en=FONT_EN, size_pt=12, bold=False, italic=False, color_rgb=None):
    """设置 run 的中西文字体、字号、加粗、斜体和颜色"""
    run.font.name = font_name_en
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
    
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name_en}" w:hAnsi="{font_name_en}" w:eastAsia="{font_name_cn}" w:cs="{font_name_en}"/>')
        rPr.append(rFonts)
    else:
        rFonts.set(qn('w:ascii'), font_name_en)
        rFonts.set(qn('w:hAnsi'), font_name_en)
        rFonts.set(qn('w:eastAsia'), font_name_cn)
        rFonts.set(qn('w:cs'), font_name_en)

def set_paragraph_spacing(paragraph, line_spacing=1.25, space_before_pt=0, space_after_pt=2):
    """设置段落行距与段前段后间距"""
    p_format = paragraph.paragraph_format
    p_format.line_spacing = line_spacing
    p_format.space_before = Pt(space_before_pt)
    p_format.space_after = Pt(space_after_pt)

def add_styled_paragraph(doc, text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, size_pt=12, bold=False, italic=False, line_spacing=1.25, space_before_pt=0, space_after_pt=2, color_rgb=None):
    """添加格式化段落"""
    p = doc.add_paragraph()
    p.alignment = align
    set_paragraph_spacing(p, line_spacing, space_before_pt, space_after_pt)
    if text:
        run = p.add_run(text)
        set_run_font(run, size_pt=size_pt, bold=bold, italic=italic, color_rgb=color_rgb)
    return p

def add_heading_1(doc, text):
    """一级标题：小三 15pt 加粗"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_spacing(p, line_spacing=1.25, space_before_pt=10, space_after_pt=4)
    run = p.add_run(text)
    set_run_font(run, size_pt=15, bold=True)
    return p

def add_heading_2(doc, text):
    """二级标题：四号 14pt 加粗"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_spacing(p, line_spacing=1.25, space_before_pt=6, space_after_pt=3)
    run = p.add_run(text)
    set_run_font(run, size_pt=14, bold=True)
    return p

def add_heading_3(doc, text):
    """三级标题：小四 12pt 加粗"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_paragraph_spacing(p, line_spacing=1.25, space_before_pt=4, space_after_pt=2)
    run = p.add_run(text)
    set_run_font(run, size_pt=12, bold=True)
    return p

def set_cell_border(cell, **kwargs):
    """
    设置单元格边框
    kwargs: top, bottom, left, right, insideH, insideV
    values: dict(val='single', sz='12', color='000000') 或 'none'
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key, val in edge_data.items():
                element.set(qn('w:{}'.format(key)), str(val))

def add_three_line_table(doc, headers: List[str], rows: List[List[str]], title: str = None):
    """
    添加标准学术三线表：
    - 顶线：1.5 磅 (sz=12)
    - 底线：1.5 磅 (sz=12)
    - 栏目线 (表头底线)：0.75 磅 (sz=6)
    - 内部横线/竖线：无
    """
    if title:
        p_title = add_styled_paragraph(doc, text=title, align=WD_ALIGN_PARAGRAPH.CENTER, size_pt=10.5, bold=True, space_before_pt=6, space_after_pt=3)
    
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    
    # 格式化表头
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        cell = hdr_cells[i]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p, line_spacing=1.15, space_before_pt=2, space_after_pt=2)
        run = p.add_run(str(header_text))
        set_run_font(run, size_pt=10.5, bold=True)
        # 表头边框：顶线 1.5pt，底线 0.75pt
        set_cell_border(cell, 
            top=dict(val='single', sz='12', color='000000'),
            bottom=dict(val='single', sz='6', color='000000'),
            left=dict(val='none'),
            right=dict(val='none')
        )
    
    # 格式化数据行
    for r_idx, row_data in enumerate(rows):
        row_cells = table.rows[r_idx + 1].cells
        is_last_row = (r_idx == len(rows) - 1)
        for c_idx, cell_value in enumerate(row_data):
            cell = row_cells[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_spacing(p, line_spacing=1.15, space_before_pt=1.5, space_after_pt=1.5)
            run = p.add_run(str(cell_value))
            set_run_font(run, size_pt=10.5, bold=False)
            
            # 边框设置
            border_kwargs = dict(left=dict(val='none'), right=dict(val='none'), top=dict(val='none'))
            if is_last_row:
                border_kwargs['bottom'] = dict(val='single', sz='12', color='000000') # 底线 1.5pt
            else:
                border_kwargs['bottom'] = dict(val='none')
            set_cell_border(cell, **border_kwargs)
            
    # 加一个空行缓冲
    add_styled_paragraph(doc, text="", space_before_pt=0, space_after_pt=4)

def add_image_or_placeholder(doc, image_path: str = None, caption: str = "", placeholder: str = ""):
    """插入真实图片或居中图占位说明，并在下方添加居中五号图题"""
    # 1. 尝试插入图片
    image_inserted = False
    if image_path and os.path.isfile(image_path):
        try:
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_spacing(p_img, line_spacing=1.0, space_before_pt=6, space_after_pt=2)
            run = p_img.add_run()
            # 限制最大宽度 14cm
            run.add_picture(image_path, width=Cm(13.5))
            image_inserted = True
        except Exception as e:
            print(f"[Warning] Failed to insert image {image_path}: {e}", file=sys.stderr)
            image_inserted = False

    # 2. 如果没有成功插入图片，则渲染居中图占位说明
    if not image_inserted:
        ph_text = placeholder if placeholder else f"【{caption}】"
        p_ph = doc.add_paragraph()
        p_ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p_ph, line_spacing=1.15, space_before_pt=6, space_after_pt=2)
        run_ph = p_ph.add_run(ph_text)
        set_run_font(run_ph, size_pt=10.5, bold=False, italic=True, color_rgb=RGBColor(80, 80, 80))

    # 3. 添加居中图题 (五号 10.5pt)
    if caption:
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p_cap, line_spacing=1.15, space_before_pt=2, space_after_pt=6)
        run_cap = p_cap.add_run(caption)
        set_run_font(run_cap, size_pt=10.5, bold=False)

def build_document_from_dict(data: Dict[str, Any], output_path: str):
    """根据结构化字典生成 Word 文档"""
    doc = Document()
    
    # 页面边距：标准 A4，上下 2.54cm，左右 3.18cm (或 2.5cm)
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.8)
        section.right_margin = Cm(2.8)
        section.header_distance = Cm(1.5)
        section.footer_distance = Cm(1.5)
        
    # 文档大标题
    title = data.get("title", "实验工作总结")
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(p_title, line_spacing=1.25, space_before_pt=12, space_after_pt=6)
    r_title = p_title.add_run(title)
    set_run_font(r_title, size_pt=18, bold=True)
    
    # 日期或副标题（如果有）
    subtitle = data.get("subtitle") or data.get("date")
    if subtitle:
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_paragraph_spacing(p_sub, line_spacing=1.15, space_before_pt=0, space_after_pt=10)
        r_sub = p_sub.add_run(str(subtitle))
        set_run_font(r_sub, size_pt=10.5, bold=False, color_rgb=RGBColor(100, 100, 100))
        
    # 遍历处理各个章节
    sections = data.get("sections", [])
    for sec in sections:
        sec_title = sec.get("title", "")
        level = sec.get("level", 1)
        
        if sec_title:
            if level == 1:
                add_heading_1(doc, sec_title)
            elif level == 2:
                add_heading_2(doc, sec_title)
            elif level == 3:
                add_heading_3(doc, sec_title)
            else:
                add_heading_1(doc, sec_title)
                
        # 章节内的内容块
        blocks = sec.get("content", [])
        for blk in blocks:
            b_type = blk.get("type", "paragraph")
            
            if b_type == "paragraph":
                text = blk.get("text", "")
                bold = blk.get("bold", False)
                align_str = blk.get("align", "justify")
                align = WD_ALIGN_PARAGRAPH.CENTER if align_str == "center" else WD_ALIGN_PARAGRAPH.JUSTIFY
                add_styled_paragraph(doc, text=text, align=align, size_pt=12, bold=bold)
                
            elif b_type == "list_item":
                text = blk.get("text", "")
                add_styled_paragraph(doc, text=text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size_pt=12, space_before_pt=1, space_after_pt=2)
                
            elif b_type == "table":
                headers = blk.get("headers", [])
                rows = blk.get("rows", [])
                t_title = blk.get("title", "")
                add_three_line_table(doc, headers=headers, rows=rows, title=t_title)
                
            elif b_type == "image":
                img_path = blk.get("image_path")
                caption = blk.get("caption", "")
                placeholder = blk.get("placeholder", "")
                add_image_or_placeholder(doc, image_path=img_path, caption=caption, placeholder=placeholder)
                
    # 确保目标目录存在
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    print(f"[Success] Generated document: {output_path}")

def main():
    parser = argparse.ArgumentParser(description="Generate styled docx for chemical experiment summaries.")
    parser.add_argument("--input-json", "-i", required=True, help="Path to input json structure.")
    parser.add_argument("--output", "-o", required=True, help="Path to output docx file.")
    args = parser.parse_args()
    
    if not os.path.exists(args.input_json):
        print(f"Error: input file '{args.input_json}' not found.", file=sys.stderr)
        sys.exit(1)
        
    with open(args.input_json, "r", encoding="utf-8") as f:
        data = json.load(f)
        
    build_document_from_dict(data, args.output)

if __name__ == "__main__":
    main()
