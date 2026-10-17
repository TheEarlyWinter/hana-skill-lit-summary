#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_figures.py - 从文献 PDF 中抽取关键图，或按需渲染指定页

用法：
  # 抽取内嵌位图（默认，过滤过小图标）
  python extract_figures.py --pdf paper.pdf --output-dir ./figs

  # 图片是矢量时，只渲染指定页（如机理图在第 3、5 页）
  python extract_figures.py --pdf paper.pdf --output-dir ./figs --pages 3,5 --dpi 300

输出：逐条打印 JSON（path / page / type / width / height），便于直接塞进 generate_docx.py 的 image 块。
"""

import argparse
import json
import os
import sys

try:
    import fitz  # PyMuPDF
except ImportError:
    sys.exit("需要 PyMuPDF：python -m pip install pymupdf")


def extract_embedded(doc, output_dir, base, min_w, min_h):
    items = []
    idx = 0
    for page_num in range(len(doc)):
        page = doc[page_num]
        for img_info in page.get_images(full=True):
            xref = img_info[0]
            try:
                base_image = doc.extract_image(xref)
            except Exception:
                continue
            w, h = base_image.get("width", 0), base_image.get("height", 0)
            if w < min_w or h < min_h:
                continue
            idx += 1
            ext = base_image.get("ext", "png")
            path = os.path.join(output_dir, f"{base}_p{page_num + 1}_fig{idx}.{ext}")
            with open(path, "wb") as f:
                f.write(base_image["image"])
            items.append({"path": os.path.abspath(path), "page": page_num + 1,
                          "type": "embedded", "width": w, "height": h})
    return items


def render_pages(doc, output_dir, base, pages, dpi):
    items = []
    zoom = dpi / 72.0
    for page_num in pages:
        if page_num < 1 or page_num > len(doc):
            print(f"[skip] 第 {page_num} 页超出范围（共 {len(doc)} 页）", file=sys.stderr)
            continue
        page = doc[page_num - 1]
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        path = os.path.join(output_dir, f"{base}_page{page_num}.png")
        pix.save(path)
        items.append({"path": os.path.abspath(path), "page": page_num,
                      "type": "rendered", "width": pix.width, "height": pix.height})
    return items


def main():
    ap = argparse.ArgumentParser(description="Extract figures or render pages from a paper PDF.")
    ap.add_argument("--pdf", "-p", required=True)
    ap.add_argument("--output-dir", "-o", required=True)
    ap.add_argument("--pages", default="", help="逗号分隔页码，只渲染这些页，例如 3,5")
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--min-width", type=int, default=200)
    ap.add_argument("--min-height", type=int, default=200)
    ap.add_argument("--no-embedded", action="store_true", help="不抽内嵌位图，只渲染指定页")
    args = ap.parse_args()

    if not os.path.isfile(args.pdf):
        sys.exit(f"找不到 PDF：{args.pdf}")
    os.makedirs(args.output_dir, exist_ok=True)

    doc = fitz.open(args.pdf)
    base = os.path.splitext(os.path.basename(args.pdf))[0]
    base = "".join(c for c in base if c.isalnum() or c in "-_")[:40] or "paper"

    items = []
    if not args.no_embedded:
        items += extract_embedded(doc, args.output_dir, base, args.min_width, args.min_height)

    if args.pages.strip():
        pages = [int(x) for x in args.pages.replace("，", ",").split(",") if x.strip()]
        items += render_pages(doc, args.output_dir, base, pages, args.dpi)

    doc.close()

    if not items:
        print("未抽到图。可能是矢量图或扫描页，试加 --pages 指定页码。", file=sys.stderr)

    print(json.dumps(items, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
