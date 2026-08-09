#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gao-science-ppt-skill: PPTX 排版统一应用脚本
按方案 JSON 执行: 标题重写 / 字体统一(中英分离) / 字号统一 / 图片统一(等宽或裁剪) / 教学设计页填充。
只改排版不改文字事实。默认输出新文件, 不覆盖原件。

用法:
  python3 pptx_apply.py -i deck.pptx -o deck-优化.pptx --plan plan.json [--append-teaching teaching.json]

方案 JSON 结构:
{
  "titles": {"<页号>": "<新标题>"},                 # 页号 = 诊断 JSON 的 slide.index
  "fonts": {"latin": "Calibri", "ea": "微软雅黑"},  # latin=西文/数字, ea=中文
  "size_map": {"18": 20, "16": 20},                 # 旧字号pt → 新字号pt
  "images": {"mode": "equal_width",                 # equal_width | crop_to_ratio | none
             "width_in": 5.5, "ratio": "4:3",
             "align": "center", "border": true},
  "teaching": {...}                                  # 或单独 --append-teaching
}
"""
import argparse
import json
import os
import sys

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.oxml.ns import qn
except ImportError:
    sys.exit("缺少 python-pptx，请先: python3 -m pip install --user python-pptx")

GRAY = RGBColor(0xD9, 0xD9, 0xD9)
DARK = RGBColor(0x33, 0x33, 0x33)


def set_run_fonts(run, latin=None, ea=None):
    """同时设置 latin(西文) 与 ea(中文) 字体。"""
    if latin:
        run.font.name = latin
    if ea:
        rPr = run._r.get_or_add_rPr()
        for tag in (qn("a:ea"), qn("a:cs")):
            el = rPr.find(tag)
            if el is None:
                el = rPr.makeelement(tag, {})
                rPr.append(el)
            el.set("typeface", ea)


def iter_runs(shape):
    if not getattr(shape, "has_text_frame", False):
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            yield run


def replace_text(shape, new_text):
    """把 shape 文本替换为新文本(保留首 run 格式)。"""
    tf = shape.text_frame
    runs = tf.paragraphs[0].runs
    if runs:
        runs[0].text = new_text
        for extra in runs[1:]:
            extra.text = ""
    else:
        tf.text = new_text
    # 清掉其余段落的文字
    for para in tf.paragraphs[1:]:
        for run in para.runs:
            run.text = ""


def apply_titles(prs, title_map):
    """按页号替换标题: 找该页字号最大的文本 shape 替换。"""
    for idx, slide in enumerate(prs.slides, 1):
        new_title = title_map.get(str(idx))
        if not new_title:
            continue
        best = None
        best_sz = -1
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            sz = 0
            for run in iter_runs(shape):
                if run.font.size and run.font.size.pt > sz:
                    sz = run.font.size.pt
            if sz > best_sz:
                best, best_sz = shape, sz
        if best is not None:
            replace_text(best, new_title)


def apply_fonts(prs, fonts):
    latin = fonts.get("latin")
    ea = fonts.get("ea")
    for slide in prs.slides:
        for shape in slide.shapes:
            for run in iter_runs(shape):
                set_run_fonts(run, latin=latin, ea=ea)


def apply_sizes(prs, size_map):
    smap = {float(k): float(v) for k, v in size_map.items()}
    for slide in prs.slides:
        for shape in slide.shapes:
            for run in iter_runs(shape):
                if run.font.size and run.font.size.pt in smap:
                    run.font.size = Pt(smap[run.font.size.pt])


def apply_images(prs, cfg):
    mode = cfg.get("mode", "equal_width")
    if mode == "none":
        return
    width_in = cfg.get("width_in", 5.5)
    ratio_spec = cfg.get("ratio", "4:3")
    try:
        rw, rh = (float(x) for x in ratio_spec.split(":"))
        target_ratio = rw / rh
    except Exception:
        target_ratio = 4 / 3
    slide_w = prs.slide_width  # EMU
    for slide in prs.slides:
        pics = [s for s in slide.shapes if s.shape_type == 13]
        if not pics:
            continue
        max_w = max(p.width for p in pics)
        width = min(Inches(width_in), max_w)
        for p in pics:
            orig_ratio = (p.width / p.height) if p.height else target_ratio
            p.width = width
            if mode == "equal_width":
                p.height = int(width / orig_ratio)
            else:  # crop_to_ratio: 等宽 + 中心裁剪到目标比例
                p.height = int(width / target_ratio)
                _set_crop(p, orig_ratio, target_ratio)
            if cfg.get("align") == "center":
                p.left = int((slide_w - width) / 2)
            if cfg.get("border"):
                p.line.color.rgb = GRAY
                p.line.width = Pt(1)


def _set_crop(pic, orig_ratio, target_ratio):
    """中心裁剪: 用 crop 属性切到目标比例(不拉伸)。"""
    if orig_ratio > target_ratio:  # 太宽 → 裁左右
        total = 1 - target_ratio / orig_ratio
        pic.crop_left = total / 2
        pic.crop_right = total / 2
        pic.crop_top = 0
        pic.crop_bottom = 0
    else:  # 太高 → 裁上下
        total = 1 - orig_ratio / target_ratio
        pic.crop_top = total / 2
        pic.crop_bottom = total / 2
        pic.crop_left = 0
        pic.crop_right = 0


def add_teaching_slide(prs, teaching):
    """在末尾插入「教学设计」页。"""
    layout = None
    for l in prs.slide_layouts:
        if l.name and ("blank" in l.name.lower() or "空白" in l.name):
            layout = l
            break
    if layout is None:
        layout = prs.slide_layouts[-1]
    slide = prs.slides.add_slide(layout)
    sw = prs.slide_width
    # 标题
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3),
                                  sw - Inches(1.0), Inches(0.8))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = teaching.get("title", "教学设计")
    run.font.size = Pt(32)
    run.font.bold = True
    run.font.color.rgb = DARK
    # 内容
    body = slide.shapes.add_textbox(Inches(0.8), Inches(1.3),
                                    sw - Inches(1.6), Inches(5.0))
    bf = body.text_frame
    bf.word_wrap = True
    sections = [
        ("教学目标", teaching.get("goals", [])),
        ("教学重难点", teaching.get("key_points", [])
         + teaching.get("difficult_points", [])),
        ("教学环节", teaching.get("stages", [])),
    ]
    first = True
    for label, items in sections:
        if not items:
            continue
        p = bf.paragraphs[0] if first else bf.add_paragraph()
        first = False
        run = p.add_run()
        run.text = f"【{label}】"
        run.font.size = Pt(20)
        run.font.bold = True
        run.font.color.rgb = DARK
        for item in items:
            p = bf.add_paragraph()
            run = p.add_run()
            run.text = f"· {item}"
            run.font.size = Pt(18)


def main():
    ap = argparse.ArgumentParser(description="PPTX 排版统一应用")
    ap.add_argument("-i", "--input", required=True, help="输入 .pptx")
    ap.add_argument("-o", "--output", required=True, help="输出 .pptx")
    ap.add_argument("--plan", help="方案 JSON")
    ap.add_argument("--append-teaching", help="教学设计 JSON(单独)或留空用 plan.teaching")
    args = ap.parse_args()

    if not args.input.endswith(".pptx"):
        sys.exit("仅支持 .pptx（.ppt 请先用 Office/LibreOffice 另存为 .pptx）")
    if args.output == args.input:
        sys.exit("输出与输入相同, 会覆盖原件; 请换输出名或确认覆盖意图")

    plan = {}
    if args.plan:
        with open(args.plan, encoding="utf-8") as f:
            plan = json.load(f)

    prs = Presentation(args.input)

    if plan.get("titles"):
        apply_titles(prs, plan["titles"])
    if plan.get("fonts"):
        apply_fonts(prs, plan["fonts"])
    if plan.get("size_map"):
        apply_sizes(prs, plan["size_map"])
    if plan.get("images"):
        apply_images(prs, plan["images"])

    teaching = plan.get("teaching")
    if args.append_teaching:
        with open(args.append_teaching, encoding="utf-8") as f:
            teaching = json.load(f)
    if teaching:
        add_teaching_slide(prs, teaching)

    prs.save(args.output)
    print(f"✓ 已保存: {args.output}")


if __name__ == "__main__":
    main()
