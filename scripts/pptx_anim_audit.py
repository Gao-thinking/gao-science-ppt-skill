#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gao-science-ppt-skill: PPTX 动画排查脚本（四类可复用改造之一）
逐页分析幻灯片 XML 中的动画/切换设置：
  - 有没有动画（含 <p:timing>/<p> 元素）
  - 动画数量、类型分布
  - 是否为"无动画"页 / "疑似重复动画"页
  - 页面切换（transition）有无
输出 JSON 报告 + 可读摘要。只读不改，供 skill 交互动画处置。

用法:
  python3 pptx_anim_audit.py -i deck.pptx -o /tmp/ppt-diag/ [--cwd 输出JSON名]
"""
import argparse
import json
import os
import re
import sys
import zipfile

try:
    from pptx import Presentation
    from pptx.oxml.ns import nsmap
except ImportError:
    sys.exit("缺少 python-pptx，请先: python3 -m pip install --user python-pptx")


# 常见动画类型关键词（p:cTn 的子元素 tag 后缀，用于归类）
ANIM_TYPES = {
    "p:anim": "属性/位移",
    "p:animClr": "颜色",
    "p:animEffect": "效果(淡出等)",
    "p:animScale": "缩放",
    "p:set": "瞬间设置",
    "p:par": "组",
}


def audit(path):
    """逐页统计动画/切换。返回 (slides_report, summary_lines)。"""
    slides = []
    total_has = 0
    with zipfile.ZipFile(path) as z:
        # 收集 slideN.xml 的序号与文件名
        slides_idx = []
        for n in z.namelist():
            m = re.match(r"ppt/slides/slide(\d+)\.xml$", n)
            if m:
                slides_idx.append((int(m.group(1)), n))
        slides_idx.sort()
        for idx, name in slides_idx:
            xml = z.read(name).decode("utf-8", "ignore")
            has_timing = "<p:timing>" in xml
            has_transition = "<p:transition>" in xml
            anim_tags = {}
            for tag in ANIM_TYPES:
                cnt = xml.count("<" + tag)
                if cnt:
                    anim_tags[tag] = cnt
            # 粗略统计动画节点数
            n_anim = xml.count("<p:par>") + xml.count("<p:cTn>")
            # 疑似重复：同一页动画节点数明显高于该页图片/文本数量（启发式，仅供参考）
            n_objs = xml.count("<p:sp>") + xml.count("<p:pic>") + xml.count("<p:graphicFrame>")
            suspected_dup = (n_anim > 0 and n_anim > max(1, n_objs * 2))
            flag = "无动画" if not n_anim else ("疑似重复动画" if suspected_dup else "正常")
            has_anim = n_anim > 0
            total_has += 1 if has_anim else 0
            slides.append({
                "index": idx,
                "slide_file": name,
                "has_timing": has_timing,
                "has_transition": has_transition,
                "anim_types": anim_tags,
                "n_anim_nodes": n_anim,
                "n_objects": n_objs,
                "suspected_duplicate": suspected_dup,
                "flag": flag,
            })
    # 摘要
    no_anim = [s["index"] for s in slides if not s["has_transition"] and not s["has_timing"]]
    dup = [s["index"] for s in slides if s["suspected_duplicate"]]
    lines = [
        f"共 {len(slides)} 页；有动画/切换 {total_has} 页",
        f"无动画页: {no_anim or '无'}",
        f"疑似重复动画页: {dup or '无'}",
    ]
    return slides, lines


def xml_tag(xml, tag):
    # 纯文本标签计数（含命名空间容错）
    return xml.count("<" + tag) + xml.count("<" + tag.replace("p:name", ""))


def main():
    ap = argparse.ArgumentParser(description="PPTX 动画排查")
    ap.add_argument("-i", "--input", required=True, help="输入 .pptx")
    ap.add_argument("-o", "--outdir", default="/tmp/ppt-diag/", help="输出目录")
    args = ap.parse_args()
    if not args.input.endswith(".pptx"):
        sys.exit("仅支持 .pptx")
    os.makedirs(args.outdir, exist_ok=True)
    base = os.path.splitext(os.path.basename(args.input))[0]
    slides, lines = audit(args.input)
    report = {"source": args.input, "pages": len(slides), "slides": slides}
    json_path = os.path.join(args.outdir, f"{base}-动画排查.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=1)
    print("\n".join(lines))
    print(f"\n→ {json_path}")


if __name__ == "__main__":
    main()