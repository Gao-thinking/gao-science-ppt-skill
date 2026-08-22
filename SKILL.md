# gao-science-ppt-skill

科学/物理等理科课件 PPT 排版优化全流程：输入课件 PPTX → 四原理（第一性原理+贝叶斯+JTBD+奥卡姆）驱动 → 内容诊断 → 统一标题/图片/字体 → 学习目标/重点难点页按本课内容重写 → 封面/导入/目录配图按课时内容替换（具象壁纸级）+ 导入文案适配 + 目录黄序号白文字 + 全片 Morph 过渡 → 每次调用后复盘自升级。仅 3 次必打断（目标确认/方案确认/教学设计确认），使用前自动检查更新。

用户说"优化这份课件 PPT / 统一一下标题和字体 / 给这个课件补上学习目标重难点" → 执行以下流程。**只在 ⬜ 处弹窗**（用 `request_user_input`），其余全部自动完成。

**架构（2026-08 四次升级）：零脚本。所有 PPT 操作 = 会话内直接运行 python-pptx 内联代码，代码模式全部内置在 §6 模式库，不编写、不依赖、不维护任何 .py 脚本文件。**

## TL;DR（老手速查）

```
自动热更新(静默) → 需求确认[打断1] → 内容诊断(M1内联) → 方案确认[打断2]
→ 执行统一(M2-M9内联：字体/目录配色/残留改写/表格重写/配图裁剪替换/删小图/Morph)
→ 教学设计填充[打断3]（原位替换学习目标/重点难点页文字）→ 交付(不覆盖原件)
→ 复盘自升级(有升级项才弹窗)
```

## 前置准备

| 项 | 要求 | 缺失处理 |
|----|------|----------|
| python3 | ≥3.9 | `brew install python3` |
| python-pptx | 含 Pillow | `python3 -m pip install --user python-pptx Pillow` |
| lxml | 随 python-pptx 安装 | 同上 |
| 本 skill 仓库 | `~/.agents/skills/gao-science-ppt-skill` | `git clone` 后同步 |

交付产物：`{课件名}-优化.pptx`（**默认不覆盖原件**，同目录/output 输出副本）。

---

## §1 四原理总纲（每次决策必过）

| 原理 | 落在哪一步 | 怎么用 |
|------|-----------|--------|
| **第一性原理** | 全程，尤其诊断与标题改写 | 抛开"原 PPT 长什么样"的惯性，从**这页要让学生学到什么**倒推每个元素是否必要。排版的本质 = 降低认知负荷 |
| **贝叶斯** | 诊断、§7 复盘 | 先验 = 常见课件病；证据 = 实际数据（字体种类/比例分布/文字量）→ 分级：**必改 / 建议改 / 不动**。2+ 次复现才升级 skill |
| **JTBD** | 全流程+交互 | 老师 JTBD："打磨成能直接上课的成品，不丢内容、不破坏原件、不让我手动改"。只打断 3 次，每问给推荐项 |
| **奥卡姆** | 全程 | 能不改就不改；内容不增删——排版只动格式，教学设计=原位替换表格文字不动版式 |

### 1.1 排版统一铁律

1. **标题**：每页至多 1 个视觉标题，「动宾短语/概念名」式；层级：封面(40-44) > 节(32-36) > 页题(24-28) > 正文(18-20) > 注释(14-16)pt。
2. **图片**：配图一律"按素材缩放替换"——换 blob 不增 shape；新图按**原图框比例裁剪**（绝不拉伸压缩）；**替换时必须清空旧 srcRect 裁剪属性**（残留会二次裁切→变形）。原生未替换图的自带 srcRect 保留。
3. **字体**：中文微软雅黑/苹方，西文 Calibri/Helvetica；latin→ea→cs 按 schema 顺序写入（乱序会触发 PowerPoint repair）。
4. **目录概览配色**：序号黄 FFC000、概览文字显式白 FFFFFF。
5. **每课时配图差异化**：封面/导入新课/目录配图按**本课时内容**取图，跨课时互不相同且不与已交付课时重复（md5 核对）。
6. **全片 Morph 过渡**：所有页面切换用 Morph（p159:morph + fade fallback 兼容旧版）。
7. **角标一致（§1.1b）**：同一小节左上角标逐字一致；序号无空格、无上一节残留、章节序号连续（一/二/三/四）。

### 1.2 教学设计填充（替换学习目标/重点难点页文字，不新增页）

定位既有「学习目标」「重点难点」表（常为上一节物理残留），把文字**替换为本课实际内容推导的文案**：

- 学习目标表：科学观念/科学思维/探究实践/态度责任 四维（items 形式保「序号金+内容白」run 格式）
- 重点难点表：重点=核心知识点，难点=抽象推导/"易错"点
- 文案全部从课件既有内容推导，禁止编造；补充需标注「(补充建议)」并在打断3确认
- 兜底：课件无这两类页时才在末尾追加教学设计页

---

## §2 输入与输出

| 项 | 说明 |
|----|------|
| 输入 | 1-N 份 `.pptx`（`.ppt` 先转 `.pptx`） |
| 输出 | `{原名}-优化.pptx`（副本，默认 output/ 或同目录，不覆盖原件） |

---

## §3 工作流

**交互原则**：只打断 3 次（目标确认/方案确认/教学设计确认），异常才追加弹窗，选项都标 `(Recommended)`。

### Step 0: 自动热更新（静默）

```bash
cd ~/.agents/skills/gao-science-ppt-skill && git fetch origin -q 2>/dev/null || true
BEHIND=$(git rev-list --count HEAD..origin/main 2>/dev/null || echo 0)
[ "$BEHIND" -gt 0 ] && git pull --ff-only origin main -q && echo "[hot-update] 已更新 ${BEHIND} 提交"
# 有本地改动(git status非空)则跳过pull，交付总结注明
```

### Step 1: 需求确认 ⬜（第 1 次打断）

文件已给不追问路径；学段学科影响目标动词；偏好缺省=默认规范。
弹窗：「将优化 {文件}。使用默认规范（微软雅黑/Calibri，正文≥18pt，配图按课时内容替换，目录黄序号白文字，全片 Morph，学习目标/重点难点页按本课内容重写）？」选项：用默认规范直接开始(Recommended) / 调整偏好 / 只排版不动表格文字 / 直接覆盖原件。

### Step 2: 内容诊断（自动，M1 内联）

用 §6 M1 提取：每页标题候选、文本(字号/字体/粗体)、图片(尺寸/比例)、文字量。分级判定：
- 同页 ≥3 字体或同层字号差 ≥2 档 → 必改（字体统一）
- 单元格/段落含上一节课内容（如物理公式出现在生物课件）→ 必改（重写）
- 图片比例混乱但为知识图示 → 不裁剪不换（防破坏标注）；仅门面图（封面/导入/目录）参与替换
- 单页 >150 字 → 建议拆页提示，不自动拆

### Step 3: 方案确认 ⬜（第 2 次打断）

弹窗列改动清单：N 处必改（目录 X/分隔页 Y/角标 Z/表格 W 张/配图 G 张/Morph）+ M 项建议。选项：全部执行(Recommended) / 只必改 / 手动挑改 / 调整方案。

### Step 4: 执行统一（自动，M2-M9 内联）

顺序固定：分隔页/角标改写 → 删残留形状 → 配图替换(裁剪+srcRect清理) → 删多余小图 → 目录重建(带配色) → 表格重写 → 导入文案 → 字体统一(最后跑，覆盖新建 run) → Morph 全片 → 保存副本。

### Step 5: 教学设计填充 ⬜（第 3 次打断）

按 §1.2 从诊断推导四维目标+重难点文案（M4 写入表格），弹窗确认后随 Step 4 一并执行或单独补跑。

### Step 6: 交付总结（不弹框）

```
✓ 优化完成（未覆盖原件）
  输出：{原名}-优化.pptx
  目录重建：X 页    分隔页/角标修复：Y 处    目标/重难点重写：W 张
  配图替换：G 张（方向匹配/去重验证）   导入文案：已适配   Morph：n/n 页
  验证：srcRect清零 ✓ rPr顺序 ✓ zip完整 ✓ md5去重 ✓ 回归 ✓
```

---

## §3A 可复用改造要点

- **目录&节名适配**：封面读本课节名；核对目录页/节分隔页/正文左上角标三处残留（如「功和机械能」混入生物课），M3 整段改写；目录末尾保留「课堂总结/练习与应用」并顺延序号。
- **导入新课适配**：右侧文案框按本课时内容写 1-3 行引导问题（M6 多行填充）；左侧**只留一张大卡片图**，多余小横幅/小图用 M8 删除。
- **配图替换策略（选图三步）**：
  1. **易得性降维**：窄概念→素材丰富的典型情境（膝盖→足球踢球；条件反射→狗听指令；高级神经活动→下棋；血糖→方糖糖果；碘与甲状腺→海带）
  2. **搜索词必须具象可拍摄**，禁抽象名词直搜（"激素"→"注射胰岛素的手"）；候选过滤学术图表类（标题含 diagram/chart/graph/labeled/scheme/svg 一律跳过），只用照片级具象图
  3. **方向匹配**：竖框只选竖源图、横框选横源图（Commons 上限约 1.78 时可用 crop_anchor=top 保主体）；分辨率 ≥1400px；md5 与历史课时去重
- **动画排查**（可选）：遍历 slide XML `<p:timing>` 报告无动画/疑似重复页，只排查建议不擅自改。

---

## §6 python-pptx 直接操作模式库（核心，替代一切脚本）

> 使用方式：bash heredoc 内联运行（`python3 - <<'PYEOF' ... PYEOF`），不落地 .py 文件。以下模式均为验证过的正确实现，直接拷贝组合。

### M1 全量诊断提取

```python
from pptx import Presentation
def iter_shapes(shapes):
    for s in shapes:
        yield s
        if s.shape_type == 6: yield from iter_shapes(s.shapes)
prs = Presentation(path)
for i, slide in enumerate(prs.slides, 1):
    for sh in iter_shapes(slide.shapes):
        if getattr(sh, "has_table", False):
            for r_i,row in enumerate(sh.table.rows):
                for c_i,c in enumerate(row.cells): print(i, "表", r_i, c_i, c.text[:40])
        elif sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    sz = r.font.size.pt if r.font.size else None
                    print(i, sh.name, sz, r.font.name, r.text[:30])
```

### M2 字体统一（rPr 顺序安全，最后执行）

```python
from pptx.oxml.ns import qn
def set_run_fonts(run, latin="Calibri", ea="微软雅黑"):
    run.font.name = latin
    rPr = run._r.get_or_add_rPr()
    for tag in (qn("a:ea"), qn("a:cs")):
        el = rPr.find(tag)
        if el is not None: rPr.remove(el)
    ea_el = rPr.makeelement(qn("a:ea"), {"typeface": ea})
    cs_el = rPr.makeelement(qn("a:cs"), {"typeface": ea})
    latin_el = rPr.find(qn("a:latin"))
    if latin_el is not None:
        latin_el.addnext(cs_el); latin_el.addnext(ea_el)
    else:
        anchor = next((rPr.find(t) for t in map(qn,("a:sym","a:hlinkClick","a:hlinkMouseOver","a:rtl","a:extLst"))
                       if rPr.find(t) is not None), None)
        if anchor is not None: anchor.addprevious(cs_el); anchor.addprevious(ea_el)
        else: rPr.append(ea_el); rPr.append(cs_el)
# 遍历须含 GROUP 与表格：
for slide in prs.slides:
    for sh in iter_shapes(slide.shapes):
        targets = []
        if getattr(sh,"has_table",False):
            for row in sh.table.rows:
                for c in row.cells:
                    for p in c.text_frame.paragraphs: targets += p.runs
        elif sh.has_text_frame:
            for p in sh.text_frame.paragraphs: targets += p.runs
        for run in targets: set_run_fonts(run)
```

### M3 段落改写（节名/角标残留）

```python
def para_replace(slide, rules):  # rules=[{"contains":"功和机械能","text":"第1节   神经调节"}]
    for sh in iter_shapes(slide.shapes):
        if not getattr(sh,"has_text_frame",False): continue
        for para in sh.text_frame.paragraphs:
            full = "".join(r.text for r in para.runs)
            for rule in rules:
                if rule["contains"] in full:
                    if not para.runs: para.add_run()
                    para.runs[0].text = rule["text"]
                    for r in para.runs[1:]: r.text = ""
```

### M4 表格重写（text/items 两式，items 保「序号金+内容白」）

```python
def table_fill(slide, spec_rows):  # {(r,c): "纯文本" 或 [("1.","条目一"),("2.","条目二")]}
    for sh in iter_shapes(slide.shapes):
        if not getattr(sh,"has_table",False): continue
        tbl = sh.table
        for (r,c), val in spec_rows.items():
            tf = tbl.cell(r,c).text_frame
            if isinstance(val, str):
                if tf.paragraphs and tf.paragraphs[0].runs:
                    tf.paragraphs[0].runs[0].text = val
                    for para in tf.paragraphs:
                        for run in para.runs[1:]: run.text=""
                else: tf.text = val
            else:
                for pi,(num,txt) in enumerate(val):
                    para = tf.paragraphs[pi] if pi < len(tf.paragraphs) else tf.add_paragraph()
                    runs = para.runs or [para.add_run()]
                    runs[0].text = num
                    if len(runs)>1: runs[1].text = txt
                    for ex in runs[2:]: ex.text=""
                for ep in tf.paragraphs[len(val):]:
                    for run in ep.runs: run.text=""
```

### M5 目录重建（黄序号 FFC000 / 白文字 FFFFFF）

⚠ 必须在清空 runs **之前**记录每段「最后一个非空 run」位置，否则定位失效、内容误继承序号色。

```python
def rebuild_toc(slide, shape_name, items, num_color="FFC000", text_color="FFFFFF"):
    from pptx.dml.color import RGBColor
    shape = next(s for s in iter_shapes(slide.shapes) if s.name==shape_name)
    tf = shape.text_frame; paras = tf.paragraphs
    orig_last=[]
    for p in paras:
        ln=-1
        for j,r in enumerate(p.runs):
            if r.text.strip(): ln=j
        orig_last.append(ln)
    for p in paras:
        for r in p.runs: r.text=""
    for i,(num,text) in enumerate(items):
        if i >= len(paras):
            nr = tf.add_paragraph().add_run(); nr.text=f"{num} {text}".strip(); continue
        runs = paras[i].runs; last = orig_last[i]
        if runs and last >= 1:
            runs[0].text = num
            for j,r in enumerate(runs):
                if j==0: continue
                r.text = (" "+text) if j==last else ""
            runs[0].font.color.rgb = RGBColor.from_string(num_color)
            runs[last].font.color.rgb = RGBColor.from_string(text_color)
        elif runs:
            runs[0].text=f"{num} {text}".strip()
            for ex in runs[1:]: ex.text=""
```

### M6 文本框多行填充（导入新课文案）

```python
def box_write(slide, shape_name, text):  # text 支持 \n 多行
    shape = next(s for s in iter_shapes(slide.shapes) if s.name==shape_name)
    tf = shape.text_frame; lines = text.split("\n")
    p0 = tf.paragraphs[0]
    if p0.runs:
        p0.runs[0].text = lines[0]
        for r in p0.runs[1:]: r.text=""
    else: p0.add_run().text = lines[0]
    src = p0.runs[0] if p0.runs else None
    for ln in lines[1:]:
        np = tf.add_paragraph(); nr = np.add_run(); nr.text = ln
        if src is not None:
            try: nr.font.size = src.font.size
            except Exception: pass
            try: nr.font.bold = src.font.bold
            except Exception: pass
```

### M7 配图替换（裁剪防变形 + 清 srcRect）

```python
import io
from PIL import Image, ImageOps
def replace_image(pic, slide, path, anchor="center"):
    target = (pic.width/pic.height) if pic.height else 1.0
    im = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
    w,h = im.size; cur = w/h
    keep = min(target/cur, cur/target)          # 方向失配检查：<0.6 应换向
    if keep < 0.6: print(f"[warn] 方向失配仅保留{keep:.0%}，请换{'竖' if target<1 else '横'}版源图")
    if cur > target+0.01:
        nw=int(h*target); x0=(w-nw)//2; im=im.crop((x0,0,x0+nw,h))
    elif cur < target-0.01:
        nh=int(w/target); y0={"top":0,"bottom":h-nh}.get(anchor,(h-nh)//2)
        im=im.crop((0,y0,w,y0+nh))
    buf=io.BytesIO(); im.save(buf,format="JPEG",quality=88)
    open("/tmp/ppt-assets/_tmp_repl.jpg","wb").write(buf.getvalue())
    part,rId = slide.part.get_or_add_image_part("/tmp/ppt-assets/_tmp_repl.jpg")
    blip = pic._element.blipFill.blip; blip.set(qn("r:embed"), rId)
    bf = pic._element.blipFill                   # 关键：清残留 srcRect 防二次裁切变形
    for sr in bf.findall(qn("a:srcRect")): bf.remove(sr)
    try: pic.crop_left=pic.crop_right=pic.crop_top=pic.crop_bottom=0.0
    except Exception: pass
# pics = collect_pics(list(slide.shapes),[])  # 见 M8 的递归收集；索引按文档序
```

### M8 形状/图片删除（按名字，递归）

```python
def delete_shape(slide, name):
    for sh in iter_shapes(slide.shapes):
        if sh.name == name:
            sh._element.getparent().remove(sh._element); return True
    return False
def collect_pics(shapes, out):
    for s in shapes:
        if s.shape_type == 13: out.append(s)
        elif s.shape_type == 6: collect_pics(s.shapes, out)
    return out
```

### M9 Morph 过渡全片

```python
from lxml import etree
P_NS="http://schemas.openxmlformats.org/presentationml/2006/main"
MC_NS="http://schemas.openxmlformats.org/markup-compatibility/2006"
def morph_all(prs, dur=2000):
    for slide in prs.slides:
        sld = slide._element
        for el in sld.findall(f"{{{P_NS}}}transition"): sld.remove(el)
        for ac in sld.findall(f"{{{MC_NS}}}AlternateContent"):
            if ac.find(f".//{{{P_NS}}}transition") is not None: sld.remove(ac)
        xml=(f'<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:p="{P_NS}">'
             f'<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
             f'<p:transition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="slow" p14:dur="{dur}">'
             f'<p159:morph option="byObject"/></p:transition></mc:Choice>'
             f'<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>'
             f'</mc:AlternateContent>')
        el = etree.fromstring(xml)
        timing = sld.find(f"{{{P_NS}}}timing"); clr = sld.find(f"{{{P_NS}}}clrMapOvr")
        if timing is not None: timing.addprevious(el)      # CT_Slide 顺序: cSld→clrMapOvr→transition→timing
        elif clr is not None: clr.addnext(el)
        else: sld.find(f"{{{P_NS}}}cSld").addnext(el)
```

### M10 配图下载（Wikimedia Commons，具象词+方向匹配+图表过滤+去重）

```python
import json, urllib.request, urllib.parse, hashlib
def api(params):
    params.update({"format":"json"})
    q = urllib.parse.urlencode(params)
    req = urllib.request.Request(f"https://commons.wikimedia.org/w/api.php?{q}",
                                 headers={"User-Agent":"courseware-bot/1.0"})
    with urllib.request.urlopen(req, timeout=25) as r: return json.load(r)
def pick(term, out, target_ratio, used_md5=set(), must_words=(), ban_words=("diagram","labeled","scheme","chart","graph")):
    data = api({"action":"query","generator":"search","gsrsearch":f"filetype:bitmap {term}",
                "gsrnamespace":6,"gsrlimit":12,"prop":"imageinfo",
                "iiprop":"url|size|mime","iiurlwidth":1600})
    cands=[]
    for p in ((data.get("query") or {}).get("pages") or {}).values():
        t=p.get("title","").lower()
        if any(w in t for w in ban_words): continue          # 图表过滤
        if must_words and not any(w in t for w in must_words): continue  # 标题强校验防跑题
        for ii in p.get("imageinfo", []):
            if ii.get("mime") not in ("image/jpeg","image/png"): continue
            w,h = ii.get("width",0), ii.get("height",0); r=w/h
            if w<1300 or abs(r-target_ratio)/max(target_ratio,.01)>0.42: continue  # 方向/比例匹配
            cands.append((abs(r-target_ratio),w,h,p.get("title",""),ii.get("thumburl") or ii["url"]))
    for _,w,h,title,u in sorted(cands)[:4]:
        try:
            req=urllib.request.Request(u,headers={"User-Agent":"courseware-bot/1.0"})
            with urllib.request.urlopen(req,timeout=30) as r2: b=r2.read()
            md5=hashlib.md5(b).hexdigest()[:8]
            if len(b)<60000 or md5 in used_md5: continue     # 质量+跨课时去重
            open(out,"wb").write(b); print(f"✓ {out} <-「{title[:40]}」{w}x{h}")
            return hashlib.md5(b).hexdigest()[:8]
        except Exception: continue
    return None
```

### M11 交付验证清单（每次必跑）

```python
# ① zip 完整性: zipfile.ZipFile(out).testzip() is None
# ② python-pptx 全量遍历无异常
# ③ rPr 子元素顺序: 正则扫 slide XML, latin→ea→cs→sym 升序, 违规数==0
# ④ transition 位置: <mc:AlternateContent> 在 </p:cSld> 之后、<p:timing> 之前
# ⑤ 替换图 srcRect==0 且 blob 比==图框比; 原生未替换图的 srcRect 保留不动
# ⑥ Morph 覆盖 == 总页数
# ⑦ 目录 run 色: 序号 FFC000 / 文字 FFFFFF
# ⑧ 门面图 md5: 两课时互斥 + 与历史课时互斥
# ⑨ 回归: 分隔页/角标/表格无旧课残留, 导入文案在位, 残留胶囊已删
```

---

## §4 踩坑精华

- **srcRect 二次裁切变形（2026-08）**：替换图 blob 时不清旧 `a:srcRect`，旧裁剪参数叠加在新图上导致压缩变形 → 替换后必须移除 srcRect 并将 crop_* 归零；原生未替换图的 srcRect 是作者有意设计，保留。
- **toc 清空时序**：先 `r.text=""` 再找非空 run 会全部判空 → 内容写进单 run 继承序号金色。必须先记录 orig_last_ne 再清空。
- **rPr 乱序 repair**：ea/cs 直接 append 到尾部会破坏 schema 顺序 → 按 latin 后 addnext 插入；验证必解压查 XML 顺序，不能只看能否打开。
- **字体验证假阳性**：endParaRPr/defRPr/空 run 残留旧 typeface 属正常（不影响渲染），验证以「非空可见 run」分布为准。
- **GROUP/表格遗漏**：iter 遍历必须递归 GROUP 与表格单元格，否则节分隔页矩形、目标表格漏统一样式。
- **zipfile 重写丢条目**：手工重打包 pptx 时所有 zip 条目都要写回（未改动的用原 data），否则 KeyError rIdN；改完 unzip -t 校验。
- **知识图示不可裁剪**：生物学示意图/实验装置图带标注，禁止套 4:3 裁剪；只有门面图（封面/导入/目录）参与替换。
- **多课时模板复用**：同节多课时共用模板，封面/导入/目录/目标表格全是上一课时残留 → 按各课时内容分别适配，配图 md5 必须互斥。
- **Commons 无超宽原图**：横幅位(≈2.84)在该图源常缺真宽幅 → 优先找全景/背景类题材（体育场夜景、图案平铺），次选 crop_anchor=top 保人物主体。

---

## §5 元规则

- **零脚本架构**：本仓库不含也不维护 .py 脚本；所有操作用 §6 模式库内联执行。升级=改本文档的模式代码。
- 不覆盖老师原件；不提交老师的课件到 git；升级只改 SKILL.md 自身。
- 提交邮箱：`293544754+Gao-thinking@users.noreply.github.com`。

---

## §7 复盘与自升级（每次调用完成后必做）

1. **收集**：弹窗非推荐项、返工步骤、脚本报错、口头纠正、环境变化
2. **四原理过滤**：贴近降认知负荷？真实复现 2+ 次（或有机制根因）？更少弹窗等待？最小改动？
3. **有升级项才弹窗**（全部升级(Recommended)/部分/暂不）→ commit + push，独立 commit 便于 revert
4. 运行中的问题先修本次输出，规则升级由复盘决定
