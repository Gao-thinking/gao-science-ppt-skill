# gao-science-ppt-skill

科学（物理/化学/生物等）课件 PPT 排版优化 AI Skill。输入课件 PPTX，自动完成：

- **统一标题**：按「动宾短语 / 概念名」式重写页标题，建立封面 > 节标题 > 页标题的层级
- **统一图片**：同页等宽/等高对齐、统一比例（默认 4:3 中心裁剪）、统一浅灰细边框
- **统一字体**：中文/西文分离统一（默认微软雅黑 + Calibri），正文字号统一（默认 20pt）
- **填充教学设计**：依据课件内容推导教学目标、重难点、教学环节，写入课件末尾新增页

另外提供**四类可复用改造**（§3A，逐项与用户交互确认后执行，方便下次编辑别的课件 PPT）：
1. **目录 & 各节标题/左上角节名适配**——按封面节名重排目录，正文页角标与节分隔页残留旧节名一并改写；「课堂总结/练习与应用/提升训练」保留在目录末尾并顺延序号
2. **导入新课右侧引导框填充**——按全课核心内容生成课程引导文案
3. **配图替换**——封面（第1页多图）/第2页（导入页）/第5页（目录页）按页内容换高清图（Pexels 直链优先，被反爬时自动换 Wikimedia Commons 高清图源）
4. **动画排查**——逐页输出「无动画页 / 疑似重复动画页」，供课堂播放优化

驱动方法论：**第一性原理 + 贝叶斯原理 + JTBD 理论 + 奥卡姆剃刀原则**（详见 SKILL.md §1）。

## 使用

对 AI 助手说：`优化这份课件 PPT / 统一标题和字体 / 补上教学设计和重难点` + 课件路径。

流程（SKILL.md 定义）：自动热更新 → 需求确认 → 内容诊断 → 方案确认 → 执行统一 → 教学设计确认 → 交付 → 复盘自升级。仅 3 次必打断（目标/方案/教学设计确认），其余自动。每次调用后自动复盘，2+ 次复现的问题才升级 skill 规则。

## 安装

```bash
git clone git@github.com:Gao-thinking/gao-science-ppt-skill.git ~/.agents/skills/gao-science-ppt-skill
python3 -m pip install --user python-pptx
```

每次调用开始时 skill 会自动 `git pull` 检查更新（本地有未提交改动时跳过并提示）。

## 目录结构

```
gao-science-ppt-skill/
├── SKILL.md                  # 主规则（工作流/四原理/四类改造/复盘自升级）
├── README.md
└── scripts/
    ├── pptx_analyze.py       # 诊断：标题/字体/字号/图片/文字量 → JSON
    ├── pptx_anim_audit.py    # 动画排查：逐页 无动画/疑似重复/有动画 → JSON
    ├── pptx_dump_shapes.py   # 辅助：逐页打印 shape 名/位置/文本（toc/box 按名定位用）
    └── pptx_apply.py         # 应用：标题/字体/图片统一 + 四类改造 + 教学设计页
```

## 脚本

```bash
# 诊断（输出 JSON + 可读摘要）
python3 scripts/pptx_analyze.py -i deck.pptx -o /tmp/ppt-diag/

# 动画排查（输出 无动画/重复 逐页报告）
python3 scripts/pptx_anim_audit.py -i deck.pptx -o /tmp/ppt-diag/

# 需要按形状名定位时（如 toc/box 指定 shape），先查页内形状名：
python3 scripts/pptx_dump_shapes.py -i deck.pptx --pages 1,2,5

# 应用方案（方案 JSON 由 AI 按 SKILL.md §1.1 / §3A 生成；示例见 examples/plan.example.json）
python3 scripts/pptx_apply.py -i deck.pptx -o deck-优化.pptx --plan plan.json
```

默认输出 `{原名}-优化.pptx` 副本，**不覆盖原件**。

## 原理速览

| 原理 | 落地 |
|------|------|
| 第一性原理 | 从「这页要让学生学到什么」倒推每个元素是否必要，排版本质 = 降低认知负荷 |
| 贝叶斯 | 用脚本统计（字体种类/字号档位/图片比例分布/超载页）作为证据分级：必改/建议改/不动；复盘单次偶发不升级 |
| JTBD | 老师雇佣它「把课件打磨成能直接上课的成品」；每页回答学生三问；只打断 3 次 |
| 奥卡姆 | 能不改就不改；一个改动服务多个指标；只改排版不改文字事实 |

## 维护

- 仓库：`Gao-thinking/gao-science-ppt-skill`（GitHub 上按需设 public/private；如需私有，在 GitHub → Settings → General → Danger Zone 里改为 Private 即可，不改代码）
- 升级：每次调用后复盘 → 弹窗确认 → 独立 commit + push（SKILL.md §5）
- 提交邮箱：`293544754+Gao-thinking@users.noreply.github.com`

## License

MIT
