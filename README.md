# gao-science-ppt-skill

科学（物理/化学/生物等）课件 PPT 排版优化 AI Skill。输入课件 PPTX，自动完成：

- **排版统一**：字体中英分离统一（微软雅黑 + Calibri）、标题层级、分隔页/角标残留改写
- **学习目标/重点难点重写**：定位既有目标/重难点表格，按本课实际内容原位替换四维目标与重难点文案（保留原表格 run 级格式，不新增页）
- **配图升级**：封面/导入新课/目录门面图按课时内容替换——具象搜索词 + 易得性降维（膝盖→足球）+ 方向匹配裁剪（绝不拉伸）+ 跨课时 md5 去重；导入新课只留一张大卡片图
- **目录配色**：序号黄 FFC000、概览文字白 FFFFFF
- **全片 Morph 过渡**：所有页面 Morph 切换（旧版 PowerPoint 自动降级淡入淡出）
- **导入新课文案适配**：按本课时内容生成引导问题

驱动方法论：**第一性原理 + 贝叶斯原理 + JTBD 理论 + 奥卡姆剃刀原则**（详见 SKILL.md §1）。

## 架构（零脚本）

本 skill **不包含、不维护任何 .py 脚本文件**。所有 PPT 操作由 AI 在会话内用 python-pptx 内联代码直接完成；验证过的正确实现以「操作模式库」（SKILL.md §6 M1-M11）形式内置在文档中：诊断提取 / 字体统一 / 目录重建 / 表格重写 / 段落改写 / 配图裁剪替换（清 srcRect 防变形）/ 形状删除 / Morph 过渡 / Commons 选图下载 / 交付验证清单。

## 使用

对 AI 助手说：`优化这份课件 PPT` + 课件路径。

流程（SKILL.md 定义）：自动热更新 → 需求确认 → 内容诊断 → 方案确认 → 执行统一 → 学习目标/重点难点确认 → 交付 → 复盘自升级。仅 3 次必打断（目标/方案/教学设计确认），其余自动。每次调用后自动复盘，2+ 次复现的问题才升级 skill 规则。

## 安装

```bash
git clone git@github.com:Gao-thinking/gao-science-ppt-skill.git ~/.agents/skills/gao-science-ppt-skill
python3 -m pip install --user python-pptx Pillow
```

依赖仅 `python-pptx` 与 `Pillow`。
