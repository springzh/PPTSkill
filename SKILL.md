---
name: ppt-skill
description: >-
  用 19 套内置中文 PPT 模板（或用户自带的 .pptx 模板）生成与编辑 PowerPoint 演示文稿。支持跨模板混搭幻灯片（slide-level composability），通过 slide-registry 按标签/版式类型检索 640+ 张独立幻灯片，搭配可分离的 styles 色彩/字体系统进行跨风格家族的颜色重映射。只替换文字、不破坏原排版/配色/字号，内置按文本框尺寸的出框检测与同级标题字号一致校验；也支持完全原创的简洁版式。当用户要"做 / 生成 / 制作 / 编辑一份 PPT / 演示文稿 / 幻灯片 / .pptx"，或需要工作汇报、年终与季度总结、述职竞聘、项目复盘、开题答辩、商务提案、教学课件、数据可视化等成品 PPTX 时使用。Use when the user wants to create or edit a PowerPoint / PPT / slides / .pptx deck, pick from built-in templates, mix slides across templates, or apply their own template without breaking the layout.
---

# ppt-skill

> **版本** 见同目录 [`VERSION`](./VERSION)。

---

## 🚨 启用本 Skill 后做的第一件事：跑一次自动更新

任何 AI 在本会话**第一次用到这个 Skill 时**，**必须**先把下面这条命令作为它的第一个 tool call：

```bash
python3 scripts/apply_update.py
```

这条命令会自动：
1. 从远端 `update_source`（git+ 仓库）拉取最新 `updates.json`
2. 比对本地 `VERSION` 与远端 `latest_version`
3. 如果有新版本 → 只下载真正变动的文件（增量更新，无 LFS 流量浪费）
4. 如果已是最新 → 几秒内退出，输出 `OK   Nothing to do`

⚠️ 本 Skill **不会自动 push 更新**，必须由 AI 主动 pull。如果跳过这一步，可能在用过期版本，行为、规则、模板都可能与文档不符。

✅ 同一会话内只需做一次，后续 tool call 不必重复检查。

---

> **⚠️ 非商业使用**：本 Skill 与其内置 PPT 模板**仅供个人学习和研究**，**严禁用于任何商业用途**（含商业演示、销售、培训分发、客户提案、企业内部以营利为目的的使用等）。模板素材来自第三方设计师作品，二次商用需要获取原作者授权。

## 何时使用本 Skill

当用户的需求满足下列任何一项就调用本 Skill：

- 需要"做一份 PPT / 演示文稿 / 幻灯片"，无论是工作汇报、年终总结、季度复盘、项目提案、述职竞聘、教学课件、开题答辩、读书报告……
- 用户给了文字大纲或一段描述，希望"做成 PPT"
- 用户拿来一份 .pptx 文件，希望"按这个模板做一份新的"
- 用户希望"调整 / 编辑 PPT 的某些文字"，且要求"不破坏排版"
- 用户希望对比多个 PPT 模板，选一个合适的

## 更新机制说明

完整的更新工具有两个脚本：

| 脚本 | 干嘛 | 何时用 |
|---|---|---|
| `scripts/apply_update.py` | **检查 + 应用**（一步到位） | **每次启用 Skill 第一件事**（见顶部红框） |
| `scripts/check_update.py` | 仅检查、不应用，列出会变的文件 | 当你只想预览改了什么、不想立即更新时 |

如果只想先看变化再决定要不要升级：

```bash
python3 scripts/check_update.py     # 列出 added / modified / removed
```

`updates.json` 的 `update_source` 已配置为 `git+https://github.com/GordenSun/GordenPPTSkill.git#main`，开箱即用，无需修改。

## 架构概述：Slide-Level Templates

本 Skill 采用**可组合幻灯片架构**，核心概念：

```
slide-registry.json  →  640+ 张独立幻灯片的统一索引（按标签/版式/风格家族检索）
styles/<slug>.json   →  可分离的色彩 + 字体定义（含 color_map 用于跨家族重映射）
templates/<slug>/    →  19 套原始模板（不变），每套含 template.pptx + detail.json + intro.md
```

**关键能力**：你可以从不同模板中挑选幻灯片混搭成一份演示文稿。例如：从 `report-massive-models` 取 SWOT 分析页 + 从 `minimal-business-summary` 取 4 卡片对比页 + 从 `report-savior` 取鱼骨图页，统一应用 `deep-blue-minimal` 风格。

### 风格家族

| 风格家族 | 成员模板 | 幻灯片数 | 状态 |
|---|---|---|---|
| `deep-blue-business-massive` | report-massive-models, report-massive-charts, report-massive-reports | 113 | ✅ 可用 |
| `deep-blue-minimal` | minimal-business-summary | 16 | ✅ 可用 |
| `deep-blue-corp-extended` | data-viz-deck, report-savior, operations-deck, competition-speech, architecture-deck | 233 | ✅ 可用 |
| `bordeaux-corp` | premium-corp, mckinsey-style | 72 | ✅ 可用 |
| `patriotic-red` | red-patriot-youth, red-patriot-general | 41 | 🔜 即将支持 |
| `warm-academic` | cute-orange-class, thesis-novice, thesis-formula | 88 | 🔜 即将支持 |
| `wine-academic` | top-thesis | 39 | 🔜 即将支持 |
| `standalone` | quarterly-illust, geometric-summary | 40 | 🔜 即将支持 |

同一风格家族内的幻灯片配色统一，混搭无需颜色重映射。跨家族混搭时，构建脚本会自动通过 `color_map` 将源幻灯片的颜色重映射到目标风格。

## 三种模式

收到用户需求后，先判断走哪种模式：

### 模式 A：从内置模板里挑（默认）

**默认走这条路。** 19 套内置模板覆盖了绝大多数中文场景。现在支持两种路径：

#### 路径 A1：单模板（传统，仍然可用）

适合用户需求明确匹配某一套模板。流程与 v1.x 相同：

1. 读 [`templates/INDEX.md`](./templates/INDEX.md) 筛选候选
2. 读候选模板的 `intro.md` + `detail.json`
3. 选页、写传统 `edits.json`（`template_slug` + `selected_slides`）
4. 跑单模板构建命令

#### 路径 A2：跨模板混搭（推荐，新架构）

适合用户需要多种版式类型（如"封面 + SWOT + 时间线 + KPI 卡片 + 结束页"），单模板无法满足。

1. **理解内容需求** → 抽取用户需要的版式类型（cover / swot / timeline / kpi / ending …）
2. **查询 slide-registry.json** → 按 `tags` + `layout_type` + `style_family` 检索候选幻灯片：
   ```bash
   python3 -c "
   import json
   reg = json.load(open('slide-registry.json'))
   # 按标签筛选：cover + deep-blue 家族
   covers = [s for s in reg['slides'] if 'cover' in s['tags'] and s['style_family'].startswith('deep-blue')]
   for s in covers:
       print(f\"{s['slide_id']:40s} | {s['style_family']:30s} | {s['use_for']}\")
   "
   ```
3. **选幻灯片** → 优先在同一风格家族内挑选（无需颜色重映射）；如需跨家族，构建时会自动重映射颜色
4. **选择目标风格** → 从 `styles/` 目录选一个风格 slug，或让用户从风格家族中挑选
5. **写多源 `edits.json`** → 使用新格式（见下方）
6. **跑多源构建命令**

#### 模板选择决策规则（两条路径通用）

- **用户已明确指定模板** → 直接用。
- **你高度确信只有 1 个模板最合适**（场景 + 风格 + 主色都强匹配，且明显优于其它）→ 可直接用，但开工前一句话告诉用户你选了哪个、为什么，给用户一个否决的机会。
- **其余所有情况（用户没指定，或你不能完全把握哪个最合适）→ 必须让用户来选**：
  - 用 AskQuestion 提供 **正好 3 个**候选模板，每个附上一句话理由（风格 / 适用场景 / 页数），并**把对应的 `templates/<slug>/preview.png` 一并展示**给用户看图决策。
  - 选项里**始终额外带一个「都不满意，换一批」**。用户选它时，再按匹配度给出**另外 3 个**没出现过的候选（同样附预览图）。可反复换，直到用户选定或候选用尽。
  - 候选都用尽仍不满意 → 询问用户更具体的偏好（风格 / 颜色 / 场景），或转模式 C 原创。
  - ⚠️ 不要在没让用户看预览图的情况下，仅凭模糊匹配就自作主张定一个模板。

### 模式 B：用户自己带 PPT 模板

当用户提供了 .pptx 文件且明确希望以它作模板时：

1. 把用户的 pptx 当作"未知模板"
2. 用 `scripts/render_slides.py` 把每页渲染成 PNG，再用 `python-pptx` 现场探查每页的 shape / paragraph / run 结构（无需额外脚本）
3. 自己看每页（PNG + shape 输出）：
   - 推断每页是什么角色（封面 / 目录 / 章节扉页 / 内容页 / 结束 / 模板宣传）
   - 推断每页适合放什么内容
   - 跳过模板宣传 / "稻壳儿" / 感谢下载 之类
4. 按 [模式 B 工作流](./references/custom-template-workflow.md) 用 explicit `address` 写 `edits.json` 选页 + 文字替换
5. **不要修改用户模板原文件**；所有改动写到新的 output.pptx

### 模式 C：完全原创（不基于任何模板）

当用户明确要求"原创设计 / 不用模板 / 简单干净的样式"时：

1. **创建尽量简洁的版式** —— 大量留白、对齐严格、装饰极少
2. 单页元素 ≤ 4 个，避免复杂图形 / 图标群组
3. 字体：英文 Arial / Helvetica；中文 微软雅黑 / 思源黑体；标题加粗即可
4. 主色 1 个 + 灰阶 + 白底；不要拼凑多种风格
5. 用 `python-pptx` 直接代码生成，参考 [`references/original-design-guide.md`](./references/original-design-guide.md)

## edits.json 格式

### 传统单模板格式（仍支持）

```jsonc
{
  "template_slug": "minimal-business-summary",
  "selected_slides": [1, 2, 5, 9, 16],
  "edits": [
    {"slide": 1, "slot_id": "cover_title_cn", "new_text": "2026 年度复盘"}
  ]
}
```

### 新多源格式（跨模板混搭）

```jsonc
{
  "style": "deep-blue-minimal",
  "slides": [
    {"slide_id": "minimal-cover-1",      "source_slide": 1},
    {"slide_id": "models-swot-12",       "source_slide": 12},
    {"slide_id": "minimal-4card-5",      "source_slide": 5},
    {"slide_id": "minimal-wheel-9",      "source_slide": 9},
    {"slide_id": "minimal-ending-16",    "source_slide": 16}
  ],
  "edits": [
    {"slide_id": "minimal-cover-1",  "slot_id": "cover_title_cn", "new_text": "2026 战略复盘"},
    {"slide_id": "models-swot-12",   "slot_id": "swot_strength_title", "new_text": "技术壁垒"}
  ]
}
```

**格式检测**：`"template_slug"` 存在 → 传统路径；`"style"` 存在 → 多源路径。

## 构建命令

### 传统单模板构建

```bash
python3 scripts/build_pptx.py \
    templates/<slug>/template.pptx \
    edits.json \
    out/final.pptx \
    --detail templates/<slug>/detail.json \
    --strict
```

### 新多源构建

```bash
python3 scripts/build_pptx.py \
    edits.json \
    out/final.pptx \
    --registry slide-registry.json \
    --styles styles/ \
    --style deep-blue-minimal \
    --strict
```

构建流水线（多源模式）：**Resolve**（slide_id → 物理位置）→ **Assemble**（从多源克隆幻灯片）→ **Edit**（文字替换）→ **Apply Style**（跨家族颜色/字体重映射）

### 自检

```bash
# 渲染最终 pptx 给用户预览 / 自检（每页一张 PNG）
python3 scripts/render_slides.py out/final.pptx out/renders --dpi 144
```

## 编辑铁律（所有模式通用）

1. **不改排版** —— 只改文字。形状的位置、大小、颜色、字体、字号、行距，都不动。
2. **所有占位文字必须替换** —— 模板里的 "Question 1" / "Vivamus..." / "Key Words Here" / "项目名称" 等占位文本都必须用真实内容替换。一份完成的 PPT 里不应出现任何示例占位词。
3. **字数受 `max_chars` 约束（按文本框真实尺寸精算，含 20% 余量）** —— detail.json 每个 slot 的 `max_chars` 是用"文本框宽高 + 字号"算出的容量（单位：视觉宽度，中文 1 字=1、英文/数字≈0.5）。还提供 `chars_per_line`（每行可容字数）和 `max_lines`（可容行数）。**写文字前先看这三个值**，确保不超过 `max_chars`。`capacity_unknown:true` 的槽（自适应/组合内异形框）测不准，谨慎短写。
   - 构建时 `build_pptx.py` 会**自动检测出框**：超容量会告警；加 `--strict` 时直接拒绝保存，必须缩短后重试。出框时**只缩短文字、不要改字号**（改字号会破坏同级一致，见第 8 条）。
4. **数字 / 序号默认不动** —— `editable: false` 的 slot 是装饰性 "01/02/1/2/%" 之类，除非用户明确要求改顺序，否则保持。
5. **图形 / 图表通常无法同步** —— 装饰性的进度条 / 圆环 / 旗帜路径 / 流程箭头是固定形状；改了百分比文字不会改弧长。每个模板 detail.json 里如有这类页面会在 `cautions` 字段列出。
6. **真实数据图表才能改数据** —— 如果某页有 PPT 原生 chart（`shape.has_chart=True`），可以用 `build_pptx.py --chart-data` 同步更新；详见 [`references/chart-editing.md`](./references/chart-editing.md)。
7. **章节名前后呼应** —— 改了目录章节名，对应的分章扉页 + 内容页面包屑文字都要同步改。
8. **封面 / 致谢页按模板能力来，不要硬造** ——
   - 读 `detail.json` 的 `page_roles`，如果 `cover` 数组为空，**直接从第一张内容页开始**，不要拿一张内容页当封面用，更不要从其他模板临时凑一张封面页。
   - 如果 `ending` 数组为空，**直接以最后一张内容页收尾**，不要硬造"感谢聆听"。
   - 同理，`agenda` 空 → 不强加目录；`section_divider` 空 → 不强加分章扉页。
   - 也就是说：**模板有什么角色就用什么角色**，少一个角色就少一页，不要破坏视觉一致性去拼凑。这条规则在 v1.0.3 起对所有模板生效。
   - **例外（多源模式）**：跨模板混搭时，可以从不同模板分别取 cover、content、ending 页——这正是新架构的核心价值。但仍需确保所有幻灯片属于同一风格家族，或让构建脚本进行颜色重映射。
9. **同级标题字号必须一致，靠"控制长度"而非"缩字号"达成** ——
   - detail.json 顶部有 `type_scale`（字号层级表，level 1 = 最大），每个 slot 标了 `level`。**同一 level 的文字必须保持模板原字号，不要改字号。**
   - 当某处文字太长放不下时，**永远是缩短文字**（换更精炼的表达），**绝不是把这一处的字号改小** —— 那会让本该同级的标题大小不一，非常难看。
   - 选多页拼一份 PPT 时，确认各页同 level 的标题用词长度相近，整体才齐整。

## 目录结构

```
GordenPPTSkill/
├── SKILL.md                    ← 本文件
├── VERSION                     ← 当前版本号
├── CHANGELOG.md                ← 人类可读变更日志
├── updates.json                ← 机器可读版本增量索引
├── manifest.json               ← 所有文件的 sha256 与版本归属
├── README.md                   ← 仓库概览（用户阅读）
├── slide-registry.json         ← 640+ 张幻灯片的统一索引（按标签/版式/风格家族检索）
├── styles/                     ← 可分离风格定义（色彩 + 字体 + color_map）
│   ├── deep-blue-minimal.json
│   ├── deep-blue-business-massive.json
│   ├── deep-blue-corp-extended.json
│   ├── bordeaux-corp.json
│   ├── patriotic-red.json
│   ├── warm-academic.json
│   └── wine-academic.json
├── scripts/
│   ├── build_pptx.py           ← 按 edits.json 选页 + 换字 → 输出 pptx（含出框检测 + 多源组装 + 风格重映射）
│   ├── build_registry.py       ← 扫描 templates/*/detail.json 生成 slide-registry.json
│   ├── render_slides.py        ← pptx → PDF → 每页 PNG（预览/自检）
│   ├── compute_capacity.py     ← 由 template.pptx 计算每个 slot 的容量字段（数据准备）
│   ├── check_update.py         ← 检查远端是否有更新
│   ├── apply_update.py         ← 增量更新本地文件
│   └── build_manifest.py       ← 重建 manifest.json
├── references/
│   ├── workflow.md
│   ├── pptx-edit-schema.md
│   ├── custom-template-workflow.md
│   ├── chart-editing.md
│   └── original-design-guide.md
└── templates/                  ← 19 套原始模板（不变）
    ├── INDEX.md
    └── <slug>/
        ├── template.pptx
        ├── intro.md            ← 高度浓缩简介
        ├── detail.json         ← 详细页面 / slot 数据（含容量字段 + type_scale）
        └── preview.png         ← 4 页 2×2 拼接预览图
```

## 关键脚本一句话说明

| 脚本 | 干嘛用 |
|---|---|
| `build_pptx.py` | 按 `edits.json`（选页 + 文字替换）从模板生成最终 pptx；支持单模板和多源两种模式；带出框检测 + 跨家族风格重映射，`--strict` 时出框拒绝保存 |
| `build_registry.py` | 扫描所有 `templates/*/detail.json` 生成 `slide-registry.json`；新增或修改模板后运行 |
| `render_slides.py` | 把任意 pptx 渲染成每页一张 PNG（用 LibreOffice + pdftoppm），用于预览 / 自检 |
| `compute_capacity.py` | 由 template.pptx 算出每个 slot 的 `chars_per_line/max_lines/max_chars` 等容量字段（自带模板已算好，仅在加新模板时需要） |
| `check_update.py` | 对比本地 VERSION 和远端 updates.json，告诉你要不要更新 |
| `apply_update.py` | 按 updates.json 的 delta 列表只下载变动文件 |
| `build_manifest.py` | 重新计算 manifest.json |

## 字体说明

模板 XML 里大量使用 `微软雅黑`。如果运行环境没有该字体，配合 `~/.config/fontconfig/fonts.conf` 把它别名到本地已安装的字体（推荐顺序：WenQuanYi Micro Hei → DengXian → Noto Sans SC → PingFang SC）。预览图正是用这条 fallback 链渲染的。

最终交给用户的 .pptx 在 PowerPoint / WPS / Keynote 里打开时会自动用宿主机的字体渲染，因此不需要担心字体丢失问题。

## 一些常见误区

- **不要**只改文字不改章节名一致性 —— 改 agenda 必须同步改分章扉页
- **不要**在装饰图上加文字以为是真图表
- **不要**把 lorem ipsum 留在最终 PPT 里
- **不要**为了塞下文字而忽略 `max_chars`
- **不要**修改用户原始模板文件 —— 所有产出都到新文件
- **不要**用本 Skill 做商业项目 —— 见顶部声明
- **不要**在跨家族混搭时忽略风格重映射 —— 如果源幻灯片和目标风格色差大，先确认 `color_map` 覆盖了所有关键颜色
