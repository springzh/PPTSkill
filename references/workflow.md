# 标准工作流详解

本文档说明三种模式的具体执行步骤，是 `SKILL.md` 的延伸。

## 通用前置：理解用户需求

不论走哪条路，先抽取下列信息：

| 维度 | 例子 |
|---|---|
| 主题/场景 | 工作总结、开题答辩、述职竞聘、教学课件、商务提案 |
| 主色偏好 | 红色（党政/爱国）、蓝色（商务）、橙色（活泼）、绿色（环保）、黑白（高级） |
| 风格关键词 | 简约、商务、卡通、复古、麦肯锡、扁平、插画、几何 |
| 受众 | 领导、同事、客户、学生、答辩老师 |
| 篇幅 | 5-10 页（短）、15-25 页（中）、30+ 页（长） |
| 数据图表需求 | 有没有真实数据需要可视化 |
| 时间紧迫度 | 是否需要先给预览图让用户确认 |

如果用户没说清楚，先用 1-2 个简短问题搞清楚，再决定模式。

## <a id="mode-a"></a>模式 A：从内置模板挑

模式 A 现在有两种路径：

- **路径 A1（传统单模板）**：用户需求明确匹配某一套模板，走传统单模板流程
- **路径 A2（跨模板混搭，推荐）**：用户需要多种版式类型（如"封面 + SWOT + 时间线 + KPI 卡片 + 结束页"），单模板无法满足时走多源混搭流程

---

### 路径 A1：传统单模板

#### A1.1. 读 INDEX

```bash
cat templates/INDEX.md
```

INDEX 把所有内置模板按风格分类列出，每个条目带一句 30 字内的特点描述。先按用户输入的"主色 + 风格 + 场景"在 INDEX 里大致缩小候选。

#### A1.2. 读候选模板的 intro.md

```bash
cat templates/<slug>/intro.md
```

intro.md 是高度浓缩简介，包含：
- 一句话定位
- 风格标签 + 排版复杂度
- 主题色（hex）
- 字体
- 适用场景（**故意写得宽泛**，不要被模板原标题局限）
- 包含的页面结构 + 版式类型
- 不要选用的页面

#### A1.3. 决定 / 跟用户确认

- **用户已指定模板** → 直接用。
- **你高度确信唯一最佳**（场景+风格+主色强匹配且明显胜出）→ 可直接用，但先一句话告知选了哪个 + 理由，留否决余地。
- **用户没指定，或你拿不准哪个最合适（默认情形）** → 让用户选：
  - 用 AskQuestion 给 **正好 3 个**候选，每个附一句话理由 + 展示其 `templates/<slug>/preview.png`。
  - 选项里**永远多放一个「都不满意，换一批」**；用户选它就再给**另外 3 个**没出现过的候选（带预览图），可反复换。
  - 候选用尽仍不满意 → 追问更具体偏好，或转模式 C 原创。
- ⚠️ 不要不看预览图就凭模糊匹配自作主张定模板。

#### A1.4. 读 detail.json，选页

```bash
python3 -c "import json; d=json.load(open('templates/<slug>/detail.json')); print(json.dumps([{'n':p['slide_number'],'role':p['role'],'use_for':p['use_for']} for p in d['pages']], ensure_ascii=False, indent=2))"
```

按用户内容规模选页：
- **封面 / 致谢页：按模板能力来** —— 读 `page_roles`：
  - `cover` 非空 → 加一张封面页；`cover` 为空 → **直接从第一张内容页开始**，不要把别的页当封面用
  - `ending` 非空 → 加一张致谢/结束页；`ending` 为空 → **直接以最后一张内容页收尾**，不要硬造"感谢聆听"
- **目录页**：`agenda` 非空时按需选；多章节才有意义；章节数要和你后面要选的章节扉页数一致
- **章节扉页 + 内容页**：`section_divider` 非空时按章节大纲挑；为空就直接列内容页
- **跳过的页面**：`skip_pages` 数组列出的 + 任何 `auto_promo_flag=true` 的页面

⚠️ 简单说："模板有什么角色就用什么角色"。不要为了凑齐"封面 / 目录 / 内容 / 致谢"四段式而从模板里强行挑出一张内容页冒充封面 —— 那会破坏视觉一致性。绝大多数 v1.0.2+ 的内置模板都**只有内容页**，请直接接受这种"纯内容"deck 形态。

把选定页面的原模板编号（1-based）按演示顺序放到 `selected_slides` 数组。

#### A1.5. 写 edits.json（传统格式）

完整 schema 见 [`pptx-edit-schema.md`](./pptx-edit-schema.md)。核心要点：

```json
{
  "template_slug": "<slug>",
  "selected_slides": [1, 2, 3, 5, 7, 9, 10, 12, 13, 14, 16],
  "edits": [
    {"slide": 1, "slot_id": "cover_title_cn", "new_text": "..."},
    {"slide": 2, "slot_id": "agenda_ch1_cn", "new_text": "..."},
    ...
  ]
}
```

强约束：
- **每个 editable=true 的 slot 都必须出现在 edits 里**（不留占位文字）
- **每条 new_text 的视觉宽度不能超过 slot 的 `max_chars`**（中文 1 字=1，英文/数字≈0.5）。参考同槽的 `chars_per_line` / `max_lines` 判断会不会换行出框。
- **太长就缩短文字，绝不改字号**：同一 `level` 的标题必须同字号（见 `type_scale`），靠精炼用词控制长度，不要把某处字号改小。
- **改了 agenda 的章节文字，必须同步改对应分章扉页 / 面包屑**
- editable=false 的 slot（"01"、"02"、"%" 之类）通常不出现在 edits 里
- 构建时务必带 `--detail`，让 `build_pptx.py` 能跑出框检测；建议用 `--strict` 让出框直接报错，便于一轮修好。

#### A1.6. 跑构建

```bash
python3 scripts/build_pptx.py \
    templates/<slug>/template.pptx \
    edits.json \
    out/<name>.pptx \
    --detail templates/<slug>/detail.json \
    --strict
```

`--strict` 会在 expected_text 不匹配或 slot 找不到时报错退出。默认会自动 sibling-lookup detail.json，但显式 `--detail` 更稳。

---

### 路径 A2：跨模板混搭（新架构）

当用户需要多种版式类型（如"封面 + SWOT 分析 + 时间线 + KPI 卡片 + 结束页"），单模板无法满足所有需求时，使用多源混搭流程。

#### A2.1. 理解内容需求 → 抽取版式类型

从用户描述中抽取所需的版式类型（layout_type）：

| 用户需求示例 | layout_type / tags |
|---|---|
| "封面" | `cover` |
| "SWOT 分析" | `swot`, `quadrant` |
| "时间线 / 里程碑" | `timeline` |
| "4 个关键指标" | `kpi`, `4-up`, `metrics` |
| "对比分析" | `comparison`, `vs`, `two-column` |
| "流程图" | `flow`, `process` |
| "鱼骨图" | `fishbone`, `root-cause` |
| "结束页" | `ending` |

#### A2.2. 查询 slide-registry.json

```bash
# 按标签查询候选幻灯片
python3 -c "
import json
reg = json.load(open('slide-registry.json'))

# 示例：查找 SWOT 相关幻灯片，限定 deep-blue 家族
for s in reg['slides']:
    if 'swot' in s['tags'] and s['style_family'].startswith('deep-blue'):
        print(f\"{s['slide_id']:45s} | {s['style_family']:35s} | {s['use_for']}\")
"
```

**查询策略**：
1. 优先在同一风格家族内查找（`style_family` 前缀匹配）—— 同家族内混搭无需颜色重映射
2. 如果同家族找不到，扩大到全部家族 —— 构建时会自动通过 `color_map` 进行颜色重映射
3. 可以用 `tags`、`layout_type`、`role`、`style_family` 任意组合筛选

**查看某张幻灯片的详细信息**：
```bash
python3 -c "
import json
reg = json.load(open('slide-registry.json'))
s = next(s for s in reg['slides'] if s['slide_id'] == 'models-swot-12')
# 读取对应的 detail.json 获取完整 slot 信息
detail = json.load(open(s['detail_ref']))
page = next(p for p in detail['pages'] if p['slide_number'] == s['source_slide'])
print(json.dumps(page, ensure_ascii=False, indent=2))
"
```

#### A2.3. 选择目标风格

从 `styles/` 目录选择目标风格。如果所有幻灯片来自同一风格家族，直接使用该家族的风格 slug。如果跨家族混搭，选择其中一个家族作为目标风格。

列出可用风格：
```bash
ls styles/
# deep-blue-minimal.json  deep-blue-business-massive.json  bordeaux-corp.json  ...
```

读风格定义：
```bash
python3 -c "import json; s=json.load(open('styles/deep-blue-minimal.json')); print(s['name'], '-', s['description'])"
```

#### A2.4. 选幻灯片 + 排序

从查询结果中挑选幻灯片，按演示逻辑排序（封面 → 目录 → 内容 → 结束），生成幻灯片列表：

```json
[
  {"slide_id": "minimal-cover-1",      "source_slide": 1},
  {"slide_id": "models-swot-12",       "source_slide": 12},
  {"slide_id": "minimal-4card-5",      "source_slide": 5},
  {"slide_id": "minimal-ending-16",    "source_slide": 16}
]
```

**选片原则**：
- 封面对封面、内容对内容 —— 保持角色匹配
- 优先同家族（无颜色重映射开销）
- 幻灯片数控制在 5-25 张（按用户篇幅需求）

#### A2.5. 写 edits.json（多源格式）

完整 schema 见 [`pptx-edit-schema.md`](./pptx-edit-schema.md)。多源格式核心要点：

```jsonc
{
  "style": "deep-blue-minimal",
  "slides": [
    {"slide_id": "minimal-cover-1",   "source_slide": 1},
    {"slide_id": "models-swot-12",    "source_slide": 12},
    {"slide_id": "minimal-4card-5",   "source_slide": 5},
    {"slide_id": "minimal-ending-16", "source_slide": 16}
  ],
  "edits": [
    {"slide_id": "minimal-cover-1",  "slot_id": "cover_title_cn", "new_text": "2026 战略复盘"},
    {"slide_id": "models-swot-12",   "slot_id": "swot_strength_title", "new_text": "技术壁垒"}
  ]
}
```

与单模板格式的关键区别：
- 用 `"style"` 替代 `"template_slug"`（用于格式检测）
- 用 `"slides"` 数组替代 `"selected_slides"`（每项含 `slide_id` + `source_slide`）
- edits 中用 `"slide_id"` 替代 `"slide"`（跨 source_pptx 的唯一标识）

强约束与单模板相同（覆盖所有 editable slot、遵守 max_chars、不改字号等）。

#### A2.6. 跑多源构建

```bash
python3 scripts/build_pptx.py \
    edits.json \
    out/<name>.pptx \
    --registry slide-registry.json \
    --styles styles/ \
    --style deep-blue-minimal \
    --strict
```

构建流水线：
1. **Resolve** — 通过 slide-registry 将每个 `slide_id` 映射到物理位置（source_pptx + source_slide）
2. **Assemble** — 从多个源 .pptx 文件中克隆幻灯片到单一输出演示文稿
3. **Edit** — 应用文字替换（与单模板模式相同的逻辑）
4. **Apply Style** — 跨家族时通过 `color_map` 将源颜色重映射到目标风格；同家族内跳过

`--strict` 在文字溢出、slot 找不到、或 expected_text 不匹配时报错退出。

#### A2.7. 跨家族混搭的颜色重映射

当幻灯片来自不同风格家族时，构建脚本会自动进行颜色重映射：

- **工作原理**：源风格的 `color_map` 定义"这个颜色是什么角色（primary/secondary/accent/…）"，目标风格定义"这个角色应该是什么颜色"
- **只重映射已映射的颜色**：源 `color_map` 中列出的颜色才参与映射，未列出的颜色（设计师刻意选择的装饰色）保留原样
- **字体也会重映射**：源 `font_map` → 目标 `fonts`
- **已知限制**：图片/图标中的嵌入颜色无法重映射；跨家族混搭图标较多的幻灯片可能有轻微视觉不一致

---

### A3. 自检（两条路径通用）

跑 `render_slides.py` 渲染最终 pptx，目测：
- 没有遗留占位文字
- 文字没溢出 / 换行错乱
- 章节名前后一致
- 顺序符合大纲
- （多源模式）跨家族幻灯片的颜色是否协调

发现问题 → 修改 edits.json → 再跑一次。

## 模式 B：用户带 PPT

详见 [`custom-template-workflow.md`](./custom-template-workflow.md)。

## 模式 C：完全原创

详见 [`original-design-guide.md`](./original-design-guide.md)。

## 失败兜底

| 现象 | 处理 |
|---|---|
| `expected_text mismatch` | 模板被改过；放弃 --strict 或重新核对 detail.json 的 current_text |
| `shape_id not found` | 模板被替换了；要么用其他模板，要么重新提取 detail.json |
| 文字溢出 | 看 `max_chars`；缩短文字或换更宽的版式 |
| 字体不一样 | 见 [`SKILL.md` 字体说明](../SKILL.md)，配置 fontconfig 别名 |
| 图表数据没改 | 见 [`chart-editing.md`](./chart-editing.md) |
