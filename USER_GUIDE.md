# Gorden PPT Skill — 用户指南

> 用 AI 做原生专业 PPT 的终极指南。读完本文你会掌握：如何用一句话生成整份 PPT、如何跨模板混搭 640+ 张幻灯片、如何在 5 分钟内做出让领导满意的演示文稿。

---

## 目录

1. [30 秒快速开始](#30-秒快速开始)
2. [核心概念：模板、幻灯片、插槽、风格](#核心概念模板幻灯片插槽风格)
3. [跨模板混搭精讲（v2 新架构）](#跨模板混搭精讲v2-新架构)
   - [示例 1：年度战略复盘](#示例-1年度战略复盘)
   - [示例 2：产品发布会](#示例-2产品发布会)
   - [示例 3：述职竞聘答辩](#示例-3述职竞聘答辩)
   - [示例 4：开题答辩](#示例-4开题答辩)
   - [示例 5：季度运营复盘](#示例-5季度运营复盘)
4. [单模板模式（快速参考）](#单模板模式快速参考)
5. [自建模板模式（快速参考）](#自建模板模式快速参考)
6. [提示词速查表](#提示词速查表)
7. [常见问题](#常见问题)

---

## 30 秒快速开始

**第一步：把这句话发给你的 AI 助手（Claude / DeepSeek / ChatGPT 均可）：**

> 安装这个 Skill：https://github.com/springzh/PPTSkill
> 然后使用这个 Skill 做一个复杂、豪华的 PPT，读取本地 XXX 文件，来介绍 XXX 项目

**第二步：等着 AI 把 .pptx 文件交给你。**

就这么简单。下面是原理和进阶玩法。

---

## 核心概念：模板、幻灯片、插槽、风格

理解这 4 个概念，你就理解了整个系统的运作方式。

### 模板（Template）

模板是一套完整的 PPT 设计，包含配色、字体、版式。本 Skill 内置 **19 套中文模板**，覆盖深蓝商务、党政红、学术、麦肯锡风等主流风格。每套模板提供 `.pptx` 文件 + 预览图 + 插槽地图。

### 幻灯片（Slide）

每一张幻灯片（比如封面页、SWOT 分析页、KPI 数据页）就是一个**独立的拼图块**。19 套模板合计提供 **640+ 张**这样的拼图块，每张都编入了统一的 `slide-registry.json` 索引，可以用关键词检索。

### 插槽（Slot）

插槽是幻灯片上的**可替换文字位置**。比如封面页上有一个写标题的文本框，它在 `detail.json` 里被标记为 `cover_title_cn`，同时记录了：

- 文本框能装多少字（`max_chars`）—— 超出会检测报错
- 字号层级（`level`）—— 保证同级标题字号一致
- 当前占位文字（`current_text`）—— 用于验证

AI 的工作就是：往每个插槽里填上你的真实内容，不改字号、不改颜色、不改位置。

### 风格（Style）

风格是**从模板中分离出来的配色 + 字体定义**。比如 `deep-blue-minimal` 风格定义了主色 `#485275`、字体用微软雅黑 + Arial。

**这是 v2 架构的关键创新**：当你从不同模板混搭幻灯片时，构建脚本会自动将源幻灯片的颜色重映射到目标风格。同一风格家族内混搭无需重映射；跨家族混搭会自动转换颜色。

### 一张图看懂

```
你的需求："帮我做一份年度战略复盘PPT"
                │
                ▼
    AI 分析需要哪些版式：封面 → SWOT → 时间线 → KPI卡片 → 鱼骨图 → 结束页
                │
                ▼
    查 slide-registry.json → 从 3 套模板中挑出 6 张幻灯片
                │
                ▼
    写 edits.json → 填上你的真实内容
                │
                ▼
    build_pptx.py 四阶段流水线：
      Resolve → Assemble → Edit → Apply Style
                │
                ▼
    输出：统一风格的 output.pptx ✓
```

---

## 跨模板混搭精讲（v2 新架构）

这是本 Skill 最强大的功能：单套模板不可能覆盖所有版式类型（封面的极简设计 + SWOT 的四象限 + 鱼骨图的问题分析 + 时间线的里程碑展示……），但你可以从不同模板里各自挑出最合适的幻灯片，拼成一份风格统一的 PPT。

**同一风格家族内混搭**：颜色天然统一，不需要任何转换。

**跨风格家族混搭**：构建脚本通过 `color_map` 自动将颜色重映射到目标风格。比如把波尔多红模板里的 `#A52524` 自动映射到深蓝风格的 `#485275`。

> ⚠️ **非商业使用声明**：本 Skill 及内置模板仅供个人学习与研究，禁止任何商业用途。

---

### 示例 1：年度战略复盘

**场景**：你是某互联网公司的 PM，要给 CEO 做一份"2026 年度战略复盘"PPT。你需要：封面、SWOT 分析、时间线里程碑、KPI 卡片、鱼骨图问题诊断、结束页。

**发给 AI 的提示词**：

```
安装这个 Skill：https://github.com/springzh/PPTSkill

帮我做一份"微澜科技 2026 年度战略复盘"PPT，要求：

1. 封面：极简风格，中英文标题，深蓝配色。中文标题"微澜科技 2026 年度战略复盘"，英文"Strategic Review 2026"
2. SWOT 分析：从 S（优势）、W（劣势）、O（机会）、T（威胁）四个维度分析。S：AI 算法壁垒深厚，团队专利 47 项；W：国际化交付效率偏低，平均周期 45 天；O：大模型带来的企业级工具链需求爆发；T：竞品以 60% 价格打价格战
3. 时间线：4 个关键里程碑。Q1 完成 Agent 框架重构；Q2 首批 30 家企业客户迁移上线；Q3 日均调用量突破 1000 万次；Q4 完成 B 轮融资并启动海外拓展
4. KPI 数据卡片：4 项核心指标。营收 2.8 亿（同比增长 134%）；客户留存率 92%；NPS 净值 68；人效提升 2.1 倍
5. 鱼骨图：分析"交付延期"的根因。从人员、流程、工具、需求四个维度展开
6. 结束页："谢谢！欢迎提问"

整份 PPT 保持统一的深蓝商务风格，信息密度高，排版豪华。
```

**AI 做了什么（幕后）**：

| 版式需求 | 来源模板 | slides_id | 风格家族 |
|---------|----------|-----------|---------|
| 封面 | `minimal-business-summary` | `minimal-cover-1` | deep-blue-minimal |
| SWOT 分析 | `report-massive-models` | `report-massive-models-swot-13` | deep-blue-business-massive |
| 时间线 | `report-savior` | `report-savior-gantt-19` | deep-blue-corp-extended |
| KPI 卡片 | `minimal-business-summary` | `minimal-4card-5` | deep-blue-minimal |
| 鱼骨图 | `report-massive-models` | `report-massive-models-fishbone-31` | deep-blue-business-massive |
| 结束页 | `minimal-business-summary` | `minimal-ending-16` | deep-blue-minimal |

6 张幻灯片来自 **3 套不同模板**。全部属于 deep-blue 色系家族，构建时无需颜色重映射。

**实际填写的关键 slot**：

```
封面:
  cover_title_cn  → "微澜科技 2026 年度战略复盘"
  cover_title_en  → "Strategic Review 2026"

SWOT:
  s13_sh10_p0r0   → "优势（Strength）"
  s13_sh11_p0r0   → "AI 算法壁垒深厚，团队持有核心专利 47 项，关键技术指标领先竞品 18 个月"
  s13_sh12_p0r0   → "劣势（Weakness）"
  s13_sh13_p0r0   → "国际化交付效率偏低，平均交付周期 45 天，海外客户满意度仅 72%"
  （机会 O / 威胁 T 对应的 slot 继续填写）

时间线（甘特图改造为里程碑时间轴）:
  s19_sh52_p0r0   → "Q1: Agent 框架重构完成"
  s19_sh51_p0r0   → "Q2: 首批 30 家企业客户迁移上线"
  s19_sh53_p0r0   → "Q3: 日均调用量突破 1000 万次"
  s19_sh60_p0r0   → "Q4: B 轮融资完成，海外拓展启动"

KPI 卡片:
  p5_card1_title  → "营收 2.8 亿"
  p5_card1_body   → "同比增长 134%，超额完成年初目标的 127%"
  p5_card2_title  → "客户留存率 92%"
  p5_card2_body   → "NRR 连续 6 个季度保持在 90% 以上"
  p5_card3_title  → "NPS 净值 68"
  p5_card3_body   → "较去年同期提升 12 个百分点，行业前 5%"
  p5_card4_title  → "人效提升 2.1x"
  p5_card4_body   → "AI 辅助开发使人均产出翻倍，研发成本下降 37%"

鱼骨图:
  s31_sh5_p0r0    → "交付延期根因分析"
  s31_sh43_p0r0   → "需求变更频繁（人员）"
  s31_sh44_p0r0   → "需求确认流程冗长，平均 3 轮评审才定稿"
  s31_sh47_p0r0   → "自动化测试覆盖不足（流程）"
  s31_sh48_p0r0   → "回归测试仍需 12 人天，CI 流水线缺少冒烟测试"
  s31_sh50_p0r0   → "多系统切换耗时（工具）"
  （继续填写更多根因分支）

结束页:
  end_cn          → "谢谢！欢迎提问"
  end_en          → "Thank You"
```

**构建命令**：

```bash
python3 scripts/build_pptx.py \
    edits.json \
    out/微澜科技_2026战略复盘.pptx \
    --registry slide-registry.json \
    --styles styles/ \
    --style deep-blue-minimal \
    --strict
```

> 💡 **提示**：这个示例中的幻灯片全部属于 deep-blue 色系（minimal / business-massive / corp-extended），配色高度一致，最终输出浑然一体，看不出是从 3 套模板拼出来的。

---

### 示例 2：产品发布会

**场景**：你是 SaaS 创业公司的产品负责人，要在年度发布会上介绍全新产品"星图 AI 分析引擎"。你需要：吸睛封面、产品架构图、核心数据看板、竞品对比矩阵、结束页。

**发给 AI 的提示词**：

```
安装 Skill：https://github.com/springzh/PPTSkill

帮我做一份产品发布会 PPT："星图 AI 分析引擎"。要求：

1. 封面：高级感，酱红色调。主标题"星图 AI 分析引擎"，副标题"重新定义企业数据分析"，日期 2026.06
2. 产品架构图：展示四层架构——数据接入层（MySQL/ClickHouse/Kafka/API）、计算引擎层（流批一体/向量检索/语义理解）、分析模型层（归因分析/预测建模/异常检测）、应用层（管理驾驶舱/智能预警/自动报告）
3. 数据看板：4 个 KPI——日均处理数据量 12TB、查询响应 < 800ms、客户数 217 家、ARR ¥4,600 万
4. 竞品对比矩阵：从部署方式（SaaS/私有化/混合云）、核心能力（自然语言查询/自动洞察/多源联邦）、定价模式（按量/按席/混合）三个维度对比国际竞品（Tableau / Looker / ThoughtSpot），突出星图的优势
5. 结束页："即刻预约演示 → starchart.io/trial"

风格用波尔多酱红商务风（premium-corp 方向），高级感、留白充足。
```

**AI 做了什么（幕后）**：

| 版式需求 | 来源模板 | slides_id | 风格家族 |
|---------|----------|-----------|---------|
| 封面 | — | — | —（premium-corp 无 cover，需从 minimal 借用，构建时 color_map 重映射） |
| 产品架构 | `architecture-deck` | `architecture-deck-arch-3` | deep-blue-corp-extended → **重映射到 bordeaux-corp** |
| 数据看板 | `operations-deck` | `operations-deck-dashboard-19` | deep-blue-corp-extended → **重映射到 bordeaux-corp** |
| 竞品对比 | `mckinsey-style` | `mckinsey-compare-24` | bordeaux-corp ✓ 同族 |
| 结束页 | — | — | —（同上，借用 minimal-ending 重映射） |

这里的亮点是**跨风格家族混搭**：`architecture-deck` 和 `operations-deck` 是深蓝色系的，但发布会要求酱红色调（bordeaux-corp）。构建时 `color_map` 自动将源幻灯片中的 `#1F3A93`（深蓝）重映射为 `#A52524`（酱红），`#2C3E70` 重映射为对应的辅助色。

**构建命令**：

```bash
python3 scripts/build_pptx.py \
    edits.json \
    out/星图AI_产品发布会.pptx \
    --registry slide-registry.json \
    --styles styles/ \
    --style bordeaux-corp \
    --strict
```

> 💡 **提示**：`--style bordeaux-corp` 告诉构建脚本目标风格是波尔多酱红。所有来自 deep-blue 家族的幻灯片会自动完成颜色重映射。图标内的颜色无法重映射，但文字和形状颜色都能完美转换。

---

### 示例 3：述职竞聘答辩

**场景**：你是某大厂的运营总监，要参加晋升答辩，汇报过去一年的工作成果和方法论。你需要：个人成果展示、PDCA 复盘、KPI 看板、KISS 复盘、项目时间线。

**发给 AI 的提示词**：

```
安装这个 Skill：https://github.com/springzh/PPTSkill

帮我做一份"运营总监晋升述职"PPT。内容包括：

1. 工作成果展示（5 项）：地图驱动技能提升项目（学习地图覆盖 12 个岗位序列，培训完成率 91%）；
   薪酬健康度诊断（发现 3 个关键岗位薪酬低于市场 25 分位，已调整）；人才梯队搭建（内部晋升率从 18% 提升到 34%）；
   绩效体系重构（OKR + KPI 双轨制落地，目标对齐率从 56% 提升到 89%）；
   文化价值观落地（员工 eNPS 从 23 提升到 47）

2. PDCA 复盘：以"绩效体系重构"为例。Plan（设计 OKR+KPI 双轨方案，选 3 个部门试点）、
   Do（Q1 培训 + 系统上线，Q2 全公司推广）、Check（目标对齐率 56% → 89%，员工满意度 +22%）、
   Act（固化季度 Review 机制，纳入新员工入职培训）

3. KPI 数据看板：团队规模 47 人、年度预算执行率 94%、管理幅度 8 个直接下属、360 评分 4.7/5.0

4. KISS 复盘：Keep（双周 1on1 机制、数据驱动决策习惯）、Improve（跨部门协调效率、向上汇报频次）、
   Start（建立行业人脉网络、每季度输出 1 份竞品分析）、Stop（过度 micromanagement、频繁临时会议）

5. 项目时间线：2025.06 入职 → 2025.09 完成诊断 → 2025.12 体系重构上线 → 2026.03 全面推广 → 2026.06 述职答辩

整体用深蓝商务风格，信息密度高，适合给 VP 级别的评委看。
```

**AI 做了什么（幕后）**：

| 版式需求 | 来源模板 | slides_id | 风格家族 |
|---------|----------|-----------|---------|
| 工作成果 5 列 | `competition-speech` | `competition-speech-results-1` | deep-blue-corp-extended ✓ |
| PDCA 复盘 | `report-massive-models` | `report-massive-models-pdca-34` | deep-blue-business-massive |
| KPI 看板 | `operations-deck` | `operations-deck-dashboard-19` | deep-blue-corp-extended ✓ |
| KISS 复盘 | `competition-speech` | `competition-speech-kiss-24` | deep-blue-corp-extended ✓ |
| 项目时间线 | `report-savior` | `report-savior-gantt-19` | deep-blue-corp-extended ✓ |

5 张幻灯片来自 **4 套模板**，全部属于 deep-blue 色系，颜色天然统一。`competition-speech` 提供了述职场景的专业版式（成果展示、KISS 复盘），`report-massive-models` 提供了 PDCA 模板，`operations-deck` 提供了数据看板。

**关键 slot 填写示例（PDCA 页）**：

```
s34_sh5_p0r0     → "PDCA 复盘"
s34_sh5_p0r3     → "绩效体系重构专项"

P（Plan 计划层）:
  s34_sh11_p0r0  → "设计 OKR+KPI 双轨方案"
  s34_sh12_p0r0  → "选取产研、运营、市场 3 个部门试点，设计分岗位序列的考核模板"
  s34_sh13_p0r0  → "制定 Q1-Q2 分阶段推广计划，明确各阶段里程碑和责任人"

D（Do 执行层）:
  s34_sh19_p0r0  → "Q1 完成全员培训 + 系统上线（飞书+自研绩效系统打通）"
  s34_sh20_p0r0  → "Q2 全公司 12 个部门推广，双周 Office Hour 答疑"

C（Check 检查层）:
  s34_sh27_p0r0  → "目标对齐率 56% → 89%（+33pp）"
  s34_sh28_p0r0  → "员工对绩效公平感满意度 +22%，管理者反馈会议时长减少 35%"

A（Act 改进层）:
  s34_sh35_p0r0  → "固化季度 OKR Review + 半年度校准会机制"
  s34_sh36_p0r0  → "纳入新员工入职培训 Day 1，作为管理基本功必修课"
```

> 💡 **提示**：`competition-speech` 有 59 页，是述职/竞聘场景的素材库。虽然没有独立的 cover/ending 页，但内容页非常专业。如果需要封面，从 `minimal-business-summary` 借用一张即可。

---

### 示例 4：开题答辩

**场景**：你是计算机专业研二学生，要做硕士开题答辩。课题："基于大语言模型的代码缺陷检测方法研究"。需要学术风格的 PPT。

**发给 AI 的提示词**：

```
安装 Skill：https://github.com/springzh/PPTSkill

帮我做一份硕士开题答辩 PPT。课题："基于大语言模型的代码缺陷检测方法研究"。

内容规划：
1. 封面：课题名称 + 学生姓名"张明远" + 导师"李志鸿教授" + 学校"清北大学计算机学院"
2. 研究背景与意义（2 页）：
   - 第 1 页：软件缺陷造成的经济损失（NIST 报告：美国每年损失 $2.08 万亿）+ 传统静态分析工具的局限性（误报率 > 70%）+ LLM 在代码理解上的突破
   - 第 2 页：现有方法综述——基于规则的（Coverity/SonarQube）→ 基于机器学习的（DefectPrediction）→ 基于预训练模型的（CodeBERT/GraphCodeBERT）→ LLM-based（GPT-4/Claude），指出空白：缺乏面向 LLM 的缺陷检测系统性框架
3. 研究方法（2 页）：
   - 第 1 页：技术路线总图。代码仓库 → AST+CFG+DFG 多粒度代码表示 → 基于检索增强（RAG）的缺陷上下文注入 → LLM 多维度缺陷检测（功能/安全/性能/可维护性）→ 人工复核闭环
   - 第 2 页：关键技术细节。RAG 检索策略（混合检索：BM25 + 代码嵌入向量）、Prompt 设计（Chain-of-Thought + Few-shot + Structured Output）、评估基准构建（从 GitHub/BigVul/Defects4J 构建 15,000 样本测试集）
4. 预期成果与进度安排：
   - 预期成果：① 一个面向代码缺陷检测的 RAG+LLM 框架（开源）② 15,000 样本的多维度评测基准 ③ 在 5 个主流 LLM 上的系统对比实验 ④ 投稿 ICSE 2027
   - 甘特图：文献调研（2026.06-08）→ 框架设计与实现（2026.09-12）→ 实验与评估（2027.01-03）→ 论文撰写（2027.04-06）

风格用学术风，酒红色调（top-thesis 方向）。
```

**AI 做了什么（幕后）**：

| 版式需求 | 来源模板 | slides_id | 风格家族 |
|---------|----------|-----------|---------|
| 封面 | `top-thesis` | — | wine-academic（需要查 detail.json 确认） |
| 研究背景（多段内容） | `thesis-novice` | 背景/意义方法论模板 | warm-academic → **重映射到 wine-academic** |
| 文献综述（流程/对比） | `thesis-formula` | 现状综述模板 | warm-academic → **重映射到 wine-academic** |
| 技术路线（架构图） | `architecture-deck` | 多层架构图 | deep-blue-corp-extended → **重映射到 wine-academic** |
| 甘特图 | `report-massive-models` | `report-massive-models-gantt-27` | deep-blue-business-massive → **重映射到 wine-academic** |
| 结束页 | `top-thesis` | — | wine-academic ✓ |

这里涉及**跨 3 个风格家族的混搭**：warm-academic（墨绿）+ wine-academic（酒红）+ deep-blue（深蓝架构图）。构建时全部重映射到目标风格 `wine-academic`（酒红 `#7A2B22`）。

> 💡 **提示**：学术类模板（`thesis-novice` / `thesis-formula` / `top-thesis`）各有特色：`thesis-novice` 有 32 套各专业研究方法的填好范例，`thesis-formula` 提供 4 段开题公式（背景→意义→现状→方法），`top-thesis` 是酒红色名校风格。三套混搭可以覆盖开题答辩的完整流程。

---

### 示例 5：季度运营复盘

**场景**：你是某电商公司的运营负责人，要汇报 Q2 运营数据和 Q3 规划。需要：数据看板首页、流量漏斗分析、PEST 宏观环境、渠道对比矩阵、Q3 行动计划。

**发给 AI 的提示词**：

```
安装这个 Skill：https://github.com/springzh/PPTSkill

帮我做一份"2026 Q2 运营复盘 & Q3 规划"PPT。

内容规划：
1. 封面：极简商务风，标题"2026 Q2 运营复盘"，副标题"数据驱动 · 精益增长"
2. Q2 核心数据看板：GMV ¥8,760 万（环比 +23%）、订单量 142 万单（环比 +18%）、客单价 ¥61.7（环比 +4%）、退货率 5.3%（环比 -1.2pp）、营销 ROI 1:4.2（环比 +0.6）
3. 流量漏斗分析：各渠道流量 → 商品详情页到达 → 加购 → 下单 → 支付完成，转化率对比 Q1 和 Q2
4. PEST 宏观环境分析：政策（跨境电商税收新政 2026.09 实施）、经济（消费降级趋势持续，性价比需求上升）、社会（短视频+直播电商渗透率达 58%）、技术（AI 推荐算法使转化率提升约 15%）
5. 4 个渠道对比：抖音电商（GMV 占比 38%，获客成本 ¥23）、天猫（GMV 占比 29%，获客成本 ¥41）、拼多多（GMV 占比 21%，获客成本 ¥15）、微信私域（GMV 占比 12%，获客成本 ¥8）
6. Q3 行动计划甘特图：7 月（跨境新平台接入开发 + 达人矩阵扩到 200 人）、8 月（AI 选品系统上线 + 私域会员体系改版）、9 月（双 11 预热 + 全渠道价格协同系统上线）

全部用深蓝商务风格，数据页要信息密度高。
```

**AI 做了什么（幕后）**：

| 版式需求 | 来源模板 | slides_id | 风格家族 |
|---------|----------|-----------|---------|
| 封面 | `minimal-business-summary` | `minimal-cover-1` | deep-blue-minimal ✓ |
| 数据看板 | `operations-deck` | `operations-deck-dashboard-19` | deep-blue-corp-extended |
| 流量漏斗 | `report-massive-charts` | funnel 分析页 | deep-blue-business-massive |
| PEST 分析 | `report-savior` | PEST 宏观分析页 | deep-blue-corp-extended |
| 渠道对比 | `operations-deck` | 多渠道路对比页 | deep-blue-corp-extended |
| 甘特图 | `report-massive-models` | `report-massive-models-gantt-27` | deep-blue-business-massive |
| 结束页 | `minimal-business-summary` | `minimal-ending-16` | deep-blue-minimal ✓ |

7 张幻灯片来自 **4 套模板**，全部属于 deep-blue 色系，无需颜色重映射。

> 💡 **提示**：`operations-deck`（52 页）是运营场景的"百宝箱"——从 KPI 看板（A-N 共 14 种变体）到私域运营、产品生命周期、客户画像，几乎覆盖了运营汇报的所有版式需求。配合 `report-massive-charts` 的漏斗/树状图和 `report-savior` 的 PEST/SWOT，可以搭出非常专业的运营复盘 PPT。

---

## 单模板模式（快速参考）

如果某套模板恰好匹配你的需求（比如做一份简约商务总结，`minimal-business-summary` 16 页全覆盖），直接用传统单模板模式更简单：

**发给 AI 的提示词**：

```
安装 Skill：https://github.com/springzh/PPTSkill

用内置的"简约商务总结汇报"模板，帮我做一份 Q2 工作汇报 PPT。
内容包括：完成了 3 个项目、关键指标提升了 20%、下季度计划推进 2 个新方向。
大概 12 页，深蓝白配色。
```

AI 会自动：
1. 从 `templates/minimal-business-summary/` 读模板
2. 按 `detail.json` 的 `page_roles` 选页（封面→目录→章节扉页→内容页→结束页）
3. 把真实内容填入每个 slot
4. 用 `build_pptx.py` 输出最终 pptx

---

## 自建模板模式（快速参考）

如果你有自己的公司 PPT 模板（.pptx），可以直接拿来用：

**发给 AI 的提示词**：

```
安装 Skill：https://github.com/springzh/PPTSkill

用我自己的 PPT 模板（/path/to/公司模板.pptx）做一份项目汇报。
内容是：[你的内容描述]。不要改原模板文件，输出一份新的 pptx。
```

AI 会自动：
1. 把你的模板渲染成预览图
2. 探查每页的 shape 结构
3. 推断封面/目录/章节/内容/结束页
4. 按地址（shape_id + paragraph + run）替换文字
5. **绝不修改原模板文件**

---

## 提示词速查表

以下是可以直接复制使用的提示词模板。把 `[...]` 替换成你的实际内容即可。

### 基础场景

| 场景 | 提示词（精简版） |
|------|----------------|
| 年度总结 | 用 GordenPPTSkill 做一份 [年份] 年度工作总结 PPT，内容包括 [3-5 个要点]，大约 [N] 页，用深蓝商务风格 |
| 季度复盘 | 用 GordenPPTSkill 做一份 [季度] 复盘 PPT，包含数据看板 + SWOT 分析 + 下季度计划，信息密度高 |
| 述职竞聘 | 用 GordenPPTSkill 做一份述职竞聘 PPT，职位 [XX]，包含个人成果 + PDCA 复盘 + KPI 看板 |
| 产品发布 | 用 GordenPPTSkill 做一份产品发布会 PPT，产品 [XX]，包含架构图 + 数据看板 + 竞品对比 + 里程碑 |
| 开题答辩 | 用 GordenPPTSkill 做一份开题答辩 PPT，课题 [XX]，包含研究背景 + 技术路线 + 进度安排，学术风格 |
| 项目复盘 | 用 GordenPPTSkill 做一份项目复盘 PPT，包含 KISS 复盘 + 鱼骨图 + 甘特图 + 经验教训 |
| 商务提案 | 用 GordenPPTSkill 做一份商务提案 PPT，方案 [XX]，麦肯锡风格，包含问题诊断 + 方案框架 + 实施路径 |

### 跨模板混搭专属提示词（v2）

以下提示词会触发跨模板混搭路径，效果更豪华：

| 场景 | 跨模板混搭提示词 |
|------|----------------|
| 战略汇报 | 用 GordenPPTSkill 做一份 [主题] 战略汇报 PPT。需要封面、SWOT、时间线、KPI 卡片、鱼骨图分析、结束页。从不同模板混搭最佳版式，统一深蓝风格 |
| 运营复盘 | 用 GordenPPTSkill 做一份运营复盘 PPT。需要数据看板、流量漏斗、渠道对比、PEST 分析、甘特图。运营数据页用 operations-deck 的版式，分析页用 report-savior，统筹混搭 |
| 学术答辩 | 用 GordenPPTSkill 做一份学术答辩 PPT。研究方法用 thesis-novice 版式，架构图用 architecture-deck，甘特图用 report-massive-models，统一学术风格 |
| 竞聘答辩 | 用 GordenPPTSkill 做一份竞聘答辩 PPT。成果展示用 competition-speech 的 5 列卡片，PDCA 用 report-massive-models，KPI 看板用 operations-deck，KISS 复盘用 competition-speech，封面和结束页用 minimal-business-summary。全部深蓝商务 |

### 风格指定关键词

在提示词里加入这些关键词，AI 会自动匹配对应模板：

| 关键词 | 匹配方向 |
|--------|---------|
| 深蓝商务 / 大厂风 | `deep-blue-minimal` 或 `deep-blue-corp-extended` |
| 信息密集 / 排版豪华 | `deep-blue-business-massive`（report-massive 系列） |
| 酱红高级 / 咨询风 | `bordeaux-corp`（premium-corp / mckinsey-style） |
| 党政红 / 党课 | `patriotic-red`（red-patriot-youth / red-patriot-general） |
| 学术 / 论文答辩 | `warm-academic` 或 `wine-academic` |
| 可爱 / 少儿培训 | `cute-orange-class` |
| 互联网 / Y2K | `quarterly-illust` |
| 麦肯锡风 | `mckinsey-style` |

---

## 常见问题

### Q1：支持哪些 AI 模型？

实测支持 **DeepSeek**、**Claude**、**GPT**、**小米 Mimo**，国产模型也能完成得很好。核心要求是模型能读 JSON、写 JSON，能执行 bash 命令。

### Q2：生成一份 PPT 要多久？

取决于页数和内容复杂度。一份 10-15 页的专业 PPT，AI 通常在 **3-8 分钟内**完成——包括模板匹配、内容填写、构建和渲染预览。

### Q3：生成的效果能直接使用吗？

可以。输出的 .pptx 文件在 PowerPoint / WPS / Keynote / LibreOffice 中都能正常打开和编辑。文字、配色、排版均已完成，你可以在此基础上手动微调。

### Q4：编码问题怎么处理？

模板使用 `微软雅黑` 字体。如果你在 Mac 上打开，系统会自动 fallback 到 PingFang SC 或 Noto Sans SC。如需在 Mac 上完美预览，参考 README 的字体配置章节。

### Q5：跨模板混搭会不会"四不像"？

不会。同一风格家族内的幻灯片配色统一，混搭天然和谐。跨家族混搭时，构建脚本会自动进行颜色重映射（`color_map`），将源幻灯片的配色转换为目标风格。唯一需要注意的是图标和图片内的颜色无法重映射——所以同类风格的混搭效果最好。

### Q6：单模板 vs 跨模板混搭，什么时候用哪个？

| 场景 | 推荐模式 |
|------|---------|
| 需求简单，某套模板恰好覆盖全部版式 | 单模板（路径 A1） |
| 需要多种特殊版式（SWOT + 鱼骨 + KPI 卡片 + 时间线） | 跨模板混搭（路径 A2） |
| 做述职/竞聘（有专门的 competition-speech 模板） | 单模板为主，缺封面/结束页时跨模板补 |
| 做产品发布（需要架构图 + 数据看板 + 竞品矩阵） | 跨模板混搭 |
| 学术答辩（研究方法 + 技术路线 + 甘特图） | 跨模板混搭 |

### Q7：可以在公司/团队内部使用吗？

⚠️ **非商业使用**：本仓库及内置模板**仅供个人学习与研究**，禁止任何形式的商业用途（包括但不限于商业演示、销售材料、培训分发、客户提案、企业内部以营利为目的的使用等）。如需定制私有化模板（按你公司的 PPT 模板规范生成），可联系作者微信 **duge360**。

### Q8：生成的 PPT 中图表会自动更新数据吗？

原生 PPT 图表（`shape.has_chart == True`）可以通过 `build_pptx.py --chart-data` 同步更新数据。装饰性图形（如环形进度条、弧形仪表盘）是固定形状，改文字不会改变图形弧长。具体参考 `references/chart-editing.md`。

### Q9：我的内容太长，超出文本框容量怎么办？

`build_pptx.py --strict` 会检测文本溢出并拒绝保存。解决方案：**缩短文字，不要缩字号**。把长句改成短句，用更精炼的表达。`detail.json` 中每个 slot 都有 `max_chars`（最大容量，含 20% 余量）和 `guidance`（写什么内容的建议）。

---

## 进阶：命令行参考

如果你想了解 AI 背后的实际执行过程：

### 查询幻灯片注册表

```bash
# 查询所有带 "swot" 标签的 deep-blue 家族幻灯片
python3 -c "
import json
reg = json.load(open('slide-registry.json'))
for s in reg['slides']:
    if 'swot' in s['tags'] and s['style_family'].startswith('deep-blue'):
        print(f\"{s['slide_id']:45s} | {s['style_family']:35s} | {s['use_for']}\")
"

# 查询所有封面页
python3 -c "
import json
reg = json.load(open('slide-registry.json'))
for s in reg['slides']:
    if s['role'] == 'cover':
        print(f\"{s['slide_id']:40s} | {s['style_family']:30s} | {s['use_for']}\")
"
```

### 跨模板多源构建

```bash
python3 scripts/build_pptx.py \
    edits.json \
    out/final.pptx \
    --registry slide-registry.json \
    --styles styles/ \
    --style deep-blue-minimal \
    --strict
```

### 渲染预览

```bash
python3 scripts/render_slides.py out/final.pptx out/preview --dpi 144
```

---

## 总结

Gorden PPT Skill 的核心能力：

1. **19 套模板，640+ 张幻灯片** — 覆盖商务、学术、党政、咨询等主流场景
2. **跨模板混搭** — 从不同模板挑选最佳版式，统一风格输出（v2 核心特性）
3. **不改排版只换文字** — 保证专业设计师的版面品质不丢失
4. **自动容量检测** — 超出文本框自动警告，`--strict` 模式拒绝保存
5. **所有模型兼容** — DeepSeek / Claude / GPT / 国产模型均可

有问题或建议？欢迎加微信 **duge360** 或来 [Issues](https://github.com/springzh/PPTSkill/issues) 反馈。

---

> 📖 相关文档：
> - [SKILL.md](./SKILL.md) — AI 入口文档（技术规范）
> - [DEVELOPER_GUIDE.md](./DEVELOPER_GUIDE.md) — 开发者指南（架构与流水线）
> - [templates/INDEX.md](./templates/INDEX.md) — 模板分类索引（按色系/场景/规模筛选）
> - [README.md](./README.md) — 仓库概览
