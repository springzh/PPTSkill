# Changelog

按版本倒序列出可读变更。机器读取请用 [`updates.json`](./updates.json)；只读哪些文件变动请用 [`manifest.json`](./manifest.json) 的 `last_modified` 字段。

## 1.0.1 — 2026-05-28

**清理 `premium-corp` 模板自带的宣传页 + 加速增量更新流程。**

模板变更：

- `templates/premium-corp/template.pptx`：删除原模板前 3 页（稻壳儿封面 + "35 页精选" 介绍 + 缩略图墙），总页数 38 → 35
- `templates/premium-corp/detail.json`：移除被删页面的 slot 数据，剩余 35 页全部重新编号；slot_id、page_roles、skip_pages、shape_caution_pages、data_charts 等所有 slide-number 引用同步重映射
- `templates/premium-corp/intro.md`：重写"页面结构"小节以反映 35 页真实内容案例
- `templates/premium-corp/preview.png`：重新拼接 2×2 预览图，4 张全部来自真实内容页，不再展示宣传页

脚本改进：

- `scripts/check_update.py`：新增对 `git+` update_source 的支持（之前只支持 HTTP）。现在直接 `git clone --filter=blob:none --no-checkout` 后只拿 `updates.json`，体积小、速度快。
- `scripts/apply_update.py`：克隆时设置 `GIT_LFS_SKIP_SMUDGE=1`，**只对实际变动的 `.pptx` 选择性 `git lfs pull --include=...`**，把更新流量从 ~535MB 降到 ~10MB。
- `scripts/build_manifest.py`：排除 Office 临时锁文件 `~$*.*`。

升级方式（v1.0.0 → v1.0.1）：

```bash
python3 scripts/check_update.py        # 看到 1.0.0 → 1.0.1
python3 scripts/apply_update.py        # 只下载 ~11 个改动文件（含 1 个 LFS pptx）
```

注：v1.0.0 用户首次升级时，本地 `apply_update.py` 仍是旧版（会拉全部 LFS）；从 v1.0.1 起每次升级仅下载真正变动的文件。

## 1.0.0 — 2026-05-27

**首次发布。**

新增：

- 核心脚本：`render_slides.py` / `extract_template.py` / `stitch_preview.py` / `build_pptx.py` / `batch_extract.py` / `scaffold_detail.py` / `check_update.py` / `apply_update.py`
- 17 个内置模板（每个 = `template.pptx` + `intro.md` + `detail.json` + `preview.png`）：
  - `minimal-business-summary` — 简约商务总结汇报（深蓝白，16 页）
  - `red-patriot-youth` — 新时代新青年红色教育（党政红，16 页）
  - `cute-orange-class` — 橙色可爱卡通教学（暖橙，17 页）
  - `quarterly-illust` — 蓝灰酸性插画季度总结（Y2K 亮蓝，19 页）
  - `geometric-summary` — 多彩几何工作总结（4 主色，21 页）
  - `red-patriot-general` — 红色爱国主题教育通用（党政红，25 页）
  - `thesis-novice` — 多专业开题方法论库（墨绿，34 页）
  - `premium-corp` — 高级感大厂 PPT 合辑（酱红，38 页）
  - `architecture-deck` — 领导爱的架构图合辑（深蓝，40 页）
  - `thesis-formula` — 开题报告万能公式（暖米色，41 页）
  - `data-viz-deck` — 数据可视化合辑（深蓝砖红，41 页，含原生 chart）
  - `top-thesis` — 名校开题报告合辑（酒红，41 页）
  - `mckinsey-style` — 麦肯锡风专业模板（酱红，42 页）
  - `report-savior` — 汇报救命合辑（深蓝，49 页）
  - `operations-deck` — 运营 PPT 合辑（深蓝，52 页）
  - `competition-speech` — 竞聘述职合辑（深蓝砖红，59 页）
  - `report-template-massive` — 118 页超大工作汇报合辑（深蓝）
- 参考文档 `references/` 系列
- 版本机制：`VERSION` / `CHANGELOG.md` / `updates.json` / `manifest.json`
- 字体 fallback 链：WenQuanYi Micro Hei → DengXian → Noto Sans SC → PingFang SC
- 非商业使用声明在 SKILL.md 顶部明确化
