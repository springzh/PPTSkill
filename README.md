# Gorden PPT Skill

> 史上最强原生PPT Skill，更适合中国宝宝。生成的效果不让你震惊，你来打我。
>
> ⚠️ **非商业使用**：本仓库及内置模板**仅供个人学习与研究**，禁止任何商业用途。
> 
> **支持定制私有化模板**：如果你想要Agent能按你公司的PPT模板来生成PPT，可以加我微信**duge360**定制。

## 交流群

扫码加入「PPT Skill 交流群」一起讨论 / 反馈问题 / 看新版本：

<p align="center">
  <img src="./assets/group-qr.jpg" alt="PPT Skill 交流群二维码" width="320" />
</p>

> ⚠️ 群二维码 7 天内有效；过期请来 [Issues](https://github.com/GordenSun/GordenPPTSkill/issues) 留言，我会贴新的。

## 几大特色
1、能生成信息密度高、排版复杂、看起来高大上的PPT，也支持生成简约、商务风格的PPT。适合国企、互联网大厂使用。

2、**跨模板混搭幻灯片**（v2 新架构）：19 套模板的 640+ 张幻灯片可按需自由组合。比如从 `report-massive-models` 取 SWOT 分析页，从 `report-savior` 取鱼骨图页，从 `minimal-business-summary` 取封面和结束页，统一渲染为 `deep-blue-minimal` 风格——一份 PPT 融合多套模板的精华版式。

3、兼容所有模型，DeepSeek、小米Mimo、Claude、GPT均实测过，国产模型也能完成的非常好。

4、技能自动更新机制：如果我更新了可选用的PPT模板，使用技能时会自动更新技能。技能像软件一样可以更新

## 效果展示

以下是用本 Skill 生成的实际页面（信息密度高、排版复杂、商务质感）：

<p align="center">
  <img src="./assets/showcase-1.jpg" alt="效果展示 1 - 战略定位分析" width="720" />
</p>

<p align="center">
  <img src="./assets/showcase-2.jpg" alt="效果展示 2 - 价值冰山模型" width="720" />
</p>

<p align="center">
  <img src="./assets/showcase-3.jpg" alt="效果展示 3 - 产品化闭环" width="720" />
</p>

## 使用方法
极其简单，把这段提示词发给你的Agent即可：

安装这个Skill：https://github.com/GordenSun/GordenPPTSkill
然后使用这个Skill做一个复杂、豪华的PPT，读取本地XXX文件，来介绍XXX项目

## 谁要看这个

- 想给自己用的 AI 助手装一个"做 PPT"技能的人：**请读 [SKILL.md](./SKILL.md)**
- 想看本项目目录怎么组织：继续往下看本文件
- 想理解模板分类 / 推荐：**请读 [templates/INDEX.md](./templates/INDEX.md)**

## 快速开始（命令行）

### 传统单模板方式

```bash
# 1. 确认依赖
python3 -c "import pptx; print(pptx.__version__)"   # python-pptx 1.0+
soffice --version    # LibreOffice (仅渲染预览时需要)
which pdftoppm       # poppler   (仅渲染预览时需要)

# 2. 选定模板 + 写 edits.json，跑构建
python3 scripts/build_pptx.py \
    templates/minimal-business-summary/template.pptx \
    edits.json \
    out/final.pptx \
    --detail templates/minimal-business-summary/detail.json

# 3. (可选) 渲染最终预览图
python3 scripts/render_slides.py out/final.pptx out/preview --dpi 144
```

### 跨模板混搭方式（新架构）：一个真实示例

假设你要做一份"2026 年度战略复盘"PPT，需要以下版式：封面、SWOT 分析、时间线、4 个 KPI 卡片、鱼骨图问题分析、结束页。单套模板无法同时覆盖这些版式——走跨模板混搭路径。

**Step 1: 查询 slide-registry 找候选幻灯片**

```bash
# 查 SWOT 相关幻灯片（deep-blue 家族）
python3 -c "
import json
reg = json.load(open('slide-registry.json'))
for s in reg['slides']:
    if 'swot' in s['tags'] and s['style_family'].startswith('deep-blue'):
        print(f\"{s['slide_id']:45s} | {s['style_family']:35s} | {s['use_for']}\")
"
# 输出示例：
# report-massive-models-swot-12                | deep-blue-business-massive          | SWOT 战略分析四象限
# report-savior-swot-8                         | deep-blue-corp-extended             | SWOT 分析（带说明栏）

# 查时间线、KPI、鱼骨图、封面、结束页……
```

**Step 2: 写多源 edits.json**

```jsonc
{
  "style": "deep-blue-minimal",
  "slides": [
    {"slide_id": "minimal-cover-1",              "source_slide": 1},
    {"slide_id": "report-massive-models-swot-12", "source_slide": 12},
    {"slide_id": "report-savior-timeline-22",     "source_slide": 22},
    {"slide_id": "minimal-4card-5",               "source_slide": 5},
    {"slide_id": "report-savior-fishbone-15",     "source_slide": 15},
    {"slide_id": "minimal-ending-16",             "source_slide": 16}
  ],
  "edits": [
    {"slide_id": "minimal-cover-1",              "slot_id": "cover_title_cn",  "new_text": "2026 年度战略复盘"},
    {"slide_id": "minimal-cover-1",              "slot_id": "cover_title_en",  "new_text": "Strategic Review 2026"},
    {"slide_id": "report-massive-models-swot-12", "slot_id": "swot_strength_title", "new_text": "技术壁垒"},
    {"slide_id": "report-massive-models-swot-12", "slot_id": "swot_weakness_title", "new_text": "交付效率"},
    {"slide_id": "report-savior-timeline-22",     "slot_id": "timeline_m1_cn", "new_text": "Q1 产品重构完成"},
    {"slide_id": "minimal-4card-5",               "slot_id": "card1_title_cn", "new_text": "营收增长 34%"},
    {"slide_id": "report-savior-fishbone-15",     "slot_id": "fishbone_head_cn", "new_text": "交付延期根因分析"},
    {"slide_id": "minimal-ending-16",             "slot_id": "ending_thanks_cn", "new_text": "谢谢！欢迎提问"}
  ]
}
```

**Step 3: 多源构建**

```bash
python3 scripts/build_pptx.py \
    edits.json \
    out/战略复盘_2026.pptx \
    --registry slide-registry.json \
    --styles styles/ \
    --style deep-blue-minimal \
    --strict
```

构建流水线自动完成：**Resolve**（slide_id → 物理文件+页码）→ **Assemble**（从 3 个不同 .pptx 克隆幻灯片）→ **Edit**（替换文字）→ **Apply Style**（跨家族颜色自动重映射到 deep-blue-minimal 风格）

这 6 张幻灯片来自 3 套不同模板（`minimal-business-summary`、`report-massive-models`、`report-savior`），但因为都属于 deep-blue 色系家族，最终输出浑然一体。

## 字体环境

模板大量使用 `微软雅黑`。如果你的机器没装它，配 `~/.config/fontconfig/fonts.conf` 加一条 alias：

```xml
<alias binding="strong">
  <family>微软雅黑</family>
  <accept>
    <family>WenQuanYi Micro Hei</family>
    <family>DengXian</family>
    <family>Noto Sans SC</family>
    <family>PingFang SC</family>
  </accept>
</alias>
```

(`brew install --cask font-noto-sans-sc`，或下载 WenQuanYi Micro Hei 放进 `~/Library/Fonts/` 并 `fc-cache -f`。)

## 目录速览

```
SKILL.md              # AI 入口文档
VERSION               # 当前版本
CHANGELOG.md          # 人读变更
updates.json          # 机读变更
manifest.json         # 每文件版本 + sha256
slide-registry.json   # 640+ 张幻灯片统一索引（按标签/版式/风格家族检索）
styles/               # 可分离风格定义（色彩 + 字体 + color_map）
scripts/              # 7 个脚本（build / build_registry / render / compute_capacity / update / manifest）
references/           # 编辑规则、Schema、工作流参考
templates/            # 19 套模板（每套 4 文件）
```

## 致谢与版权

- 本仓库没有PPT模板的版权
- **禁止任何二次分发 / 商业使用**
- 用到的开源工具：[LibreOffice](https://www.libreoffice.org/)、[python-pptx](https://python-pptx.readthedocs.io/)、[Poppler](https://poppler.freedesktop.org/)、[WenQuanYi Micro Hei](http://wenq.org/)
- 感谢 [LinuxDO](https://linux.do) 社区的支持
