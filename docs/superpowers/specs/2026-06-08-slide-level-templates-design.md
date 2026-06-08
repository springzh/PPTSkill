# Slide-Level Template Architecture — Design Spec

**Date:** 2026-06-08
**Status:** Design approved, awaiting implementation plan
**Scope:** Refactor GordenPPTSkill from monolithic template files to composable slide-level templates with separable styles

---

## Problem

Today each template is a monolithic `.pptx` file. Users pick one template and are locked into its slide set. Problems:

1. **No cross-template mixing.** A user who needs "SWOT analysis" + "4-card comparison" + "wheel method" can't get it — SWOT lives in `report-massive-models`, 4-card in `minimal-business-summary`, wheel in `minimal-business-summary`. Different templates, different styles.
2. **Overkill for simple decks.** Many templates are 35-59 slides of curated material. A user who needs 5 slides still carries the baggage of picking from a large template.
3. **Duplicate near-identical slides.** The 9 "Deep Blue" templates share a color palette and font scheme but exist as separate `.pptx` files with no shared infrastructure. A SWOT slide in one template is structurally identical to a SWOT slide in another — only colors differ.
4. **Style is baked into slides.** Colors and fonts are hardcoded per shape, not inherited from a theme. There is no way to say "take this slide layout but render it in red."

---

## Design Decisions

| Decision | Choice | Rationale |
|---|---|---|
| Architecture approach | **Additive Registry Layer** (not full rewrite) | Zero migration of existing files, backward compatible, ships incrementally |
| Style scope | **Colors + Fonts only** | These give the biggest visual impact for least complexity; shapes/decor/layout stay with the slide |
| Slide organization | **Slide Registry** (unified index with source references) | Clean unified view of all ~640 slides; within-family and cross-family mixing both supported |
| Style application | **Programmatic color/font remapping** via color_map | No template rebuild required; works by walking shapes at build time |
| Rollout | **3-phase incremental** | Phase 1 delivers core value immediately (mix slides across 4 Deep Blue templates, 129 slides); Phase 2 expands to 9 templates (418 slides) |

---

## New Files

### 1. `slide-registry.json` — Unified Slide Index

One flat array indexing every non-promo slide across all 19 templates. Each entry is a lightweight reference — slot details remain in the existing `detail.json` files.

```jsonc
{
  "$schema": "slide-registry/v1",
  "version": "1.0.0",
  "build_date": "2026-06-08",
  "slides": [
    {
      "slide_id": "minimal-cover-1",
      "source_pptx": "templates/minimal-business-summary/template.pptx",
      "source_slide": 1,
      "style_family": "deep-blue-minimal",
      "layout_type": "cover",
      "role": "cover",
      "use_for": "极简封面，中央对齐中英文标题，四角细线装饰",
      "tags": ["cover", "minimal", "centered", "bilingual"],
      "detail_ref": "templates/minimal-business-summary/detail.json",
      "slots_count": 2,
      "text_density": "low"
    }
  ]
}
```

**Key fields:**
- `slide_id` — universal identifier used in `edits.json` (format: `{template-slug}-{role}-{slide_number}`)
- `style_family` — determines within-family vs cross-family behavior at build time
- `detail_ref` — points to the existing `detail.json` for full slot metadata (no duplication)
- `tags` — searchable keywords for AI-driven slide selection (e.g., `["swot", "quadrant", "strategy"]`)
- `layout_type` — normalized layout category across templates

### 2. `styles/<slug>.json` — Style Definitions

Each style captures the visual identity: palette, fonts, and the critical `color_map` that enables cross-family remapping.

```jsonc
{
  "$schema": "ppt-style/v1",
  "slug": "deep-blue-minimal",
  "name": "深蓝极简",
  "description": "深蓝主色 + 大量留白，适合季度/年度汇报",
  "palette": {
    "primary":        "#485275",
    "secondary":      "#44546A",
    "accent":         "#E7E6E6",
    "background":     "#FFFFFF",
    "text_on_light":  "#334155",
    "text_on_dark":   "#F8FAFC"
  },
  "fonts": {
    "cn_heading": "微软雅黑",
    "cn_body":    "微软雅黑",
    "en_heading": "Arial",
    "en_body":    "Arial"
  },
  "color_map": {
    "#485275": "primary",
    "#44546A": "secondary",
    "#E7E6E6": "accent",
    "#FFFFFF": "background",
    "#334155": "text_on_light",
    "#F8FAFC": "text_on_dark"
  },
  "font_map": {
    "微软雅黑": "cn_body",
    "Arial": "en_body",
    "思源黑体": "cn_body"
  }
}
```

**`color_map` is the key enabler:** it tells the build script "when you see `#485275` in a slide from this style family, it's the `primary` role. If the target style defines `primary` as `#A91F1F`, swap it." Only colors listed in `color_map` are remapped — everything else is treated as intentional designer choice.

### 3. Updated `edits.json` — Multi-Source Format (backward compatible)

**Old format** (still works):
```jsonc
{
  "template_slug": "minimal-business-summary",
  "selected_slides": [1, 2, 5, 9, 16],
  "edits": [
    {"slide": 1, "slot_id": "cover_title_cn", "new_text": "2026 复盘"}
  ]
}
```

**New format** (multi-source):
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

**Detection:** `"template_slug"` present → legacy path. `"style"` present → new multi-source path.

---

## Build Pipeline (build_pptx.py changes)

The refactored script adds two phases. Existing phases 1 and 3 are adapted, not rewritten.

```
edits.json → Phase 1: RESOLVE → Phase 2: ASSEMBLE → Phase 3: EDIT → Phase 4: APPLY STYLE → output.pptx
```

### Phase 1: Resolve (NEW logic for multi-source)

Map each `slide_id` to its physical location via the registry:

```python
def resolve_slides(edits_spec, registry):
    resolved = []
    sources = set()
    for entry in edits_spec["slides"]:
        slide = registry.lookup(entry["slide_id"])
        resolved.append({
            "slide_id": entry["slide_id"],
            "source_pptx": slide["source_pptx"],
            "source_slide": slide["source_slide"],
            "style_family": slide["style_family"],
            "detail_ref": slide.get("detail_ref")
        })
        sources.add(slide["source_pptx"])
    return resolved, sources
```

Fails fast if a `slide_id` is not found in the registry.

### Phase 2: Assemble (NEW — replaces single-source prune_slides)

Clone slides from potentially multiple source `.pptx` files into a single output presentation. Uses XML-level slide cloning (python-pptx has no built-in `slide.clone()`). Source presentations are cached to avoid repeated file opens.

### Phase 3: Edit Text (adapted from existing)

The existing `apply_edit_to_shape()` logic is preserved. The change: `slot_id` resolution now queries a merged index built from multiple `detail.json` files (one per source template), keyed by `(slide_id, slot_id)`.

### Phase 4: Apply Style (NEW)

Walk every shape in the assembled presentation. For each slide:
1. Look up its source style family
2. If source style == target style → skip (within-family, no remapping needed)
3. If cross-family → build color_map from source_style → target_style
4. Walk all shapes recursively, remapping fill colors, line colors, and font colors
5. Remap font names via font_map

```python
def apply_style(prs, target_style, resolved_slides, style_registry):
    target = style_registry.load(target_style)
    for i, entry in enumerate(resolved_slides):
        source = style_registry.load(entry["style_family"])
        if source.slug == target.slug:
            continue
        color_map = build_color_map(source, target)
        for shape in all_shapes(prs.slides[i]):
            remap_shape(shape, color_map, source.font_map, target.fonts)
```

### New CLI

```bash
# Legacy mode (unchanged):
python3 scripts/build_pptx.py template.pptx edits.json out.pptx --detail detail.json --strict

# New multi-source mode:
python3 scripts/build_pptx.py edits.json out.pptx \
    --registry slide-registry.json \
    --styles styles/ \
    --style deep-blue-minimal \
    --strict
```

### New Script: `build_registry.py`

Scans all `templates/*/detail.json` files and generates `slide-registry.json`. Run whenever templates are added or modified. Uses a predefined `FAMILY_MAP` to assign `style_family` to each template slug.

---

## Style Application Edge Cases

| Edge Case | Risk | Handling |
|---|---|---|
| Shape uses a color NOT in source `color_map` | High | Skip it. Only remap explicitly mapped colors. Designer accent choices are preserved. |
| Source and target have different palette role counts | Medium | Fallback chain: accent → secondary → primary. If no match, keep original. |
| Gradient fills with multiple stops | Medium | Remap each stop individually. Preserve stop positions and transparency. |
| Font change causes text overflow (Arial → 思源宋体) | High | Re-run overflow check after font swap. In `--strict` mode, refuse to save if overflow detected. |
| Images/icons with embedded colors | Low | Cannot remap. Document as known limitation — minor visual inconsistency possible for cross-family icon-heavy decks. |
| `capacity_unknown: true` slots | Low | Skip overflow check for these slots (same as current behavior). |

---

## Style Families

Clustered from the existing 19 templates based on palette similarity:

| Style Family | Member Templates | Slides | Rollout |
|---|---|---|---|
| `deep-blue-business-massive` | report-massive-models, report-massive-charts, report-massive-reports | 113 | **Phase 1** |
| `deep-blue-minimal` | minimal-business-summary | 16 | **Phase 1** |
| `deep-blue-corp-extended` | data-viz-deck, report-savior, operations-deck, competition-speech, architecture-deck | 233 | Phase 2 |
| `bordeaux-corp` | premium-corp, mckinsey-style | 72 | Phase 2 |
| `patriotic-red` | red-patriot-youth, red-patriot-general | 41 | Phase 3 |
| `warm-academic` | cute-orange-class, thesis-novice, thesis-formula | 88 | Phase 3 |
| `wine-academic` | top-thesis | 39 | Phase 3 |
| `standalone` | quarterly-illust, geometric-summary | 40 | Phase 3 |

**Phase 1 alone** enables mixing slides across 4 templates (129 slides) in the Deep Blue family — this covers the most common business use cases.

---

## AI Workflow Changes

### Mode A: Built-in Templates (upgraded)

**Before:** Pick one template → read its intro.md + detail.json → select slides from that template → build.

**After:** Understand content needs → query slide-registry.json by tags/layout_type → select slides from any matching template (prefer same style family) → write multi-source edits.json → build.

Slide selection example:
```
User: "年度复盘，需要封面、SWOT分析、时间线、4个关键指标、结束页"
AI queries registry:
  - tags:"cover" + style_family:"deep-blue-*" → 4 candidates
  - tags:"swot" + style_family:"deep-blue-*" → 2 candidates
  - tags:"timeline" + style_family:"deep-blue-*" → 3 candidates
  - tags:"kpi" OR tags:"4-up" → 5 candidates
  - tags:"ending" → 2 candidates
AI selects best 5 slides → writes edits.json → builds
```

### Mode B: User-Provided Template — unchanged
### Mode C: Original Design — unchanged

---

## File Layout (Post-Refactor)

```
GordenPPTSkill/
├── SKILL.md                       ← updated Mode A workflow
├── slide-registry.json            ← NEW
├── styles/                        ← NEW
│   ├── deep-blue-minimal.json
│   ├── deep-blue-business-massive.json
│   ├── deep-blue-corp-extended.json
│   ├── bordeaux-corp.json
│   ├── patriotic-red.json
│   ├── warm-academic.json
│   └── wine-academic.json
├── scripts/
│   ├── build_pptx.py              ← MODIFIED: multi-source + style
│   ├── build_registry.py          ← NEW
│   ├── render_slides.py           ← unchanged
│   ├── compute_capacity.py        ← unchanged
│   ├── check_update.py            ← unchanged
│   ├── apply_update.py            ← unchanged
│   └── build_manifest.py          ← unchanged
├── templates/                     ← UNCHANGED (all 19)
│   └── ...
└── references/
    ├── workflow.md                ← updated
    └── ...
```

---

## Backward Compatibility

| Feature | Status |
|---|---|
| Old `edits.json` format (template_slug + slide numbers) | ✅ Fully supported |
| `--detail` flag for single-template mode | ✅ Fully supported |
| All 19 template directories | ✅ Unchanged |
| `render_slides.py`, `compute_capacity.py` | ✅ Unchanged |
| `editing_rules` in detail.json | ✅ Still enforced |
| `--strict` mode | ✅ Works in both old and new paths |
| `--dry-run` | ✅ Works in both paths |

---

## Non-Goals (explicitly out of scope)

- Rebuilding slides to use PowerPoint-native Slide Layouts with theme color references
- Extracting individual slides into standalone `.pptx` files
- WYSIWYG slide editor or browser-based slide composer
- AI-powered slide layout generation
- Real-time style preview before build
- Cross-family font metrics compensation (just detect overflow, don't auto-fix)
