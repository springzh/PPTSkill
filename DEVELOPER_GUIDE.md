# GordenPPTSkill — Developer Guide: Template-to-PPTX Pipeline

This guide explains the architecture and data flow of the GordenPPTSkill system, with a focus on how templates are selected, introspected, and programmatically composed into final `.pptx` output files.

---

## Table of Contents

1. [System Overview](#system-overview)
2. [Template Anatomy](#template-anatomy)
3. [The detail.json Contract](#the-detailjson-contract)
4. [The edits.json Edit Specification](#the-editsjson-edit-specification)
5. [build_pptx.py — The Composition Engine](#build_pptxpy--the-composition-engine)
6. [Capacity System: Text-Fits-the-Box](#capacity-system-text-fits-the-box)
7. [Type Scale & Heading Consistency](#type-scale--heading-consistency)
8. [Supporting Scripts](#supporting-scripts)
9. [Three Composition Modes](#three-composition-modes)
10. [End-to-End Data Flow](#end-to-end-data-flow)

---

## System Overview

```
┌──────────────┐     ┌──────────────┐     ┌──────────────┐
│  template.pptx │     │  detail.json  │     │  edits.json   │
│  (visual base) │     │  (slot map)   │     │  (text data)  │
└──────┬───────┘     └──────┬───────┘     └──────┬───────┘
       │                    │                    │
       │           ┌────────┴────────┐           │
       │           │ slide-registry  │           │
       │           │ (unified index) │           │
       │           └────────┬────────┘           │
       │                    │                    │
       │           ┌────────┴────────┐           │
       │           │   styles/*.json  │           │
       │           │ (color + fonts)  │           │
       │           └────────┬────────┘           │
       │                    │                    │
       └────────────────────┼────────────────────┘
                            │
                   ┌────────▼────────┐
                   │  build_pptx.py  │
                   │  4-phase pipe:  │
                   │  Resolve→       │
                   │  Assemble→      │
                   │  Edit→          │
                   │  Apply Style    │
                   └────────┬────────┘
                            │
                   ┌────────▼────────┐
                   │   output.pptx   │
                   │  (final deck)   │
                   └────────┬────────┘
                            │
                   ┌────────▼────────┐
                   │ render_slides.py│
                   │   (preview)     │
                   └────────┬────────┘
                            │
                   ┌────────▼────────┐
                   │  slide-XX.png   │
                   │  (per-page PNG) │
                   └─────────────────┘
```

The system is **template-driven and edit-specification-based**: a template `.pptx` provides the visual design, a `detail.json` maps every editable text position to a stable identifier, and an `edits.json` specifies which slides to keep and what text to place in each slot. The `build_pptx.py` script combines all three to produce the final output — it never modifies the original template file.

**New in v2:** `slide-registry.json` provides a unified index of 640+ slides across all 19 templates, enabling cross-template slide mixing. `styles/*.json` defines separable color palettes and font sets with `color_map` entries that enable automatic color remapping when slides from different style families are combined into one deck.

---

## Template Anatomy

Each template lives in `templates/<slug>/` and consists of exactly 4 files:

| File | Role | Created by |
|---|---|---|
| `template.pptx` | The raw PowerPoint file — all slides, shapes, colors, fonts, animations intact | Designer |
| `intro.md` | Human + AI-readable summary: style tags, color palette, page structure, use-cases | Template maintainer |
| `detail.json` | Machine-readable map of every text slot on every page with addressing info + capacity metadata | `compute_capacity.py` |
| `preview.png` | 2×2 grid preview of 4 representative slides for visual selection | Screenshot tool |

### Template Directory Structure

```
templates/minimal-business-summary/
├── template.pptx      # 16 slides, 16:9, 微软雅黑 + Arial
├── intro.md           # "极简商务，留白多，4 章节，季度/年度汇报首选"
├── detail.json        # ~4000 lines: full slot map with capacity data
└── preview.png        # 2×2 preview for AskQuestion selection UI
```

### Page Roles

Every slide in a template is classified into one of these roles (tracked in `detail.json → page_roles`):

| Role | Meaning | Optional? |
|---|---|---|
| `cover` | Title slide with presentation name | Can be empty → no cover |
| `agenda` | Table of contents with chapter list | Can be empty → no TOC |
| `section_divider` | Chapter separator with large number + title | Can be empty → no dividers |
| `content` | Body slides with text slots for real content | Required — the payload |
| `ending` | Closing slide ("Thank you" etc.) | Can be empty → ends on last content slide |
| `template_promo` | Advertisement for the template author | Always excluded via `skip_pages` |

**Key principle**: a template only has the roles it was designed with. If `cover` is empty, the output starts directly from the first content slide — no improvisation.

---

## Slide-Level Template Architecture (v2)

The v2 architecture decouples three concerns that were previously bundled inside each monolithic template:

### slide-registry.json — Unified Slide Index

A flat JSON array indexing every non-promo slide across all 19 templates. Each entry is a lightweight reference — full slot details remain in the existing per-template `detail.json` files.

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

Key fields:
- `slide_id` — universal identifier (format: `{template-slug}-{role}-{slide_number}`)
- `style_family` — determines within-family vs cross-family behavior at build time
- `detail_ref` — points to the existing `detail.json` for full slot metadata (no duplication)
- `tags` — searchable keywords for AI-driven slide selection
- `layout_type` — normalized layout category across templates

Generated by `scripts/build_registry.py`, which scans all `templates/*/detail.json` files.

### styles/<slug>.json — Separable Style Definitions

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
    "Arial": "en_body"
  }
}
```

**`color_map` is the key enabler:** it tells the build script "when you see `#485275` in a slide from this style family, it's the `primary` role. If the target style defines `primary` as `#1E3A5F`, swap it." Only colors listed in `color_map` are remapped — everything else is treated as intentional designer choice.

### Style Families

Templates are clustered by palette similarity into 8 style families:

| Style Family | Member Templates | Slides | Rollout |
|---|---|---|---|
| `deep-blue-business-massive` | report-massive-models, report-massive-charts, report-massive-reports | 113 | Phase 1 |
| `deep-blue-minimal` | minimal-business-summary | 16 | Phase 1 |
| `deep-blue-corp-extended` | data-viz-deck, report-savior, operations-deck, competition-speech, architecture-deck | 233 | Phase 2 |
| `bordeaux-corp` | premium-corp, mckinsey-style | 72 | Phase 2 |
| `patriotic-red` | red-patriot-youth, red-patriot-general | 41 | Phase 3 |
| `warm-academic` | cute-orange-class, thesis-novice, thesis-formula | 88 | Phase 3 |
| `wine-academic` | top-thesis | 39 | Phase 3 |
| `standalone` | quarterly-illust, geometric-summary | 40 | Phase 3 |

Within-family mixing requires no color remapping. Cross-family mixing triggers automatic color + font remapping via `color_map` and `font_map`.

### Additive Design

The registry + styles layer is **purely additive** — all 19 existing template directories, `detail.json` files, and `intro.md` files remain unchanged. The legacy single-template build path is fully backward compatible.

---

## The detail.json Contract

`detail.json` is the bridge between the visual `.pptx` and the programmatic edit layer. It is versioned under schema `ppt-template-detail/v1`.

### Top-Level Structure

```jsonc
{
  "$schema": "ppt-template-detail/v1",
  "slug": "minimal-business-summary",
  "name": "简约商务总结汇报",
  "slide_count": 16,
  "slide_size_cm": [33.87, 19.05],       // 16:9 landscape
  "aspect": "16:9",
  "style": "minimalist-business",
  "theme_colors": ["#485275", "#44546A", "#E7E6E6", "#FFFFFF"],
  "fonts": {"cn": "微软雅黑", "en": "Arial"},

  "editing_rules": ["..."],               // Human-readable constraints
  "skip_pages": [],                       // Slides to never include
  "shape_caution_pages": [15],            // Slides with decorative-only graphics
  "data_charts": [],                      // Native PPT charts (can be data-driven)

  "page_roles": {                         // Index: which slides serve which role
    "cover": [1],
    "agenda": [2],
    "section_divider": [3, 7, 10, 13],
    "content": [4, 5, 6, 8, 9, 11, 12, 14, 15],
    "ending": [16]
  },

  "pages": [ /* ... per-page slot arrays ... */ ]
}
```

### Per-Page Structure

Each page entry in the `pages` array:

```jsonc
{
  "slide_number": 4,                  // 1-based, matches template.pptx order
  "role": "content",                  // cover / agenda / section_divider / content / ending
  "layout": "顶部章节面包屑 + 左侧 3 块图片菱形拼接...",
  "use_for": "用 3 个并列要点展示一个主题...",
  "cautions": ["..."],                // Optional: known limitations of this page
  "preview_slot_in_2x2": "左上",      // For preview.png mapping (null if not in preview)
  "title_slot_id": "p4_breadcrumb_cn",
  "body_slot_ids": ["p4_item1_body", "p4_item2_body", "p4_item3_body"],
  "decorative_slot_ids": ["p4_breadcrumb_number", "p4_item1_num", ...],
  "text_density": "high",
  "text_slots": [ /* ... */ ]
}
```

### Per-Slot Structure (the core addressing unit)

```jsonc
{
  "slot_id": "cover_title_cn",        // GLOBALLY UNIQUE within this detail.json
  "role": "封面主标题",                // Human-readable semantic role
  "position": "中部居中",              // Rough position hint
  "address": {                        // PHYSICAL ADDRESS in python-pptx terms
    "shape_id": 46,                   //   Shape ID (unique per slide)
    "paragraph": 1,                   //   Paragraph index (0-based)
    "run": 0                          //   Run index within paragraph (0-based)
  },
  "current_text": "工作计划模板",       // For sanity check / --strict matching
  "max_chars": 14,                    // Capacity budget (visual-width units, includes 20% tolerance)
  "chars_per_line": 6,                // How many visual-width units fit per line
  "max_lines": 2,                     // How many lines the box can hold
  "font_size_pt": 54.0,
  "font_name": "微软雅黑",
  "wrap": true,                       // Does PowerPoint wrap text in this box?
  "autofit": false,                   // Does PowerPoint shrink-to-fit?
  "level": 2,                         // Type scale level (1 = largest heading)
  "language": "zh",                   // zh / en / number
  "editable": true,                   // false = decorative, don't touch
  "guidance": "6-12 字中文主标题，例如 '2026 年度工作汇报'"
}
```

### Address Resolution

There are two address forms, both using python-pptx's coordinate system:

| Form | JSON | Behavior |
|---|---|---|
| **Exact run** | `{"shape_id": 46, "paragraph": 1, "run": 0}` | Replace only that run's text; other runs in the same paragraph untouched |
| **Whole paragraph** | `{"shape_id": 46, "paragraph": 1}` | Replace run 0 with new text, clear all other runs. Run 0's formatting (font, color, size) is preserved |

The shape lookup is recursive: `find_shape()` walks into `MSO_SHAPE_TYPE.GROUP` children, so nested shapes in group hierarchies are reachable.

---

## The edits.json Edit Specification

`edits.json` is the input that drives composition. It is authored by the AI (or human) based on user content. Two formats are supported, detected by the presence of `template_slug` (legacy single-template) or `style` (multi-source).

### Single-Template Format (v1, backward compatible)

```jsonc
{
  "$schema": "ppt-edit-spec/v1",
  "template_slug": "minimal-business-summary",     // Optional, for audit trail
  "selected_slides": [1, 2, 3, 5, 7, 9, 10, 12, 13, 14, 16],
                                                   // 1-based template slide numbers, in presentation order
  "edits": [
    // Form A: via slot_id (requires --detail for resolution)
    {"slide": 1, "slot_id": "cover_title_cn", "new_text": "2026 年度复盘"},

    // Form B: via explicit address (no detail.json needed)
    {
      "slide": 1,
      "address": {"shape_id": 46, "paragraph": 1, "run": 0},
      "expected_text": "工作计划模板",              // Optional sanity check
      "new_text": "2026 年度复盘"
    }
  ]
}
```

### Multi-Source Format (v2, cross-template mixing)

```jsonc
{
  "$schema": "ppt-edit-spec/v2",
  "style": "deep-blue-minimal",                   // Target style slug; also used for format detection
  "slides": [                                      // Slides from potentially different source .pptx files
    {"slide_id": "minimal-cover-1",              "source_slide": 1},
    {"slide_id": "report-massive-models-swot-12", "source_slide": 12},
    {"slide_id": "minimal-4card-5",               "source_slide": 5},
    {"slide_id": "minimal-ending-16",             "source_slide": 16}
  ],
  "edits": [
    // slide_id replaces the numeric "slide" field
    {"slide_id": "minimal-cover-1", "slot_id": "cover_title_cn", "new_text": "2026 战略复盘"},

    // Address-based edits also work (shape_id must be unique within its source slide)
    {
      "slide_id": "report-massive-models-swot-12",
      "address": {"shape_id": 9, "paragraph": 0, "run": 0},
      "expected_text": "STRENGTHS",
      "new_text": "核心优势"
    }
  ]
}
```

**Format detection**: `"template_slug"` present → legacy single-template path. `"style"` present → multi-source path.

**Slot resolution in multi-source mode**: The build script uses `slide_id` → slide-registry → `detail_ref` → corresponding `detail.json` to look up slot metadata. Each slide resolves against its own source template's detail.json.

### Critical Rules (both formats)

1. **Every `editable: true` slot must appear in edits** — no placeholder text left behind
2. **`editable: false` slots are decorative** (serial numbers, "01"/"02", "%" symbols) — usually excluded from edits
3. **Order is irrelevant** — duplicates resolve to last-write-wins
4. **Slide numbers are 1-based and refer to the original template** — not the output order

### Chapter Name Consistency

When you change a chapter name in the agenda (e.g., `agenda_ch1_cn`), you MUST also update:
- The corresponding section divider (`div1_cn`)
- All content-page breadcrumbs in that chapter (`p4_breadcrumb_cn`, `p5_breadcrumb_cn`, etc.)

This is enforced by convention, not by the tool — the `guidance` field in each dependent slot reminds the editor.

---

## build_pptx.py — The Composition Engine

`scripts/build_pptx.py` is the core script. In legacy single-template mode it orchestrates three phases. In multi-source mode (v2), it adds a fourth phase and a preceding resolve step.

### Single-Template Mode (legacy, 3 phases)

#### Phase 1: Slot Resolution

```python
slot_index = load_slot_index(args.detail)
# Returns: {(slide_number, slot_id): slot_metadata}
```

When `--detail detail.json` is provided, the script builds a lookup table mapping every `(slide_number, slot_id)` tuple to its full slot metadata — including the `address`, capacity fields, and `expected_text` for sanity checking.

#### Phase 2: Apply Edits (Pre-Prune)

Edits are applied to the **original template slides** (1-based indexing) *before* slide pruning. This is critical because slide indices shift after pruning.

For each edit:
1. Resolve the physical address — either from `slot_id` via the detail.json index, or directly from the edit's `address` field
2. Locate the shape by `shape_id` using recursive group-aware search
3. Mutate the run/paragraph text while preserving all formatting (font, color, size, bold, italic)
4. Run overflow detection (see [Capacity System](#capacity-system-text-fits-the-box))
5. In `--strict` mode: fail on `expected_text` mismatch, missing shape, or overflow

```python
def apply_edit_to_shape(shape, address, new_text, expected, strict):
    # Paragraph-level: replace run 0, clear others → preserves run 0 formatting
    # Run-level: replace exact run → other runs untouched
```

#### Phase 3: Slide Pruning

```python
def prune_slides(prs, selected_1indexed):
    # Manipulates the slide ID list directly via lxml
    # Keeps slides in selected_slides order, removes all others
    # Preserves relationship IDs (r:id) to avoid reference breakage
```

After pruning, `prs.save(args.output)` writes the final `.pptx`.

### Multi-Source Mode (v2, 4 phases)

When `--registry` is provided, the script enters multi-source mode:

```
edits.json → Phase 1: RESOLVE → Phase 2: ASSEMBLE → Phase 3: EDIT → Phase 4: APPLY STYLE → output.pptx
```

#### Phase 1: Resolve

Map each `slide_id` to its physical location via the slide registry:

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

#### Phase 2: Assemble

Clone slides from potentially multiple source `.pptx` files into a single output presentation. Uses XML-level slide cloning (python-pptx has no built-in `slide.clone()`). Source presentations are cached to avoid repeated file opens.

#### Phase 3: Edit Text

The existing `apply_edit_to_shape()` logic is preserved. The change: `slot_id` resolution now queries a merged index built from multiple `detail.json` files (one per source template), keyed by `(slide_id, slot_id)`.

#### Phase 4: Apply Style

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

### CLI Usage

```bash
# Legacy single-template mode:
python3 scripts/build_pptx.py \
    templates/minimal-business-summary/template.pptx \
    edits.json \
    out/final.pptx \
    --detail templates/minimal-business-summary/detail.json \
    --strict

# New multi-source mode:
python3 scripts/build_pptx.py \
    edits.json \
    out/final.pptx \
    --registry slide-registry.json \
    --styles styles/ \
    --style deep-blue-minimal \
    --strict

# Common flags (both modes):
#   --strict          Fail on any mismatch/overflow/missing shape
#   --no-lint         Skip overflow detection
#   --dry-run         Print planned edits without writing output
```

### Auto-Detection of detail.json

If `--detail` is omitted in single-template mode, `build_pptx.py` checks for a `detail.json` sibling next to the template file:

```python
guess = args.template.with_name("detail.json")
if guess.exists():
    args.detail = guess
```

---

## Capacity System: Text-Fits-the-Box

The capacity system prevents text overflow — a common failure mode in template-based PPT generation where replacement text is longer than the original placeholder.

### How Capacity Is Calculated

`compute_capacity.py` measures each text box's dimensions from the template `.pptx`:

```
box_cm (width × height) + font_size_pt → chars_per_line, max_lines
                                            ↓
                              max_chars = chars_per_line × max_lines × 1.2
                                          (includes 20% calibration tolerance)
```

`max_chars` uses **visual width units**, not raw character count:
- Chinese character = 1.0 unit
- ASCII letter/digit = 0.5 unit
- Space = 0.35 unit
- Other Unicode = 0.8 unit

This mixed-width model means "ABC" (1.5 units) and "中文" (2.0 units) are measured correctly relative to each other.

### Overflow Detection at Build Time

`check_overflow()` in `build_pptx.py`:

```python
def check_overflow(new_text, meta):
    total_vw = _visual_width(new_text)
    if total_vw <= meta["max_chars"]:
        return True, ""                    # Fits
    # Estimate needed lines
    needed = sum(ceil(visual_width(segment) / cpl) for segment in split_lines)
    if meta.get("autofit"):
        return True, "AUTOFIT ..."         # PowerPoint will shrink, but warn
    return False, f"视觉宽度 {total_vw} > 容量 {max_chars}"
```

### Behavior Modes

| Mode | Overflow detected | Behavior |
|---|---|---|
| Default | Yes | Prints warning to stderr, saves anyway |
| `--strict` | Yes | **Refuses to save** — must shorten text and retry |
| `--no-lint` | — | Skips all overflow checks |
| `autofit: true` | Yes | Soft note only — PowerPoint handles shrink-to-fit |
| `capacity_unknown: true` | — | Skipped entirely (unreliable geometry) |

### The Iron Rule

**Shorten the text, never shrink the font.** Changing font size on one heading while another at the same type-scale `level` keeps the original size creates visual inconsistency. See [Type Scale](#type-scale--heading-consistency).

---

## Type Scale & Heading Consistency

Each `detail.json` includes a `type_scale` at the top level:

```jsonc
"type_scale": {
  "description": "数字越小字号越大。同一 level 的文字必须保持同字号。",
  "levels": [
    {"level": 1, "role": "页面主标题",      "font_size_pt": 60, "example": "目录"},
    {"level": 2, "role": "章节标题/封面主标", "font_size_pt": 54, "example": "工作完成情况"},
    {"level": 3, "role": "卡片大标题",       "font_size_pt": 40, "example": "Question 1"},
    // ... down to level 14 for fine print
  ]
}
```

Every slot carries a `level` field. The invariant is:

> **All slots at the same `level` share the same `font_size_pt`. Never change any font size.**

When a heading is too long:
1. Check `max_chars`, `chars_per_line`, `max_lines` for that slot
2. Rewrite the heading to be shorter / more concise
3. Run `build_pptx.py --strict` to verify it fits
4. If it still doesn't fit, **pick a different page layout** with a wider text box

This is enforced socially (through the `guidance` text and SKILL.md rules) rather than programmatically, because `build_pptx.py` intentionally never changes font sizes.

---

## Supporting Scripts

### render_slides.py

Converts any `.pptx` to per-page PNG images for visual inspection.

```bash
python3 scripts/render_slides.py input.pptx output_dir/ --dpi 144
# Produces: output_dir/slide-01.png, slide-02.png, ...
```

Internally uses LibreOffice for PPTX→PDF conversion, then `pdftoppm` for PDF→PNG rasterization. Used for:
- Previewing template candidates (mode A template selection)
- Self-checking final output before delivery
- Mode B template introspection (user-provided templates)

### compute_capacity.py

Pre-computes capacity metadata for a template. Run once when adding a new template:

```bash
python3 scripts/compute_capacity.py templates/new-template/template.pptx \
    --output templates/new-template/detail.json
```

Measures each text box's dimensions, calculates `chars_per_line`, `max_lines`, `max_chars`, and writes them into a `text_slots[].capacity` block that gets merged into the final `detail.json`.

### apply_update.py / check_update.py

Version management — compare local `VERSION` against remote `updates.json`, incrementally download changed files.

### build_registry.py

Generates `slide-registry.json` by scanning all `templates/*/detail.json` files. Run whenever templates are added or modified:

```bash
python3 scripts/build_registry.py
```

Uses a predefined `FAMILY_MAP` to assign `style_family` to each template slug. The output `slide-registry.json` is the unified index used by the multi-source build path.

---

## Three Composition Modes

### Mode A: Built-in Template (Default)

Mode A now has two sub-paths:

#### Path A1: Single Template (traditional, backward compatible)

```
User describes need
       ↓
Read templates/INDEX.md → narrow to candidates
       ↓
Read candidate intro.md files
       ↓
AskQuestion with preview.png for top 3 matches
       ↓
User selects template
       ↓
Read detail.json → select slides by role
       ↓
Write edits.json (template_slug format) with slot_id references
       ↓
run build_pptx.py --detail ... --strict (single-template CLI)
       ↓
run render_slides.py → visual check
       ↓
Deliver output.pptx
```

#### Path A2: Cross-Template Mixing (new in v2)

This path is triggered when the user's needs span layout types that no single template covers. For example: *"年度战略复盘，需要封面、SWOT 分析、时间线、KPI 卡片、鱼骨图、结束页"* — SWOT lives in `report-massive-models`, fishbone in `report-savior`, KPI cards in `minimal-business-summary`.

```
User describes need with diverse layout types
       ↓
Extract layout_type / tags from user's content plan
       ↓
Query slide-registry.json by tags + style_family
  → python3 -c "import json; reg=json.load(open('slide-registry.json'));
     [s for s in reg['slides'] if 'swot' in s['tags']]"
       ↓
Select slides from registry results (prefer same style family)
       ↓
Choose target style from styles/ directory
       ↓
Write edits.json (multi-source format: "style" + "slides" array)
       ↓
run build_pptx.py --registry ... --styles ... --style ... --strict
       ↓
  Pipeline: Resolve(slide_id→physical) → Assemble(clone from N sources)
          → Edit(text replacement) → Apply Style(color/font remapping)
       ↓
run render_slides.py → visual check (pay attention to cross-family color harmony)
       ↓
Deliver output.pptx
```

**Concrete example — edits.json for cross-template mixing:**

User request: "2026 年度战略复盘，需要封面、SWOT 分析、时间线、4 个 KPI 卡片、鱼骨图分析、结束页"

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
    {"slide_id": "report-massive-models-swot-12", "slot_id": "swot_opportunity_title", "new_text": "AI 原生转型"},
    {"slide_id": "report-massive-models-swot-12", "slot_id": "swot_threat_title", "new_text": "竞品低价策略"},
    {"slide_id": "report-savior-timeline-22",     "slot_id": "timeline_m1_cn", "new_text": "Q1 产品重构完成"},
    {"slide_id": "report-savior-timeline-22",     "slot_id": "timeline_m2_cn", "new_text": "Q2 客户迁移启动"},
    {"slide_id": "report-savior-timeline-22",     "slot_id": "timeline_m3_cn", "new_text": "Q3 规模化上线"},
    {"slide_id": "report-savior-timeline-22",     "slot_id": "timeline_m4_cn", "new_text": "Q4 全面推广"},
    {"slide_id": "minimal-4card-5",               "slot_id": "card1_title_cn", "new_text": "营收增长 34%"},
    {"slide_id": "minimal-4card-5",               "slot_id": "card2_title_cn", "new_text": "客户留存 92%"},
    {"slide_id": "minimal-4card-5",               "slot_id": "card3_title_cn", "new_text": "NPS 净值 68"},
    {"slide_id": "minimal-4card-5",               "slot_id": "card4_title_cn", "new_text": "人效提升 2.1x"},
    {"slide_id": "report-savior-fishbone-15",     "slot_id": "fishbone_head_cn", "new_text": "交付延期根因分析"},
    {"slide_id": "minimal-ending-16",             "slot_id": "ending_thanks_cn", "new_text": "谢谢！欢迎提问"}
  ]
}
```

This single `edits.json` pulls 6 slides from 3 different templates — `minimal-business-summary` (cover, 4-card, ending), `report-massive-models` (SWOT), `report-savior` (timeline, fishbone) — all unified under the `deep-blue-minimal` style. The build pipeline automatically handles color remapping for any cross-family slides.

### Mode B: User-Provided Template

```
User provides custom .pptx
       ↓
render_slides.py → PNG previews of every page
       ↓
python-pptx walk script → shape_id/paragraph/run dump
       ↓
AI infers page roles from visual + structural data
       ↓
Write edits.json with explicit address (no slot_id)
       ↓
run build_pptx.py --strict (no --detail needed)
       ↓
Deliver output.pptx (NEVER overwrite user's original)
```

### Mode C: Original (No Template)

```
User wants clean, minimal design
       ↓
python-pptx direct API calls
       ↓
Strict rules: ≤4 elements per slide, 1 accent color,
              Arial/Helvetica or 微软雅黑/思源黑体,
              heavy whitespace, no decoration
```

---

## End-to-End Data Flow

Here is the complete sequence from user request to delivered file, annotated with the data structures at each step:

```
USER: "帮我做一份 2026 年度工作汇报 PPT"
                    │
                    ▼
┌─────────────────────────────────────────────┐
│ 1. NEEDS EXTRACTION                          │
│    Theme: 年度工作汇报                        │
│    Color: 深蓝商务                            │
│    Audience: 领导                             │
│    Length: ~15 pages                          │
└──────────────────┬──────────────────────────┘
                   ▼
┌─────────────────────────────────────────────┐
│ 2. TEMPLATE MATCHING (INDEX.md)              │
│    Matches: minimal-business-summary          │
│    Reason: 极简商务, 4章节, 季度/年度汇报首选  │
└──────────────────┬──────────────────────────┘
                   ▼
┌─────────────────────────────────────────────┐
│ 3. SLIDE SELECTION (detail.json)             │
│    cover: [1]                                 │
│    agenda: [2]                                │
│    section_dividers: [3, 7, 10, 13]          │
│    content: [4, 5, 8, 9, 11, 12, 14]         │
│    ending: [16]                               │
│    → selected_slides: [1,2,3,4,5,7,8,9,10,   │
│                        11,12,13,14,16]        │
└──────────────────┬──────────────────────────┘
                   ▼
┌─────────────────────────────────────────────┐
│ 4. CONTENT AUTHORING (edits.json)            │
│                                              │
│    For each selected slide, for each          │
│    editable=true slot:                        │
│      - Read slot.role + slot.guidance         │
│      - Write real content                     │
│      - Check content width ≤ slot.max_chars   │
│      - Ensure same-level headings match       │
│      - Sync chapter names across pages        │
│                                              │
│    Result: ~60+ edit entries                  │
└──────────────────┬──────────────────────────┘
                   ▼
┌─────────────────────────────────────────────┐
│ 5. BUILD (build_pptx.py --strict)            │
│                                              │
│    Phase 1: Load template.pptx via python-pptx│
│    Phase 2: For each edit:                    │
│      a) Resolve address from slot_id          │
│      b) Locate shape (recursive group search) │
│      c) Replace run text (preserve fmt)       │
│      d) Check overflow vs capacity            │
│    Phase 3: Prune unselected slides           │
│    Phase 4: Save output.pptx                  │
└──────────────────┬──────────────────────────┘
                   ▼
┌─────────────────────────────────────────────┐
│ 6. VERIFY (render_slides.py)                 │
│    → output PNG per page                      │
│    → Check: no placeholder text, no overflow, │
│      chapter names consistent, correct order  │
└──────────────────┬──────────────────────────┘
                   ▼
              output.pptx delivered
```

---

## Key Design Decisions

### Why slot_id Instead of Direct Shape Addressing?

Templates change — shapes get reordered, IDs shift when the designer edits the `.pptx`. The `slot_id` is a semantic identifier that remains stable across template revisions as long as the `detail.json` is regenerated. This decouples the edit specification from the physical layout.

### Why Edit Before Pruning?

Slide numbers in `edits.json` refer to the **original template** (1-based). If pruning happened first, the same slide would have a different index after removal of earlier slides. Editing pre-prune keeps the addressing simple and deterministic.

### Why Visual Width Units for Capacity?

Character-count-based limits fail for mixed Chinese/English content. "项目ABC复盘" is 6 characters but only ~4.5 visual width units — while "项目ABC复盘报告总结" is 8 characters / ~6.5 units. The visual width model correctly predicts which fits and which overflows.

### Why --strict Refuses to Save?

The alternative — silently producing a broken PPT with overflowing text — wastes the user's time. By refusing to save, `--strict` forces the AI to fix the problem at authoring time, which is when the context to rewrite concisely is still available.

### Why an Additive Registry Layer Instead of a Full Rewrite?

The 19 existing templates, their `detail.json` files, and the single-template build path all work correctly and are well-tested. The registry + styles layer is purely additive — it sits on top of the existing infrastructure without modifying any of it. This means:
- Zero migration of existing files
- Full backward compatibility for all v1 workflows
- Incremental rollout (Phase 1 covers the most-used Deep Blue templates immediately)
- The legacy single-template CLI continues to work unchanged

### Why color_map Instead of PowerPoint-Native Theme Colors?

PowerPoint's native theme color system (`<a:theme>`, `<a:schemeClr>`) would require rebuilding every template to use theme color references instead of hardcoded hex values — a massive undertaking across 640+ slides. The `color_map` approach achieves the same result (programmatic color remapping) without touching any template file: at build time, walk every shape and remap known colors to their target equivalents. The trade-off is that images and icons with embedded colors can't be remapped, but this is a minor visual inconsistency for cross-family decks compared to the cost of rebuilding all templates.

### Why Separate Styles Instead of Embedding Style in the Registry?

Styles change infrequently (when a new template family is added or a palette is refined), while the slide registry is regenerated whenever templates change. Keeping them separate avoids coupling these two update cadences. It also allows users to define custom styles without touching the registry.

---

## Adding a New Template

1. Place `template.pptx` in `templates/<new-slug>/`
2. Create `intro.md` following the existing templates' format
3. Run `compute_capacity.py` to generate initial slot data
4. Manually curate the output:
   - Assign semantic `slot_id` names
   - Set `editable: true/false` for each slot
   - Add `guidance` text in Chinese
   - Tag `decorative_slot_ids`, `title_slot_id`, `body_slot_ids`
   - Fill `page_roles`, `skip_pages`, `shape_caution_pages`
   - Define `type_scale` levels
5. Generate `preview.png` (4-page 2×2 grid)
6. Add entry to `templates/INDEX.md`
7. **If the new template belongs to a new style family**, create a corresponding `styles/<family-slug>.json` with its palette, fonts, and `color_map`
8. Run `build_registry.py` to regenerate `slide-registry.json`
9. Validate: `build_pptx.py --strict --dry-run` with a sample edits.json (both single-template and multi-source paths)
