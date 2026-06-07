# Slide-Level Template Architecture — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add slide-level template composition to GordenPPTSkill — separate styles from slides, allow mixing slides across templates, with full backward compatibility.

**Architecture:** Additive Registry Layer. Three new file types (slide-registry.json, style definitions, build_registry.py) sit alongside the existing template structure. build_pptx.py gains a multi-source path but the legacy single-template path is untouched. Style application is programmatic: walk shapes at build time and remap colors/fonts via color_map.

**Tech Stack:** Python 3.13+, python-pptx, lxml (for XML-level slide cloning), existing scripts (render_slides.py, compute_capacity.py unchanged)

---

## File Structure

```
New:
  slide-registry.json                          ← auto-generated index of all slides
  styles/deep-blue-business-massive.json       ← Phase 1
  styles/deep-blue-minimal.json                ← Phase 1
  styles/deep-blue-corp-extended.json          ← Phase 2
  styles/bordeaux-corp.json                    ← Phase 2
  styles/patriotic-red.json                    ← Phase 3
  styles/warm-academic.json                    ← Phase 3
  styles/wine-academic.json                    ← Phase 3
  scripts/build_registry.py                    ← generates slide-registry.json

Modified:
  scripts/build_pptx.py                        ← add multi-source + style phases
  SKILL.md                                     ← update Mode A workflow
  references/workflow.md                       ← update Mode A details
```

---

## Phase 1: Core Infrastructure + P0 Styles

Phase 1 delivers the registry, two style definitions, and the multi-source build pipeline. After Phase 1, any slide within the `deep-blue-business-massive` family (report-massive-models, report-massive-charts, report-massive-reports — 113 slides across 3 templates) can be mixed in a single deck. The `deep-blue-minimal` style is defined but benefits from cross-family mixing in Phase 2. All legacy functionality remains intact.

### Task 1: Create `scripts/build_registry.py`

**Files:**
- Create: `scripts/build_registry.py`
- Reference: `templates/*/detail.json`

- [ ] **Step 1: Write `build_registry.py` — the registry generator**

This script scans all template `detail.json` files and produces `slide-registry.json`. The `FAMILY_MAP` is hardcoded here as the single source of truth for template-to-family assignment.

```python
#!/usr/bin/env python3
"""Generate slide-registry.json by scanning all template detail.json files."""
from __future__ import annotations
import json
import sys
from pathlib import Path

# Single source of truth: template slug → style family
# Updated as new style families are defined in Phases 2 and 3.
FAMILY_MAP: dict[str, str] = {
    # Phase 1
    "minimal-business-summary":  "deep-blue-minimal",
    "report-massive-models":     "deep-blue-business-massive",
    "report-massive-charts":     "deep-blue-business-massive",
    "report-massive-reports":    "deep-blue-business-massive",
    # Phase 2 (commented until style files exist)
    # "data-viz-deck":             "deep-blue-corp-extended",
    # "report-savior":             "deep-blue-corp-extended",
    # "operations-deck":           "deep-blue-corp-extended",
    # "competition-speech":        "deep-blue-corp-extended",
    # "architecture-deck":         "deep-blue-corp-extended",
    # "premium-corp":              "bordeaux-corp",
    # "mckinsey-style":            "bordeaux-corp",
    # Phase 3 (commented until style files exist)
    # "red-patriot-youth":         "patriotic-red",
    # "red-patriot-general":       "patriotic-red",
    # "cute-orange-class":         "warm-academic",
    # "thesis-novice":             "warm-academic",
    # "thesis-formula":            "warm-academic",
    # "top-thesis":                "wine-academic",
    # "quarterly-illust":          "standalone",
    # "geometric-summary":         "standalone",
}

# Layout type inference: mapping from page role + text patterns to normalized types
def infer_layout_type(page: dict) -> str:
    """Heuristic: derive a normalized layout_type from page role and slot patterns."""
    role = page.get("role", "content")
    if role in ("cover", "agenda", "section_divider", "ending"):
        return role

    # Check slot_id patterns for known layout types
    slot_ids = [s.get("slot_id", "") for s in page.get("text_slots", [])]
    joined = " ".join(slot_ids)

    patterns = [
        ("swot",        ["swot", "strength", "weakness", "opportunity", "threat"]),
        ("timeline",    ["timeline", "milestone", "phase"]),
        ("4-card-grid", ["card", "col1", "col2", "col3", "col4"]),
        ("4-card-grid", ["q1", "q2", "q3", "q4"]),
        ("3-col",       ["item1", "item2", "item3"]),
        ("2x2-grid",    ["block1", "block2"]),
        ("wheel",       ["wheel", "method", "center"]),
        ("pdca",        ["pdca", "plan", "do", "check", "act"]),
        ("fishbone",    ["fishbone", "cause", "effect"]),
        ("kpi-cards",   ["kpi", "metric", "indicator"]),
        ("chart",       ["chart", "bar", "line", "pie"]),
        ("comparison",  ["compare", "versus", "vs", "left", "right"]),
        ("list",        ["list", "bullet", "item"]),
    ]
    for layout, keywords in patterns:
        if any(kw in joined.lower() for kw in keywords):
            return layout

    # Fallback: use role + number of slots as type
    return f"{role}-{len(slot_ids)}slots"


def generate_tags(page: dict) -> list[str]:
    """Derive searchable tags from page metadata."""
    tags = []
    role = page.get("role", "")
    use_for = page.get("use_for", "")
    layout_type = infer_layout_type(page)

    # Role-based tags
    if role:
        tags.append(role)

    # Layout-type tag
    tags.append(layout_type)

    # Keyword extraction from use_for
    keyword_map = {
        "封面": "cover", "目录": "agenda", "致谢": "ending",
        "分析": "analysis", "对比": "comparison", "流程": "process",
        "时间": "timeline", "指标": "kpi", "数据": "data",
        "图表": "chart", "总结": "summary", "汇报": "report",
        "答辩": "defense", "教学": "teaching", "培训": "training",
        "竞聘": "competition", "述职": "review",
    }
    for cn, en in keyword_map.items():
        if cn in use_for:
            tags.append(en)

    # Deduplicate
    return list(dict.fromkeys(tags))


def build_registry(templates_dir: Path) -> dict:
    """Scan all templates/*/detail.json and build the registry."""
    slides: list[dict] = []
    detail_paths = sorted(templates_dir.glob("*/detail.json"))

    for detail_path in detail_paths:
        slug = detail_path.parent.name
        if slug == "INDEX.md":
            continue
        style_family = FAMILY_MAP.get(slug, "standalone")
        detail = json.loads(detail_path.read_text(encoding="utf-8"))
        skip_pages = set(detail.get("skip_pages", []))

        for page in detail.get("pages", []):
            slide_num = page["slide_number"]
            if slide_num in skip_pages:
                continue
            role = page.get("role", "content")
            if role == "template_promo":
                continue

            slide_id = f"{slug}-{role}-{slide_num}"
            layout_type = infer_layout_type(page)

            slides.append({
                "slide_id": slide_id,
                "source_pptx": f"templates/{slug}/template.pptx",
                "source_slide": slide_num,
                "style_family": style_family,
                "layout_type": layout_type,
                "role": role,
                "use_for": page.get("use_for", ""),
                "cautions": page.get("cautions", []),
                "tags": generate_tags(page),
                "detail_ref": f"templates/{slug}/detail.json",
                "slots_count": len(page.get("text_slots", [])),
                "text_density": page.get("text_density", "unknown"),
            })

    return {
        "$schema": "slide-registry/v1",
        "version": "1.0.0",
        "build_date": "",  # filled below
        "total_slides": len(slides),
        "slides": slides,
    }


def main() -> int:
    from datetime import date
    root = Path(__file__).resolve().parent.parent
    templates_dir = root / "templates"
    if not templates_dir.is_dir():
        print(f"ERROR: templates directory not found at {templates_dir}", file=sys.stderr)
        return 1

    registry = build_registry(templates_dir)
    registry["build_date"] = date.today().isoformat()

    output_path = root / "slide-registry.json"
    output_path.write_text(
        json.dumps(registry, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(f"Wrote {output_path} ({registry['total_slides']} slides from {len(FAMILY_MAP)} templates)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Run it to generate the initial registry**

```bash
cd /Users/spring/Dev/ai_coding/GordenPPTSkill && python3 scripts/build_registry.py
```

Expected output: `Wrote .../slide-registry.json (N slides from 4 templates)`

- [ ] **Step 3: Verify the output**

```bash
python3 -c "
import json
r = json.load(open('slide-registry.json'))
print(f'Total: {r[\"total_slides\"]} slides')
families = {}
for s in r['slides']:
    families.setdefault(s['style_family'], 0)
    families[s['style_family']] += 1
for f, c in sorted(families.items()):
    print(f'  {f}: {c} slides')
# Spot check a known slide
swot = [s for s in r['slides'] if 'swot' in s['tags']]
print(f'SWOT slides: {len(swot)}')
for s in swot:
    print(f'  {s[\"slide_id\"]} — {s[\"use_for\"][:40]}')
"
```

Expected: 4 families shown, at least some SWOT slides found.

- [ ] **Step 4: Commit**

```bash
git add scripts/build_registry.py slide-registry.json
git commit -m "feat: add build_registry.py and initial slide-registry.json (Phase 1)"
```

---

### Task 2: Create Phase 1 Style Definitions

**Files:**
- Create: `styles/deep-blue-business-massive.json`
- Create: `styles/deep-blue-minimal.json`
- Reference: `templates/report-massive-models/detail.json` (theme_colors, fonts)
- Reference: `templates/minimal-business-summary/detail.json` (theme_colors, fonts)

- [ ] **Step 1: Create styles directory**

```bash
mkdir -p /Users/spring/Dev/ai_coding/GordenPPTSkill/styles
```

- [ ] **Step 2: Create `styles/deep-blue-business-massive.json`**

Based on the actual colors from `report-massive-models/detail.json`:
- `theme_colors`: `["#1E3A5F", "#2F6EBA", "#A6B5C8"]` plus white/grays
- `fonts`: 思源黑体 / 微软雅黑

```jsonc
{
  "$schema": "ppt-style/v1",
  "slug": "deep-blue-business-massive",
  "name": "深蓝商务汇报合辑",
  "description": "深蓝主色商务风，适合工作汇报、竞聘、思维模型、数据图表等场景",
  "palette": {
    "primary":        "#1E3A5F",
    "secondary":      "#2F6EBA",
    "accent":         "#A6B5C8",
    "background":     "#FFFFFF",
    "bg_alt":         "#F4F6F9",
    "text_on_light":  "#1E293B",
    "text_on_dark":   "#FFFFFF",
    "text_muted":     "#64748B"
  },
  "fonts": {
    "cn_heading": "思源黑体",
    "cn_body":    "微软雅黑",
    "en_heading": "Arial",
    "en_body":    "Arial"
  },
  "color_map": {
    "#1E3A5F": "primary",
    "#2F6EBA": "secondary",
    "#A6B5C8": "accent",
    "#FFFFFF": "background",
    "#F4F6F9": "bg_alt",
    "#1E293B": "text_on_light",
    "#64748B": "text_muted"
  },
  "font_map": {
    "思源黑体": "cn_heading",
    "微软雅黑": "cn_body",
    "Arial":       "en_body"
  }
}
```

- [ ] **Step 3: Create `styles/deep-blue-minimal.json`**

Based on actual colors from `minimal-business-summary/detail.json`:
- `theme_colors`: `["#485275", "#44546A", "#E7E6E6"]`
- `fonts`: 微软雅黑 / Arial

```jsonc
{
  "$schema": "ppt-style/v1",
  "slug": "deep-blue-minimal",
  "name": "深蓝极简商务",
  "description": "深蓝主色 + 大量留白，极简商务风，适合季度/年度汇报",
  "palette": {
    "primary":        "#485275",
    "secondary":      "#44546A",
    "accent":         "#E7E6E6",
    "background":     "#FFFFFF",
    "bg_alt":         "#F7F8FA",
    "text_on_light":  "#334155",
    "text_on_dark":   "#F8FAFC",
    "text_muted":     "#94A3B8"
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
    "Arial":    "en_body"
  }
}
```

- [ ] **Step 4: Verify both files are valid JSON**

```bash
python3 -c "import json; json.load(open('styles/deep-blue-business-massive.json')); print('OK')"
python3 -c "import json; json.load(open('styles/deep-blue-minimal.json')); print('OK')"
```

- [ ] **Step 5: Commit**

```bash
git add styles/
git commit -m "feat: add Phase 1 style definitions (deep-blue-business-massive, deep-blue-minimal)"
```

---

### Task 3: Add Slide Cloning to `build_pptx.py`

**Files:**
- Modify: `scripts/build_pptx.py` — add `clone_slide()` and `assemble_slides()`

- [ ] **Step 1: Add the slide cloning function**

Add this function after the existing `prune_slides()` function (around line 197). This uses lxml to deep-copy a slide element, clone its relationships, and fix up IDs to avoid conflicts.

```python
# ------------------------------------------------------------------
# Slide cloning for multi-source assembly
# ------------------------------------------------------------------
import copy as _copy
from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.slide import SlidePart
from pptx.util import Emu


def _clone_slide(src_slide, dst_prs):
    """Clone a slide from a source presentation into a destination presentation.

    Returns the new slide in the destination presentation.

    This is the core primitive for multi-source assembly. It copies the
    slide XML, its relationships (images, charts, diagrams), and the
    slide layout reference.
    """
    # 1. Add a new blank slide to the destination (we'll replace its content)
    blank_layout = dst_prs.slide_layouts[6]  # blank layout — index 6 is standard
    dst_slide = dst_prs.slides.add_slide(blank_layout)

    # 2. Copy the source slide's XML content onto the destination slide
    src_elem = src_slide._element
    dst_elem = dst_slide._element

    # Remove default blank-layout shapes from destination
    sp_tree = dst_elem.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree")
    if sp_tree is not None:
        dst_elem.remove(sp_tree)

    # Deep-copy the source's shape tree (contains all shapes, groups, etc.)
    src_sp_tree = src_elem.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree")
    if src_sp_tree is not None:
        imported = _copy.deepcopy(src_sp_tree)
        dst_elem.append(imported)

    # 3. Clone relationships (images, charts, etc.)
    src_part = src_slide.part
    dst_part = dst_slide.part

    # Track rId remapping: old_rId → new_rId
    rid_map: dict[str, str] = {}

    for rel in src_part.rels.values():
        if rel.is_external:
            new_rel = dst_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            # Internal relationship — need to copy the target part
            try:
                blob = rel.target_part.blob
            except Exception:
                continue  # skip relationships we can't resolve
            new_rel = dst_part.rels.get_or_add(rel.reltype, rel.target_part)
        rid_map[rel.rId] = new_rel.rId

    # 4. Rewrite rIds in the copied XML to use new relationship IDs
    if rid_map:
        _rewrite_rids(dst_elem, rid_map)

    return dst_slide


def _rewrite_rids(element, rid_map: dict[str, str]):
    """Recursively rewrite rId attributes in copied XML to match new relationships."""
    NSMAP = {
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    }
    for attr_name in ("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed",
                       "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link",
                       "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}picture"):
        for el in element.iter():
            val = el.get(attr_name)
            if val and val in rid_map:
                el.set(attr_name, rid_map[val])


def assemble_slides(resolved_slides: list[dict], dst_prs) -> list:
    """Copy slides from multiple source .pptx files into destination presentation.

    resolved_slides: list of dicts with source_pptx, source_slide (1-based)
    Returns: list of (slide_index, style_family) tuples for style application phase
    """
    source_cache: dict[str, Presentation] = {}
    slide_meta: list[tuple[int, str]] = []

    for entry in resolved_slides:
        src_path = entry["source_pptx"]
        if src_path not in source_cache:
            source_cache[src_path] = Presentation(src_path)
        src_prs = source_cache[src_path]
        src_slide = src_prs.slides[entry["source_slide"] - 1]

        dst_slide = _clone_slide(src_slide, dst_prs)
        slide_meta.append((len(dst_prs.slides) - 1, entry["style_family"]))

    return slide_meta
```

- [ ] **Step 2: Verify the import is already available**

The file already imports `copy` (line 42) and `Presentation` (line 48). The `lxml.etree` import is new:

```bash
python3 -c "from lxml import etree; print('lxml available')"
```

If this fails: `pip install lxml` (should already be installed as a python-pptx dependency).

- [ ] **Step 3: Commit**

```bash
git add scripts/build_pptx.py
git commit -m "feat: add slide cloning for multi-source assembly"
```

---

### Task 4: Add Style Loading & Registry Resolution to `build_pptx.py`

**Files:**
- Modify: `scripts/build_pptx.py` — add `load_registry()`, `load_style()`, `resolve_slides()`, merged slot index

- [ ] **Step 1: Add registry and style loading functions**

Add after the `load_slot_index()` function (around line 74):

```python
# ------------------------------------------------------------------
# Registry loading for multi-source mode
# ------------------------------------------------------------------
def load_registry(registry_path: Path) -> dict:
    """Load slide-registry.json. Returns {slide_id: slide_entry}."""
    if not registry_path.exists():
        raise FileNotFoundError(f"Registry not found: {registry_path}")
    data = json.loads(registry_path.read_text(encoding="utf-8"))
    index: dict[str, dict] = {}
    for slide in data.get("slides", []):
        index[slide["slide_id"]] = slide
    return index


def load_style(styles_dir: Path, style_slug: str) -> dict:
    """Load a style definition from styles/<slug>.json."""
    style_path = styles_dir / f"{style_slug}.json"
    if not style_path.exists():
        raise FileNotFoundError(f"Style not found: {style_path}")
    return json.loads(style_path.read_text(encoding="utf-8"))


def load_all_styles(styles_dir: Path) -> dict[str, dict]:
    """Load all style definitions. Returns {slug: style_dict}."""
    styles: dict[str, dict] = {}
    if not styles_dir.is_dir():
        return styles
    for style_path in sorted(styles_dir.glob("*.json")):
        style = json.loads(style_path.read_text(encoding="utf-8"))
        styles[style["slug"]] = style
    return styles


def resolve_slides(edits_spec: dict, registry: dict) -> tuple[list[dict], set[str]]:
    """Resolve slide_id entries to physical locations.

    Returns: (resolved_slides, set_of_source_pptx_paths)
    Each resolved entry: {slide_id, source_pptx, source_slide, style_family, detail_ref}
    """
    resolved: list[dict] = []
    sources: set[str] = set()
    for entry in edits_spec.get("slides", []):
        sid = entry["slide_id"]
        if sid not in registry:
            raise KeyError(f"slide_id {sid!r} not found in registry")
        slide = registry[sid]
        resolved.append({
            "slide_id": sid,
            "source_pptx": slide["source_pptx"],
            "source_slide": slide["source_slide"],
            "style_family": slide["style_family"],
            "detail_ref": slide.get("detail_ref"),
        })
        sources.add(slide["source_pptx"])
    return resolved, sources


def build_merged_slot_index(resolved_slides: list[dict]) -> dict:
    """Build a merged slot index from multiple detail.json files.

    Returns: {(slide_id, slot_id): slot_metadata}
    """
    index: dict = {}
    detail_cache: dict[str, dict] = {}
    for entry in resolved_slides:
        ref = entry.get("detail_ref")
        if not ref:
            continue
        if ref not in detail_cache:
            detail_cache[ref] = json.loads(Path(ref).read_text(encoding="utf-8"))
        detail = detail_cache[ref]
        for page in detail.get("pages", []):
            if page["slide_number"] != entry["source_slide"]:
                continue
            for slot in page.get("text_slots", []):
                slot.setdefault("expected_text", slot.get("current_text"))
                index[(entry["slide_id"], slot["slot_id"])] = slot
            break
    return index
```

- [ ] **Step 2: Commit**

```bash
git add scripts/build_pptx.py
git commit -m "feat: add registry/ style loading and slide resolution functions"
```

---

### Task 5: Add Style Application to `build_pptx.py`

**Files:**
- Modify: `scripts/build_pptx.py` — add `apply_style()` and color/font remapping functions

- [ ] **Step 1: Add the color remapping engine**

Add after the new functions from Task 4:

```python
# ------------------------------------------------------------------
# Style application (color + font remapping)
# ------------------------------------------------------------------
from pptx.dml.color import RGBColor


def _hex_to_rgb(hex_str: str) -> RGBColor:
    """Convert '#1E3A5F' → RGBColor(0x1E, 0x3A, 0x5F)."""
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _rgb_to_hex(rgb) -> str:
    """Convert RGBColor → '#1E3A5F'."""
    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def build_color_map(source_style: dict, target_style: dict) -> dict[str, str]:
    """Build a hex→hex remap table from source style to target style.

    Only colors listed in source_style.color_map are candidates for remapping.
    Returns: {source_hex: target_hex}
    """
    src_cmap = source_style.get("color_map", {})
    tgt_palette = target_style.get("palette", {})
    result: dict[str, str] = {}
    for hex_val, role in src_cmap.items():
        hex_upper = hex_val.upper()
        if role in tgt_palette:
            result[hex_upper] = tgt_palette[role].upper()
        else:
            # Fallback chain: accent → secondary → primary
            for fallback in ("accent_alt", "accent", "secondary", "primary"):
                if fallback in tgt_palette and fallback != role:
                    result[hex_upper] = tgt_palette[fallback].upper()
                    break
    return result


def _collect_all_shapes(slide):
    """Recursively yield all shapes including those inside groups."""
    for shape in slide.shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _collect_all_shapes(shape)


def remap_shape_colors(shape, color_map: dict[str, str]):
    """Remap fill, line, and font colors on a shape if they match the color_map."""
    # Text runs
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if run.font.color and run.font.color.rgb:
                    current = _rgb_to_hex(run.font.color.rgb)
                    if current in color_map:
                        run.font.color.rgb = _hex_to_rgb(color_map[current])

    # Solid fill
    try:
        fill = shape.fill
        if fill.type is not None:
            # Solid fill
            if hasattr(fill, 'fore_color') and fill.fore_color.rgb:
                current = _rgb_to_hex(fill.fore_color.rgb)
                if current in color_map:
                    fill.fore_color.rgb = _hex_to_rgb(color_map[current])
            # Gradient stops
            if hasattr(fill, 'gradient_stops'):
                for stop in fill.gradient_stops:
                    if stop.color.rgb:
                        current = _rgb_to_hex(stop.color.rgb)
                        if current in color_map:
                            stop.color.rgb = _hex_to_rgb(color_map[current])
    except Exception:
        pass  # Some fill types don't support color reading

    # Line color
    try:
        if hasattr(shape, 'line') and shape.line.fill.type is not None:
            if hasattr(shape.line, 'color') and shape.line.color.rgb:
                current = _rgb_to_hex(shape.line.color.rgb)
                if current in color_map:
                    shape.line.color.rgb = _hex_to_rgb(color_map[current])
    except Exception:
        pass


def remap_shape_fonts(shape, source_font_map: dict[str, str], target_fonts: dict[str, str]):
    """Remap font names on a shape according to font_map → target_fonts."""
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            current_font = run.font.name
            if not current_font:
                continue
            # If the current font is in the font_map, find its role and remap
            role = source_font_map.get(current_font)
            if role:
                # Map role to target font
                if role == "cn_heading" or role == "cn_body":
                    new_font = target_fonts.get(role, current_font)
                    if new_font and new_font != current_font:
                        run.font.name = new_font
                elif role == "en_heading" or role == "en_body":
                    new_font = target_fonts.get(role, current_font)
                    if new_font and new_font != current_font:
                        run.font.name = new_font


def apply_style(prs, target_style_slug: str, slide_meta: list[tuple[int, str]],
                styles: dict[str, dict]):
    """Apply target style to assembled slides.

    slide_meta: list of (slide_index_in_prs, source_style_family)
    styles: {slug: style_dict} — all loaded style definitions
    """
    target = styles.get(target_style_slug)
    if target is None:
        print(f"WARN: target style {target_style_slug!r} not found, skipping style application",
              file=sys.stderr)
        return

    for slide_idx, source_family in slide_meta:
        source = styles.get(source_family)
        if source is None:
            print(f"WARN: source style {source_family!r} not found for slide {slide_idx}",
                  file=sys.stderr)
            continue
        if source["slug"] == target["slug"]:
            continue  # Same family — no remapping needed

        color_map = build_color_map(source, target)
        src_font_map = source.get("font_map", {})
        tgt_fonts = target.get("fonts", {})

        slide = prs.slides[slide_idx]
        for shape in _collect_all_shapes(slide):
            if color_map:
                remap_shape_colors(shape, color_map)
            if src_font_map and tgt_fonts:
                remap_shape_fonts(shape, src_font_map, tgt_fonts)
```

- [ ] **Step 2: Verify RGBColor import works**

```bash
python3 -c "from pptx.dml.color import RGBColor; c = RGBColor(0x1E, 0x3A, 0x5F); print(str(c))"
```

- [ ] **Step 3: Commit**

```bash
git add scripts/build_pptx.py
git commit -m "feat: add style application engine (color + font remapping)"
```

---

### Task 6: Wire Multi-Source Mode into `build_pptx.py` Main Entry

**Files:**
- Modify: `scripts/build_pptx.py` — modify `run()` and `main()` to dispatch multi-source mode

- [ ] **Step 1: Add the multi-source runner function**

Add before `run()`:

```python
def run_multi_source(edits_spec: dict, args) -> int:
    """Execute multi-source mode: resolve → assemble → edit → style."""
    registry = load_registry(args.registry)
    resolved, sources = resolve_slides(edits_spec, registry)
    target_style = edits_spec.get("style")

    # Require --registry flag
    if args.registry is None:
        print("ERROR: multi-source mode requires --registry slide-registry.json", file=sys.stderr)
        return 2

    # Phase 1: Assemble slides from source(s)
    prs = Presentation()  # Start with empty presentation (no template)

    # Set slide size from first source (or default 16:9)
    first_source = list(sources)[0]
    src_prs = Presentation(first_source)
    prs.slide_width = src_prs.slide_width
    prs.slide_height = src_prs.slide_height

    slide_meta = assemble_slides(resolved, prs)

    # Phase 2: Resolve edits (slot_id via merged detail.json index)
    slot_index = build_merged_slot_index(resolved)
    raw_edits = list(edits_spec.get("edits") or [])

    planned: list[dict] = []
    for e in raw_edits:
        slide_id = e.get("slide_id")
        if slide_id is None:
            print(f"ERROR: edit {e} missing slide_id", file=sys.stderr)
            return 2
        # Map slide_id to output slide index (order in slides array)
        output_idx = None
        for idx, entry in enumerate(resolved):
            if entry["slide_id"] == slide_id:
                output_idx = idx
                break
        if output_idx is None:
            print(f"ERROR: slide_id {slide_id!r} not in slides list", file=sys.stderr)
            return 2

        if "address" in e:
            addr = dict(e["address"])
            expected = e.get("expected_text") or addr.get("expected_text")
            meta = None
        else:
            slot_id = e.get("slot_id")
            if slot_id is None:
                print(f"ERROR: edit {e} missing both address and slot_id", file=sys.stderr)
                return 2
            key = (slide_id, slot_id)
            if key not in slot_index:
                print(f"ERROR: slot_id {slot_id!r} not found for slide {slide_id} in detail.json",
                      file=sys.stderr)
                return 2
            meta = slot_index[key]
            addr = dict(meta["address"])
            expected = e.get("expected_text") or meta.get("expected_text")
        planned.append({
            "output_slide": output_idx,
            "address": addr,
            "expected": expected,
            "new_text": e["new_text"],
            "meta": meta,
            "raw": e,
        })

    if args.dry_run:
        for p in planned:
            print(f"slide {p['output_slide']:>3} addr={p['address']} → {p['new_text']!r}")
        return 0

    # Phase 3: Apply text edits
    overflow_issues: list[str] = []
    for p in planned:
        slide = prs.slides[p["output_slide"]]
        shape_id = p["address"].get("shape_id")
        shape = find_shape(slide.shapes, shape_id)
        if shape is None:
            msg = f"slide {p['output_slide']}: shape_id {shape_id} not found"
            if args.strict:
                print(f"ERROR: {msg}", file=sys.stderr)
                return 3
            print(f"WARN:  {msg}", file=sys.stderr)
            continue
        ok, info = apply_edit_to_shape(
            shape, p["address"], p["new_text"], p["expected"], args.strict
        )
        prefix = "OK  " if ok else ("ERR " if args.strict else "WARN")
        print(f"{prefix} output slide {p['output_slide']:>3} shape {shape_id}: {info}")
        if not ok and args.strict:
            return 4

        if not args.no_lint and p["meta"]:
            fits, omsg = check_overflow(p["new_text"], p["meta"])
            if omsg and not fits:
                slot_id = p["meta"].get("slot_id", "?")
                line = (f"slide {p['output_slide']:>3} {slot_id}: 文字过长 {omsg} "
                        f"→ {p['new_text'][:24]!r}")
                overflow_issues.append(line)
                print(f"OVERFLOW {line}", file=sys.stderr)

    if overflow_issues:
        print(f"\n{len(overflow_issues)} 处文字可能出框:", file=sys.stderr)
        for line in overflow_issues:
            print(f"  - {line}", file=sys.stderr)
        if args.strict:
            print("\n--strict: 因出框风险拒绝保存。", file=sys.stderr)
            return 5

    # Phase 4: Apply style
    if target_style and args.styles:
        all_styles = load_all_styles(args.styles)
        apply_style(prs, target_style, slide_meta, all_styles)

    prs.save(args.output)
    note = f"\nSaved {args.output} with {len(resolved)} slides from {len(sources)} source(s) and {len(planned)} edits"
    if overflow_issues:
        note += f"  (⚠️ {len(overflow_issues)} 处可能出框，见上)"
    print(note)
    return 0
```

- [ ] **Step 2: Modify `main()` to detect and dispatch multi-source mode**

Update the `main()` function (around line 303) to add new CLI flags and detect the mode:

```python
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("arg1", type=Path, help="template.pptx (legacy) or edits.json (multi-source)")
    ap.add_argument("arg2", type=Path, nargs="?", default=None,
                    help="edits.json (legacy) or output.pptx (multi-source)")
    ap.add_argument("arg3", type=Path, nargs="?", default=None,
                    help="output.pptx (legacy only)")
    ap.add_argument("--detail", type=Path, default=None)
    ap.add_argument("--registry", type=Path, default=None,
                    help="slide-registry.json (required for multi-source mode)")
    ap.add_argument("--styles", type=Path, default=None,
                    help="styles/ directory (for multi-source style application)")
    ap.add_argument("--no-lint", action="store_true")
    ap.add_argument("--strict", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    # Auto-detect mode from the first positional argument
    # If it's an edits.json with a "style" field → multi-source mode
    # Otherwise → legacy mode (template.pptx is the first arg)
    if args.arg1.suffix == ".json":
        spec = json.loads(args.arg1.read_text(encoding="utf-8"))
        if "style" in spec:
            # Multi-source mode: arg1=edits.json, arg2=output.pptx
            if args.arg2 is None:
                print("ERROR: multi-source mode requires output path as second argument",
                      file=sys.stderr)
                return 2
            args.edits = args.arg1
            args.output = args.arg2
            return run_multi_source(spec, args)

    # Legacy mode: arg1=template.pptx, arg2=edits.json, arg3=output.pptx
    args.template = args.arg1
    if args.arg2 is None or args.arg3 is None:
        print("ERROR: legacy mode requires: template.pptx edits.json output.pptx",
              file=sys.stderr)
        return 2
    args.edits = args.arg2
    args.output = args.arg3

    if args.detail is None:
        guess = args.template.with_name("detail.json")
        if guess.exists():
            args.detail = guess
            print(f"(auto-detected detail.json: {guess})", file=sys.stderr)

    return run(args)
```

- [ ] **Step 3: Test legacy mode still works**

```bash
cd /Users/spring/Dev/ai_coding/GordenPPTSkill
python3 scripts/build_pptx.py \
    templates/minimal-business-summary/template.pptx \
    /dev/null \
    /tmp/test-legacy.pptx \
    --detail templates/minimal-business-summary/detail.json \
    --dry-run 2>&1 || true
# Should not crash with new argument parsing
```

- [ ] **Step 4: Test multi-source mode with a simple case**

Create a test `edits.json` that uses two slides from the same family:

```bash
cat > /tmp/test-multi.json << 'EOF'
{
  "style": "deep-blue-business-massive",
  "slides": [
    {"slide_id": "report-massive-models-content-3"},
    {"slide_id": "report-massive-models-content-5"}
  ],
  "edits": []
}
EOF

python3 scripts/build_pptx.py /tmp/test-multi.json /tmp/test-multi.pptx \
    --registry slide-registry.json \
    --styles styles/ \
    --dry-run
```

- [ ] **Step 5: Commit**

```bash
git add scripts/build_pptx.py
git commit -m "feat: wire multi-source mode into build_pptx.py main entry"
```

---

### Task 7: Update SKILL.md and workflow.md for Mode A

**Files:**
- Modify: `SKILL.md` — update Mode A workflow description
- Modify: `references/workflow.md` — update Mode A detailed steps

- [ ] **Step 1: Update Mode A in SKILL.md**

Replace the Mode A workflow block (lines 128-159 in current file) with the updated version that includes both legacy and multi-source paths:

The change is additive — add a new subsection "模式 A-New: 跨模板自由选页" after the existing Mode A block, and note that the legacy path still works.

Add after the existing Mode A command block (around line 159):

```markdown
### 模式 A-New: 跨模板自由选页（v2.0+）

当用户需求涉及多种版式（如 SWOT + 时间线 + 卡片网格），且它们分布在不同的模板里时：

1. **确定目标 style** — 从 `styles/` 目录选择一个 style family。Phase 1 支持：
   - `deep-blue-business-massive`（report-massive-* 三套，113 页）
   - `deep-blue-minimal`（minimal-business-summary，16 页）

2. **查 slide-registry.json 选页** — 按 `layout_type` / `tags` / `use_for` 字段搜索匹配的页：
   ```bash
   python3 -c "
   import json
   r = json.load(open('slide-registry.json'))
   for s in r['slides']:
       if 'swot' in s['tags'] and s['style_family'] == 'deep-blue-business-massive':
           print(f\"{s['slide_id']}: {s['use_for']}\")
   "
   ```

3. **写 multi-source edits.json**（见 `references/pptx-edit-schema.md` 的新格式）

4. **构建**：
   ```bash
   python3 scripts/build_pptx.py edits.json out/final.pptx \
       --registry slide-registry.json \
       --styles styles/ \
       --strict
   ```

5. **预览**（同原有流程）：
   ```bash
   python3 scripts/render_slides.py out/final.pptx out/renders --dpi 144
   ```

**跨 family 混用**：如果选了不同 style_family 的页，build_pptx.py 会自动用 color_map 做颜色映射。Phase 1 暂不支持（deep-blue-minimal 和 deep-blue-business-massive 是独立 family），Phase 2 起启用完整的跨 family 映射。
```

- [ ] **Step 2: Update references/workflow.md**

Add a new section "## <a id="mode-a-multi"></a>模式 A-New: 跨模板选页详解" after the existing Mode A section (around line 137 before the Mode B heading). Include the registry query examples and the new edits.json format reference.

- [ ] **Step 3: Commit**

```bash
git add SKILL.md references/workflow.md
git commit -m "docs: update Mode A workflow for multi-source slide selection (Phase 1)"
```

---

### Task 8: End-to-End Integration Test

**Files:**
- Create: (temporary) `/tmp/test-phase1.json` — test edits.json
- Verify: output file renders correctly

- [ ] **Step 1: Create a real multi-source test**

Pick 5 slides from the report-massive-* family: a cover-like page, a SWOT, a timeline, a 4-card grid, and a content page. Look up actual slide_ids from the registry:

```bash
cd /Users/spring/Dev/ai_coding/GordenPPTSkill

# Find candidate slides
python3 -c "
import json
r = json.load(open('slide-registry.json'))
# Show all slides in deep-blue-business-massive family, grouped by layout_type
from collections import defaultdict
by_type = defaultdict(list)
for s in r['slides']:
    if s['style_family'] == 'deep-blue-business-massive':
        by_type[s['layout_type']].append(s['slide_id'])
for lt, ids in sorted(by_type.items()):
    print(f'{lt}: {len(ids)} slides — {ids[:3]}...')
"
```

- [ ] **Step 2: Write a test edits.json using real slide_ids**

Pick 4 slides from the output above (at least one with editable slots) and create a test file:

```bash
cat > /tmp/test-phase1.json << 'EOF'
{
  "style": "deep-blue-business-massive",
  "slides": [
    {"slide_id": "REPLACE_WITH_ACTUAL_ID_1"},
    {"slide_id": "REPLACE_WITH_ACTUAL_ID_2"},
    {"slide_id": "REPLACE_WITH_ACTUAL_ID_3"},
    {"slide_id": "REPLACE_WITH_ACTUAL_ID_4"}
  ],
  "edits": []
}
EOF
```

Then run with `--dry-run` first, then without:

```bash
python3 scripts/build_pptx.py /tmp/test-phase1.json /tmp/test-phase1.pptx \
    --registry slide-registry.json --styles styles/ --dry-run

python3 scripts/build_pptx.py /tmp/test-phase1.json /tmp/test-phase1.pptx \
    --registry slide-registry.json --styles styles/
```

- [ ] **Step 3: Verify the output**

```bash
# Check it's a valid pptx
python3 -c "
from pptx import Presentation
p = Presentation('/tmp/test-phase1.pptx')
print(f'Slides: {len(p.slides)}')
for i, s in enumerate(p.slides):
    shapes = len(s.shapes)
    print(f'  Slide {i+1}: {shapes} shapes')
"

# Check file size is reasonable
ls -lh /tmp/test-phase1.pptx
```

- [ ] **Step 4: Render to verify visually**

```bash
python3 scripts/render_slides.py /tmp/test-phase1.pptx /tmp/test-phase1-renders --dpi 72
ls /tmp/test-phase1-renders/
```

- [ ] **Step 5: Test backward compatibility — legacy mode still works**

```bash
# Create a minimal valid edits.json for legacy mode
cat > /tmp/test-legacy.json << 'EOF'
{
  "template_slug": "minimal-business-summary",
  "selected_slides": [1, 16],
  "edits": [
    {"slide": 1, "slot_id": "cover_title_cn", "new_text": "兼容性测试"}
  ]
}
EOF

python3 scripts/build_pptx.py \
    templates/minimal-business-summary/template.pptx \
    /tmp/test-legacy.json \
    /tmp/test-legacy.pptx \
    --detail templates/minimal-business-summary/detail.json \
    --strict
echo "Exit code: $?"
```

- [ ] **Step 6: Commit if all passes, or fix issues and re-test**

```bash
git add -A
git commit -m "test: add end-to-end integration test for Phase 1 multi-source build"
```

---

## Phase 2: P1 Styles + Cross-Family Remapping

### Task 9: Add P1 Style Definitions

**Files:**
- Create: `styles/deep-blue-corp-extended.json`
- Create: `styles/bordeaux-corp.json`

- [ ] **Step 1: Create `styles/deep-blue-corp-extended.json`**

Based on data-viz-deck (#2C3E70), report-savior (#1F3A93), operations-deck (#1F3A93), competition-speech (#1B3464), architecture-deck (#1F3A93). These share a deep blue primary with slight variations. Use the most common primary `#1F3A93` and include the variant blues in color_map.

```jsonc
{
  "$schema": "ppt-style/v1",
  "slug": "deep-blue-corp-extended",
  "name": "深蓝企业扩展",
  "description": "深蓝企业风，覆盖数据可视化、运营、竞聘、架构图等场景",
  "palette": {
    "primary":        "#1F3A93",
    "secondary":      "#3498DB",
    "accent":         "#C0392B",
    "background":     "#FFFFFF",
    "bg_alt":         "#ECF0F1",
    "text_on_light":  "#1E293B",
    "text_on_dark":   "#FFFFFF",
    "text_muted":     "#64748B"
  },
  "fonts": {
    "cn_heading": "思源黑体",
    "cn_body":    "微软雅黑",
    "en_heading": "Arial",
    "en_body":    "Arial"
  },
  "color_map": {
    "#1F3A93": "primary",
    "#1B3464": "primary",
    "#2C3E70": "primary",
    "#2D4F8A": "primary",
    "#3498DB": "secondary",
    "#3F6FCC": "secondary",
    "#C0392B": "accent",
    "#B73E3E": "accent",
    "#FFFFFF": "background",
    "#ECF0F1": "bg_alt",
    "#E1ECF7": "bg_alt",
    "#E8E8E8": "bg_alt",
    "#34495E": "text_on_light"
  },
  "font_map": {
    "思源黑体": "cn_heading",
    "微软雅黑": "cn_body",
    "Arial":    "en_body"
  }
}
```

- [ ] **Step 2: Create `styles/bordeaux-corp.json`**

Based on premium-corp and mckinsey-style — both use #A52524 primary and #2D3955 secondary.

```jsonc
{
  "$schema": "ppt-style/v1",
  "slug": "bordeaux-corp",
  "name": "酱红企业风",
  "description": "酱红主色 + 深蓝灰辅色，高级感大厂风，适合咨询、战略分析",
  "palette": {
    "primary":        "#A52524",
    "secondary":      "#2D3955",
    "accent":         "#D14D4A",
    "background":     "#FFFFFF",
    "bg_alt":         "#F5F0EB",
    "text_on_light":  "#1E293B",
    "text_on_dark":   "#FFFFFF",
    "text_muted":     "#78716C"
  },
  "fonts": {
    "cn_heading": "思源黑体",
    "cn_body":    "微软雅黑",
    "en_heading": "Arial",
    "en_body":    "Arial"
  },
  "color_map": {
    "#A52524": "primary",
    "#2D3955": "secondary",
    "#D14D4A": "accent",
    "#FFFFFF": "background",
    "#1E293B": "text_on_light"
  },
  "font_map": {
    "思源黑体": "cn_heading",
    "微软雅黑": "cn_body",
    "Arial":    "en_body"
  }
}
```

- [ ] **Step 3: Uncomment Phase 2 entries in build_registry.py FAMILY_MAP**

```python
# Uncomment these lines in FAMILY_MAP:
"data-viz-deck":             "deep-blue-corp-extended",
"report-savior":             "deep-blue-corp-extended",
"operations-deck":           "deep-blue-corp-extended",
"competition-speech":        "deep-blue-corp-extended",
"architecture-deck":         "deep-blue-corp-extended",
"premium-corp":              "bordeaux-corp",
"mckinsey-style":            "bordeaux-corp",
```

- [ ] **Step 4: Regenerate the registry**

```bash
python3 scripts/build_registry.py
```

- [ ] **Step 5: Test cross-family mixing (deep-blue-minimal → deep-blue-business-massive)**

Create a test with slides from both families and verify color remapping:

```bash
python3 scripts/build_pptx.py /tmp/test-cross-family.json /tmp/test-cross.pptx \
    --registry slide-registry.json --styles styles/ --strict
```

- [ ] **Step 6: Commit**

```bash
git add styles/ scripts/build_registry.py slide-registry.json
git commit -m "feat: add Phase 2 style definitions and uncomment P1 registry entries"
```

---

## Phase 3: Remaining Styles + Polish

### Task 10: Add P2/P3 Style Definitions

**Files:**
- Create: `styles/patriotic-red.json`
- Create: `styles/warm-academic.json`
- Create: `styles/wine-academic.json`
- Modify: `scripts/build_registry.py` — uncomment remaining FAMILY_MAP entries

- [ ] **Step 1: Create all three style files**

Each follows the same `ppt-style/v1` schema. Extract colors and fonts from the respective template detail.json files (see the table in Task 1's exploration output).

**`styles/patriotic-red.json`:**
```jsonc
{
  "$schema": "ppt-style/v1",
  "slug": "patriotic-red",
  "name": "党政红",
  "description": "红色主调 + 金色点缀，庄重党政风，适合党课、爱国教育、思政课件",
  "palette": {
    "primary":        "#A91F1F",
    "secondary":      "#F4D9B4",
    "accent":         "#F2C656",
    "background":     "#FFFFFF",
    "bg_alt":         "#FFFDF7",
    "text_on_light":  "#1A1A1A",
    "text_on_dark":   "#FFFFFF",
    "text_muted":     "#8B7355"
  },
  "fonts": {
    "cn_heading": "微软雅黑",
    "cn_body":    "微软雅黑",
    "en_heading": "Arial",
    "en_body":    "Arial"
  },
  "color_map": {
    "#A91F1F": "primary",
    "#A8181C": "primary",
    "#F4D9B4": "secondary",
    "#F2C656": "accent",
    "#FFFFFF": "background",
    "#FFFDF7": "bg_alt"
  },
  "font_map": {
    "微软雅黑": "cn_body",
    "Arial":    "en_body"
  }
}
```

**`styles/warm-academic.json`:**
```jsonc
{
  "$schema": "ppt-style/v1",
  "slug": "warm-academic",
  "name": "暖色学术",
  "description": "暖色系学术风，覆盖卡通教学、开题报告、方法论等场景",
  "palette": {
    "primary":        "#4F6E4F",
    "secondary":      "#F5C97E",
    "accent":         "#C09548",
    "background":     "#FFFFFF",
    "bg_alt":         "#F6F0DC",
    "text_on_light":  "#1E293B",
    "text_on_dark":   "#FFFFFF",
    "text_muted":     "#64748B"
  },
  "fonts": {
    "cn_heading": "思源黑体",
    "cn_body":    "微软雅黑",
    "en_heading": "Arial",
    "en_body":    "Arial"
  },
  "color_map": {
    "#4F6E4F": "primary",
    "#F5C97E": "secondary",
    "#C09548": "accent",
    "#FFFFFF": "background",
    "#F6F0DC": "bg_alt",
    "#2A3F75": "primary"
  },
  "font_map": {
    "思源黑体": "cn_heading",
    "微软雅黑": "cn_body",
    "Arial":    "en_body"
  }
}
```

**`styles/wine-academic.json`:**
```jsonc
{
  "$schema": "ppt-style/v1",
  "slug": "wine-academic",
  "name": "酒红学术",
  "description": "酒红学术风，适合开题答辩、课题汇报",
  "palette": {
    "primary":        "#7A2B22",
    "secondary":      "#A55248",
    "accent":         "#D4A574",
    "background":     "#FAF1E4",
    "bg_alt":         "#FFFFFF",
    "text_on_light":  "#1E293B",
    "text_on_dark":   "#FFFFFF",
    "text_muted":     "#78716C"
  },
  "fonts": {
    "cn_heading": "思源黑体",
    "cn_body":    "微软雅黑",
    "en_heading": "Arial",
    "en_body":    "Arial"
  },
  "color_map": {
    "#7A2B22": "primary",
    "#A55248": "secondary",
    "#FAF1E4": "background",
    "#1E293B": "text_on_light"
  },
  "font_map": {
    "思源黑体": "cn_heading",
    "微软雅黑": "cn_body",
    "Arial":    "en_body"
  }
}
```

- [ ] **Step 2: Uncomment remaining FAMILY_MAP entries in build_registry.py**

```python
"red-patriot-youth":         "patriotic-red",
"red-patriot-general":       "patriotic-red",
"cute-orange-class":         "warm-academic",
"thesis-novice":             "warm-academic",
"thesis-formula":            "warm-academic",
"top-thesis":                "wine-academic",
"quarterly-illust":          "standalone",
"geometric-summary":         "standalone",
```

- [ ] **Step 3: Regenerate and verify full registry**

```bash
python3 scripts/build_registry.py
python3 -c "
import json
r = json.load(open('slide-registry.json'))
print(f'Total slides: {r[\"total_slides\"]}')
from collections import Counter
c = Counter(s['style_family'] for s in r['slides'])
for f, n in sorted(c.items()):
    print(f'  {f}: {n}')
"
```

Expected: 7-8 families shown, ~600+ total slides.

- [ ] **Step 4: Commit**

```bash
git add styles/ scripts/build_registry.py slide-registry.json
git commit -m "feat: add Phase 3 style definitions, complete registry with all 19 templates"
```

---

### Task 11: Final Cleanup and Documentation Polish

**Files:**
- Modify: `SKILL.md` — update for full Phase 3 capability
- Modify: `references/workflow.md` — final Mode A-New section
- Modify: `templates/INDEX.md` — add style family column (optional)

- [ ] **Step 1: Final SKILL.md Mode A update**

Update the Mode A-New section to list all 7 style families and their coverage. Remove the "Phase 1 limitation" note.

- [ ] **Step 2: Add `.superpowers/` to `.gitignore`**

```bash
echo ".superpowers/" >> .gitignore
```

- [ ] **Step 3: Final commit**

```bash
git add SKILL.md references/workflow.md .gitignore
git commit -m "docs: finalize multi-source documentation for all 7 style families"
```

---

## Summary

| Phase | Tasks | New Files | Modified Files | Slides Unlocked |
|---|---|---|---|---|
| 1 | 1-8 | slide-registry.json, 2 styles, build_registry.py | build_pptx.py, SKILL.md, workflow.md | 129 (4 templates) |
| 2 | 9 | 2 styles | build_registry.py, slide-registry.json | 434 (11 templates) |
| 3 | 10-11 | 3 styles | build_registry.py, slide-registry.json, docs | 642 (19 templates) |

**Backward compatibility:** The legacy `build_pptx.py template.pptx edits.json out.pptx --detail ... --strict` path is preserved unchanged. The new multi-source path is activated only when `edits.json` contains a `"style"` field.
