#!/usr/bin/env python3
"""Generate slide-registry.json by scanning all template detail.json files."""
from __future__ import annotations
import json
import sys
from pathlib import Path

# Single source of truth: template slug → style family
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


def infer_layout_type(page: dict) -> str:
    """Heuristic: derive a normalized layout_type from page role and slot patterns."""
    role = page.get("role", "content")
    if role in ("cover", "agenda", "section_divider", "ending"):
        return role

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

    return f"{role}-{len(slot_ids)}slots"


def generate_tags(page: dict) -> list[str]:
    """Derive searchable tags from page metadata."""
    tags = []
    role = page.get("role", "")
    use_for = page.get("use_for", "")
    layout_type = infer_layout_type(page)

    if role:
        tags.append(role)
    tags.append(layout_type)

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
        "build_date": "",
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
