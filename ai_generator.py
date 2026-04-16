import json
import os
import re
import time
from typing import Dict, List

from groq import Groq

ALLOWED_THEMES = {"modern", "corporate", "dark", "academic"}
ALLOWED_SLIDE_TYPES = {
    "title",
    "agenda",
    "section",
    "bullet",
    "two-column",
    "comparison",
    "chart",
    "timeline",
    "conclusion",
    "summary",
    "thank-you",
}

THEME_DEFAULT_ACCENTS = {
    "modern": "#4F46E5",
    "corporate": "#1E3A8A",
    "dark": "#22D3EE",
    "academic": "#0F766E",
}

ICON_BY_TYPE = {
    "title": "rocket",
    "agenda": "clipboard",
    "section": "spark",
    "bullet": "check",
    "two-column": "columns",
    "comparison": "scale",
    "chart": "chart",
    "timeline": "timeline",
    "conclusion": "target",
    "summary": "target",
    "thank-you": "handshake",
}

THEME_ACCENT_SEQUENCES = {
    "modern": ["#4F46E5", "#7C3AED", "#0EA5E9", "#14B8A6", "#F59E0B"],
    "corporate": ["#1E3A8A", "#1D4ED8", "#0EA5E9", "#0369A1", "#334155"],
    "dark": ["#22D3EE", "#0EA5E9", "#A78BFA", "#38BDF8", "#14B8A6"],
    "academic": ["#0F766E", "#14B8A6", "#0891B2", "#2563EB", "#D97706"],
}

DOMAIN_ICON_HINTS = [
    (["strategy", "strategic", "plan", "positioning", "framework"], "target"),
    (["growth", "scale", "traction", "pipeline"], "chart"),
    (["business", "sales", "revenue", "market", "finance"], "briefcase"),
    (["education", "school", "college", "learning", "student", "course", "teaching"], "book"),
    (["technology", "software", "ai", "automation", "data", "cloud", "engineering", "platform"], "ai"),
    (["startup", "pitch", "investor", "funding"], "rocket"),
    (["operations", "plan", "roadmap", "timeline", "process"], "timeline"),
]

PLACEHOLDER_TERMS = [
    "visual idea",
    "image ideas",
    "use this space",
    "placeholder",
    "insert image",
    "add icon",
]

RETRYABLE_ERROR_TERMS = [
    "503",
    "unavailable",
    "timeout",
    "timed out",
    "service unavailable",
    "server error",
    "connection reset",
]


def build_prompt(topic: str, description: str, slides: int, tone: str, audience: str, theme: str, presenter_name: str = "") -> str:
    safe_theme = _normalize_theme(theme)
    safe_presenter = _normalize_string(presenter_name)
    base_prompt = (
        "Generate a premium business presentation JSON for PowerPoint. "
        f"Topic: {topic}. Tone: {tone}. Audience: {audience}. Requested slides: {slides}. Theme: {safe_theme}. "
        f"Presenter name: {safe_presenter}. "
        "You are acting as a presentation designer, UX designer, visual designer, and content strategist. "
        "Output valid JSON only with no markdown and no extra text. "
        "Do not use placeholder text like 'Visual idea', 'Image ideas', or 'Use this space'. "
        "Never return empty content slides: bullets/points/milestones/chart data must be populated. "
        "Keep each bullet concise: max 5 bullets per slide and max 12 words per bullet. "
        "For chart slides, always return 3-5 numeric points between 20 and 95. "
        "Smart layout engine rules: title->centered, bullet->one-column, comparison->comparison, process->timeline, data-heavy->chart. "
        "Use slide types intelligently from: title, agenda, section, bullet, two-column, comparison, chart, timeline, conclusion, thank-you. "
        "Ensure at least one visual slide among chart/timeline/two-column/comparison for decks with 6+ slides. "
        "Maintain design consistency for typography, spacing, icon style, and accent hierarchy across all slides. "
        "Every slide must include design metadata with accent color, icon, background type, and shape. "
        "Return this exact schema:\n"
        "{\n"
        "  \"title\": \"Presentation Title\",\n"
        "  \"theme\": \"modern|corporate|dark|academic\",\n"
        "  \"slides\": [\n"
        "    {\n"
        "      \"slide_type\": \"title|agenda|section|bullet|two-column|comparison|chart|timeline|conclusion|thank-you\",\n"
        "      \"title\": \"...\",\n"
        "      \"subtitle\": \"...\",\n"
        "      \"bullets\": [\"...\"],\n"
        "      \"left_title\": \"...\",\n"
        "      \"left_points\": [\"...\"],\n"
        "      \"right_title\": \"...\",\n"
        "      \"right_points\": [\"...\"],\n"
        "      \"milestones\": [\"...\"],\n"
        "      \"chart_title\": \"...\",\n"
        "      \"chart_type\": \"bar|pie|trend\",\n"
        "      \"chart_points\": [{\"label\": \"...\", \"value\": 40}],\n"
        "      \"layout\": \"centered|one-column|two-column|comparison|timeline|chart\",\n"
        "      \"design\": {\"background\": \"solid|gradient\", \"accent_color\": \"#RRGGBB\", \"icon\": \"...\", \"shape\": \"rounded\"}\n"
        "    }\n"
        "  ]\n"
        "}\n"
    )
    clean_description = _normalize_string(description)
    if clean_description:
        return f"{base_prompt}Additional context: {clean_description}"
    return base_prompt


def extract_response_text(response) -> str:
    if hasattr(response, "choices"):
        try:
            choice = response.choices[0]
            if hasattr(choice, "message") and hasattr(choice.message, "content"):
                return str(choice.message.content).strip()
            if isinstance(choice, dict):
                message = choice.get("message")
                if isinstance(message, dict):
                    return str(message.get("content", "")).strip()
        except Exception:
            pass

    if hasattr(response, "output"):
        output = response.output
        if isinstance(output, list) and output:
            first = output[0]
            if isinstance(first, dict):
                for fragment in first.get("content", []):
                    if fragment.get("type") == "output_text":
                        return fragment.get("text", "").strip()
                if "text" in first:
                    return first.get("text", "").strip()
            elif isinstance(first, str):
                return first.strip()

    if hasattr(response, "last"):
        return str(response.last).strip()

    return str(response).strip()


def _normalize_string(value, default: str = "") -> str:
    if isinstance(value, str):
        return value.strip()
    if value is None:
        return default
    if isinstance(value, (int, float, bool)):
        return str(value)
    if isinstance(value, list):
        parts = [str(item).strip() for item in value if str(item).strip()]
        return ", ".join(parts) if parts else default
    if isinstance(value, dict):
        for candidate in value.values():
            text = _normalize_string(candidate)
            if text:
                return text
        return default
    return str(value).strip()


def _normalize_theme(value: str) -> str:
    theme = _normalize_string(value, default="modern").lower()
    return theme if theme in ALLOWED_THEMES else "modern"


def _accent_for_index(theme: str, index: int) -> str:
    safe_theme = _normalize_theme(theme)
    accents = THEME_ACCENT_SEQUENCES.get(safe_theme) or [THEME_DEFAULT_ACCENTS[safe_theme]]
    if not accents:
        return THEME_DEFAULT_ACCENTS[safe_theme]
    return accents[max(index - 1, 0) % len(accents)]


def _sanitize_accent_color(value: str, fallback: str) -> str:
    candidate = _normalize_string(value)
    if not candidate.startswith("#") or len(candidate) not in {4, 7}:
        return fallback

    clean = candidate.lstrip("#")
    if len(clean) == 3:
        clean = "".join(ch * 2 for ch in clean)

    try:
        r = int(clean[0:2], 16)
        g = int(clean[2:4], 16)
        b = int(clean[4:6], 16)
    except Exception:
        return fallback

    # Very light accents look almost invisible in chart and badge surfaces.
    if (r + g + b) > 700:
        return fallback
    return f"#{clean.upper()}"


def _icon_from_context(slide_type: str, title: str, topic: str, audience: str) -> str:
    fallback = ICON_BY_TYPE.get(slide_type, "spark")
    context = f"{title} {topic} {audience}".lower()
    for keywords, icon in DOMAIN_ICON_HINTS:
        if any(word in context for word in keywords):
            return icon
    return fallback


def _normalize_chart_points(value) -> List[Dict]:
    if isinstance(value, list):
        raw_points = value
    elif isinstance(value, dict):
        raw_points = [{"label": key, "value": val} for key, val in value.items()]
    elif isinstance(value, str):
        parts = [part.strip() for part in value.replace("\n", ";").split(";") if part.strip()]
        raw_points = parts
    elif value is None:
        raw_points = []
    else:
        raw_points = [value]

    points = []
    for index, item in enumerate(raw_points, start=1):
        label = ""
        numeric_value = None
        if isinstance(item, dict):
            label = _clean_text(_normalize_string(item.get("label") or item.get("name")))
            numeric_raw = item.get("value")
            if numeric_raw is None:
                numeric_raw = item.get("score")
            if numeric_raw is None:
                numeric_raw = item.get("amount")
            if numeric_raw is None:
                numeric_raw = item.get("percent")

            numeric_match = re.search(r"-?\d+(?:\.\d+)?", str(numeric_raw)) if numeric_raw is not None else None
            if numeric_match:
                try:
                    numeric_value = int(float(numeric_match.group(0)))
                except Exception:
                    numeric_value = None
        else:
            text = _clean_text(_normalize_string(item))
            if not text:
                continue
            if ":" in text:
                left, right = text.split(":", 1)
                label = _clean_text(left)
                numeric_match = re.search(r"-?\d+(?:\.\d+)?", right)
                if numeric_match:
                    try:
                        numeric_value = int(float(numeric_match.group(0)))
                    except Exception:
                        numeric_value = None
            else:
                label = text

        if not label:
            continue
        if numeric_value is None:
            numeric_value = 24 + index * 13
        numeric_value = max(min(numeric_value, 100), 10)
        points.append({"label": _limit_words(label, max_words=4), "value": numeric_value})
        if len(points) >= 5:
            break

    if points and len({point["value"] for point in points}) == 1:
        for idx, point in enumerate(points, start=1):
            point["value"] = max(min(point["value"] + idx * 4, 100), 10)

    return points


def _design_checklist_for_slide(slide_type: str) -> Dict:
    return {
        "readability": "pass",
        "visual_balance": "pass",
        "spacing": "pass",
        "typography": "pass",
        "recommended_motion": "subtle fade" if slide_type in {"title", "section", "agenda"} else "subtle rise",
    }


def _clean_text(text: str) -> str:
    cleaned = _normalize_string(text)
    lowered = cleaned.lower()
    if any(term in lowered for term in PLACEHOLDER_TERMS):
        return ""
    cleaned = cleaned.replace("\n", " ").replace("  ", " ").strip(" -•\t")
    return cleaned


def _limit_words(text: str, max_words: int = 12) -> str:
    words = _clean_text(text).split()
    if not words:
        return ""
    if len(words) <= max_words:
        return " ".join(words)
    return " ".join(words[:max_words]).rstrip(".,;:")


def _normalize_list(value) -> List[str]:
    if isinstance(value, list):
        raw_items = value
    elif isinstance(value, str):
        normalized_text = value.replace("\n", ";")
        if ";" in normalized_text:
            raw_items = [part.strip() for part in normalized_text.split(";") if part.strip()]
        else:
            sentence_candidates = [segment.strip() for segment in normalized_text.split(".") if segment.strip()]
            raw_items = sentence_candidates if sentence_candidates else [normalized_text.strip()]
    elif isinstance(value, dict):
        raw_items = list(value.values())
    elif value is None:
        raw_items = []
    else:
        raw_items = [str(value)]

    normalized = []
    seen = set()
    for item in raw_items:
        text = _limit_words(_normalize_string(item))
        if not text:
            continue
        key = text.lower()
        if key in seen:
            continue
        seen.add(key)
        normalized.append(text)
    return normalized


def _compress_bullets(items: List[str], max_bullets: int = 5) -> List[str]:
    compressed = []
    for item in items:
        text = _limit_words(item)
        if text:
            compressed.append(text)
        if len(compressed) >= max_bullets:
            break
    return compressed


def _extract_json_payload(raw_text: str) -> Dict:
    if not isinstance(raw_text, str):
        return {}

    candidate_texts = [raw_text.strip()]

    if "```" in raw_text:
        segments = raw_text.split("```")
        for segment in segments:
            cleaned = segment.replace("json", "", 1).strip()
            if cleaned:
                candidate_texts.append(cleaned)

    start = raw_text.find("{")
    end = raw_text.rfind("}")
    if start != -1 and end != -1 and end > start:
        candidate_texts.append(raw_text[start : end + 1].strip())

    for candidate in candidate_texts:
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict):
                return parsed
            if isinstance(parsed, list) and parsed and isinstance(parsed[0], dict):
                return parsed[0]
        except json.JSONDecodeError:
            continue

    return {}


def _default_layout(slide_type: str) -> str:
    if slide_type in {"title", "section", "agenda", "thank-you"}:
        return "centered"
    if slide_type == "two-column":
        return "two-column"
    if slide_type == "comparison":
        return "comparison"
    if slide_type == "chart":
        return "chart"
    if slide_type == "timeline":
        return "timeline"
    return "one-column"


def _normalize_design(design, theme: str, slide_type: str, index: int, title: str, topic: str, audience: str) -> Dict:
    safe_theme = _normalize_theme(theme)
    if not isinstance(design, dict):
        design = {}

    background_default = "gradient" if slide_type in {"title", "section", "agenda", "thank-you"} else "solid"
    background = _normalize_string(design.get("background"), default=background_default).lower()
    if background not in {"solid", "gradient"}:
        background = background_default

    accent_color = _sanitize_accent_color(
        _normalize_string(design.get("accent_color"), default=_accent_for_index(safe_theme, index)),
        _accent_for_index(safe_theme, index),
    )

    icon = _normalize_string(design.get("icon"), default=_icon_from_context(slide_type, title, topic, audience)).lower()
    shape = _normalize_string(design.get("shape"), default="rounded").lower()
    visual_priority = _normalize_string(design.get("visual_priority"), default="balanced").lower()
    if visual_priority not in {"balanced", "title-first", "data-first", "action-first"}:
        visual_priority = "balanced"

    return {
        "background": background,
        "accent_color": accent_color,
        "icon": icon,
        "shape": shape,
        "typography": {
            "title_size": 42 if slide_type == "title" else 34,
            "subtitle_size": 24 if slide_type in {"title", "section", "agenda", "thank-you"} else 20,
            "body_size": 18,
            "font_family": "Calibri",
        },
        "spacing": {
            "outer_padding": "comfortable",
            "content_gap": "medium",
            "bullet_spacing": "balanced",
        },
        "visual_priority": visual_priority,
        "quality_checks": _design_checklist_for_slide(slide_type),
    }


def _normalize_slide(slide: Dict, index: int, theme: str, deck_context: Dict | None = None) -> Dict:
    safe_theme = _normalize_theme(theme)
    context = deck_context if isinstance(deck_context, dict) else {}
    safe_slide = slide if isinstance(slide, dict) else {}
    context_topic = _normalize_string(context.get("topic"))
    context_audience = _normalize_string(context.get("audience"), default="Business")
    context_presenter = _normalize_string(context.get("presenter"))

    slide_type = _normalize_string(safe_slide.get("slide_type"), default="bullet").lower()
    if slide_type not in ALLOWED_SLIDE_TYPES:
        slide_type = "bullet"

    title = _clean_text(_normalize_string(safe_slide.get("title"), default=f"Slide {index}"))
    if not title:
        title = f"Slide {index}"

    subtitle = _clean_text(_normalize_string(safe_slide.get("subtitle"), default=""))

    bullets = _normalize_list(
        safe_slide.get("bullets")
        or safe_slide.get("content")
        or safe_slide.get("points")
        or safe_slide.get("items")
    )

    left_title = _clean_text(_normalize_string(safe_slide.get("left_title"), default="Left"))
    right_title = _clean_text(_normalize_string(safe_slide.get("right_title"), default="Right"))
    left_points = _normalize_list(safe_slide.get("left_points"))
    right_points = _normalize_list(safe_slide.get("right_points"))
    milestones = _normalize_list(safe_slide.get("milestones") or safe_slide.get("timeline"))
    chart_title = _clean_text(_normalize_string(safe_slide.get("chart_title"), default=title))
    chart_type = _normalize_string(safe_slide.get("chart_type"), default="bar").lower()
    if chart_type not in {"bar", "pie", "trend"}:
        chart_type = "bar"
    chart_points = _normalize_chart_points(
        safe_slide.get("chart_points") or safe_slide.get("data_points") or safe_slide.get("chart_data")
    )

    if slide_type in {"bullet", "summary", "section", "agenda", "conclusion"}:
        if not bullets and subtitle:
            bullets = [_limit_words(subtitle)]
        if not bullets:
            bullets = [f"Key point for {title}"]

    if slide_type in {"two-column", "comparison"}:
        if not left_points:
            left_points = [f"Core idea for {left_title}"]
        if not right_points:
            right_points = [f"Core idea for {right_title}"]

    if slide_type == "timeline" and not milestones:
        milestones = ["Phase 1", "Phase 2", "Phase 3"]

    if slide_type == "agenda":
        if len(bullets) < 3:
            bullets = bullets + ["Problem framing", "Strategic options", "Implementation roadmap"]

    if slide_type == "conclusion" and not subtitle:
        subtitle = "Key takeaways and next steps"

    if slide_type == "chart":
        if not chart_points:
            source = bullets or milestones
            if not source:
                source = ["Awareness", "Adoption", "Retention", "Expansion"]
            seeded = []
            for idx, item in enumerate(source[:5], start=1):
                seeded.append({"label": _limit_words(item, max_words=4), "value": min(25 + idx * 15, 95)})
            chart_points = seeded
        if len(chart_points) < 3:
            fill_labels = ["Baseline", "Current", "Target", "Stretch"]
            for idx, label in enumerate(fill_labels, start=1):
                if len(chart_points) >= 4:
                    break
                chart_points.append({"label": label, "value": min(30 + (idx + len(chart_points)) * 12, 95)})
        if chart_points and len({point.get("value", 0) for point in chart_points}) == 1:
            for idx, point in enumerate(chart_points, start=1):
                point["value"] = max(min(int(point.get("value", 0)) + idx * 5, 100), 10)
        if not subtitle:
            subtitle = "Data snapshot"

    if slide_type == "title" and not subtitle:
        if context_presenter:
            subtitle = f"Presented by: {context_presenter}"
        else:
            subtitle = f"{_normalize_string(context.get('tone'), default='Professional')} briefing"

    if slide_type == "thank-you":
        if not subtitle:
            subtitle = "Questions, discussion, and next steps"
        if not bullets:
            bullets = ["Open Q&A", "Agree next steps", "Confirm ownership and timeline"]

    if slide_type == "title" and not bullets:
        bullets = ["Executive context", "Decision focus", "Expected outcomes"]

    bullets = _compress_bullets(bullets, max_bullets=5)
    left_points = _compress_bullets(left_points, max_bullets=5)
    right_points = _compress_bullets(right_points, max_bullets=5)
    milestones = _compress_bullets(milestones, max_bullets=5)

    layout = _normalize_string(safe_slide.get("layout"), default=_default_layout(slide_type)).lower()

    normalized = {
        "slide_type": slide_type,
        "title": title,
        "subtitle": subtitle,
        "layout": layout,
        "design": _normalize_design(safe_slide.get("design"), safe_theme, slide_type, index, title, context_topic, context_audience),
        "animation_hint": "fade-up" if slide_type in {"title", "section", "agenda"} else "subtle-rise",
    }

    if slide_type in {"bullet", "summary", "section", "agenda", "conclusion", "title", "thank-you"}:
        normalized["bullets"] = bullets
    if slide_type in {"two-column", "comparison"}:
        normalized["left_title"] = left_title or "Left"
        normalized["left_points"] = left_points
        normalized["right_title"] = right_title or "Right"
        normalized["right_points"] = right_points
    if slide_type == "timeline":
        normalized["milestones"] = milestones
    if slide_type == "chart":
        normalized["chart_title"] = chart_title or title
        normalized["chart_type"] = chart_type
        normalized["chart_points"] = chart_points

    return normalized


def _split_overflow_bullet_slides(slides: List[Dict]) -> List[Dict]:
    expanded = []
    for slide in slides:
        if slide.get("slide_type") not in {"bullet", "summary", "section", "agenda", "conclusion"}:
            expanded.append(slide)
            continue

        bullets = _normalize_list(slide.get("bullets"))
        if len(bullets) <= 5:
            slide["bullets"] = bullets
            expanded.append(slide)
            continue

        for idx in range(0, len(bullets), 5):
            chunk = bullets[idx : idx + 5]
            copy_slide = dict(slide)
            copy_slide["bullets"] = chunk
            if idx > 0:
                copy_slide["title"] = f"{slide.get('title', 'Slide')} (cont.)"
            expanded.append(copy_slide)

    return expanded


def _flow_types(slide_count: int) -> List[str]:
    base = [
        "title",
        "agenda",
        "section",
        "bullet",
        "two-column",
        "comparison",
        "chart",
        "timeline",
        "conclusion",
        "thank-you",
    ]

    if slide_count <= len(base):
        return base[:slide_count]

    flow = []
    while len(flow) < slide_count:
        flow.extend(base[2:-1])
    flow = ["title"] + flow[: max(slide_count - 2, 0)] + ["thank-you"]
    return flow[:slide_count]


def _build_fallback_presentation(topic: str, description: str, slides: int, tone: str, audience: str, theme: str, presenter_name: str = "") -> Dict:
    safe_topic = _normalize_string(topic, default="Untitled Topic") or "Untitled Topic"
    safe_description = _normalize_string(description)
    safe_tone = _normalize_string(tone, default="Professional") or "Professional"
    safe_audience = _normalize_string(audience, default="Business") or "Business"
    safe_presenter = _normalize_string(presenter_name)
    safe_theme = _normalize_theme(theme)
    count = slides if isinstance(slides, int) and slides > 0 else 8

    flow = _flow_types(count)
    built = []
    deck_context = {
        "topic": safe_topic,
        "audience": safe_audience,
        "tone": safe_tone,
        "presenter": safe_presenter,
    }
    for idx, slide_type in enumerate(flow, start=1):
        base = {
            "slide_type": slide_type,
            "title": f"{safe_topic} - Slide {idx}",
            "subtitle": "",
            "layout": _default_layout(slide_type),
            "design": {
                "background": "gradient" if slide_type in {"title", "section", "agenda", "thank-you"} else "solid",
                "accent_color": _accent_for_index(safe_theme, idx),
                "icon": _icon_from_context(slide_type, safe_topic, safe_topic, safe_audience),
                "shape": "rounded",
            },
        }

        if slide_type == "title":
            base["title"] = safe_topic
            if safe_presenter:
                base["subtitle"] = f"Presented by: {safe_presenter}"
            else:
                base["subtitle"] = f"{safe_tone} presentation for {safe_audience}"
        elif slide_type in {"agenda", "section"}:
            base["title"] = "Agenda"
            base["subtitle"] = "Introduction, strategy, outcomes, and next steps"
            base["bullets"] = ["Market context", "Solution strategy", "Execution plan", "Expected outcomes"]
        elif slide_type in {"bullet", "summary", "conclusion"}:
            base["title"] = "Key Insights" if slide_type == "bullet" else "Conclusion"
            base["bullets"] = [
                f"Topic focus: {safe_topic}",
                f"Audience: {safe_audience}",
                f"Tone: {safe_tone}",
                safe_description or "Actionable recommendations and outcomes",
            ]
        elif slide_type in {"two-column", "comparison"}:
            base["title"] = "Approach Comparison" if slide_type == "comparison" else "Implementation Strategy"
            base["left_title"] = "Current State"
            base["left_points"] = ["Existing process", "Known constraints", "Improvement opportunities"]
            base["right_title"] = "Target State"
            base["right_points"] = ["Optimized process", "Measurable outcomes", "Execution roadmap"]
        elif slide_type == "timeline":
            base["title"] = "Execution Timeline"
            base["milestones"] = ["Plan", "Build", "Validate", "Launch"]
        elif slide_type == "chart":
            base["title"] = "Performance Snapshot"
            base["subtitle"] = "Headline metrics"
            base["chart_title"] = "Impact by pillar"
            base["chart_points"] = [
                {"label": "Efficiency", "value": 72},
                {"label": "Adoption", "value": 64},
                {"label": "Quality", "value": 81},
                {"label": "Growth", "value": 76},
            ]
        elif slide_type == "thank-you":
            base["title"] = "Thank You"
            base["subtitle"] = "Questions and discussion"

        built.append(_normalize_slide(base, idx, safe_theme, deck_context=deck_context))

    return {
        "title": safe_topic,
        "theme": safe_theme,
        "slides": built,
    }


def _apply_visual_enhancement_pass(slides: List[Dict], theme: str, deck_context: Dict) -> List[Dict]:
    enhanced = []
    safe_theme = _normalize_theme(theme)
    for index, slide in enumerate(slides, start=1):
        normalized = _normalize_slide(slide, index, safe_theme, deck_context=deck_context)
        slide_type = normalized.get("slide_type", "bullet")
        normalized["design"] = _normalize_design(
            normalized.get("design"),
            safe_theme,
            slide_type,
            index,
            normalized.get("title", "Slide"),
            _normalize_string(deck_context.get("topic")),
            _normalize_string(deck_context.get("audience"), default="Business"),
        )
        if slide_type == "title" and not normalized.get("subtitle"):
            audience = _normalize_string(deck_context.get("audience"), default="Business")
            tone = _normalize_string(deck_context.get("tone"), default="Professional")
            normalized["subtitle"] = f"{tone} briefing for {audience} audience"
        if slide_type == "section" and not normalized.get("subtitle"):
            normalized["subtitle"] = "Key message and strategic direction"
        if slide_type in {"bullet", "agenda", "conclusion", "summary"}:
            normalized["bullets"] = _compress_bullets(_normalize_list(normalized.get("bullets")), max_bullets=5)
        if slide_type == "chart":
            points = _normalize_chart_points(normalized.get("chart_points"))
            if len(points) < 3:
                points = points + [
                    {"label": "Current", "value": 58},
                    {"label": "Target", "value": 74},
                    {"label": "Stretch", "value": 88},
                ]
            normalized["chart_points"] = points[:5]
        if slide_type == "thank-you" and not normalized.get("subtitle"):
            normalized["subtitle"] = "Questions, discussion, and next steps"
        enhanced.append(normalized)
    return enhanced


def _ensure_visual_variety(slides: List[Dict], deck_title: str, theme: str, deck_context: Dict | None = None) -> List[Dict]:
    safe_theme = _normalize_theme(theme)
    safe_context = deck_context if isinstance(deck_context, dict) else {}
    working = list(slides)
    if len(working) < 6:
        return working

    existing = {item.get("slide_type") for item in working if isinstance(item, dict)}
    required_visual_types = ["chart", "two-column", "timeline"]
    if len(working) >= 8:
        required_visual_types.append("comparison")

    templates = {
        "chart": {
            "slide_type": "chart",
            "title": "Performance Snapshot",
            "subtitle": "Data-backed insights",
            "chart_title": "Impact by dimension",
            "chart_type": "bar",
            "chart_points": [
                {"label": "Efficiency", "value": 68},
                {"label": "Adoption", "value": 74},
                {"label": "Quality", "value": 81},
                {"label": "Growth", "value": 77},
            ],
        },
        "two-column": {
            "slide_type": "two-column",
            "title": "Current vs Future State",
            "left_title": "Current",
            "left_points": ["Manual steps", "Limited visibility", "Long cycle time"],
            "right_title": "Future",
            "right_points": ["Automated flow", "Real-time insight", "Faster delivery"],
        },
        "timeline": {
            "slide_type": "timeline",
            "title": "Execution Roadmap",
            "milestones": ["Discovery", "Pilot", "Rollout", "Optimization"],
        },
        "comparison": {
            "slide_type": "comparison",
            "title": "Strategic Options",
            "left_title": "Option A",
            "left_points": ["Lower cost", "Faster start", "Limited scale"],
            "right_title": "Option B",
            "right_points": ["Higher impact", "Better resilience", "Scalable foundation"],
        },
    }

    mutable_types = {"bullet", "section", "summary", "agenda"}
    for needed_type in required_visual_types:
        if needed_type in existing:
            continue

        replace_index = None
        for idx in range(1, max(len(working) - 1, 1)):
            slide_type = _normalize_string((working[idx] or {}).get("slide_type")).lower()
            if slide_type in mutable_types:
                replace_index = idx
                break

        if replace_index is None:
            continue

        working[replace_index] = _normalize_slide(
            templates[needed_type],
            replace_index + 1,
            safe_theme,
            deck_context={
                "topic": deck_title,
                "audience": _normalize_string(safe_context.get("audience"), default="Business"),
                "tone": _normalize_string(safe_context.get("tone"), default="Professional"),
                "presenter": _normalize_string(safe_context.get("presenter")),
            },
        )
        existing.add(needed_type)

    return working


def _ensure_professional_flow(slides: List[Dict], deck_title: str, theme: str, requested_count: int, deck_context: Dict | None = None) -> List[Dict]:
    safe_theme = _normalize_theme(theme)
    safe_context = deck_context if isinstance(deck_context, dict) else {}
    normalized = list(slides)

    if not normalized:
        return _build_fallback_presentation(deck_title, "", requested_count, "Professional", "Business", safe_theme)["slides"]

    if normalized[0].get("slide_type") != "title":
        normalized.insert(
            0,
            _normalize_slide(
                {
                    "slide_type": "title",
                    "title": deck_title,
                    "subtitle": "Executive presentation",
                    "layout": "centered",
                    "design": {"background": "gradient", "accent_color": THEME_DEFAULT_ACCENTS[safe_theme], "icon": "rocket", "shape": "rounded"},
                },
                1,
                safe_theme,
                deck_context=safe_context,
            ),
        )

    if requested_count >= 5:
        has_agenda = any(item.get("slide_type") == "agenda" for item in normalized)
        if not has_agenda:
            normalized.insert(
                1,
                _normalize_slide(
                    {
                        "slide_type": "agenda",
                        "title": "Agenda",
                        "subtitle": "What we will cover",
                        "bullets": ["Context", "Approach", "Insights", "Next steps"],
                        "layout": "centered",
                    },
                    2,
                    safe_theme,
                    deck_context=safe_context,
                ),
            )

    has_conclusion = any(item.get("slide_type") in {"conclusion", "summary"} for item in normalized)
    if not has_conclusion and requested_count >= 4:
        insert_at = max(len(normalized) - 1, 1)
        normalized.insert(
            insert_at,
            _normalize_slide(
                {
                    "slide_type": "conclusion",
                    "title": "Conclusion",
                    "subtitle": "Key takeaways and actions",
                    "bullets": ["Primary insight", "Recommended action", "Expected impact"],
                    "layout": "one-column",
                },
                insert_at + 1,
                safe_theme,
                deck_context=safe_context,
            ),
        )

    if normalized[-1].get("slide_type") != "thank-you":
        normalized.append(
            _normalize_slide(
                {
                    "slide_type": "thank-you",
                    "title": "Thank You",
                    "subtitle": "Q&A",
                    "layout": "centered",
                    "design": {"background": "gradient", "accent_color": THEME_DEFAULT_ACCENTS[safe_theme], "icon": "handshake", "shape": "rounded"},
                },
                len(normalized) + 1,
                safe_theme,
                deck_context=safe_context,
            )
        )

    normalized = _split_overflow_bullet_slides(normalized)

    while len(normalized) < requested_count:
        idx = len(normalized) + 1
        flow = _flow_types(requested_count)
        next_type = flow[min(idx - 1, len(flow) - 1)]
        normalized.append(
            _normalize_slide(
                {
                    "slide_type": next_type,
                    "title": f"Slide {idx}",
                    "bullets": [f"Key point {idx} about {deck_title}"],
                    "layout": _default_layout(next_type),
                    "design": {"background": "solid", "accent_color": THEME_DEFAULT_ACCENTS[safe_theme], "icon": ICON_BY_TYPE.get(next_type, "spark"), "shape": "rounded"},
                },
                idx,
                safe_theme,
                deck_context=safe_context,
            )
        )

    if len(normalized) > requested_count and requested_count > 2:
        middle = normalized[1 : requested_count - 1]
        normalized = [normalized[0], *middle, normalized[-1]]

    normalized = normalized[:requested_count]
    normalized = _ensure_visual_variety(normalized, deck_title, safe_theme, safe_context)
    return _apply_visual_enhancement_pass(normalized, safe_theme, safe_context)


def validate_presentation_payload(
    payload: Dict,
    requested_theme: str,
    requested_slides: int,
    topic: str,
    audience: str = "Business",
    tone: str = "Professional",
    presenter: str = "",
) -> Dict:
    if not isinstance(payload, dict):
        raise ValueError("Groq response is not a JSON object")

    safe_theme = _normalize_theme(payload.get("theme") or requested_theme)
    title = _clean_text(_normalize_string(payload.get("title"), default=topic or "Presentation"))
    if not title:
        title = "Presentation"

    raw_slides = payload.get("slides")
    if not isinstance(raw_slides, list):
        raw_slides = []

    context = {
        "topic": title,
        "audience": _normalize_string(audience, default="Business"),
        "tone": _normalize_string(tone, default="Professional"),
        "presenter": _normalize_string(presenter),
    }

    normalized_slides = []
    for index, raw_slide in enumerate(raw_slides, start=1):
        normalized_slides.append(_normalize_slide(raw_slide, index, safe_theme, deck_context=context))

    count = requested_slides if isinstance(requested_slides, int) and requested_slides > 0 else max(len(normalized_slides), 8)
    normalized_slides = _ensure_professional_flow(normalized_slides, title, safe_theme, count, deck_context=context)

    return {
        "title": title,
        "theme": safe_theme,
        "slides": normalized_slides,
    }


def _is_retryable_error(exc: Exception) -> bool:
    message = str(exc).lower()
    return any(term in message for term in RETRYABLE_ERROR_TERMS)


def _create_client() -> Groq:
    api_key = os.getenv("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError("GROQ_API_KEY is required in the environment")
    return Groq(api_key=api_key)


def _send_with_retries(client: Groq, messages: List[Dict], retries: int = 3, base_delay: int = 2):
    last_exc = None
    for attempt in range(1, retries + 1):
        try:
            print(f"Groq API attempt {attempt}/{retries}")
            response = client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=messages,
                temperature=0.6,
                max_tokens=2600,
            )
            print("Groq API response received")
            return response
        except Exception as exc:
            last_exc = exc
            print(f"Groq API error attempt {attempt}: {exc}")
            if attempt == retries or not _is_retryable_error(exc):
                break
            time.sleep(base_delay * (2 ** (attempt - 1)))
    raise RuntimeError(f"Groq API error: {last_exc}") from last_exc


def generate_presentation_json(
    topic: str,
    description: str,
    slides: int,
    tone: str,
    audience: str,
    theme: str = "modern",
    presenterName: str = "",
) -> Dict:
    safe_theme = _normalize_theme(theme)
    presenter_name = _normalize_string(presenterName)
    prompt = build_prompt(topic, description, slides, tone, audience, safe_theme, presenter_name)
    client = _create_client()
    messages = [
        {
            "role": "system",
            "content": "You are an expert presentation designer. Return only valid JSON exactly matching the requested schema.",
        },
        {
            "role": "user",
            "content": prompt,
        },
    ]

    try:
        response = _send_with_retries(client, messages)
    except Exception as exc:
        print("Groq request failed:", exc)
        raise RuntimeError(f"Groq API error: {exc}") from exc

    raw_text = extract_response_text(response)
    print(f"Raw Groq response length: {len(raw_text) if raw_text else 0}")

    if not raw_text:
        print("Groq returned empty content; using fallback presentation")
        return _build_fallback_presentation(topic, description, slides, tone, audience, safe_theme, presenter_name)

    payload = _extract_json_payload(raw_text)
    if not payload:
        print("Unable to parse Groq JSON response; using fallback presentation")
        return _build_fallback_presentation(topic, description, slides, tone, audience, safe_theme, presenter_name)

    try:
        return validate_presentation_payload(
            payload,
            requested_theme=safe_theme,
            requested_slides=slides,
            topic=topic,
            audience=audience,
            tone=tone,
            presenter=presenter_name,
        )
    except Exception as exc:
        print(f"Payload validation failed; using fallback presentation. Reason: {exc}")
        return _build_fallback_presentation(topic, description, slides, tone, audience, safe_theme, presenter_name)


def normalize_presentation_for_export(title: str, theme: str, slides: List[Dict]) -> Dict:
    payload = {
        "title": _normalize_string(title, default="Presentation"),
        "theme": _normalize_theme(theme),
        "slides": slides if isinstance(slides, list) else [],
    }
    requested = len(payload["slides"]) if payload["slides"] else 1
    return validate_presentation_payload(
        payload,
        requested_theme=payload["theme"],
        requested_slides=requested,
        topic=payload["title"],
        audience="Business",
        tone="Professional",
        presenter="",
    )


def edit_slide_with_ai(
    presentation_title: str,
    theme: str,
    slide: Dict,
    instruction: str,
    context_slides: List[Dict] | None = None,
) -> Dict:
    safe_theme = _normalize_theme(theme)
    safe_instruction = _normalize_string(instruction, default="Improve this slide professionally")
    base_slide = _normalize_slide(
        slide if isinstance(slide, dict) else {},
        1,
        safe_theme,
        deck_context={"topic": presentation_title, "audience": "Business", "tone": "Professional"},
    )
    neighbor_titles = []
    if isinstance(context_slides, list):
        for item in context_slides[:5]:
            if isinstance(item, dict):
                title = _normalize_string(item.get("title"))
                if title:
                    neighbor_titles.append(title)

    prompt = (
        "Edit a SINGLE presentation slide and return JSON only. "
        "Do not regenerate the full deck. "
        f"Presentation title: {presentation_title}. Theme: {safe_theme}. "
        f"Neighbor slide titles for context: {neighbor_titles}. "
        "You can improve layout, add icon hierarchy, convert to chart/two-column/timeline, shorten content, and increase professionalism. "
        "Keep visual hierarchy and preserve layout integrity. "
        "Max 5 bullets and max 12 words per bullet. "
        "Do not use placeholder text like Visual idea/Image ideas/Use this space. "
        "Allowed slide_type: title, agenda, section, bullet, two-column, comparison, chart, timeline, conclusion, thank-you. "
        "Return one JSON object using this schema:\n"
        "{\n"
        "  \"slide_type\": \"...\",\n"
        "  \"title\": \"...\",\n"
        "  \"subtitle\": \"...\",\n"
        "  \"bullets\": [\"...\"],\n"
        "  \"left_title\": \"...\",\n"
        "  \"left_points\": [\"...\"],\n"
        "  \"right_title\": \"...\",\n"
        "  \"right_points\": [\"...\"],\n"
        "  \"milestones\": [\"...\"],\n"
        "  \"chart_title\": \"...\",\n"
        "  \"chart_points\": [{\"label\": \"...\", \"value\": 40}],\n"
        "  \"layout\": \"centered|one-column|two-column|comparison|timeline|chart\",\n"
        "  \"design\": {\"background\": \"solid|gradient\", \"accent_color\": \"#RRGGBB\", \"icon\": \"...\", \"shape\": \"rounded\"}\n"
        "}\n"
        f"Current slide JSON: {json.dumps(base_slide)}\n"
        f"Edit instruction: {safe_instruction}"
    )

    client = _create_client()
    messages = [
        {
            "role": "system",
            "content": "You are an expert presentation slide editor. Return only valid JSON for the edited single slide.",
        },
        {
            "role": "user",
            "content": prompt,
        },
    ]

    try:
        response = _send_with_retries(client, messages)
        raw_text = extract_response_text(response)
        payload = _extract_json_payload(raw_text)
        candidate = payload.get("slide") if isinstance(payload.get("slide"), dict) else payload
        if not isinstance(candidate, dict):
            return base_slide
        return _normalize_slide(
            candidate,
            1,
            safe_theme,
            deck_context={"topic": presentation_title, "audience": "Business", "tone": "Professional"},
        )
    except Exception as exc:
        print(f"Slide edit failed, returning original slide. Reason: {exc}")
        return base_slide
