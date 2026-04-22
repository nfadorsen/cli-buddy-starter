"""
analyze_exports.py — read slides.json from an exports dir and produce
an analysis_report.json plus an exec-ready markdown summary on stdout.

Usage:
    python analyze_exports.py <exportsDir>

No third-party dependencies; standard library only.
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter
from pathlib import Path

STOPWORDS = {
    "the", "a", "an", "and", "or", "but", "if", "then", "else", "of", "in",
    "on", "to", "for", "with", "at", "by", "from", "as", "is", "are", "was",
    "were", "be", "been", "being", "it", "its", "this", "that", "these",
    "those", "we", "you", "they", "he", "she", "them", "our", "your", "their",
    "will", "would", "should", "could", "can", "may", "might", "have", "has",
    "had", "do", "does", "did", "not", "no", "so", "than", "too", "very",
    "just", "also", "into", "about", "over", "under", "via", "per", "up",
    "down", "out", "off", "more", "most", "some", "any", "all", "each",
    "such", "only", "own", "same", "other", "new", "now", "one", "two",
    "three",
}

WORD_RE = re.compile(r"[A-Za-z][A-Za-z0-9\-]{2,}")
SENT_RE = re.compile(r"(?<=[.!?])\s+")


def split_sentences(text: str) -> list[str]:
    if not text:
        return []
    parts = [p.strip() for p in SENT_RE.split(text) if p.strip()]
    return parts


def summarize_text(text: str, max_bullets: int = 3) -> list[str]:
    sents = split_sentences(text)
    if not sents:
        lines = [l.strip() for l in (text or "").splitlines() if l.strip()]
        return lines[:max_bullets]
    return sents[:max_bullets]


def key_phrases(text: str, top: int = 8) -> list[str]:
    if not text:
        return []
    tokens = [w.lower() for w in WORD_RE.findall(text)]
    tokens = [t for t in tokens if t not in STOPWORDS and not t.isdigit()]
    counts = Counter(tokens)
    return [w for w, _ in counts.most_common(top)]


def shingles(text: str, size: int = 6) -> set[str]:
    tokens = [w.lower() for w in WORD_RE.findall(text or "")]
    if len(tokens) < size:
        return set()
    return {" ".join(tokens[i : i + size]) for i in range(len(tokens) - size + 1)}


def detect_duplication(slides: list[dict]) -> list[dict]:
    shingle_sets = [shingles(s.get("allText", "")) for s in slides]
    dups = []
    for i in range(len(slides)):
        for j in range(i + 1, len(slides)):
            a, b = shingle_sets[i], shingle_sets[j]
            if not a or not b:
                continue
            inter = a & b
            union = a | b
            if not union:
                continue
            jaccard = len(inter) / len(union)
            if jaccard >= 0.25:
                dups.append(
                    {
                        "slideA": slides[i].get("slideNumber"),
                        "slideB": slides[j].get("slideNumber"),
                        "jaccard": round(jaccard, 3),
                    }
                )
    return dups


def find_ask_slide(slides: list[dict]) -> int | None:
    ask_re = re.compile(r"\b(the ask|ask:|approvals?|we are asking|decision needed|request(ed)? approval)\b", re.IGNORECASE)
    for s in slides:
        hay = " ".join([s.get("title", "") or "", s.get("allText", "") or ""])
        if ask_re.search(hay):
            return int(s.get("slideNumber"))
    return None


def top_insights(slides: list[dict], dups: list[dict]) -> list[str]:
    insights: list[str] = []
    n = len(slides)
    insights.append(f"Deck has {n} slide{'s' if n != 1 else ''}.")
    with_notes = sum(1 for s in slides if (s.get("notesText") or "").strip())
    insights.append(f"{with_notes} of {n} slides have speaker notes.")
    long_slides = [s["slideNumber"] for s in slides if len((s.get("allText") or "")) > 1200]
    if long_slides:
        insights.append(f"Dense slides (>1200 chars): {long_slides}.")
    empty_slides = [s["slideNumber"] for s in slides if not (s.get("allText") or "").strip()]
    if empty_slides:
        insights.append(f"Slides with no body text: {empty_slides}.")
    if dups:
        pairs = ", ".join(f"{d['slideA']}~{d['slideB']}" for d in dups[:5])
        insights.append(f"Potential duplication (Jaccard>=0.25): {pairs}.")
    return insights[:5]


def top_improvements(slides: list[dict], dups: list[dict], ask_slide: int | None) -> list[str]:
    recs: list[str] = []
    if ask_slide is None:
        recs.append("No clear 'Ask' detected. If this is a decision deck, add 'The Ask' as slide 2.")
    elif ask_slide != 2:
        recs.append(f"'The Ask' appears to be on slide {ask_slide}. Move to slide 2 unless purely educational.")
    long_slides = [s["slideNumber"] for s in slides if len((s.get("allText") or "")) > 1200]
    if long_slides:
        recs.append(f"Trim dense slides {long_slides}: favor a big stat callout + 2-3 short bullets.")
    topic_titles = [s for s in slides if s.get("title") and len(s["title"].split()) <= 2]
    if topic_titles:
        nums = [s["slideNumber"] for s in topic_titles[:5]]
        recs.append(f"Titles on slides {nums} look topic-label style; rewrite as headlines stating the insight.")
    if dups:
        recs.append("Consolidate duplicated content across the flagged slide pairs.")
    no_notes = [s["slideNumber"] for s in slides if not (s.get("notesText") or "").strip()]
    if no_notes and len(no_notes) > len(slides) / 2:
        recs.append("Most slides lack speaker notes; add brief narration to support exec delivery.")
    if not recs:
        recs.append("Structure looks reasonable; review for plain-language tone and single clear call-to-action.")
    return recs[:5]


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python analyze_exports.py <exportsDir>", file=sys.stderr)
        return 2
    exports_dir = Path(sys.argv[1])
    slides_json = exports_dir / "slides.json"
    if not slides_json.exists():
        print(f"slides.json not found under {exports_dir}", file=sys.stderr)
        return 2

    data = json.loads(slides_json.read_text(encoding="utf-8"))
    slides = data.get("slides") or []

    per_slide = []
    for s in slides:
        per_slide.append(
            {
                "slideNumber": s.get("slideNumber"),
                "title": (s.get("title") or "").strip(),
                "summary": summarize_text(s.get("allText") or "", 3),
                "notesSummary": summarize_text(s.get("notesText") or "", 2),
                "keyPhrases": key_phrases(s.get("allText") or ""),
                "charCount": len(s.get("allText") or ""),
                "hasNotes": bool((s.get("notesText") or "").strip()),
            }
        )

    dups = detect_duplication(slides)
    ask_slide = find_ask_slide(slides)
    insights = top_insights(slides, dups)
    improvements = top_improvements(slides, dups, ask_slide)

    report = {
        "sourcePath": data.get("sourcePath"),
        "slideCount": len(slides),
        "askSlide": ask_slide,
        "duplicates": dups,
        "perSlide": per_slide,
        "topInsights": insights,
        "topImprovements": improvements,
    }

    out_path = exports_dir / "analysis_report.json"
    out_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")

    # Exec-ready markdown to stdout
    print(f"# Deck Analysis — {Path(data.get('sourcePath') or '').name or '(unknown)'}")
    print()
    print(f"- Slides: **{len(slides)}**")
    print(f"- 'Ask' slide detected: **{ask_slide if ask_slide else 'none'}**")
    print(f"- Duplicate slide pairs flagged: **{len(dups)}**")
    print()
    print("## Top insights")
    for i, x in enumerate(insights, 1):
        print(f"{i}. {x}")
    print()
    print("## Top recommended improvements")
    for i, x in enumerate(improvements, 1):
        print(f"{i}. {x}")
    print()
    print("## Per-slide headlines")
    for ps in per_slide:
        title = ps["title"] or "(no title)"
        print(f"- **{ps['slideNumber']}.** {title}")
    print()
    print(f"Full report: `{out_path}`")
    return 0


if __name__ == "__main__":
    sys.exit(main())
