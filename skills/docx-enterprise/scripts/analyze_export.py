"""Summarize a docx-enterprise export (document.json) as exec-ready markdown.

Usage:
    python analyze_export.py <exportsDir>

Expects <exportsDir>/document.json produced by export_docx_com.ps1.
Prints a short markdown summary to stdout: headings outline, revision stats
by author, open comments, and a heuristic flag for high-risk contract phrases.
"""
from __future__ import annotations

import json
import re
import sys
from collections import Counter
from pathlib import Path

REVISION_TYPE_NAMES = {
    1: "insert",
    2: "delete",
    3: "property",
    4: "paragraph-number",
    5: "display-field",
    6: "reconcile",
    7: "conflict",
    8: "style",
    9: "replace",
    10: "paragraph-property",
    11: "table-property",
    12: "section-property",
    13: "style-definition",
    14: "moved-from",
    15: "moved-to",
    16: "cell-insertion",
    17: "cell-deletion",
    18: "cell-merge",
}

# Heuristic phrases that often warrant a second read in contracts / MSAs.
RISK_PATTERNS = [
    r"\bindemnif(y|ies|ication)\b",
    r"\blimitation of liability\b",
    r"\bconsequential damages\b",
    r"\bauto[- ]?renew(al)?\b",
    r"\bexclusivity\b",
    r"\bterminat(e|ion) for convenience\b",
    r"\bmost favored (nation|customer)\b",
    r"\bassignment\b",
    r"\bgoverning law\b",
    r"\bcap\b.{0,40}\bliability\b",
]
RISK_REGEX = re.compile("|".join(RISK_PATTERNS), re.IGNORECASE)


def main(argv: list[str]) -> int:
    if len(argv) != 2:
        print(__doc__, file=sys.stderr)
        return 2

    exports_dir = Path(argv[1])
    json_path = exports_dir / "document.json"
    if not json_path.exists():
        print(f"document.json not found at {json_path}", file=sys.stderr)
        return 2

    data = json.loads(json_path.read_text(encoding="utf-8"))

    paragraphs = data.get("paragraphs", [])
    revisions = data.get("revisions", [])
    comments = data.get("comments", [])
    label = data.get("sensitivityLabel", {}) or {}

    lines: list[str] = []
    lines.append(f"# Document summary — {Path(data.get('sourcePath', '')).name}")
    lines.append("")
    lines.append(
        f"- **Pages:** {data.get('pageCount', 0)}  "
        f"**Words:** {data.get('wordCount', 0)}  "
        f"**Paragraphs:** {data.get('paragraphCount', len(paragraphs))}"
    )
    if label.get("name"):
        lines.append(f"- **Sensitivity label:** {label.get('name')}")
    lines.append("")

    # --- Headings outline ---
    headings = [
        p for p in paragraphs
        if (p.get("outlineLevel", 10) or 10) <= 4 and (p.get("text") or "").strip()
        and str(p.get("style", "")).lower().startswith(("heading", "title"))
    ]
    if headings:
        lines.append("## Outline")
        for h in headings[:40]:
            level = int(h.get("outlineLevel") or 1)
            indent = "  " * max(0, level - 1)
            text = (h.get("text") or "").strip().rstrip(".")
            if len(text) > 120:
                text = text[:117] + "..."
            lines.append(f"{indent}- {text}")
        if len(headings) > 40:
            lines.append(f"  _(...{len(headings) - 40} more headings omitted)_")
        lines.append("")

    # --- Revision stats ---
    lines.append(f"## Tracked changes ({len(revisions)})")
    if revisions:
        by_author = Counter((r.get("author") or "(unknown)") for r in revisions)
        by_type = Counter(REVISION_TYPE_NAMES.get(int(r.get("type") or 0), f"type-{r.get('type')}") for r in revisions)
        lines.append("")
        lines.append("**By author:**")
        for author, n in by_author.most_common():
            lines.append(f"- {author}: {n}")
        lines.append("")
        lines.append("**By type:**")
        for t, n in by_type.most_common():
            lines.append(f"- {t}: {n}")
    else:
        lines.append("- No tracked changes.")
    lines.append("")

    # --- Comments ---
    lines.append(f"## Comments ({len(comments)})")
    if comments:
        for c in comments:
            author = c.get("author") or "(unknown)"
            anchor = (c.get("anchor") or "").strip().replace("\n", " ")
            if len(anchor) > 80:
                anchor = anchor[:77] + "..."
            text = (c.get("text") or "").strip().replace("\n", " ")
            if len(text) > 200:
                text = text[:197] + "..."
            lines.append(f"- **{author}** on \"{anchor}\": {text}")
    else:
        lines.append("- No comments.")
    lines.append("")

    # --- Risk-phrase heuristic scan ---
    hits: list[tuple[int, str, str]] = []
    for p in paragraphs:
        text = p.get("text") or ""
        m = RISK_REGEX.search(text)
        if m:
            hits.append((int(p.get("index") or 0), m.group(0), text.strip()))
    if hits:
        lines.append(f"## Heuristic risk-phrase hits ({len(hits)})")
        lines.append("_These are keyword flags, not a legal review. Confirm manually._")
        lines.append("")
        for idx, kw, text in hits[:25]:
            snippet = text[:200] + ("..." if len(text) > 200 else "")
            lines.append(f"- p{idx} (**{kw}**): {snippet}")
        if len(hits) > 25:
            lines.append(f"  _(...{len(hits) - 25} more hits omitted)_")
        lines.append("")

    print("\n".join(lines))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
