#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import ssl
import sys
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Iterable


USER_AGENT = "LogFallHeadlineScout/1.0 (+https://logfall.com)"

FEEDS = [
    {"name": "Google News Top Stories", "url": "https://news.google.com/rss?hl=en-US&gl=US&ceid=US:en"},
    {"name": "Google News Politics", "url": "https://news.google.com/rss/search?q=politics+when:7d&hl=en-US&gl=US&ceid=US:en"},
    {"name": "Google News Courts", "url": "https://news.google.com/rss/search?q=court+when:7d&hl=en-US&gl=US&ceid=US:en"},
    {"name": "Google News Elections", "url": "https://news.google.com/rss/search?q=election+when:7d&hl=en-US&gl=US&ceid=US:en"},
    {"name": "CBS News Latest", "url": "https://www.cbsnews.com/latest/rss/main"},
]

CAUSE_PATTERNS = [
    (re.compile(r"\b(because|due to|after|as|since|why)\b", re.I), "Single cause fallacy", 4, "Headline leans on compact causal framing."),
    (re.compile(r"\b(causes?|caused|led to|leads to|drives?|driven by|forces?|forced|results? in|turns? into|makes?|means?|signals?|reveals?|transforms?|reshapes?|haunts?|backfires?)\b", re.I), "Correlation is not causation", 4, "Headline suggests causation in a compressed form."),
]

BINARY_PATTERNS = [
    (re.compile(r"\b(either|only|just|nothing but|must choose|choice between|two paths|no choice)\b", re.I), "False dilemma", 3, "Headline may collapse the issue into too few live options."),
]

GROUP_PATTERNS = [
    (re.compile(r"\b(black|white|democrat|democratic|republican|immigrant|students|teachers|men|women|parents|voters)\b", re.I), "Category or group label may be doing too much explanatory work.", 2, "Group label could invite overgeneralized or proxy reasoning."),
]

PROOF_PATTERNS = [
    (re.compile(r"\b(proves?|debunks?|confirms?|exposes?|shows once and for all|settles)\b", re.I), "Bare assertion / evidential overclaim", 3, "Headline sounds more conclusive than the underlying evidence may warrant."),
]

EMOTIONAL_PATTERNS = [
    (re.compile(r"\b(slams?|blasts?|destroys?|humiliates?|shocks?|chaos|panic|outrage|disaster)\b", re.I), "Appeal to emotion / loaded framing", 2, "Headline uses emotionally forceful language that may outrun the reasoning."),
]

COMPARISON_PATTERNS = [
    (re.compile(r"\b(vs\.?|same as|just like|worse than|better than)\b", re.I), "False equivalence / incomplete comparison", 2, "Headline may invite a comparison without enough qualifying context."),
]


@dataclass
class Candidate:
    title: str
    url: str
    source: str
    published: str
    score: int
    likely_fallacies: list[str]
    why_flagged: list[str]


def fetch_xml(url: str) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            return response.read()
    except Exception as exc:
        if "CERTIFICATE_VERIFY_FAILED" not in str(exc):
            raise
        insecure_context = ssl._create_unverified_context()
        with urllib.request.urlopen(request, timeout=20, context=insecure_context) as response:
            return response.read()


def text_or_empty(element: ET.Element | None, tag: str) -> str:
    if element is None:
        return ""
    child = element.find(tag)
    if child is not None and child.text:
        return child.text.strip()
    return ""


def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def parse_datetime(raw: str) -> datetime | None:
    raw = normalize_whitespace(raw)
    if not raw:
        return None
    try:
        dt = parsedate_to_datetime(raw)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        pass
    iso = raw.replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(iso)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        return None


def iter_items(xml_bytes: bytes, fallback_source: str) -> Iterable[dict[str, str]]:
    root = ET.fromstring(xml_bytes)
    channel = root.find("channel")
    if channel is not None:
        for item in channel.findall("item"):
            title = normalize_whitespace(text_or_empty(item, "title"))
            link = normalize_whitespace(text_or_empty(item, "link"))
            pub = normalize_whitespace(text_or_empty(item, "pubDate"))
            source = normalize_whitespace(text_or_empty(item, "source")) or fallback_source
            if title and link:
                yield {"title": title, "url": link, "published": pub, "source": source}
        return

    namespace = "{http://www.w3.org/2005/Atom}"
    for entry in root.findall(f"{namespace}entry"):
        title = normalize_whitespace(text_or_empty(entry, f"{namespace}title"))
        published = normalize_whitespace(text_or_empty(entry, f"{namespace}updated") or text_or_empty(entry, f"{namespace}published"))
        link = ""
        for link_el in entry.findall(f"{namespace}link"):
            href = link_el.attrib.get("href", "").strip()
            if href:
                link = href
                break
        if title and link:
            yield {"title": title, "url": link, "published": published, "source": fallback_source}


def score_title(title: str) -> tuple[int, list[str], list[str]]:
    score = 0
    fallacies: list[str] = []
    reasons: list[str] = []
    checks = CAUSE_PATTERNS + BINARY_PATTERNS + PROOF_PATTERNS + EMOTIONAL_PATTERNS + COMPARISON_PATTERNS

    for pattern, label, points, reason in checks:
        if pattern.search(title):
            score += points
            if label not in fallacies:
                fallacies.append(label)
            reasons.append(reason)

    for pattern, label, points, reason in GROUP_PATTERNS:
        if pattern.search(title):
            score += points
            if "Group-label overreach / proxy reasoning" not in fallacies:
                fallacies.append("Group-label overreach / proxy reasoning")
            reasons.append(reason)

    if ":" in title or ";" in title:
        score += 1
        reasons.append("Two-part headline structure may hide a jump from setup to conclusion.")

    if len(title.split()) <= 6:
        score += 1
        reasons.append("Very short headline may be too compressed for the claim it makes.")

    fallacies = fallacies[:4]
    reasons = list(dict.fromkeys(reasons))
    return score, fallacies, reasons


def collect_candidates(days: int, limit: int) -> list[Candidate]:
    now = datetime.now(timezone.utc)
    cutoff = now - timedelta(days=days)
    seen: set[str] = set()
    candidates: list[Candidate] = []

    for feed in FEEDS:
        try:
            xml_bytes = fetch_xml(feed["url"])
        except Exception as exc:
            print(f"[warn] failed to fetch {feed['name']}: {exc}", file=sys.stderr)
            continue

        for item in iter_items(xml_bytes, feed["name"]):
            url = item["url"]
            if not url or url in seen:
                continue
            seen.add(url)
            published_dt = parse_datetime(item["published"])
            if published_dt and published_dt < cutoff:
                continue
            score, fallacies, reasons = score_title(item["title"])
            if score < 2:
                continue
            candidates.append(
                Candidate(
                    title=item["title"],
                    url=url,
                    source=item["source"] or feed["name"],
                    published=published_dt.date().isoformat() if published_dt else "unknown",
                    score=score,
                    likely_fallacies=fallacies,
                    why_flagged=reasons,
                )
            )

    candidates.sort(key=lambda item: (-item.score, item.published, item.source, item.title))
    return candidates[:limit]


def render_markdown(candidates: list[Candidate], days: int) -> str:
    today = datetime.now(timezone.utc).date().isoformat()
    lines = [
        f"# LogFall headline scout report",
        "",
        f"- Generated: {today}",
        f"- Corpus window: last {days} days",
        f"- Feeds scanned: {len(FEEDS)}",
        f"- Candidates surfaced: {len(candidates)}",
        "",
        "> This report is a curation aid, not an auto-diagnosis engine. Every flagged headline still needs human review before it becomes a LogFall feature page.",
        "",
        "## Strongest candidates",
        "",
    ]

    if not candidates:
        lines.append("No candidates crossed the scoring threshold this run.")
        return "\n".join(lines) + "\n"

    for index, candidate in enumerate(candidates, start=1):
        lines.extend(
            [
                f"### {index}. [{candidate.title}]({candidate.url})",
                "",
                f"- Source: {candidate.source}",
                f"- Published: {candidate.published}",
                f"- Scout score: {candidate.score}",
                f"- Likely fallacy lenses: {', '.join(candidate.likely_fallacies) if candidate.likely_fallacies else 'None'}",
                f"- Why it was flagged: {'; '.join(candidate.why_flagged)}",
                "",
            ]
        )

    lines.extend(
        [
            "## How to use this report",
            "",
            "1. Open the headline and ask what it explicitly states versus what it tempts the reader to infer.",
            "2. Before diagnosing anything, ask what about the issue might make a reader especially eager or especially reluctant to call the wording fallacious.",
            "3. Read at least enough of the article to check whether the apparent fallacy is in the headline itself or only in a common reaction to it.",
            "4. Prefer feature pages built around 3 to 5 defensible fallacy lenses rather than a single overconfident diagnosis.",
            "5. Treat self-knowledge as part of the editorial test: the best weekly cases are not just public reasoning failures, but cases that expose how political sympathy or hostility can distort the reader's own judgment.",
            "",
        ]
    )

    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate a weekly LogFall headline scouting report.")
    parser.add_argument("--days", type=int, default=7, help="How many recent days to scan.")
    parser.add_argument("--limit", type=int, default=20, help="Maximum number of candidates to keep.")
    parser.add_argument("--out", type=Path, required=True, help="Markdown output path.")
    parser.add_argument("--json-out", type=Path, required=True, help="JSON output path.")
    args = parser.parse_args()

    candidates = collect_candidates(days=args.days, limit=args.limit)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.json_out.parent.mkdir(parents=True, exist_ok=True)

    args.out.write_text(render_markdown(candidates, args.days), encoding="utf8")
    args.json_out.write_text(
        json.dumps(
            {
                "generated": datetime.now(timezone.utc).isoformat(),
                "days": args.days,
                "feeds": FEEDS,
                "candidates": [asdict(candidate) for candidate in candidates],
            },
            indent=2,
        )
        + "\n",
        encoding="utf8",
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
