#!/usr/bin/env python3
"""Collect leads for reading. Wording signals are never fallacy diagnoses."""
from __future__ import annotations

import argparse
import json
import re
import sys
import urllib.request
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

USER_AGENT = "LogFallHeadlineScout/2.0 (+https://logfall.com)"
FEEDS = [
    {"name": "Google News Top Stories", "url": "https://news.google.com/rss?hl=en-US&gl=US&ceid=US:en"},
    {"name": "Google News Politics", "url": "https://news.google.com/rss/search?q=politics+when:7d&hl=en-US&gl=US&ceid=US:en"},
    {"name": "Google News Business", "url": "https://news.google.com/rss/search?q=economy+business+when:7d&hl=en-US&gl=US&ceid=US:en"},
    {"name": "Google News Science", "url": "https://news.google.com/rss/search?q=science+health+technology+when:7d&hl=en-US&gl=US&ceid=US:en"},
    {"name": "CBS News", "url": "https://www.cbsnews.com/latest/rss/main"},
    {"name": "BBC News", "url": "https://feeds.bbci.co.uk/news/rss.xml"},
    {"name": "NPR", "url": "https://feeds.npr.org/1001/rss.xml"},
    {"name": "Fox News", "url": "https://moxie.foxnews.com/google-publisher/latest.xml"},
    {"name": "The Guardian", "url": "https://www.theguardian.com/world/rss"},
]
# These cues prioritize reading, not accusations. Temporal words such as "after",
# group names, and punctuation do not earn points.
CUES = [
    (r"\b(because|due to|causes?|caused|leads? to|led to|responsible for)\b", 2,
     "What evidence supports the causal claim, and what alternative explanations were tested?"),
    (r"\b(either|only option|no choice|must choose|choice between)\b", 2,
     "Are the stated alternatives exhaustive, and does the source actually claim they are?"),
    (r"\b(proves?|debunks?|confirms?|settles|always|never)\b", 2,
     "Does the evidence justify the strength and scope of the conclusion?"),
    (r"\b(everyone|no one|all voters|all scientists|all experts)\b", 2,
     "Is the generalization supported by representative evidence?"),
    (r"\b(slams?|blasts?|destroys?|humiliates?|outrage|disaster)\b", 1,
     "Is the language merely expressive, or is emotion replacing support for a conclusion?"),
    (r"\b(same as|just like|worse than|better than)\b", 1,
     "Are the relevant similarities, differences, and comparison baseline supplied?"),
]

@dataclass
class Candidate:
    title: str
    url: str
    source: str
    publisher_url: str
    published: str
    selection_score: int
    review_questions: list[str]
    surfaced_in: list[str] = field(default_factory=list)
    status: str = "unreviewed_lead"


def fetch_xml(url: str) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    # Keep certificate verification enabled. A failed source is reported explicitly.
    with urllib.request.urlopen(request, timeout=20) as response:
        return response.read(5_000_001)


def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def parse_datetime(raw: str) -> datetime | None:
    for parser in (parsedate_to_datetime, lambda value: datetime.fromisoformat(value.replace("Z", "+00:00"))):
        try:
            dt = parser(normalize_whitespace(raw))
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc)
        except (ValueError, TypeError, OverflowError):
            continue
    return None


def safe_url(url: str) -> str:
    try:
        parts = urlsplit(url.strip())
        return url.strip() if parts.scheme in ("https", "http") and parts.netloc and not parts.username else ""
    except ValueError:
        return ""


def canonical_url(url: str) -> str:
    parts = urlsplit(url)
    query = [(key, value) for key, value in parse_qsl(parts.query, keep_blank_values=True)
             if not key.lower().startswith("utm_") and key.lower() not in {"fbclid", "gclid"}]
    return urlunsplit((parts.scheme.lower(), parts.netloc.lower(), parts.path, urlencode(query), ""))


def iter_items(xml_bytes: bytes, fallback_source: str):
    if len(xml_bytes) > 5_000_000:
        raise ValueError("Feed exceeds 5 MB limit")
    root = ET.fromstring(xml_bytes)
    channel = root.find("channel")
    if channel is not None:
        for item in channel.findall("item"):
            source_el = item.find("source")
            yield {
                "title": normalize_whitespace(item.findtext("title", "")),
                "url": safe_url(item.findtext("link", "")),
                "published": item.findtext("pubDate", ""),
                "source": normalize_whitespace(source_el.text or "") if source_el is not None else fallback_source,
                "publisher_url": safe_url(source_el.get("url", "")) if source_el is not None else "",
            }
        return
    ns = "{http://www.w3.org/2005/Atom}"
    for entry in root.findall(f"{ns}entry"):
        links = [el for el in entry.findall(f"{ns}link") if el.get("rel", "alternate") == "alternate"]
        yield {
            "title": normalize_whitespace(entry.findtext(f"{ns}title", "")),
            "url": safe_url(links[0].get("href", "")) if links else "",
            "published": entry.findtext(f"{ns}published", "") or entry.findtext(f"{ns}updated", ""),
            "source": fallback_source, "publisher_url": "",
        }


def score_title(title: str) -> tuple[int, list[str]]:
    hits = [(points, question) for pattern, points, question in CUES if re.search(pattern, title, re.I)]
    return sum(points for points, _ in hits), [question for _, question in hits]


def collect_candidates(days: int, limit: int, *, now: datetime | None = None, feeds=None, per_source: int = 3):
    now = now or datetime.now(timezone.utc)
    cutoff = now - timedelta(days=days)
    feeds = FEEDS if feeds is None else feeds
    candidates: dict[str, Candidate] = {}
    urls: dict[str, str] = {}
    feed_results = []
    skipped = Counter()
    for feed in feeds:
        try:
            items = list(iter_items(fetch_xml(feed["url"]), feed["name"]))
            feed_results.append({"name": feed["name"], "status": "ok", "items": len(items)})
        except Exception as exc:
            feed_results.append({"name": feed["name"], "status": "error", "error": str(exc)[:240]})
            print(f"[warn] {feed['name']}: {exc}", file=sys.stderr)
            continue
        for item in items:
            published = parse_datetime(item["published"])
            if not published:
                skipped["missing_or_invalid_date"] += 1
                continue
            if published < cutoff or published > now:
                skipped["outside_window"] += 1
                continue
            if not item["title"] or not item["url"]:
                skipped["missing_title_or_safe_url"] += 1
                continue
            source = item["source"] or feed["name"]
            title = item["title"]
            # Google appends the publisher name; remove only a matching suffix.
            suffix = f" - {source}"
            if title.endswith(suffix):
                title = title[:-len(suffix)]
            url = canonical_url(item["url"])
            key = source.casefold() + ":" + re.sub(r"\W+", " ", title.casefold()).strip()
            key = urls.get(url, key)
            urls[url] = key
            if key in candidates:
                if feed["name"] not in candidates[key].surfaced_in:
                    candidates[key].surfaced_in.append(feed["name"])
                # Prefer the publisher's direct URL over a Google aggregator URL.
                if urlsplit(candidates[key].url).netloc == "news.google.com" and urlsplit(url).netloc != "news.google.com":
                    candidates[key].url = url
                skipped["duplicate"] += 1
                continue
            score, questions = score_title(title)
            candidates[key] = Candidate(title, url, source, item["publisher_url"], published.isoformat(), score,
                questions or ["What argument, if any, is made? Is there enough accessible evidence to evaluate it?"], [feed["name"]])
    if not any(result["status"] == "ok" for result in feed_results):
        raise RuntimeError("All feeds failed; no current research input was collected.")
    ranked = sorted(candidates.values(), key=lambda item: (-item.selection_score, -datetime.fromisoformat(item.published).timestamp(), item.source, item.title))
    selected = []
    counts = Counter()
    for candidate in ranked:
        if counts[candidate.source.casefold()] >= per_source:
            continue
        selected.append(candidate)
        counts[candidate.source.casefold()] += 1
        if len(selected) >= limit:
            break
    return selected, feed_results, dict(skipped), len(candidates)


def markdown_text(value: str) -> str:
    return value.replace("\\", "\\\\").replace("[", "\\[").replace("]", "\\]").replace("<", "&lt;").replace(">", "&gt;")


def render_markdown(report: dict) -> str:
    ok = sum(result["status"] == "ok" for result in report["feed_results"])
    lines = ["# LogFall research leads", "", f"- Generated: {report['generated']}",
        f"- Window: {report['window_start']} to {report['generated']}",
        f"- Feeds read successfully: {ok}/{len(report['feed_results'])}",
        f"- Eligible unique leads: {report['eligible_count']}", f"- Shortlist: {len(report['candidates'])}", "",
        "> These are unreviewed leads, not accusations. Selection scores reflect wording cues only; they do not measure fallacy likelihood, truth, bias, or public importance.", "",
        "Full-article reading and independent evidence are required. Do not copy a headline into a published accusation because it appears here.", ""]
    if not report["candidates"]:
        lines.extend(["No eligible recent leads. This is not a finding that the news contains no reasoning errors.", ""])
    for i, item in enumerate(report["candidates"], 1):
        lines.extend([f"## {i}. [{markdown_text(item['title'])}](<{item['url']}>)", "",
            f"- Publisher: {markdown_text(item['source'])}", f"- Feed timestamp (verify on article): {item['published']}",
            f"- Reading-priority score: {item['selection_score']}",
            f"- Found in: {', '.join(item['surfaced_in'])}",
            "- Status: unreviewed lead; attribution, relevance, and reasoning not yet checked.", ""])
        lines.extend(f"- {question}" for question in item["review_questions"])
        lines.append("")
    lines.extend(["## Collection health", ""])
    for result in report["feed_results"]:
        lines.append(f"- {result['name']}: {result['status']}" + (f" — {result['error']}" if result['status'] == 'error' else f" ({result['items']} items)"))
    lines.extend(["", "## Editorial next step", "", "Follow `FALLACY_DETECTIVE.md`: verify dates and publisher URLs, assess salience, read the full argument, distinguish the speaker from the outlet, test a charitable reading, and reject weak cases. Do not target a fixed number of fallacies. Save drafts for review; this collector never publishes an investigation.", ""])
    return "\n".join(lines)


def positive_int(value: str) -> int:
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("must be at least 1")
    return number


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--days", type=positive_int, default=7)
    parser.add_argument("--limit", type=positive_int, default=20)
    parser.add_argument("--per-source", type=positive_int, default=3)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--json-out", type=Path, required=True)
    args = parser.parse_args()
    now = datetime.now(timezone.utc)
    try:
        candidates, results, skipped, eligible_count = collect_candidates(args.days, args.limit, now=now, per_source=args.per_source)
    except RuntimeError as exc:
        print(f"[error] {exc}", file=sys.stderr)
        return 1
    report = {"schema_version": 2, "generated": now.isoformat(), "window_start": (now - timedelta(days=args.days)).isoformat(),
        "days": args.days, "feeds": FEEDS, "feed_results": results, "skipped": skipped, "eligible_count": eligible_count,
        "candidates": [asdict(candidate) for candidate in candidates]}
    for output in (args.out, args.json_out):
        output.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(render_markdown(report), encoding="utf8")
    args.json_out.write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf8")
    print(f"Saved {len(candidates)} unreviewed leads; {sum(r['status'] == 'ok' for r in results)}/{len(results)} feeds available.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
