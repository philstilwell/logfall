import importlib.util
import sys
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch
from xml.sax.saxutils import escape

spec = importlib.util.spec_from_file_location("headline_scout", Path(__file__).resolve().parents[1] / "scripts/headline_scout.py")
scout = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = scout
spec.loader.exec_module(scout)
NOW = datetime(2026, 10, 23, 13, tzinfo=timezone.utc)


def rss(items):
    return ("<rss><channel>" + "".join(
        f'<item><title>{escape(title)}</title><link>{escape(url)}</link><pubDate>{date}</pubDate><source url="https://example.com">{source}</source></item>'
        for title, url, date, source in items) + "</channel></rss>").encode()


class ScoutTests(unittest.TestCase):
    def collect(self, payloads, **kwargs):
        feeds = [{"name": f"Feed {i}", "url": f"https://example.com/{i}"} for i in range(len(payloads))]
        with patch.object(scout, "fetch_xml", side_effect=payloads):
            return scout.collect_candidates(7, 20, now=NOW, feeds=feeds, **kwargs)

    def test_temporal_words_and_group_names_are_not_diagnoses(self):
        self.assertEqual(scout.score_title("Black voters arrive after rain as polls open: Tuesday update"), (0, []))

    def test_rejects_unknown_old_and_future_dates(self):
        items = [(date, f"https://example.com/{i}", date, "Outlet") for i, date in enumerate([
            "", "invalid", "2026-10-01T12:00:00Z", "2026-10-24T12:00:00Z", "2026-10-22T12:00:00Z"])]
        selected, _, skipped, _ = self.collect([rss(items)])
        self.assertEqual(len(selected), 1)
        self.assertEqual(skipped["missing_or_invalid_date"], 2)
        self.assertEqual(skipped["outside_window"], 2)

    def test_deduplicates_title_and_prefers_publisher_link(self):
        a = rss([("A new report - Outlet", "https://news.google.com/rss/articles/example", "2026-10-22T12:00:00Z", "Outlet")])
        b = rss([("A new report", "https://example.com/story?utm_source=feed", "2026-10-22T12:00:00Z", "Outlet")])
        selected, _, _, eligible = self.collect([a, b])
        self.assertEqual(eligible, 1)
        self.assertEqual(selected[0].url, "https://example.com/story")
        self.assertEqual(len(selected[0].surfaced_in), 2)

    def test_newest_first_on_ties_and_source_cap(self):
        selected, _, _, _ = self.collect([rss([
            (f"Report {i}", f"https://example.com/{i}", f"2026-10-{18+i}T12:00:00Z", "Outlet") for i in range(5)
        ])], per_source=2)
        self.assertEqual([item.title for item in selected], ["Report 4", "Report 3"])

    def test_partial_feed_failure_is_explicit(self):
        selected, results, _, _ = self.collect([b"not XML", rss([("A report", "https://example.com/a", "2026-10-22T12:00:00Z", "Outlet")])])
        self.assertEqual(len(selected), 1)
        self.assertEqual([item["status"] for item in results], ["error", "ok"])

    def test_all_failed_feeds_raise(self):
        with self.assertRaisesRegex(RuntimeError, "All feeds failed"):
            self.collect([OSError("offline"), b"bad XML"])

    def test_atom_prefers_article_link_and_publication_date(self):
        xml = b'<feed xmlns="http://www.w3.org/2005/Atom"><entry><title>Example</title><published>2020-01-01T00:00:00Z</published><updated>2026-10-22T00:00:00Z</updated><link rel="self" href="https://example.com/api"/><link rel="alternate" href="https://example.com/story"/></entry></feed>'
        item = list(scout.iter_items(xml, "Outlet"))[0]
        self.assertEqual(item["url"], "https://example.com/story")
        self.assertEqual(item["published"], "2020-01-01T00:00:00Z")

    def test_unsafe_links_are_not_collected(self):
        selected, _, _, _ = self.collect([rss([("Report", "javascript:alert(1)", "2026-10-22T12:00:00Z", "Outlet")])])
        self.assertEqual(selected, [])


if __name__ == "__main__":
    unittest.main()
