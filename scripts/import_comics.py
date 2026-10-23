"""Import the reviewed comic collection: lossless originals and smaller web previews.

Usage: python3 scripts/import_comics.py /path/to/LOGFALL
Requires the existing cwebp command. Source PNGs are never modified.
"""
import argparse
from concurrent.futures import ThreadPoolExecutor
import hashlib
import json
from pathlib import Path
import struct
import subprocess

ROOT = Path(__file__).resolve().parents[1]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('collection', type=Path)
    args = parser.parse_args()
    catalog = json.loads((args.collection / 'production/catalog.json').read_text())
    records = json.loads((ROOT / 'data/fallacies.json').read_text())['records']
    assert len(catalog) == len(records)
    assert {c['slug'] for c in catalog} == {r['slug'] for r in records}, 'Comic/page mapping differs'
    for size in ('full', 'preview'):
        (ROOT / 'assets/comics' / size).mkdir(parents=True, exist_ok=True)

    def convert(item):
        slug = item['slug']
        source = args.collection / f"{item['number']:03d}-{slug}.png"
        data = source.read_bytes()
        assert data[:8] == b'\x89PNG\r\n\x1a\n'
        width, height = struct.unpack('>II', data[16:24])
        assert (width, height) == (1024, 1536), source.name
        full = f'assets/comics/full/{slug}.webp'
        preview = f'assets/comics/preview/{slug}.webp'
        subprocess.run(['cwebp', '-quiet', '-lossless', '-m', '4', '-exact', str(source), '-o', str(ROOT / full)], check=True)
        subprocess.run(['cwebp', '-quiet', '-q', '88', '-m', '6', '-resize', '768', '1152', str(source), '-o', str(ROOT / preview)], check=True)
        return slug, {
            'number': item['number'], 'full': full, 'preview': preview,
            'width': width, 'height': height, 'previewWidth': 768, 'previewHeight': 1152,
            'sourceSha256': hashlib.sha256(data).hexdigest(),
            'fullBytes': (ROOT / full).stat().st_size,
            'previewBytes': (ROOT / preview).stat().st_size,
        }

    with ThreadPoolExecutor(max_workers=4) as pool:
        comics = dict(pool.map(convert, catalog))
    (ROOT / 'data/comic_illustrations.json').write_text(json.dumps(comics, indent=2) + '\n')
    print(json.dumps({'comics':len(comics), 'fullBytes':sum(c['fullBytes'] for c in comics.values()), 'previewBytes':sum(c['previewBytes'] for c in comics.values())}, indent=2))


if __name__ == '__main__':
    main()
