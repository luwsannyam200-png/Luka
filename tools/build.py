#!/usr/bin/env python3
"""Inline the libraries into one offline HTML file: tools/Heregsel.html

  python3 tools/build.py --jszip path/to/jszip.min.js --xlsx path/to/xlsx.full.min.js

Libraries (npm): jszip@3.10.1, xlsx@0.18.5.
"""
import argparse
from pathlib import Path

HERE = Path(__file__).parent


def safe(js):
    # A literal "</script" inside inlined code would end the <script> element.
    return js.replace("</script", "<\\/script")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--jszip", required=True)
    ap.add_argument("--xlsx", required=True)
    a = ap.parse_args()
    html = (HERE / "src" / "app.html").read_text(encoding="utf-8")
    parts = {
        "/*JSZIP*/": Path(a.jszip).read_text(encoding="utf-8"),
        "/*XLSX*/": Path(a.xlsx).read_text(encoding="utf-8"),
        "/*WEEKLY_CORE*/": (HERE / "src" / "weekly-core.js").read_text(encoding="utf-8"),
        "/*ZASVAR_CORE*/": (HERE / "src" / "zasvar-core.js").read_text(encoding="utf-8"),
    }
    for marker, js in parts.items():
        assert html.count(marker) == 1, marker
        html = html.replace(marker, safe(js))
    out = HERE / "Heregsel.html"
    out.write_text(html, encoding="utf-8")
    print(f"{out} ({out.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    main()
