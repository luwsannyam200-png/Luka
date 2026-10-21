#!/usr/bin/env python3
"""Inline the libraries into one offline HTML file: tools/Heregsel.html

  python3 tools/build.py --jszip path/to/jszip.min.js --xlsx path/to/xlsx.full.min.js --npm path/to/node_modules

Libraries (npm): jszip@3.10.1, xlsx@0.18.5, tesseract.js@7.0.0 (+ tesseract.js-core, @tesseract.js-data/mon@1.0.0)
— Loader PM-ийн зураг уншигч (OCR), интернэтгүй ажиллахаар бүгдийг файлд суулгана.
"""
import argparse
import base64
from pathlib import Path

HERE = Path(__file__).parent


def safe(js):
    # A literal "</script" inside inlined code would end the <script> element.
    return js.replace("</script", "<\\/script")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--jszip", required=True)
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--npm", required=True, help="node_modules with tesseract.js, tesseract.js-core, @tesseract.js-data/mon")
    a = ap.parse_args()
    nm = Path(a.npm)
    html = (HERE / "src" / "app.html").read_text(encoding="utf-8")
    parts = {
        "/*JSZIP*/": Path(a.jszip).read_text(encoding="utf-8"),
        "/*XLSX*/": Path(a.xlsx).read_text(encoding="utf-8"),
        "/*WEEKLY_CORE*/": (HERE / "src" / "weekly-core.js").read_text(encoding="utf-8"),
        "/*ZASVAR_CORE*/": (HERE / "src" / "zasvar-core.js").read_text(encoding="utf-8"),
        "/*MEDEE_CORE*/": (HERE / "src" / "medee-core.js").read_text(encoding="utf-8"),
        "/*ASSISTANT_CORE*/": (HERE / "src" / "assistant-core.js").read_text(encoding="utf-8"),
        "/*LOADER_CORE*/": (HERE / "src" / "loader-core.js").read_text(encoding="utf-8"),
        "/*TESSERACT*/": (nm / "tesseract.js/dist/tesseract.min.js").read_text(encoding="utf-8"),
        "/*OCR_CORE*/": (nm / "tesseract.js-core/tesseract-core-lstm.wasm.js").read_text(encoding="utf-8"),
        "/*OCR_WORKER*/": (nm / "tesseract.js/dist/worker.min.js").read_text(encoding="utf-8"),
        "/*OCR_LANG*/": base64.b64encode((nm / "@tesseract.js-data/mon/4.0.0_best_int/mon.traineddata.gz").read_bytes()).decode(),
    }
    for marker, js in parts.items():
        assert html.count(marker) == 1, marker
        html = html.replace(marker, safe(js))
    out = HERE / "Heregsel.html"
    out.write_text(html, encoding="utf-8")
    print(f"{out} ({out.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    main()
