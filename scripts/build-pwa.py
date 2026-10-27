"""Bentuk cache PWA dari aset lokal; tidak menyertakan dokumen pengguna.

Jalankan dari folder mana pun: python scripts/build-pwa.py
Tidak memerlukan package Python tambahan.
"""
from pathlib import Path
import hashlib
import json

ROOT = Path(__file__).resolve().parents[1]
SITE = ROOT / "docs"
TEMPLATE = ROOT / "scripts" / "service-worker.template.js"
RELEASE = "1.1.0"


def main():
    # Daftar eksplisit: file sementara dan unggahan pengguna tidak ikut cache.
    assets = [
        "index.html", "style.css", "viewer.html", "viewer.js",
        "pwa.js", "pwa-status.js", "manifest.webmanifest",
        "vendor/jszip.min.js", "vendor/docx-preview.js", "vendor/nspell.js",
    ]
    assets += sorted(str(p.relative_to(SITE)) for p in (SITE / "js").glob("*.js"))
    assets += sorted(
        str(p.relative_to(SITE)) for p in (SITE / "assets").rglob("*") if p.is_file()
    )

    digest = hashlib.sha256()
    for asset in assets:
        digest.update(asset.encode("utf-8"))
        digest.update((SITE / asset).read_bytes())
    digest.update(Path(__file__).read_bytes())
    digest.update(TEMPLATE.read_bytes())
    version = RELEASE + "-" + digest.hexdigest()[:12]

    source = TEMPLATE.read_text(encoding="utf-8")
    source = source.replace("__VERSION__", json.dumps(version))
    source = source.replace("__ASSETS__", json.dumps(assets))
    (SITE / "sw.js").write_text(source, encoding="utf-8")
    print("PWA shell:", version, len(assets), "assets")


if __name__ == "__main__":
    main()
