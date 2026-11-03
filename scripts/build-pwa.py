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
RELEASE = "1.3.5"
BUILD = "20261103.5"


def main():
    import re
    index=SITE / 'index.html'
    html=index.read_text(encoding='utf-8')
    html=re.sub(r'(<html[^>]*)(>)', lambda m: re.sub(r' data-build="[^"]*"','',m[1])+' data-build="'+BUILD+'"'+m[2],html,count=1)
    html=re.sub(r'Versi [0-9]+\.[0-9]+\.[0-9]+(?:-experimental\.[0-9]+)?(?: · Build [0-9.]+)?','Versi '+RELEASE,html)
    index.write_bytes(html.encode('utf-8'))
    # Daftar eksplisit: file sementara dan unggahan pengguna tidak ikut cache.
    assets = [
        "index.html", "style.css", "viewer.html", "viewer.js",
        "pwa.js", "pwa-status.js", "manifest.webmanifest",
        "vendor/jszip.min.js", "vendor/docx-preview.js", "vendor/nspell.js", "vendor/legacy-doc.js",
    ]
    assets += sorted(str(p.relative_to(SITE)) for p in (SITE / "js").glob("*.js"))
    assets += sorted(
        str(p.relative_to(SITE)) for p in (SITE / "assets").rglob("*") if p.is_file()
    )

    # Hash the same LF bytes that GitHub serves, even from a Windows checkout.
    text_extensions = {'.html', '.css', '.js', '.json', '.webmanifest', '.svg', '.txt', '.aff', '.dic', '.xml'}
    for asset in assets:
        path = SITE / asset
        if path.suffix.lower() in text_extensions:
            data = path.read_bytes()
            normalized = data.replace(b'\r\n', b'\n')
            if normalized != data:
                path.write_bytes(normalized)

    digest = hashlib.sha256()
    for asset in assets:
        digest.update(asset.encode("utf-8"))
        digest.update((SITE / asset).read_bytes())
    digest.update(Path(__file__).read_bytes().replace(b'\r\n', b'\n'))
    digest.update(TEMPLATE.read_bytes().replace(b'\r\n', b'\n'))
    version = RELEASE + "+" + BUILD + "-" + digest.hexdigest()[:12]

    source = TEMPLATE.read_text(encoding="utf-8")
    source = source.replace("__VERSION__", json.dumps(version))
    source = source.replace("__ASSETS__", json.dumps(assets))
    source = source.replace("__DIGESTS__", json.dumps({a:hashlib.sha256((SITE/a).read_bytes()).hexdigest() for a in assets}))
    (SITE / "sw.js").write_bytes(source.encode("utf-8"))
    print("PWA shell:", version, len(assets), "assets")


if __name__ == "__main__":
    main()
