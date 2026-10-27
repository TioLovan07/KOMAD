"""Verify exported KOMAD DOCX layout, without printing document contents.
Usage: python scripts/verify-word-layout.py path/to/Draft.docx
"""
import sys
from zipfile import ZipFile
from xml.etree import ElementTree as ET

W="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ns={"w":W}
with ZipFile(sys.argv[1]) as archive:
    doc=ET.fromstring(archive.read("word/document.xml"))
    sections=doc.findall(".//w:sectPr",ns)
    assert sections, "No Word sections"
    for section in sections:
        margins=section.find("w:pgMar",ns)
        for side in ["top","left","right","bottom"]:
            assert abs(int(margins.get("{"+W+"}"+side))-1417)<=1, "Unexpected margin"
        size=section.find("w:pgSz",ns)
        assert abs(int(size.get("{"+W+"}w"))-12191)<=1
        assert abs(int(size.get("{"+W+"}h"))-18711)<=1
    assert sections[0].find("w:titlePg",ns) is not None, "Missing first-page layout"
    assert all(s.find("w:titlePg",ns) is None for s in sections[1:]), "Title spacer repeats on a later section"
    header=ET.fromstring(archive.read("word/komad-first-page.xml"))
    spacing=header.find(".//w:spacing",ns)
    assert spacing.get("{"+W+"}line")=="3969", "Unexpected first-page spacer"
    assert spacing.get("{"+W+"}lineRule")=="exact"
    for spacing in doc.findall(".//w:pPr/w:spacing",ns):
        if spacing.get("{"+W+"}lineRule")=="auto":
            assert spacing.get("{"+W+"}line")=="240", "Body is not single-spaced"
print("PASS: F4 portrait, 2.5 cm margins, first-page-only title space, single-spaced body.")
