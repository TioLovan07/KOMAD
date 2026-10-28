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
    assert "word/komad-first-page.xml" not in archive.namelist(), "Obsolete empty header is still present"
    spacer = next((p for p in doc.findall(".//w:body/w:p",ns)
                   if any(b.get("{"+W+"}name")=="komad_first_page_spacer" for b in p.findall("w:bookmarkStart",ns))),None)
    assert spacer is not None, "Missing first-page body spacer"
    spacing=spacer.find("w:pPr/w:spacing",ns)
    assert spacing.get("{"+W+"}line")=="3118", "Unexpected first-page reserved height"
    assert spacing.get("{"+W+"}lineRule")=="exact"
    assert not "".join(t.text or "" for t in spacer.findall(".//w:t",ns)).strip(), "Spacer must not add document text"
    first_text=next(p for p in doc.findall(".//w:body//w:p",ns)
                    if p is not spacer and "".join(t.text or "" for t in p.findall(".//w:t",ns)).strip())
    assert first_text.find("w:pPr/w:spacing",ns).get("{"+W+"}before")=="0", "First text has additional spacing"
    for spacing in doc.findall(".//w:pPr/w:spacing",ns):
        if spacing.get("{"+W+"}lineRule")=="auto":
            assert spacing.get("{"+W+"}line")=="240", "Body is not single-spaced"
print("PASS: F4 portrait, 2.5 cm margins, first-page-only title space, single-spaced body.")
