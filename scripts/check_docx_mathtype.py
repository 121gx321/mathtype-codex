from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "o": "urn:schemas-microsoft-com:office:office",
}


def inspect(docx: Path, paragraphs: list[int]) -> None:
    with tempfile.TemporaryDirectory() as td:
        temp = Path(td) / "check.docx"
        shutil.copyfile(docx, temp)
        with ZipFile(temp) as z:
            root = etree.fromstring(z.read("word/document.xml"))

    text = "".join(root.xpath("//w:t/text()", namespaces=NS))
    paras = root.xpath("//w:body/w:p", namespaces=NS)
    ole = root.xpath("//o:OLEObject", namespaces=NS)
    dsmt4 = root.xpath("//o:OLEObject[@ProgID='Equation.DSMT4']", namespaces=NS)
    mt_display = root.xpath("//w:p[w:pPr/w:pStyle/@w:val='MTDisplayEquation']", namespaces=NS)

    print(f"ole_objects={len(ole)}")
    print(f"equation_dsmt4={len(dsmt4)}")
    print(f"mt_display_paragraphs={len(mt_display)}")
    print(f"placeholder_remaining={text.count('@@MT')}")
    print(f"latex_dollar_remaining={text.count('$')}")
    for idx in paragraphs:
        if idx < 0 or idx >= len(paras):
            print(f"paragraph={idx} missing=true")
            continue
        p = paras[idx]
        p_text = "".join(p.xpath(".//w:t/text()", namespaces=NS))
        p_ole = p.xpath(".//o:OLEObject", namespaces=NS)
        print(f"paragraph={idx} ole={len(p_ole)} text_preview={p_text[:160]!r}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True, type=Path)
    parser.add_argument("--paragraphs", default="", help="Comma-separated zero-based paragraph indexes to inspect.")
    args = parser.parse_args()

    paragraphs = [int(x) for x in args.paragraphs.split(",") if x.strip()]
    inspect(args.docx.resolve(), paragraphs)


if __name__ == "__main__":
    main()
