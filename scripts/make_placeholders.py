from __future__ import annotations

import argparse
import json
import re
import shutil
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
LATEX_RE = re.compile(r"\$([^$\r\n]+)\$")


def convert_docx(input_docx: Path, output_docx: Path, map_json: Path) -> int:
    tmp = output_docx.with_suffix(output_docx.suffix + ".tmp")
    mapping: list[dict[str, str]] = []
    with ZipFile(input_docx, "r") as zin, ZipFile(tmp, "w", ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                root = etree.fromstring(data)
                counter = 1

                for t in root.xpath("//w:t", namespaces=NS):
                    if not t.text or "$" not in t.text:
                        continue

                    def repl(match: re.Match[str]) -> str:
                        nonlocal counter
                        placeholder = f"@@MT{counter:04d}@@"
                        mapping.append({"placeholder": placeholder, "latex": match.group(1).strip()})
                        counter += 1
                        return placeholder

                    t.text = LATEX_RE.sub(repl, t.text)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    map_json.write_text(json.dumps(mapping, ensure_ascii=False, indent=2), encoding="utf-8")
    shutil.move(str(tmp), str(output_docx))
    return len(mapping)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--map-json", required=True, type=Path)
    args = parser.parse_args()

    count = convert_docx(args.input.resolve(), args.output.resolve(), args.map_json.resolve())
    print(f"output={args.output.resolve()}")
    print(f"map_json={args.map_json.resolve()}")
    print(f"count={count}")


if __name__ == "__main__":
    main()
