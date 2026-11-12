"""Read a PPTX in presentation order. Print a page index or text for selected pages."""

import argparse
from pathlib import Path
import posixpath
import sys
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


def slide_parts(archive):
    package_rels = ET.fromstring(archive.read("_rels/.rels"))
    main = next(
        (rel.get("Target") for rel in package_rels
         if rel.get("Type", "").endswith("/officeDocument")),
        None,
    )
    if not main:
        raise ValueError("找不到演示文稿主文档。")
    main = main.lstrip("/")
    directory, filename = posixpath.split(main)
    relname = posixpath.join(directory, "_rels", filename + ".rels")
    rels = ET.fromstring(archive.read(relname))
    targets = {
        rel.get("Id"): posixpath.normpath(posixpath.join(directory, rel.attrib["Target"])).lstrip("/")
        for rel in rels
        if rel.get("Type", "").endswith("/slide") and rel.get("TargetMode") != "External"
    }
    main_xml = ET.fromstring(archive.read(main))
    parts = []
    for slide_id in main_xml.findall("./{*}sldIdLst/{*}sldId"):
        relation = next((v for k, v in slide_id.attrib.items() if k.endswith("}id")), None)
        if relation not in targets:
            raise ValueError("幻灯片与文稿关系表不一致。")
        parts.append(targets[relation])
    return parts


def page_lines(archive, part):
    root = ET.fromstring(archive.read(part))
    shape_tree = root.find("./{*}cSld/{*}spTree")
    if shape_tree is None:
        return []
    lines = []
    for para in shape_tree.findall(".//{*}p"):
        value = "".join(node.text or "" for node in para.findall(".//{*}t")).strip()
        if value:
            lines.append(value)
    return lines


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path, help="源 PPTX 文件")
    parser.add_argument("--page", type=int, nargs="+", help="列出这些页的文字，从 1 开始计数")
    args = parser.parse_args()
    try:
        with ZipFile(args.pptx) as archive:
            parts = slide_parts(archive)
            if args.page and any(n < 1 or n > len(parts) for n in args.page):
                parser.error(f"页码范围是 1–{len(parts)}。")
            print(f"{args.pptx.name}  共 {len(parts)} 页")
            for number in args.page or range(1, len(parts) + 1):
                lines = page_lines(archive, parts[number - 1])
                if args.page:
                    print(f"\n第 {number} 页")
                    print("\n".join(lines) or "请渲染查看本页。")
                else:
                    summary = " ｜ ".join(lines[:2])[:130] or "请渲染查看本页。"
                    print(f"{number:>3}   {summary}")
    except (OSError, BadZipFile, KeyError, ValueError, ET.ParseError) as error:
        parser.exit(1, f"无法读取 PPTX：{error}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
