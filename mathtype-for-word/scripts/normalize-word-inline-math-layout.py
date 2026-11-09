from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
O = "urn:schemas-microsoft-com:office:office"
NS = {"w": W, "o": O}
DISPLAY_STYLE_PREFIX = "MTDisplayEquation"
INLINE_MATH_PROGID = "Equation.DSMT4"


def qn(local: str) -> str:
    return f"{{{W}}}{local}"


def paragraph_text(paragraph: etree._Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def ensure_child(parent: etree._Element, tag: str, order: list[str]) -> etree._Element:
    existing = parent.find(f"w:{tag}", NS)
    if existing is not None:
        return existing
    child = etree.Element(qn(tag))
    rank = order.index(tag)
    for index, sibling in enumerate(parent):
        local = etree.QName(sibling).localname
        if local in order and order.index(local) > rank:
            parent.insert(index, child)
            return child
    parent.append(child)
    return child


PPR_ORDER = [
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
]

RPR_ORDER = [
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
]

IND_KEYS = ("firstLine", "firstLineChars", "hanging", "hangingChars")


def _odd_indent_problem(ind: dict[str, str]) -> bool:
    """A 其中， paragraph is wrong when its effective first-line indent is not zero.

    firstLineChars and hangingChars are in 1/100 character units, and take precedence over the
    twip values when present and nonzero. A style indent left in effect counts, because Word
    applies ind attributes independently.
    """
    first_line_chars = ind.get("firstLineChars", "")
    hanging_chars = ind.get("hangingChars", "")
    first_line = ind.get("firstLine", "")
    hanging = ind.get("hanging", "")
    if first_line_chars not in ("", "0"):
        return True
    if hanging_chars not in ("", "0"):
        return True
    if hanging not in ("", "0") and first_line in ("", "0"):
        return True
    return first_line not in ("", "0")


def load_style_map(styles_xml: bytes) -> tuple[dict[str, dict], dict[str, str], dict[str, str]]:
    root = etree.fromstring(styles_xml)
    doc_defaults: dict[str, str] = {}
    ppr_default = root.find("./w:docDefaults/w:pPrDefault/w:pPr", NS)
    if ppr_default is not None:
        ind = ppr_default.find("./w:ind", NS)
        if ind is not None:
            for key in IND_KEYS:
                value = ind.get(qn(key))
                if value is not None:
                    doc_defaults[key] = value
    styles: dict[str, dict] = {}
    default_paragraph = ""
    for style in root.xpath(".//w:style", namespaces=NS):
        style_id = style.get(qn("styleId"), "")
        if not style_id:
            continue
        based = style.find("./w:basedOn", NS)
        ind = style.find("./w:pPr/w:ind", NS)
        entry = {"basedOn": based.get(qn("val"), "") if based is not None else "", "ind": {}}
        if ind is not None:
            for key in IND_KEYS:
                value = ind.get(qn(key))
                if value is not None:
                    entry["ind"][key] = value
        styles[style_id] = entry
        if style.get(qn("type")) == "paragraph" and style.get(qn("default")) in ("1", "true", "on"):
            default_paragraph = style_id
    return styles, doc_defaults, {"default_paragraph": default_paragraph}


def resolve_indent(paragraph: etree._Element, styles: dict[str, dict], doc_defaults: dict[str, str],
                   defaults: dict[str, str]) -> dict[str, str]:
    effective = dict(doc_defaults)
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    style_id = style.get(qn("val"), "") if style is not None else defaults["default_paragraph"]
    chain: list[dict] = []
    visited: set[str] = set()
    while style_id and style_id not in visited:
        visited.add(style_id)
        entry = styles.get(style_id)
        if entry is None:
            break
        chain.append(entry)
        style_id = entry["basedOn"]
    for entry in reversed(chain):
        effective.update(entry["ind"])
    direct = paragraph.find("./w:pPr/w:ind", NS)
    if direct is not None:
        for key in IND_KEYS:
            value = direct.get(qn(key))
            if value is not None:
                effective[key] = value
    return effective


def is_inline_math_paragraph(paragraph: etree._Element) -> bool:
    objects = paragraph.findall(".//w:object", NS)
    if not objects:
        return False
    equation = False
    for obj in objects:
        for ole in obj.findall(".//o:OLEObject", NS):
            if (ole.get("ProgID") or "") == INLINE_MATH_PROGID:
                equation = True
                break
        if equation:
            break
    if not equation:
        return False
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    style_name = style.get(qn("val"), "") if style is not None else ""
    return not style_name.startswith(DISPLAY_STYLE_PREFIX)


def effective_spacing(paragraph: etree._Element, spacing_map: dict[str, str]) -> tuple[str, str]:
    spacing = paragraph.find("./w:pPr/w:spacing", NS)
    line = spacing.get(qn("line")) if spacing is not None else None
    rule = spacing.get(qn("lineRule")) if spacing is not None else None
    if line is None:
        style = paragraph.find("./w:pPr/w:pStyle", NS)
        style_id = style.get(qn("val"), "") if style is not None else ""
        line = spacing_map.get(style_id)
    return (line or "", rule or "")


def load_spacing_map(styles_xml: bytes) -> dict[str, str]:
    root = etree.fromstring(styles_xml)
    values: dict[str, str] = {}
    for style in root.xpath(".//w:style", namespaces=NS):
        style_id = style.get(qn("styleId"), "")
        spacing = style.find("./w:pPr/w:spacing", NS)
        if style_id and spacing is not None and spacing.get(qn("line")):
            values[style_id] = spacing.get(qn("line"))
    return values


def patch_document(document_xml: bytes, styles_xml: bytes, strict: bool) -> tuple[bytes, dict]:
    root = etree.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise ValueError("document.xml has no w:body")
    styles, doc_defaults, defaults = load_style_map(styles_xml)
    spacing_map = load_spacing_map(styles_xml)

    where_changed: list[int] = []
    inline_changed: list[int] = []
    run_position_changed = 0
    line_spacing_relaxed = 0
    skipped_auto_spacing: list[int] = []

    for index, paragraph in enumerate(body.xpath(".//w:p", namespaces=NS), 1):
        text = paragraph_text(paragraph)
        ppr = paragraph.find("w:pPr", NS)
        if ppr is None and (text.startswith("其中，") or (strict and is_inline_math_paragraph(paragraph))):
            ppr = etree.Element(qn("pPr"))
            paragraph.insert(0, ppr)

        if text.startswith("其中，") and ppr is not None:
            effective = resolve_indent(paragraph, styles, doc_defaults, defaults)
            if _odd_indent_problem(effective):
                indent = ensure_child(ppr, "ind", PPR_ORDER)
                # Explicit zero values override every inherited first-line or hanging indent.
                indent.set(qn("firstLine"), "0")
                indent.set(qn("firstLineChars"), "0")
                indent.set(qn("hanging"), "0")
                indent.set(qn("hangingChars"), "0")
                where_changed.append(index)

        if not strict or ppr is None or not is_inline_math_paragraph(paragraph):
            continue

        line, rule = effective_spacing(paragraph, spacing_map)
        if line:
            spacing = ensure_child(ppr, "spacing", PPR_ORDER)
            # auto means 240ths of a line, atLeast and exact mean twips. Converting auto to
            # atLeast by copying the number would change the spacing, so leave auto alone.
            if rule == "exact":
                spacing.set(qn("line"), line)
                spacing.set(qn("lineRule"), "atLeast")
                line_spacing_relaxed += 1
            elif rule in ("", None):
                spacing.set(qn("lineRule"), "atLeast")
                line_spacing_relaxed += 1
            else:
                skipped_auto_spacing.append(index)

        text_alignment = ensure_child(ppr, "textAlignment", PPR_ORDER)
        text_alignment.set(qn("val"), "center")

        for run in paragraph.findall(".//w:r", NS):
            if run.find("w:object", NS) is None:
                continue
            rpr = run.find("w:rPr", NS)
            if rpr is None:
                rpr = etree.Element(qn("rPr"))
                run.insert(0, rpr)
            position = ensure_child(rpr, "position", RPR_ORDER)
            if position.get(qn("val")) != "0":
                run_position_changed += 1
            position.set(qn("val"), "0")
        inline_changed.append(index)

    report = {
        "where_paragraph_count": len(where_changed),
        "where_paragraph_indexes": where_changed,
        "inline_math_paragraph_count": len(inline_changed),
        "inline_math_paragraph_indexes": inline_changed,
        "inline_object_run_positions_normalized": run_position_changed,
        "inline_paragraphs_line_spacing_relaxed": line_spacing_relaxed,
        "inline_paragraphs_auto_spacing_left_alone": skipped_auto_spacing,
    }
    return (
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
        report,
    )


def inspect_document(document_xml: bytes, styles_xml: bytes, strict: bool) -> dict:
    root = etree.fromstring(document_xml)
    styles, doc_defaults, defaults = load_style_map(styles_xml)
    spacing_map = load_spacing_map(styles_xml)
    where_nonexact: list[int] = []
    inline_not_center: list[int] = []
    inline_position_nonzero: list[int] = []
    inline_not_at_least: list[int] = []

    for index, paragraph in enumerate(root.xpath(".//w:body//w:p", namespaces=NS), 1):
        text = paragraph_text(paragraph)
        if text.startswith("其中，"):
            if _odd_indent_problem(resolve_indent(paragraph, styles, doc_defaults, defaults)):
                where_nonexact.append(index)

        if not strict or not is_inline_math_paragraph(paragraph):
            continue
        ppr = paragraph.find("w:pPr", NS)
        alignment = ppr.find("w:textAlignment", NS) if ppr is not None else None
        if alignment is None or alignment.get(qn("val")) != "center":
            inline_not_center.append(index)
        line, rule = effective_spacing(paragraph, spacing_map)
        if rule == "exact":
            inline_not_at_least.append(index)
        for run in paragraph.findall(".//w:r", NS):
            if run.find("w:object", NS) is None:
                continue
            position = run.find("w:rPr/w:position", NS)
            if position is None or position.get(qn("val")) not in (None, "0"):
                inline_position_nonzero.append(index)
                break

    return {
        "where_paragraphs_with_indent": where_nonexact,
        "inline_paragraphs_not_center": inline_not_center,
        "inline_paragraphs_with_nonstandard_position": inline_position_nonzero,
        "inline_paragraphs_with_exact_line_spacing": inline_not_at_least,
        "ok": not where_nonexact
        and (not strict or (not inline_not_center and not inline_position_nonzero and not inline_not_at_least)),
    }


def rewrite_docx(input_path: Path, output_path: Path, strict: bool) -> tuple[dict, dict, dict]:
    with ZipFile(input_path) as source:
        document_xml = source.read("word/document.xml")
        styles_xml = source.read("word/styles.xml")
        patched_xml, report = patch_document(document_xml, styles_xml, strict)
        validation = inspect_document(patched_xml, styles_xml, strict)
        if not validation["ok"]:
            raise ValueError(f"Post-patch validation failed: {validation}")

        ole_before = {
            name: hashlib.sha256(source.read(name)).hexdigest()
            for name in source.namelist()
            if name.startswith("word/embeddings/")
        }
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(
            prefix=f".{output_path.stem}.",
            suffix=".tmp.docx",
            dir=output_path.parent,
            delete=False,
        ) as handle:
            temporary = Path(handle.name)
        try:
            with ZipFile(temporary, "w", ZIP_DEFLATED) as target:
                for item in source.infolist():
                    payload = (
                        patched_xml
                        if item.filename == "word/document.xml"
                        else source.read(item.filename)
                    )
                    target.writestr(item, payload)
            os.replace(temporary, output_path)
        finally:
            temporary.unlink(missing_ok=True)

    with ZipFile(output_path) as output:
        ole_after = {
            name: hashlib.sha256(output.read(name)).hexdigest()
            for name in output.namelist()
            if name.startswith("word/embeddings/")
        }
    if ole_before != ole_after:
        raise ValueError("OLE embedded parts changed during inline-math normalization")
    return report, validation, {"ole_parts_unchanged": True, "ole_count": len(ole_after)}


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Check and fix Chinese inline-math paragraph layout without rebuilding OLE objects. "
            "By default only the 其中， first-line indent is checked; --strict also audits the "
            "inline layout standard (centered text alignment, Normal character position, At Least line spacing)."
        )
    )
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path, nargs="?")
    parser.add_argument("--check", action="store_true")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Include the fork's inline layout rules. MathType's default render uses negative w:position "
             "values on purpose, so this mode reports normal output as non-conforming.",
    )
    args = parser.parse_args()

    input_path = args.input.resolve()
    if args.check:
        with ZipFile(input_path) as archive:
            result = inspect_document(
                archive.read("word/document.xml"),
                archive.read("word/styles.xml"),
                args.strict,
            )
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if not result["ok"]:
            raise SystemExit(1)
        return

    if args.output is None:
        parser.error("output is required unless --check is used")
    output_path = args.output.resolve()
    if input_path == output_path:
        parser.error("input and output must differ; use a staged file and atomic replacement")

    report, validation, preservation = rewrite_docx(input_path, output_path, args.strict)
    print(
        json.dumps(
            {
                "input": str(input_path),
                "output": str(output_path),
                "strict": args.strict,
                "changes": report,
                "validation": validation,
                "preservation": preservation,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
