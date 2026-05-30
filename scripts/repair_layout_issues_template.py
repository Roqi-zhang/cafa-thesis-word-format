#!/usr/bin/env python3
"""Parameterized repair template for CAFA thesis DOCX problem-fix passes.

This template is intentionally generic. It contains no private filenames,
student data, thesis excerpts, generated reports, or hard-coded working
directories. Use it only after the user has confirmed a repair plan.

Typical repairs covered:
- merge safely detected fragmented native footnotes into complete notes;
- add centered Arabic PAGE fields in footers;
- make major chapters start on a new page;
- normalize figure/caption/source formatting and figure numbering;
- keep legacy floating picture shapes from overlapping text.

The script relies on OOXML for footnote reconstruction and on Microsoft Word
COM for page/layout operations. If Word COM is unavailable, the script still
performs OOXML footnote repair and reports skipped layout tasks.
"""

from __future__ import annotations

import argparse
import hashlib
import re
import shutil
import sys
import time
import zipfile
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS = {"w": W_NS, "r": R_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)


def qn(name: str) -> str:
    prefix, local = name.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def wtag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


CIRCLED_MARKS = [
    "",
    "①",
    "②",
    "③",
    "④",
    "⑤",
    "⑥",
    "⑦",
    "⑧",
    "⑨",
    "⑩",
    "⑪",
    "⑫",
    "⑬",
    "⑭",
    "⑮",
    "⑯",
    "⑰",
    "⑱",
    "⑲",
    "⑳",
    "㉑",
    "㉒",
    "㉓",
    "㉔",
    "㉕",
    "㉖",
    "㉗",
    "㉘",
    "㉙",
    "㉚",
    "㉛",
    "㉜",
    "㉝",
    "㉞",
    "㉟",
    "㊱",
    "㊲",
    "㊳",
    "㊴",
    "㊵",
]

FIGURE_TITLE_RE = re.compile(r"^图\s*([0-9A-Za-z]+)[-－—](\d+)\s*(.*)$")
FIGURE_EN_RE = re.compile(r"^Figure\s*\d+\s*(.*)$", re.I)
SOURCE_RE = re.compile(r"^图片来源\s*[:：]\s*(.*)$")
HEADING_NUM_RE = re.compile(r"^([1-9])\s+")
APPENDIX_RE = re.compile(r"^附录\s*([A-ZＡ-Ｚ])")
BODY_FIG_REF_RE = re.compile(r"(如图|见图|图)\s*[0-9A-Za-z]+[-－—][0-9]+")


WD_ALIGN_LEFT = 0
WD_ALIGN_CENTER = 1
WD_FIELD_PAGE = 33
WD_HEADER_FOOTER_PRIMARY = 1
WD_HEADER_FOOTER_FIRST_PAGE = 2
WD_HEADER_FOOTER_EVEN_PAGES = 3
WD_LINE_SPACE_SINGLE = 0
WD_LINE_SPACE_EXACTLY = 4
WD_STATISTIC_PAGES = 2
WD_WRAP_TOP_BOTTOM = 4
WD_REL_H_MARGIN = 0
WD_REL_V_PARAGRAPH = 2
WD_SHAPE_CENTER = -999995
MSO_TRUE = -1
MSO_FALSE = 0


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def register_namespaces_from_xml(data: bytes) -> None:
    for prefix_b, uri_b in re.findall(br'xmlns:([A-Za-z0-9_]+)="([^"]+)"', data[:20000]):
        prefix = prefix_b.decode("utf-8", errors="ignore")
        uri = uri_b.decode("utf-8", errors="ignore")
        if prefix and prefix.lower() != "xml" and uri:
            try:
                ET.register_namespace(prefix, uri)
            except ValueError:
                pass


def trim_ignorable_prefixes(root: ET.Element, allowed: set[str]) -> None:
    ignorable = f"{{{MC_NS}}}Ignorable"
    current = root.get(ignorable)
    if not current:
        return
    kept = [prefix for prefix in current.split() if prefix in allowed]
    if kept:
        root.set(ignorable, " ".join(kept))
    else:
        root.attrib.pop(ignorable, None)


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(".//w:t", NS))


def strip_leading_marker(text: str) -> str:
    text = (text or "").strip()
    if text and text[0] in set(CIRCLED_MARKS[1:]):
        return text[1:].strip()
    return text


def is_fragment_text(text: str) -> bool:
    body = strip_leading_marker(text)
    if not body:
        return True
    if len(body) <= 2 and all(ch in "，。、；：,.》）)]】 " for ch in body):
        return True
    if len(body) <= 8:
        return True
    return False


def merge_fragment_texts(parts: Sequence[str]) -> str:
    return "".join(strip_leading_marker(part) for part in parts if strip_leading_marker(part)).strip()


def set_run_font(run: ET.Element, size_half_points: int, superscript: bool) -> None:
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(wtag("rPr"))
        run.insert(0, rpr)
    for tag in ("w:b", "w:bCs", "w:i", "w:iCs"):
        node = rpr.find(tag, NS)
        if node is not None:
            rpr.remove(node)
    fonts = rpr.find("w:rFonts", NS)
    if fonts is None:
        fonts = ET.SubElement(rpr, wtag("rFonts"))
    fonts.set(wtag("eastAsia"), "宋体")
    fonts.set(wtag("ascii"), "Times New Roman")
    fonts.set(wtag("hAnsi"), "Times New Roman")
    fonts.set(wtag("cs"), "Times New Roman")
    for tag in ("sz", "szCs"):
        node = rpr.find(f"w:{tag}", NS)
        if node is None:
            node = ET.SubElement(rpr, wtag(tag))
        node.set(wtag("val"), str(size_half_points))
    color = rpr.find("w:color", NS)
    if color is None:
        color = ET.SubElement(rpr, wtag("color"))
    color.set(wtag("val"), "000000")
    vert = rpr.find("w:vertAlign", NS)
    if superscript:
        vert = vert or ET.SubElement(rpr, wtag("vertAlign"))
        vert.set(wtag("val"), "superscript")
    elif vert is not None:
        rpr.remove(vert)


def make_text_run(text: str, preserve: bool = False) -> ET.Element:
    run = ET.Element(wtag("r"))
    set_run_font(run, 18, superscript=False)
    node = ET.SubElement(run, wtag("t"))
    if preserve or text.startswith(" ") or text.endswith(" "):
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text
    return run


def make_footnote(note_id: int, marker: str, note_text: str) -> ET.Element:
    footnote = ET.Element(wtag("footnote"))
    footnote.set(wtag("id"), str(note_id))
    paragraph = ET.SubElement(footnote, wtag("p"))
    ppr = ET.SubElement(paragraph, wtag("pPr"))
    ET.SubElement(ppr, wtag("pStyle")).set(wtag("val"), "FootnoteText")
    spacing = ET.SubElement(ppr, wtag("spacing"))
    spacing.set(wtag("before"), "0")
    spacing.set(wtag("after"), "0")
    spacing.set(wtag("line"), "240")
    spacing.set(wtag("lineRule"), "auto")
    indent = ET.SubElement(ppr, wtag("ind"))
    indent.set(wtag("left"), "0")
    indent.set(wtag("firstLine"), "0")
    ET.SubElement(ppr, wtag("jc")).set(wtag("val"), "left")
    paragraph.append(make_text_run(f"{marker} ", preserve=True))
    paragraph.append(make_text_run(note_text))
    return footnote


@dataclass
class FootnoteRepairResult:
    initial_refs: int = 0
    initial_notes: int = 0
    final_notes: int = 0
    merged_groups: List[Tuple[List[int], str]] = field(default_factory=list)
    removed_empty_refs: List[int] = field(default_factory=list)
    uncertain_groups: List[Tuple[List[int], List[str]]] = field(default_factory=list)


def parent_map(root: ET.Element) -> Dict[ET.Element, ET.Element]:
    return {child: parent for parent in root.iter() for child in parent}


def repair_fragmented_footnotes(docx_path: Path) -> FootnoteRepairResult:
    result = FootnoteRepairResult()
    with zipfile.ZipFile(docx_path, "r") as package:
        parts = {name: package.read(name) for name in package.namelist()}
    if "word/document.xml" not in parts or "word/footnotes.xml" not in parts:
        return result

    register_namespaces_from_xml(parts["word/document.xml"])
    register_namespaces_from_xml(parts["word/footnotes.xml"])
    document_root = ET.fromstring(parts["word/document.xml"])
    footnotes_root = ET.fromstring(parts["word/footnotes.xml"])
    trim_ignorable_prefixes(document_root, {"w14", "wp14"})
    trim_ignorable_prefixes(footnotes_root, {"w14"})

    note_text: Dict[int, str] = {}
    special_notes: List[ET.Element] = []
    for footnote in footnotes_root.findall("w:footnote", NS):
        raw_id = footnote.get(wtag("id"))
        if not raw_id or not raw_id.lstrip("-").isdigit():
            continue
        note_id = int(raw_id)
        if note_id <= 0 or footnote.get(wtag("type")):
            special_notes.append(deepcopy(footnote))
        else:
            note_text[note_id] = paragraph_text(footnote)

    result.initial_notes = len(note_text)
    all_refs: List[int] = []
    paragraph_ref_groups: List[List[int]] = []
    for paragraph in document_root.findall(".//w:body/w:p", NS):
        ids: List[int] = []
        for ref in paragraph.findall(".//w:footnoteReference", NS):
            raw_id = ref.get(wtag("id"))
            if raw_id and raw_id.lstrip("-").isdigit() and int(raw_id) > 0:
                ids.append(int(raw_id))
                all_refs.append(int(raw_id))
        if len(ids) > 1:
            paragraph_ref_groups.append(ids)
    result.initial_refs = len(all_refs)

    empty_ids = {note_id for note_id, text in note_text.items() if not strip_leading_marker(text)}
    merge_groups: List[List[int]] = []
    for ids in paragraph_ref_groups:
        if ids != list(range(ids[0], ids[-1] + 1)):
            result.uncertain_groups.append((ids, [note_text.get(note_id, "") for note_id in ids]))
            continue
        parts_for_group = [note_text.get(note_id, "") for note_id in ids]
        bodies = [strip_leading_marker(item) for item in parts_for_group]
        first_open = bodies[0].count("（") > bodies[0].count("）") or bodies[0].count("《") > bodies[0].count("》")
        tail_fragmenty = all(is_fragment_text(item) or strip_leading_marker(item).startswith(("，", "、", "-", "》")) for item in parts_for_group[1:])
        if bodies[0] and (first_open or tail_fragmenty):
            merged = merge_fragment_texts(parts_for_group)
            merge_groups.append(ids)
            result.merged_groups.append((ids, merged))
        else:
            result.uncertain_groups.append((ids, parts_for_group))

    group_for_old_id: Dict[int, List[int]] = {}
    for group in merge_groups:
        for old_id in group:
            group_for_old_id[old_id] = group

    final_items: List[Tuple[List[int], str]] = []
    seen_groups: set[Tuple[int, ...]] = set()
    for old_id in all_refs:
        if old_id in empty_ids:
            continue
        if old_id in group_for_old_id:
            group = tuple(group_for_old_id[old_id])
            if group in seen_groups:
                continue
            seen_groups.add(group)
            final_items.append((list(group), merge_fragment_texts(note_text.get(item, "") for item in group)))
        elif old_id in note_text:
            final_items.append(([old_id], strip_leading_marker(note_text[old_id])))

    old_to_new: Dict[int, int] = {}
    for new_id, (old_ids, _) in enumerate(final_items, start=1):
        for old_id in old_ids:
            old_to_new[old_id] = new_id

    parents = parent_map(document_root)
    runs_to_delete: List[ET.Element] = []
    for ref in document_root.findall(".//w:footnoteReference", NS):
        raw_id = ref.get(wtag("id"))
        if not raw_id or not raw_id.lstrip("-").isdigit() or int(raw_id) <= 0:
            continue
        old_id = int(raw_id)
        run = parents.get(ref)
        if run is None or run.tag != wtag("r"):
            continue
        if old_id in empty_ids:
            runs_to_delete.append(run)
            result.removed_empty_refs.append(old_id)
            continue
        if old_id in group_for_old_id and old_id != group_for_old_id[old_id][0]:
            runs_to_delete.append(run)
            continue
        new_id = old_to_new.get(old_id)
        if new_id is None:
            continue
        marker = CIRCLED_MARKS[new_id] if new_id < len(CIRCLED_MARKS) else f"({new_id})"
        ref.set(wtag("id"), str(new_id))
        ref.set(wtag("customMarkFollows"), "1")
        set_run_font(run, 18, superscript=True)
        for node in list(run.findall("w:t", NS)):
            run.remove(node)
        ET.SubElement(run, wtag("t")).text = marker

    parents = parent_map(document_root)
    for run in runs_to_delete:
        parent = parents.get(run)
        if parent is not None:
            try:
                parent.remove(run)
            except ValueError:
                pass

    new_footnotes_root = ET.Element(footnotes_root.tag, footnotes_root.attrib)
    for footnote in special_notes:
        new_footnotes_root.append(footnote)
    for new_id, (_, text) in enumerate(final_items, start=1):
        marker = CIRCLED_MARKS[new_id] if new_id < len(CIRCLED_MARKS) else f"({new_id})"
        new_footnotes_root.append(make_footnote(new_id, marker, text))
    result.final_notes = len(final_items)

    replacements = {
        "word/document.xml": ET.tostring(document_root, encoding="utf-8", xml_declaration=True),
        "word/footnotes.xml": ET.tostring(new_footnotes_root, encoding="utf-8", xml_declaration=True),
    }
    rewrite_docx(docx_path, docx_path, replacements)
    return result


def rewrite_docx(source: Path, target: Path, replacements: Dict[str, bytes]) -> None:
    temp = target.with_suffix(".tmp.docx")
    with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replacements.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)
    temp.replace(target)


def clean_word_text(text: str) -> str:
    return text.replace("\r", "").replace("\x07", "").strip()


def para_text(para) -> str:
    return clean_word_text(str(para.Range.Text))


def is_caption(text: str) -> bool:
    return bool(FIGURE_TITLE_RE.match(text) or FIGURE_EN_RE.match(text))


def is_source_line(text: str) -> bool:
    return bool(SOURCE_RE.match(text))


def set_word_font(rng, size: float, bold: bool = False) -> None:
    try:
        rng.Font.NameFarEast = "宋体"
        rng.Font.NameAscii = "Times New Roman"
        rng.Font.NameOther = "Times New Roman"
    except Exception:
        pass
    rng.Font.Size = size
    rng.Font.Bold = MSO_TRUE if bold else MSO_FALSE
    rng.Font.Italic = MSO_FALSE
    rng.Font.Color = 0


def format_caption_para(para) -> None:
    set_word_font(para.Range, 12, bold=False)
    pf = para.Range.ParagraphFormat
    pf.Alignment = WD_ALIGN_CENTER
    pf.LineSpacingRule = WD_LINE_SPACE_EXACTLY
    pf.LineSpacing = 30
    pf.SpaceBefore = 0
    pf.SpaceAfter = 0
    pf.FirstLineIndent = 0
    pf.LeftIndent = 0
    pf.RightIndent = 0


@dataclass
class WordRepairResult:
    word_com_available: bool = False
    skipped_reason: str = ""
    page_number_footers: int = 0
    major_headings: List[str] = field(default_factory=list)
    initial_shapes: int = 0
    converted_shapes: int = 0
    processed_shapes: int = 0
    processed_inline_shapes: int = 0
    resized_images: int = 0
    figure_caption_map: List[Tuple[str, str]] = field(default_factory=list)
    missing_sources: List[str] = field(default_factory=list)
    body_figure_refs: List[str] = field(default_factory=list)
    page_count: int = 0


def add_page_numbers(doc) -> int:
    edited = 0
    for section in doc.Sections:
        for footer_type in (WD_HEADER_FOOTER_PRIMARY, WD_HEADER_FOOTER_FIRST_PAGE, WD_HEADER_FOOTER_EVEN_PAGES):
            try:
                footer = section.Footers(footer_type)
                footer.LinkToPrevious = False
                footer.Range.Text = ""
                footer.Range.ParagraphFormat.Alignment = WD_ALIGN_CENTER
                footer.PageNumbers.RestartNumberingAtSection = False
                footer.PageNumbers.NumberStyle = 0
                footer.PageNumbers.StartingNumber = 1
                rng = footer.Range
                rng.Collapse(0)
                doc.Fields.Add(rng, WD_FIELD_PAGE)
                set_word_font(footer.Range, 9, bold=False)
                footer.Range.ParagraphFormat.Alignment = WD_ALIGN_CENTER
                edited += 1
            except Exception:
                pass
    return edited


def apply_major_heading_breaks(doc, start_heading: str) -> List[str]:
    changed: List[str] = []
    in_body = False
    for para in doc.Paragraphs:
        text = para_text(para)
        if not text:
            continue
        if text == start_heading or text.startswith(start_heading):
            in_body = True
        if not in_body:
            continue
        major = bool(HEADING_NUM_RE.match(text) or APPENDIX_RE.match(text) or text.startswith(("结论", "结语", "参考文献", "致谢")))
        if major:
            para.Range.ParagraphFormat.PageBreakBefore = MSO_TRUE
            para.Range.ParagraphFormat.KeepWithNext = MSO_TRUE
            changed.append(text)
    return changed


def normalize_source_line(text: str) -> str:
    match = SOURCE_RE.match(text)
    source = match.group(1).strip() if match else text.strip()
    if source.startswith("《") and source.endswith("》"):
        inner = source[1:-1].strip()
        if inner in {"作者", "作者自绘", "作者绘制", "作者自摄", "作者拍摄", "自绘", "自摄"}:
            source = inner
    return f"图片来源：{source}" if source else "图片来源："


def replace_para_text(para, new_text: str) -> None:
    rng = para.Range
    rng.End = rng.End - 1
    rng.Text = new_text


def current_caption_name(text: str) -> Tuple[Optional[str], str]:
    match = FIGURE_TITLE_RE.match(text)
    if match:
        return match.group(1).upper(), match.group(3).strip()
    match = FIGURE_EN_RE.match(text)
    if match:
        return None, match.group(1).strip()
    return None, text.strip()


def normalize_captions(doc) -> Tuple[List[Tuple[str, str]], List[str]]:
    current_chapter: Optional[str] = None
    counters: Dict[str, int] = {}
    mapping: List[Tuple[str, str]] = []
    missing_sources: List[str] = []
    full_width = str.maketrans("ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ", "ABCDEFGHIJKLMNOPQRSTUVWXYZ")

    for idx in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs(idx)
        text = para_text(para)
        if not text:
            continue
        heading = HEADING_NUM_RE.match(text)
        appendix = APPENDIX_RE.match(text)
        if heading:
            current_chapter = heading.group(1)
        elif appendix:
            current_chapter = appendix.group(1).upper().translate(full_width)
        if not is_caption(text):
            continue

        old_chapter, name = current_caption_name(text)
        chapter = current_chapter or old_chapter or "?"
        counters[chapter] = counters.get(chapter, 0) + 1
        new_text = f"图 {chapter}-{counters[chapter]} {name}".strip()
        replace_para_text(para, new_text)
        format_caption_para(para)
        para.Range.ParagraphFormat.KeepWithNext = MSO_TRUE
        mapping.append((text, new_text))

        try:
            source_para = doc.Paragraphs(idx + 1)
            source_text = para_text(source_para)
            if is_source_line(source_text):
                replace_para_text(source_para, normalize_source_line(source_text))
                format_caption_para(source_para)
                source_para.Range.ParagraphFormat.KeepTogether = MSO_TRUE
            else:
                missing_sources.append(new_text)
        except Exception:
            missing_sources.append(new_text)
    return mapping, missing_sources


def process_images(doc) -> Tuple[int, int, int, int]:
    processed_inline = 0
    processed_shapes = 0
    converted_shapes = 0
    resized = 0
    initial_shape_count = doc.Shapes.Count

    for idx in range(doc.Shapes.Count, 0, -1):
        try:
            shape = doc.Shapes(idx)
            try:
                shape.LockAspectRatio = MSO_TRUE
                shape.WrapFormat.AllowOverlap = MSO_FALSE
            except Exception:
                pass
            try:
                shape.ConvertToInlineShape()
                converted_shapes += 1
            except Exception:
                processed_shapes += 1
                shape.WrapFormat.Type = WD_WRAP_TOP_BOTTOM
                shape.WrapFormat.AllowOverlap = MSO_FALSE
                shape.RelativeHorizontalPosition = WD_REL_H_MARGIN
                shape.RelativeVerticalPosition = WD_REL_V_PARAGRAPH
                shape.Left = WD_SHAPE_CENTER
                shape.Top = 0
                shape.Anchor.Paragraphs(1).Range.ParagraphFormat.Alignment = WD_ALIGN_CENTER
                shape.Anchor.Paragraphs(1).Range.ParagraphFormat.KeepWithNext = MSO_TRUE
        except Exception:
            pass

    for idx in range(1, doc.InlineShapes.Count + 1):
        try:
            inline = doc.InlineShapes(idx)
            processed_inline += 1
            inline.LockAspectRatio = MSO_TRUE
            page_setup = inline.Range.Sections(1).PageSetup
            max_width = float(page_setup.PageWidth - page_setup.LeftMargin - page_setup.RightMargin)
            max_height = float(page_setup.PageHeight - page_setup.TopMargin - page_setup.BottomMargin - 90)
            if float(inline.Width) > max_width:
                inline.Width = max_width
                resized += 1
            if float(inline.Height) > max_height:
                inline.Height = max_height
                resized += 1
            pf = inline.Range.Paragraphs(1).Range.ParagraphFormat
            pf.Alignment = WD_ALIGN_CENTER
            pf.KeepWithNext = MSO_TRUE
            pf.KeepTogether = MSO_TRUE
        except Exception:
            pass
    return initial_shape_count, converted_shapes, processed_shapes, processed_inline + processed_shapes, resized


def detect_body_figure_refs(doc) -> List[str]:
    refs: List[str] = []
    for para in doc.Paragraphs:
        text = para_text(para)
        if not text or is_caption(text) or is_source_line(text):
            continue
        if BODY_FIG_REF_RE.search(text):
            refs.append(text[:180])
    return refs


def apply_word_repairs(docx_path: Path, args: argparse.Namespace) -> WordRepairResult:
    result = WordRepairResult()
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as exc:
        result.skipped_reason = f"pywin32/Word COM unavailable: {exc}"
        return result

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=False, AddToRecentFiles=False, ConfirmConversions=False, NoEncodingDialog=True)
        result.word_com_available = True
        if not args.skip_images:
            (
                result.initial_shapes,
                result.converted_shapes,
                result.processed_shapes,
                image_count,
                result.resized_images,
            ) = process_images(doc)
            result.processed_inline_shapes = image_count - result.processed_shapes
        if not args.skip_page_numbers:
            result.page_number_footers = add_page_numbers(doc)
        if not args.skip_heading_breaks:
            result.major_headings = apply_major_heading_breaks(doc, args.major_heading_start)
        for idx in range(1, doc.Footnotes.Count + 1):
            try:
                set_word_font(doc.Footnotes(idx).Range, 9, bold=False)
                doc.Footnotes(idx).Range.ParagraphFormat.Alignment = WD_ALIGN_LEFT
                doc.Footnotes(idx).Range.ParagraphFormat.LineSpacingRule = WD_LINE_SPACE_SINGLE
            except Exception:
                pass
        if not args.skip_captions:
            result.figure_caption_map, result.missing_sources = normalize_captions(doc)
        result.body_figure_refs = detect_body_figure_refs(doc)
        doc.Repaginate()
        result.page_count = doc.ComputeStatistics(WD_STATISTIC_PAGES)
        doc.Save()
    except Exception as exc:
        result.skipped_reason = f"Word COM repair failed: {exc}"
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return result


@dataclass
class ValidationSummary:
    source_hash_before: str
    source_hash_after: str
    target_hash: str
    body_refs: int = 0
    native_notes: int = 0
    stacked_ref_paragraphs: int = 0
    abnormal_notes: int = 0
    figure_residuals: int = 0
    footer_has_page_field: bool = False
    footer_has_literal_page_text: bool = False


def validate(source: Path, target: Path, source_hash_before: str) -> ValidationSummary:
    summary = ValidationSummary(source_hash_before, sha256(source), sha256(target))
    with zipfile.ZipFile(target, "r") as package:
        document_root = ET.fromstring(package.read("word/document.xml"))
        footnotes_root = ET.fromstring(package.read("word/footnotes.xml"))
        for idx, paragraph in enumerate(document_root.findall(".//w:body/w:p", NS), start=1):
            text = paragraph_text(paragraph).strip()
            if text.lower().startswith("figure"):
                summary.figure_residuals += 1
            refs = [
                ref.get(wtag("id"))
                for ref in paragraph.findall(".//w:footnoteReference", NS)
                if ref.get(wtag("id")) and ref.get(wtag("id")).lstrip("-").isdigit() and int(ref.get(wtag("id"))) > 0
            ]
            summary.body_refs += len(refs)
            if len(refs) > 1:
                summary.stacked_ref_paragraphs += 1
        for footnote in footnotes_root.findall("w:footnote", NS):
            raw_id = footnote.get(wtag("id"))
            if not raw_id or not raw_id.isdigit() or int(raw_id) <= 0:
                continue
            summary.native_notes += 1
            text = paragraph_text(footnote).strip()
            body = strip_leading_marker(text)
            if not body or len(body) <= 1 or all(ch in "，。、；：,.》）)]】 " for ch in body):
                summary.abnormal_notes += 1
        for name in package.namelist():
            if not (name.startswith("word/footer") and name.endswith(".xml")):
                continue
            data = package.read(name)
            if b"PAGE" in data or b"fldChar" in data:
                summary.footer_has_page_field = True
            footer_text = paragraph_text(ET.fromstring(data))
            if "页码" in footer_text:
                summary.footer_has_literal_page_text = True
    return summary


def write_report(path: Optional[Path], lines: List[str]) -> None:
    if path is None:
        print("\n".join(lines))
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def build_report(
    source: Path,
    target: Path,
    backup: Optional[Path],
    footnotes: FootnoteRepairResult,
    word: WordRepairResult,
    validation: ValidationSummary,
) -> List[str]:
    lines = [
        "# CAFA Thesis Problem Repair Report",
        "",
        f"- Generated: {time.strftime('%Y-%m-%d %H:%M:%S')}",
        f"- Source: {source}",
        f"- Target: {target}",
        f"- Backup: {backup or 'not requested'}",
        "",
        "## Footnotes",
        f"- Initial body refs: {footnotes.initial_refs}",
        f"- Initial native notes: {footnotes.initial_notes}",
        f"- Final native notes: {footnotes.final_notes}",
        f"- Merged fragment groups: {len(footnotes.merged_groups)}",
        f"- Empty refs removed: {sorted(set(footnotes.removed_empty_refs)) or 'none'}",
        f"- Uncertain groups left unchanged/reported: {len(footnotes.uncertain_groups)}",
        "",
        "## Word Layout",
        f"- Word COM available: {word.word_com_available}",
        f"- Skipped/failure reason: {word.skipped_reason or 'none'}",
        f"- Page-number footer edits: {word.page_number_footers}",
        f"- Major headings set to new page: {len(word.major_headings)}",
        f"- Initial floating shapes: {word.initial_shapes}",
        f"- Converted shapes: {word.converted_shapes}",
        f"- Conservatively processed legacy shapes: {word.processed_shapes}",
        f"- Processed inline shapes: {word.processed_inline_shapes}",
        f"- Image resize operations: {word.resized_images}",
        f"- Caption mappings: {len(word.figure_caption_map)}",
        f"- Captions missing source line: {len(word.missing_sources)}",
        f"- Body figure refs found but not changed: {len(word.body_figure_refs)}",
        f"- Word page count: {word.page_count}",
        "",
        "## Validation",
        f"- Source hash before: {validation.source_hash_before}",
        f"- Source hash after: {validation.source_hash_after}",
        f"- Source unchanged: {validation.source_hash_before == validation.source_hash_after}",
        f"- Target hash: {validation.target_hash}",
        f"- Body refs: {validation.body_refs}",
        f"- Native notes: {validation.native_notes}",
        f"- Stacked ref paragraphs: {validation.stacked_ref_paragraphs}",
        f"- Abnormal notes: {validation.abnormal_notes}",
        f"- Figure residuals: {validation.figure_residuals}",
        f"- Footer has PAGE field: {validation.footer_has_page_field}",
        f"- Footer still has literal page text: {validation.footer_has_literal_page_text}",
        "",
        "Manual visual checks remain required for actual Word pagination, image clarity, and any missing source data.",
    ]
    return lines


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Repair common DOCX layout issues in confirmed CAFA thesis workflows.")
    parser.add_argument("--source-docx", required=True, type=Path)
    parser.add_argument("--target-docx", required=True, type=Path)
    parser.add_argument("--report-md", type=Path)
    parser.add_argument("--validation-report-md", type=Path)
    parser.add_argument("--backup-docx", type=Path)
    parser.add_argument("--major-heading-start", default="1 引言")
    parser.add_argument("--skip-footnotes", action="store_true")
    parser.add_argument("--skip-page-numbers", action="store_true")
    parser.add_argument("--skip-heading-breaks", action="store_true")
    parser.add_argument("--skip-images", action="store_true")
    parser.add_argument("--skip-captions", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.source_docx.exists():
        raise FileNotFoundError(args.source_docx)
    if args.source_docx.resolve() == args.target_docx.resolve():
        raise ValueError("target-docx must not be the same file as source-docx")

    source_hash_before = sha256(args.source_docx)
    args.target_docx.parent.mkdir(parents=True, exist_ok=True)
    if args.backup_docx:
        args.backup_docx.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(args.source_docx, args.backup_docx)
    shutil.copy2(args.source_docx, args.target_docx)

    footnotes = FootnoteRepairResult()
    if not args.skip_footnotes:
        footnotes = repair_fragmented_footnotes(args.target_docx)
    word = apply_word_repairs(args.target_docx, args)
    validation = validate(args.source_docx, args.target_docx, source_hash_before)

    lines = build_report(args.source_docx, args.target_docx, args.backup_docx, footnotes, word, validation)
    write_report(args.report_md, lines)
    if args.validation_report_md:
        write_report(args.validation_report_md, lines)
    elif not args.report_md:
        print("\n".join(lines))


if __name__ == "__main__":
    sys.exit(main())
