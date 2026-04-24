"""Parse OneNote XML into markdown and structured data."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from html.parser import HTMLParser

# OneNote 2013 XML namespace
NS = {"one": "http://schemas.microsoft.com/office/onenote/2013/onenote"}

# Fallback quickStyleIndex → style name when the page has no QuickStyleDef elements.
# These match the default OneNote heading assignments (0 = normal paragraph).
_FALLBACK_STYLE_MAP: dict[int, str] = {
    1: "h1", 2: "h2", 3: "h3", 4: "h4", 5: "h5", 6: "h6",
}

# Style name → markdown heading prefix.
# The page title is already rendered as "#", so H1 starts at "##".
_HEADING_PREFIXES: dict[str, str] = {
    "h1": "##", "h2": "###", "h3": "####", "h4": "#####", "h5": "######", "h6": "######",
    "heading 1": "##", "heading 2": "###", "heading 3": "####",
    "heading 4": "#####", "heading 5": "######", "heading 6": "######",
}


class _MarkdownConverter(HTMLParser):
    """Convert HTML-like CDATA found in OneNote <T> elements to inline markdown.

    Handles: <b>/<strong> → **…**, <i>/<em> → *…*, monospace <span> → `…`.
    convert_charrefs=True lets the stdlib handle all HTML entity decoding.
    """

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self._parts: list[str] = []
        self._stack: list[str | None] = []

    def handle_starttag(self, tag: str, attrs: list) -> None:
        style = dict(attrs).get("style", "")
        kind = self._classify(tag, style)
        self._stack.append(kind)
        if kind == "bold":
            self._parts.append("**")
        elif kind == "italic":
            self._parts.append("*")
        elif kind == "code":
            self._parts.append("`")

    def handle_endtag(self, tag: str) -> None:
        if not self._stack:
            return
        kind = self._stack.pop()
        if kind == "bold":
            self._parts.append("**")
        elif kind == "italic":
            self._parts.append("*")
        elif kind == "code":
            self._parts.append("`")

    def handle_data(self, data: str) -> None:
        self._parts.append(data)

    def _classify(self, tag: str, style: str) -> str | None:
        if tag in ("b", "strong"):
            return "bold"
        if tag in ("i", "em"):
            return "italic"
        if tag == "span":
            s = style.lower()
            if "font-weight" in s and "bold" in s:
                return "bold"
            if "font-style" in s and "italic" in s:
                return "italic"
            if "font-family" in s and any(
                f in s for f in ("courier", "consolas", "monospace", "lucida console")
            ):
                return "code"
        return None

    def markdown(self) -> str:
        return "".join(self._parts)


def _html_to_markdown(text: str) -> str:
    """Convert HTML-like content from a OneNote <T> CDATA block to markdown."""
    if not text:
        return text
    conv = _MarkdownConverter()
    conv.feed(text)
    return conv.markdown()


@dataclass
class ImageRef:
    """Reference to an image in a OneNote page."""
    callback_id: str
    index: int
    width: float | None = None
    height: float | None = None
    alt_text: str | None = None


@dataclass
class PageInfo:
    """Parsed page metadata."""
    id: str
    name: str
    last_modified: str | None = None
    level: int = 0


@dataclass
class SectionInfo:
    """Parsed section metadata."""
    id: str
    name: str
    path: str | None = None
    pages: list[PageInfo] = field(default_factory=list)


@dataclass
class SectionGroupInfo:
    """Parsed section group metadata."""
    id: str
    name: str
    sections: list[SectionInfo] = field(default_factory=list)
    section_groups: list["SectionGroupInfo"] = field(default_factory=list)


@dataclass
class NotebookInfo:
    """Parsed notebook metadata."""
    id: str
    name: str
    path: str | None = None
    last_modified: str | None = None
    sections: list[SectionInfo] = field(default_factory=list)
    section_groups: list[SectionGroupInfo] = field(default_factory=list)


def parse_notebooks(xml_str: str) -> list[NotebookInfo]:
    """Parse hierarchy XML into notebook list."""
    root = ET.fromstring(xml_str)
    notebooks = []
    for nb in root.findall("one:Notebook", NS):
        notebook = NotebookInfo(
            id=nb.get("ID", ""),
            name=nb.get("name", ""),
            path=nb.get("path"),
            last_modified=nb.get("lastModifiedTime"),
        )
        notebook.sections = _parse_sections(nb)
        notebook.section_groups = _parse_section_groups(nb)
        notebooks.append(notebook)
    return notebooks


def _parse_sections(parent) -> list[SectionInfo]:
    """Parse Section elements under a parent node."""
    sections = []
    for sec in parent.findall("one:Section", NS):
        section = SectionInfo(
            id=sec.get("ID", ""),
            name=sec.get("name", ""),
            path=sec.get("path"),
        )
        for page in sec.findall("one:Page", NS):
            section.pages.append(PageInfo(
                id=page.get("ID", ""),
                name=page.get("name", ""),
                last_modified=page.get("lastModifiedTime"),
                level=int(page.get("pageLevel", "0")),
            ))
        sections.append(section)
    return sections


def _parse_section_groups(parent) -> list[SectionGroupInfo]:
    """Parse SectionGroup elements recursively."""
    groups = []
    for sg in parent.findall("one:SectionGroup", NS):
        # Skip recycle bin
        if sg.get("isRecycleBin") == "true":
            continue
        group = SectionGroupInfo(
            id=sg.get("ID", ""),
            name=sg.get("name", ""),
        )
        group.sections = _parse_sections(sg)
        group.section_groups = _parse_section_groups(sg)
        groups.append(group)
    return groups


def parse_page_to_markdown(xml_str: str) -> tuple[str, list[ImageRef]]:
    """Convert OneNote page XML to markdown text + image references.

    Returns:
        Tuple of (markdown_text, list_of_image_refs)
    """
    root = ET.fromstring(xml_str)
    title = root.get("name", root.get("ID", "Untitled"))
    lines = [f"# {title}", ""]

    images: list[ImageRef] = []
    img_counter = 0

    # Build quickStyleIndex → heading-prefix map from page-level QuickStyleDef elements.
    style_map = _build_style_map(root)

    # Process all Outline elements (main content containers)
    for outline in root.findall(".//one:Outline", NS):
        outline_lines, outline_images, img_counter = _process_outline(
            outline, img_counter, style_map
        )
        lines.extend(outline_lines)
        images.extend(outline_images)
        lines.append("")

    # Process top-level images (outside outlines)
    for img in root.findall(".//one:Image", NS):
        # Skip images already found inside outlines
        cb_id = _get_callback_id(img)
        if cb_id and not any(i.callback_id == cb_id for i in images):
            img_counter += 1
            ref = _make_image_ref(img, img_counter)
            if ref:
                images.append(ref)
                lines.append(f"[Image {ref.index}]")

    return "\n".join(lines).strip(), images


def _build_style_map(root) -> dict[int, str]:
    """Build quickStyleIndex → markdown heading prefix from page QuickStyleDef elements.

    Falls back to _FALLBACK_STYLE_MAP only for indices that have no QuickStyleDef
    definition at all.  If a page explicitly maps an index to a non-heading style
    (e.g. index 2 → "p"), the fallback will NOT overwrite it with "h2".
    """
    style_map: dict[int, str] = {}
    defined_indices: set[int] = set()
    for qsd in root.findall(".//one:QuickStyleDef", NS):
        try:
            idx = int(qsd.get("index", "-1"))
        except (ValueError, TypeError):
            continue
        if idx >= 0:
            defined_indices.add(idx)
            name = qsd.get("name", "").lower()
            prefix = _HEADING_PREFIXES.get(name)
            if prefix:
                style_map[idx] = prefix
    # Apply fallback only for indices that have NO definition on the page
    for idx, name in _FALLBACK_STYLE_MAP.items():
        if idx not in defined_indices:
            prefix = _HEADING_PREFIXES.get(name)
            if prefix:
                style_map[idx] = prefix
    return style_map


def _process_outline(
    outline, img_counter: int = 0, style_map: dict[int, str] | None = None
) -> tuple[list[str], list[ImageRef], int]:
    """Process an Outline element into markdown lines."""
    if style_map is None:
        style_map = {}
    lines: list[str] = []
    images: list[ImageRef] = []

    for oe_children in outline.findall("one:OEChildren", NS):
        child_lines, child_images, img_counter = _process_oechildren(
            oe_children, img_counter, style_map, list_depth=0
        )
        lines.extend(child_lines)
        images.extend(child_images)

    return lines, images, img_counter


def _process_oechildren(
    oe_children, img_counter: int, style_map: dict[int, str], list_depth: int
) -> tuple[list[str], list[ImageRef], int]:
    """Process an OEChildren element, iterating over its direct OE children."""
    lines: list[str] = []
    images: list[ImageRef] = []

    for oe in oe_children.findall("one:OE", NS):
        oe_lines, oe_images, img_counter = _process_oe(oe, img_counter, style_map, list_depth)
        lines.extend(oe_lines)
        images.extend(oe_images)

    return lines, images, img_counter


def _process_oe(
    oe, img_counter: int, style_map: dict[int, str], list_depth: int
) -> tuple[list[str], list[ImageRef], int]:
    """Process a single OE element, applying heading/list style and collecting content."""
    lines: list[str] = []
    images: list[ImageRef] = []

    # Heading prefix from quickStyleIndex
    heading_prefix = ""
    style_idx_str = oe.get("quickStyleIndex")
    if style_idx_str is not None:
        try:
            heading_prefix = style_map.get(int(style_idx_str), "")
        except ValueError:
            pass

    # List prefix from <one:List> child
    indent = "  " * list_depth
    list_prefix = ""
    is_list_item = False
    list_elem = oe.find("one:List", NS)
    if list_elem is not None:
        is_list_item = True
        if list_elem.find("one:Bullet", NS) is not None:
            list_prefix = indent + "- "
        elif list_elem.find("one:Number", NS) is not None:
            list_prefix = indent + "1. "
        else:
            ls = list_elem.get("listStyle", "").lower()
            if "bullet" in ls:
                list_prefix = indent + "- "
            elif ls:
                list_prefix = indent + "1. "

    # Collect text from direct <one:T> children (multiple T per OE is uncommon but valid)
    text_parts: list[str] = []
    for t in oe.findall("one:T", NS):
        converted = _html_to_markdown(t.text or "")
        if converted.strip():
            text_parts.append(converted)

    if text_parts:
        text = "".join(text_parts)
        if list_prefix:
            lines.append(f"{list_prefix}{text}")
        elif heading_prefix:
            lines.append(f"{heading_prefix} {text.strip()}")
        else:
            lines.append(text)

    # Images (direct OE children)
    for img_elem in oe.findall("one:Image", NS):
        cb_id = _get_callback_id(img_elem)
        if cb_id:
            img_counter += 1
            ref = _make_image_ref(img_elem, img_counter)
            if ref:
                images.append(ref)
                lines.append(f"[Image {ref.index}]")

    # Tables (direct OE children)
    for table in oe.findall("one:Table", NS):
        lines.extend(_process_table(table))

    # Attached files
    for attached in oe.findall("one:InsertedFile", NS):
        lines.append(f"[Attached: {attached.get('preferredName', 'file')}]")

    # Recurse into nested OEChildren; increase list_depth only when inside a list item
    nested_depth = list_depth + 1 if is_list_item else list_depth
    for child_oe_children in oe.findall("one:OEChildren", NS):
        child_lines, child_images, img_counter = _process_oechildren(
            child_oe_children, img_counter, style_map, nested_depth
        )
        lines.extend(child_lines)
        images.extend(child_images)

    return lines, images, img_counter


def _process_table(table_elem) -> list[str]:
    """Convert a OneNote table to markdown table."""
    rows = table_elem.findall("one:Row", NS)
    if not rows:
        return []

    md_rows = []
    for row in rows:
        cells = row.findall("one:Cell", NS)
        cell_texts = []
        for cell in cells:
            texts = []
            for t in cell.iter():
                if _local_tag(t.tag) == "T" and t.text:
                    texts.append(_html_to_markdown(t.text).strip())
            cell_text = " ".join(texts) if texts else ""
            cell_text = cell_text.replace("|", "\\|")
            cell_texts.append(cell_text)
        md_rows.append("| " + " | ".join(cell_texts) + " |")

    if len(md_rows) >= 1:
        # Insert header separator after first row
        col_count = md_rows[0].count("|") - 1
        separator = "| " + " | ".join(["---"] * col_count) + " |"
        md_rows.insert(1, separator)

    return md_rows


def _get_callback_id(img_elem) -> str | None:
    """Extract callbackID from an Image element.

    OneNote stores it as a child element: <one:CallbackID callbackID="..."/>
    not as an attribute on the Image tag itself.
    """
    # Check child element first (actual OneNote format)
    cb_elem = img_elem.find("one:CallbackID", NS)
    if cb_elem is not None:
        return cb_elem.get("callbackID")
    # Fallback: check as attribute (for compatibility)
    return img_elem.get("callbackID")


def _make_image_ref(img_elem, index: int) -> ImageRef | None:
    """Create an ImageRef from an Image element."""
    cb_id = _get_callback_id(img_elem)
    if not cb_id:
        return None

    # Try to get dimensions from Size child
    width = None
    height = None
    size = img_elem.find("one:Size", NS)
    if size is not None:
        w = size.get("width")
        h = size.get("height")
        if w:
            width = float(w)
        if h:
            height = float(h)

    return ImageRef(
        callback_id=cb_id,
        index=index,
        width=width,
        height=height,
    )


def _local_tag(tag: str) -> str:
    """Strip namespace from tag name."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def parse_search_results(xml_str: str) -> list[dict]:
    """Parse FindPages result XML into a list of matches."""
    root = ET.fromstring(xml_str)
    results = []

    for nb in root.findall("one:Notebook", NS):
        nb_name = nb.get("name", "")
        for sec in nb.findall(".//one:Section", NS):
            sec_name = sec.get("name", "")
            for page in sec.findall("one:Page", NS):
                results.append({
                    "page_id": page.get("ID", ""),
                    "page_name": page.get("name", ""),
                    "notebook": nb_name,
                    "section": sec_name,
                    "last_modified": page.get("lastModifiedTime", ""),
                })

    return results
