"""OneNote MCP Server — COM automation for OneNote desktop.

Provides 13 tools for navigating, reading, searching, and analyzing
OneNote content including embedded images and diagrams.
"""

import json
import re
from dataclasses import dataclass
from datetime import datetime, timezone

from mcp.server.fastmcp import FastMCP, Image

from onenote_lib import com_client
from onenote_lib.config import config
from onenote_lib.image_handler import get_all_images, get_image_base64
from onenote_lib.vision import describe_image, describe_images
from onenote_lib.xml_parser import (
    NotebookInfo,
    SectionGroupInfo,
    parse_notebooks,
    parse_page_to_markdown,
    parse_search_results,
)

mcp = FastMCP(
    "OneNote MCP",
    instructions="Access OneNote desktop notebooks via COM automation. "
    "Read, search, and analyze pages including embedded images.",
)


# ── Filter Infrastructure ────────────────────────────────────────────


@dataclass
class _HierarchyFilter:
    """Compiled filter parameters for hierarchy/list tools."""
    name_re: re.Pattern | None = None
    path_re: re.Pattern | None = None
    path_exclude_re: re.Pattern | None = None
    modified_since: datetime | None = None
    modified_before: datetime | None = None
    max_depth: int = -1
    include_pages: bool = True


def _build_filter(
    name_regex: str = "",
    path_regex: str = "",
    path_exclude_regex: str = "",
    modified_since: str = "",
    modified_before: str = "",
    max_depth: int = -1,
    include_pages: bool = True,
) -> _HierarchyFilter | None:
    """Create a compiled filter from raw string parameters.

    Returns None if no filters are active.  Raises ValueError for invalid inputs.
    """
    f = _HierarchyFilter()
    active = False

    try:
        if name_regex:
            f.name_re = re.compile(name_regex)
            active = True
        if path_regex:
            f.path_re = re.compile(path_regex)
            active = True
        if path_exclude_regex:
            f.path_exclude_re = re.compile(path_exclude_regex)
            active = True
    except re.error as e:
        raise ValueError(f"Invalid regex: {e}") from e

    if modified_since:
        f.modified_since = _parse_iso(modified_since)
        active = True
    if modified_before:
        f.modified_before = _parse_iso(modified_before)
        active = True
    if max_depth >= 0:
        f.max_depth = max_depth
        active = True
    if not include_pages:
        f.include_pages = False
        active = True

    return f if active else None


def _parse_iso(s: str) -> datetime:
    """Parse an ISO timestamp string to a timezone-aware datetime (defaults to UTC)."""
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        raise ValueError(f"Invalid ISO timestamp: {s!r}")
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def _make_path(parent: str, name: str) -> str:
    """Build a logical hierarchy path with ``/`` separator."""
    return f"{parent}/{name}" if parent else name


def _is_excluded(path: str, filt: _HierarchyFilter | None) -> bool:
    return bool(filt and filt.path_exclude_re and filt.path_exclude_re.search(path))


def _node_matches(path: str, name: str, filt: _HierarchyFilter | None) -> bool:
    """Check if a node directly matches *name_regex* and *path_regex*."""
    if not filt:
        return True
    if filt.name_re and not filt.name_re.search(name):
        return False
    if filt.path_re and not filt.path_re.search(path):
        return False
    return True


def _beyond_depth(depth: int, filt: _HierarchyFilter | None) -> bool:
    return bool(filt and filt.max_depth >= 0 and depth > filt.max_depth)


def _ts_in_range(ts: str | None, filt: _HierarchyFilter | None) -> bool:
    """Check if a timestamp string falls within the filter's date range."""
    if not filt or (not filt.modified_since and not filt.modified_before):
        return True
    if not ts:
        return False
    try:
        dt = datetime.fromisoformat(ts)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
    except (ValueError, TypeError):
        return False
    if filt.modified_since and dt < filt.modified_since:
        return False
    if filt.modified_before and dt > filt.modified_before:
        return False
    return True


# ── Navigation Tools ─────────────────────────────────────────────────


@mcp.tool()
def onenote_list_notebooks(name_regex: str = "") -> str:
    """List all open notebooks with their IDs, names, paths, and last modified times.

    Args:
        name_regex: Optional regex to filter notebooks by name.
    """
    xml = com_client.get_hierarchy("", com_client.NOTEBOOKS)
    notebooks = parse_notebooks(xml)

    try:
        name_re = re.compile(name_regex) if name_regex else None
    except re.error as e:
        return json.dumps({"error": f"Invalid regex: {e}"})

    result = []
    for nb in notebooks:
        if name_re and not name_re.search(nb.name):
            continue
        result.append({
            "id": nb.id,
            "name": nb.name,
            "path": nb.path,
            "last_modified": nb.last_modified,
        })
    return json.dumps(result, indent=2)


@mcp.tool()
def onenote_list_sections(
    notebook_id: str,
    name_regex: str = "",
    path_regex: str = "",
    path_exclude_regex: str = "",
) -> str:
    """List all sections in a notebook, including sections inside section groups.

    Args:
        notebook_id: The notebook's OneNote ID (from onenote_list_notebooks)
        name_regex: Regex matched against section names.
        path_regex: Regex matched against full logical path (Notebook/Group/.../Section).
        path_exclude_regex: Regex to exclude sections by path (e.g. "Archive").
    """
    xml = com_client.get_hierarchy(notebook_id, com_client.SECTIONS)
    notebooks = parse_notebooks(xml)
    if not notebooks:
        return json.dumps({"error": "Notebook not found"})

    nb = notebooks[0]
    result = _flatten_sections(nb.sections, nb.section_groups)

    try:
        name_re = re.compile(name_regex) if name_regex else None
        path_re = re.compile(path_regex) if path_regex else None
        exclude_re = re.compile(path_exclude_regex) if path_exclude_regex else None
    except re.error as e:
        return json.dumps({"error": f"Invalid regex: {e}"})

    if name_re or path_re or exclude_re:
        filtered = []
        for s in result:
            parts = [nb.name]
            if s.get("group"):
                parts.append(s["group"])
            parts.append(s["name"])
            full_path = "/".join(parts)

            if exclude_re and exclude_re.search(full_path):
                continue
            if name_re and not name_re.search(s["name"]):
                continue
            if path_re and not path_re.search(full_path):
                continue
            filtered.append(s)
        result = filtered

    return json.dumps(result, indent=2)


@mcp.tool()
def onenote_list_pages(
    section_id: str,
    name_regex: str = "",
    modified_since: str = "",
    modified_before: str = "",
) -> str:
    """List all pages in a section with titles and last modified times.

    Args:
        section_id: The section's OneNote ID (from onenote_list_sections)
        name_regex: Regex matched against page titles.
        modified_since: ISO timestamp — only pages modified on or after this date.
        modified_before: ISO timestamp — only pages modified on or before this date.
    """
    try:
        filt = _build_filter(
            name_regex=name_regex,
            modified_since=modified_since,
            modified_before=modified_before,
        )
    except ValueError as e:
        return json.dumps({"error": str(e)})

    xml = com_client.get_hierarchy(section_id, com_client.PAGES)
    notebooks = parse_notebooks(xml)

    pages = []
    for nb in notebooks:
        for sec in nb.sections:
            if sec.id == section_id:
                for p in sec.pages:
                    if filt:
                        if filt.name_re and not filt.name_re.search(p.name):
                            continue
                        if not _ts_in_range(p.last_modified, filt):
                            continue
                    pages.append({
                        "id": p.id,
                        "name": p.name,
                        "last_modified": p.last_modified,
                        "level": p.level,
                    })
                return json.dumps(pages, indent=2)
        found = _find_section_pages(nb.section_groups, section_id, filt)
        if found is not None:
            return json.dumps(found, indent=2)

    return json.dumps({"error": "Section not found"})


@mcp.tool()
def onenote_get_notebook_tree(
    notebook_id: str = "",
    name_regex: str = "",
    path_regex: str = "",
    path_exclude_regex: str = "",
    modified_since: str = "",
    modified_before: str = "",
    max_depth: int = -1,
    include_pages: bool = True,
) -> str:
    """Get the full hierarchy: notebooks -> section groups -> sections -> page titles.
    All filter parameters are optional and narrow the returned tree.

    Args:
        notebook_id: Optional notebook ID to scope the tree. Empty = all notebooks.
        name_regex: Regex matched (search) against entity names at every level.
        path_regex: Regex matched against full logical path (Notebook/Group/.../Section/Page).
        path_exclude_regex: Regex to exclude subtrees by path (e.g. "Archive").
        modified_since: ISO timestamp — only include pages modified on or after this date.
        modified_before: ISO timestamp — only include pages modified on or before this date.
        max_depth: Maximum hierarchy depth (0=notebooks, 1=+sections/groups, ...). -1=unlimited.
        include_pages: When False, return structure without page lists.
    """
    try:
        filt = _build_filter(
            name_regex=name_regex,
            path_regex=path_regex,
            path_exclude_regex=path_exclude_regex,
            modified_since=modified_since,
            modified_before=modified_before,
            max_depth=max_depth,
            include_pages=include_pages,
        )
    except ValueError as e:
        return json.dumps({"error": str(e)})

    xml = com_client.get_hierarchy(notebook_id, com_client.PAGES)
    notebooks = parse_notebooks(xml)
    result = []
    for nb in notebooks:
        tree = _notebook_to_tree(nb, filt=filt)
        if tree is not None:
            result.append(tree)
    return json.dumps(result, indent=2)


# ── Content Retrieval Tools ──────────────────────────────────────────


@mcp.tool()
def onenote_get_page(page_id: str) -> str:
    """Get a page's content as clean markdown. Images are listed as [Image N] references
    with callback IDs that can be retrieved with onenote_get_page_images or onenote_get_image.

    Args:
        page_id: The page's OneNote ID (from onenote_list_pages or search)
    """
    xml = com_client.get_page_content(page_id)
    markdown, images = parse_page_to_markdown(xml)

    if images:
        markdown += "\n\n---\n**Image References:**\n"
        for img in images:
            dims = ""
            if img.width and img.height:
                dims = f" ({img.width:.0f}x{img.height:.0f})"
            markdown += f"- [Image {img.index}]: callback_id=`{img.callback_id}`{dims}\n"

    return markdown


@mcp.tool()
def onenote_get_page_raw(page_id: str) -> str:
    """Get a page's raw OneNote XML content for debugging.

    Args:
        page_id: The page's OneNote ID
    """
    return com_client.get_page_content(page_id)


@mcp.tool()
def onenote_get_page_images(
    page_id: str,
    max_images: int = 10,
    max_size_kb: int = 512,
) -> list:
    """Extract all images from a page and return them as viewable images.
    Claude can see these images natively for analysis.

    Args:
        page_id: The page's OneNote ID
        max_images: Maximum number of images to extract (default 10)
        max_size_kb: Maximum size per image in KB (default 512, images are resized if larger)
    """
    xml = com_client.get_page_content(page_id)
    _, image_refs = parse_page_to_markdown(xml)

    if not image_refs:
        return ["No images found on this page."]

    images = get_all_images(page_id, image_refs, max_images, max_size_kb)

    result = []
    for img in images:
        if "error" in img:
            result.append(f"[Image {img['index']}] Error: {img['error']}")
        else:
            result.append(f"[Image {img['index']}] (callback_id: {img['callback_id']})")
            result.append(Image(data=img["base64"], media_type=img["media_type"]))

    return result


@mcp.tool()
def onenote_get_image(
    page_id: str,
    callback_id: str,
    max_size_kb: int = 512,
) -> list:
    """Get a single image by its callback ID. Returns the image for Claude to see natively.

    Args:
        page_id: The page's OneNote ID
        callback_id: The image's callback ID (from onenote_get_page output)
        max_size_kb: Maximum size in KB (default 512, image is resized if larger)
    """
    b64, media_type = get_image_base64(page_id, callback_id, max_size_kb)
    return [Image(data=b64, media_type=media_type)]


# ── Search Tools ─────────────────────────────────────────────────────


@mcp.tool()
def onenote_search(query: str) -> str:
    """Full-text search across all open notebooks. Uses Windows Search indexing.

    Args:
        query: Search query string
    """
    xml = com_client.find_pages(query)
    results = parse_search_results(xml)
    if not results:
        return json.dumps({"message": "No results found", "query": query})
    return json.dumps(results, indent=2)


@mcp.tool()
def onenote_search_in_notebook(notebook_id: str, query: str) -> str:
    """Search within a specific notebook.

    Args:
        notebook_id: The notebook's OneNote ID
        query: Search query string
    """
    xml = com_client.find_pages(query, notebook_id)
    results = parse_search_results(xml)
    if not results:
        return json.dumps({"message": "No results found", "query": query, "notebook_id": notebook_id})
    return json.dumps(results, indent=2)


# ── Vision Analysis Tools ────────────────────────────────────────────


@mcp.tool()
async def onenote_analyze_page_visuals(
    page_id: str,
    prompt: str = "",
    max_images: int = 5,
    max_size_kb: int = 512,
) -> list:
    """Fetch all images from a page, send each to a vision model for description,
    and return both the descriptions and the raw images for Claude to see.

    Requires a vision-capable LLM server (set ONENOTE_VISION_URL and ONENOTE_VISION_MODEL).

    Args:
        page_id: The page's OneNote ID
        prompt: Optional custom prompt for the vision model
        max_images: Maximum number of images to process (default 5)
        max_size_kb: Maximum size per image in KB (default 512)
    """
    xml = com_client.get_page_content(page_id)
    _, image_refs = parse_page_to_markdown(xml)

    if not image_refs:
        return ["No images found on this page."]

    images = get_all_images(page_id, image_refs, max_images, max_size_kb)
    analyzed = await describe_images(images, prompt or None)

    result = []
    for img in analyzed:
        if "error" in img:
            result.append(f"[Image {img['index']}] Error: {img['error']}")
            if "description" in img:
                result.append(f"Description: {img['description']}")
        else:
            result.append(f"[Image {img['index']}] Vision analysis: {img['description']}")
            result.append(Image(data=img["base64"], media_type=img["media_type"]))

    return result


@mcp.tool()
async def onenote_describe_image(
    page_id: str,
    callback_id: str,
    prompt: str = "Describe this image in detail. If it's a diagram, explain the structure and relationships shown.",
    max_size_kb: int = 512,
) -> list:
    """Send a single image to the vision model with a custom prompt.
    Returns both the description and the raw image.

    Args:
        page_id: The page's OneNote ID
        callback_id: The image's callback ID
        prompt: Custom prompt for the vision model
        max_size_kb: Maximum size in KB (default 512)
    """
    b64, media_type = get_image_base64(page_id, callback_id, max_size_kb)
    description = await describe_image(b64, media_type, prompt)

    return [
        f"Vision analysis: {description}",
        Image(data=b64, media_type=media_type),
    ]


# ── Write Tool ───────────────────────────────────────────────────────


@mcp.tool()
def onenote_create_page(section_id: str, title: str, content_html: str = "") -> str:
    """Create a new page in a section.

    Args:
        section_id: The section's OneNote ID where the page will be created
        title: Page title
        content_html: Optional HTML content for the page body.
            Use simple HTML: <p>, <b>, <i>, <ul>/<li>, <table>, <h1>-<h6>.
            Leave empty for a blank page.
    """
    ns = "http://schemas.microsoft.com/office/onenote/2013/onenote"
    try:
        new_page_id = com_client.create_new_page(section_id)

        if content_html or title:
            import xml.etree.ElementTree as ET
            page_xml = com_client.get_page_content(new_page_id)
            root = ET.fromstring(page_xml)
            ns_map = {"one": ns}

            title_elem = root.find(".//one:Title//one:T", ns_map)
            if title_elem is not None:
                title_elem.text = title

            if content_html:
                outline = ET.SubElement(root, f"{{{ns}}}Outline")
                oe_children = ET.SubElement(outline, f"{{{ns}}}OEChildren")
                oe = ET.SubElement(oe_children, f"{{{ns}}}OE")
                t = ET.SubElement(oe, f"{{{ns}}}T")
                t.text = content_html

            updated_xml = ET.tostring(root, encoding="unicode", xml_declaration=True)
            com_client.update_page_content(updated_xml)

        return json.dumps({
            "status": "created",
            "page_id": new_page_id,
            "title": title,
        })
    except Exception as e:
        return json.dumps({"error": str(e)})


# ── Helpers ──────────────────────────────────────────────────────────


def _flatten_sections(
    sections: list, section_groups: list, prefix: str = ""
) -> list[dict]:
    """Flatten sections and section groups into a flat list with group paths."""
    result = []
    for sec in sections:
        result.append({
            "id": sec.id,
            "name": sec.name,
            "group": prefix or None,
            "path": sec.path,
            "page_count": len(sec.pages),
        })
    for sg in section_groups:
        group_path = f"{prefix}/{sg.name}" if prefix else sg.name
        result.extend(_flatten_sections(sg.sections, sg.section_groups, group_path))
    return result


def _find_section_pages(
    section_groups: list[SectionGroupInfo],
    section_id: str,
    filt: _HierarchyFilter | None = None,
):
    """Recursively find pages in a section within section groups."""
    for sg in section_groups:
        for sec in sg.sections:
            if sec.id == section_id:
                pages = []
                for p in sec.pages:
                    if filt:
                        if filt.name_re and not filt.name_re.search(p.name):
                            continue
                        if not _ts_in_range(p.last_modified, filt):
                            continue
                    pages.append({
                        "id": p.id,
                        "name": p.name,
                        "last_modified": p.last_modified,
                        "level": p.level,
                    })
                return pages
        found = _find_section_pages(sg.section_groups, section_id, filt)
        if found is not None:
            return found
    return None


def _notebook_to_tree(
    nb: NotebookInfo,
    filt: _HierarchyFilter | None = None,
    depth: int = 0,
) -> dict | None:
    """Convert a NotebookInfo to a filtered tree dict."""
    nb_path = nb.name

    if _is_excluded(nb_path, filt):
        return None
    if _beyond_depth(depth, filt):
        return None

    direct_match = _node_matches(nb_path, nb.name, filt)

    tree = {"id": nb.id, "name": nb.name, "sections": [], "section_groups": []}

    for sec in nb.sections:
        item = _section_to_tree_item(sec, nb_path, filt, depth + 1, direct_match)
        if item is not None:
            tree["sections"].append(item)

    for sg in nb.section_groups:
        item = _section_group_to_tree(sg, nb_path, filt, depth + 1, direct_match)
        if item is not None:
            tree["section_groups"].append(item)

    if not filt or direct_match or tree["sections"] or tree["section_groups"]:
        return tree
    return None


def _section_to_tree_item(
    sec,
    parent_path: str,
    filt: _HierarchyFilter | None,
    depth: int,
    ancestor_matched: bool,
) -> dict | None:
    """Convert a SectionInfo to a filtered tree item."""
    sec_path = _make_path(parent_path, sec.name)

    if _is_excluded(sec_path, filt):
        return None
    if _beyond_depth(depth, filt):
        return None

    direct_match = ancestor_matched or _node_matches(sec_path, sec.name, filt)

    pages = []
    if filt and not filt.include_pages:
        pass  # structure only
    elif not _beyond_depth(depth + 1, filt):
        for p in sec.pages:
            p_path = _make_path(sec_path, p.name)
            if _is_excluded(p_path, filt):
                continue
            if not direct_match and not _node_matches(p_path, p.name, filt):
                continue
            if filt and (filt.modified_since or filt.modified_before):
                if not _ts_in_range(p.last_modified, filt):
                    continue
            pages.append({"id": p.id, "name": p.name, "level": p.level})

    if not filt or direct_match or pages:
        return {"id": sec.id, "name": sec.name, "pages": pages}
    return None


def _section_group_to_tree(
    sg: SectionGroupInfo,
    parent_path: str = "",
    filt: _HierarchyFilter | None = None,
    depth: int = 0,
    ancestor_matched: bool = False,
) -> dict | None:
    """Convert a SectionGroupInfo to a filtered tree dict."""
    sg_path = _make_path(parent_path, sg.name)

    if _is_excluded(sg_path, filt):
        return None
    if _beyond_depth(depth, filt):
        return None

    direct_match = ancestor_matched or _node_matches(sg_path, sg.name, filt)

    tree = {"id": sg.id, "name": sg.name, "sections": [], "section_groups": []}

    for sec in sg.sections:
        item = _section_to_tree_item(sec, sg_path, filt, depth + 1, direct_match)
        if item is not None:
            tree["sections"].append(item)

    for child in sg.section_groups:
        item = _section_group_to_tree(child, sg_path, filt, depth + 1, direct_match)
        if item is not None:
            tree["section_groups"].append(item)

    if not filt or direct_match or tree["sections"] or tree["section_groups"]:
        return tree
    return None


def main():
    mcp.run()


if __name__ == "__main__":
    main()
