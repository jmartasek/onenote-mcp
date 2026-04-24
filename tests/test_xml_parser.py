"""Unit tests for OneNote XML parser."""

import pytest

from onenote_lib.xml_parser import (
    parse_notebooks,
    parse_page_to_markdown,
    parse_search_results,
)

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


HIERARCHY_XML = f"""<?xml version="1.0"?>
<one:Notebooks xmlns:one="{NS}">
  <one:Notebook name="Work Notes" ID="nb-001" path="C:\\Users\\test\\Work Notes"
                lastModifiedTime="2026-02-15T10:00:00Z">
    <one:Section name="Meeting Notes" ID="sec-001" path="C:\\Users\\test\\Work Notes\\Meeting Notes.one">
      <one:Page ID="page-001" name="Monday Standup" lastModifiedTime="2026-02-14T09:00:00Z" pageLevel="0"/>
      <one:Page ID="page-002" name="Sprint Review" lastModifiedTime="2026-02-13T14:00:00Z" pageLevel="0"/>
    </one:Section>
    <one:SectionGroup name="Archive" ID="sg-001">
      <one:Section name="Old Notes" ID="sec-002">
        <one:Page ID="page-003" name="Archived Page" pageLevel="0"/>
      </one:Section>
    </one:SectionGroup>
    <one:SectionGroup name="Recycle Bin" ID="sg-bin" isRecycleBin="true">
      <one:Section name="Deleted" ID="sec-deleted"/>
    </one:SectionGroup>
  </one:Notebook>
  <one:Notebook name="Personal" ID="nb-002" path="C:\\Users\\test\\Personal"
                lastModifiedTime="2026-02-10T08:00:00Z">
    <one:Section name="Journal" ID="sec-003">
      <one:Page ID="page-004" name="Feb 10" pageLevel="0"/>
    </one:Section>
  </one:Notebook>
</one:Notebooks>"""


PAGE_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-001" name="Test Page">
  <one:Title>
    <one:OE><one:T><![CDATA[Test Page]]></one:T></one:OE>
  </one:Title>
  <one:Outline>
    <one:OEChildren>
      <one:OE>
        <one:T><![CDATA[This is paragraph one.]]></one:T>
      </one:OE>
      <one:OE>
        <one:T><![CDATA[This is paragraph two with <b>bold</b> text.]]></one:T>
      </one:OE>
      <one:OE>
        <one:Image>
          <one:Size width="640" height="480" isSetByUser="true"/>
          <one:CallbackID callbackID="img-001"/>
        </one:Image>
      </one:OE>
      <one:OE>
        <one:T><![CDATA[Text after image.]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
  <one:Outline>
    <one:OEChildren>
      <one:OE>
        <one:Table>
          <one:Row>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Header 1]]></one:T></one:OE></one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Header 2]]></one:T></one:OE></one:OEChildren></one:Cell>
          </one:Row>
          <one:Row>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Data 1]]></one:T></one:OE></one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Data 2]]></one:T></one:OE></one:OEChildren></one:Cell>
          </one:Row>
        </one:Table>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""


PAGE_NO_IMAGES_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-005" name="Plain Page">
  <one:Title>
    <one:OE><one:T><![CDATA[Plain Page]]></one:T></one:OE>
  </one:Title>
  <one:Outline>
    <one:OEChildren>
      <one:OE>
        <one:T><![CDATA[Just text, no images.]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""


SEARCH_XML = f"""<?xml version="1.0"?>
<one:Notebooks xmlns:one="{NS}">
  <one:Notebook name="Work Notes" ID="nb-001">
    <one:Section name="Meeting Notes" ID="sec-001">
      <one:Page ID="page-001" name="Monday Standup" lastModifiedTime="2026-02-14T09:00:00Z"/>
    </one:Section>
  </one:Notebook>
</one:Notebooks>"""


FORMATTING_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-fmt" name="Formatting Page">
  <one:QuickStyleDef index="1" name="h1"/>
  <one:QuickStyleDef index="2" name="h2"/>
  <one:QuickStyleDef index="3" name="h3"/>
  <one:Title>
    <one:OE><one:T><![CDATA[Formatting Page]]></one:T></one:OE>
  </one:Title>
  <one:Outline>
    <one:OEChildren>
      <one:OE quickStyleIndex="1">
        <one:T><![CDATA[Heading One]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="2">
        <one:T><![CDATA[Heading Two]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="3">
        <one:T><![CDATA[Heading Three]]></one:T>
      </one:OE>
      <one:OE>
        <one:T><![CDATA[Normal paragraph with <b>bold</b> and <i>italic</i> text.]]></one:T>
      </one:OE>
      <one:OE>
        <one:T><![CDATA[Inline <span style="font-family:Courier New">code span</span> here.]]></one:T>
      </one:OE>
      <one:OE>
        <one:T><![CDATA[Bold via span: <span style="font-weight:bold;">strong</span>.]]></one:T>
      </one:OE>
      <one:OE>
        <one:List><one:Bullet/></one:List>
        <one:T><![CDATA[Bullet item one]]></one:T>
      </one:OE>
      <one:OE>
        <one:List><one:Bullet/></one:List>
        <one:T><![CDATA[Bullet item two]]></one:T>
        <one:OEChildren>
          <one:OE>
            <one:List><one:Bullet/></one:List>
            <one:T><![CDATA[Nested bullet]]></one:T>
          </one:OE>
        </one:OEChildren>
      </one:OE>
      <one:OE>
        <one:List><one:Number/></one:List>
        <one:T><![CDATA[Numbered item one]]></one:T>
      </one:OE>
      <one:OE>
        <one:List><one:Number/></one:List>
        <one:T><![CDATA[Numbered item two]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""


HEADING_FALLBACK_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-hf" name="Fallback Headings">
  <one:Outline>
    <one:OEChildren>
      <one:OE quickStyleIndex="1">
        <one:T><![CDATA[Fallback H1]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="2">
        <one:T><![CDATA[Fallback H2]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""


class TestParseNotebooks:
    def test_basic_parsing(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        assert len(notebooks) == 2
        assert notebooks[0].name == "Work Notes"
        assert notebooks[0].id == "nb-001"
        assert notebooks[1].name == "Personal"

    def test_sections(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        nb = notebooks[0]
        assert len(nb.sections) == 1
        assert nb.sections[0].name == "Meeting Notes"
        assert len(nb.sections[0].pages) == 2

    def test_section_groups(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        nb = notebooks[0]
        # Should have Archive but NOT Recycle Bin
        assert len(nb.section_groups) == 1
        assert nb.section_groups[0].name == "Archive"
        assert len(nb.section_groups[0].sections) == 1

    def test_pages(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        pages = notebooks[0].sections[0].pages
        assert len(pages) == 2
        assert pages[0].name == "Monday Standup"
        assert pages[0].id == "page-001"

    def test_nested_section_group_pages(self):
        notebooks = parse_notebooks(HIERARCHY_XML)
        sg = notebooks[0].section_groups[0]
        assert sg.sections[0].pages[0].name == "Archived Page"


class TestParsePageToMarkdown:
    def test_basic_content(self):
        md, images = parse_page_to_markdown(PAGE_XML)
        assert "# Test Page" in md
        assert "This is paragraph one." in md
        assert "This is paragraph two with **bold** text." in md
        assert "Text after image." in md

    def test_image_references(self):
        md, images = parse_page_to_markdown(PAGE_XML)
        assert len(images) == 1
        assert images[0].callback_id == "img-001"
        assert images[0].width == 640.0
        assert images[0].height == 480.0
        assert "[Image 1]" in md

    def test_table_parsing(self):
        md, _ = parse_page_to_markdown(PAGE_XML)
        assert "Header 1" in md
        assert "Header 2" in md
        assert "Data 1" in md
        assert "|" in md
        assert "---" in md

    def test_no_images(self):
        md, images = parse_page_to_markdown(PAGE_NO_IMAGES_XML)
        assert len(images) == 0
        assert "Just text, no images." in md

    def test_html_stripping(self):
        md, _ = parse_page_to_markdown(PAGE_XML)
        # HTML tags must not appear raw; bold content must be markdown-formatted
        assert "<b>" not in md
        assert "**bold**" in md


class TestParseSearchResults:
    def test_basic_search(self):
        results = parse_search_results(SEARCH_XML)
        assert len(results) == 1
        assert results[0]["page_name"] == "Monday Standup"
        assert results[0]["notebook"] == "Work Notes"
        assert results[0]["section"] == "Meeting Notes"

    def test_empty_search(self):
        empty_xml = f'<one:Notebooks xmlns:one="{NS}"/>'
        results = parse_search_results(empty_xml)
        assert len(results) == 0


class TestFormatting:
    def test_headings_from_quick_style_def(self):
        md, _ = parse_page_to_markdown(FORMATTING_XML)
        assert "## Heading One" in md
        assert "### Heading Two" in md
        assert "#### Heading Three" in md

    def test_headings_fallback_no_style_def(self):
        md, _ = parse_page_to_markdown(HEADING_FALLBACK_XML)
        assert "## Fallback H1" in md
        assert "### Fallback H2" in md

    def test_bold_tag(self):
        md, _ = parse_page_to_markdown(FORMATTING_XML)
        assert "**bold**" in md
        assert "<b>" not in md

    def test_italic_tag(self):
        md, _ = parse_page_to_markdown(FORMATTING_XML)
        assert "*italic*" in md
        assert "<i>" not in md

    def test_inline_code_span(self):
        md, _ = parse_page_to_markdown(FORMATTING_XML)
        assert "`code span`" in md

    def test_bold_via_span_style(self):
        md, _ = parse_page_to_markdown(FORMATTING_XML)
        assert "**strong**" in md

    def test_bullet_list(self):
        md, _ = parse_page_to_markdown(FORMATTING_XML)
        assert "- Bullet item one" in md
        assert "- Bullet item two" in md

    def test_nested_bullet_list(self):
        md, _ = parse_page_to_markdown(FORMATTING_XML)
        assert "  - Nested bullet" in md

    def test_numbered_list(self):
        md, _ = parse_page_to_markdown(FORMATTING_XML)
        assert "1. Numbered item one" in md
        assert "1. Numbered item two" in md

    def test_no_raw_html_in_output(self):
        md, _ = parse_page_to_markdown(FORMATTING_XML)
        assert "<span" not in md
        assert "<b>" not in md
        assert "<i>" not in md

    def test_html_entities_decoded(self):
        ns = NS
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{ns}" ID="p" name="Entities">
  <one:Outline>
    <one:OEChildren>
      <one:OE><one:T><![CDATA[A &amp; B &lt;tag&gt;]]></one:T></one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "A & B <tag>" in md
