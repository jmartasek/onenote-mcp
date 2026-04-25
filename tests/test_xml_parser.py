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


REAL_STYLE_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-real" name="Real Style Page">
  <one:QuickStyleDef index="0" name="PageTitle"/>
  <one:QuickStyleDef index="1" name="h1"/>
  <one:QuickStyleDef index="2" name="p"/>
  <one:Outline>
    <one:OEChildren>
      <one:OE quickStyleIndex="1">
        <one:T><![CDATA[Date Heading]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="2">
        <one:T><![CDATA[Normal topic text]]></one:T>
        <one:OEChildren>
          <one:OE quickStyleIndex="2">
            <one:List><one:Bullet/></one:List>
            <one:T><![CDATA[A bullet detail]]></one:T>
            <one:OEChildren>
              <one:OE quickStyleIndex="2">
                <one:List><one:Bullet/></one:List>
                <one:T><![CDATA[Nested bullet]]></one:T>
              </one:OE>
            </one:OEChildren>
          </one:OE>
        </one:OEChildren>
      </one:OE>
      <one:OE quickStyleIndex="2">
        <one:T><![CDATA[Another plain paragraph]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""


TAGS_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-tags" name="Tagged Page">
  <one:TagDef index="0" name="To Do" type="0" symbol="3"/>
  <one:TagDef index="1" name="Important" type="1" symbol="13"/>
  <one:TagDef index="2" name="Question" type="2" symbol="26"/>
  <one:QuickStyleDef index="0" name="PageTitle"/>
  <one:QuickStyleDef index="1" name="p"/>
  <one:Outline>
    <one:OEChildren>
      <one:OE quickStyleIndex="1">
        <one:Tag index="0" completed="false"/>
        <one:T><![CDATA[Unchecked task]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:Tag index="0" completed="true"/>
        <one:T><![CDATA[Completed task]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:Tag index="1" completed="true"/>
        <one:T><![CDATA[Important item]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:Tag index="2" completed="true"/>
        <one:T><![CDATA[A question]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:Tag index="0" completed="false"/>
        <one:Tag index="1" completed="true"/>
        <one:T><![CDATA[Important unchecked task]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:List><one:Bullet/></one:List>
        <one:Tag index="0" completed="true"/>
        <one:Tag index="1" completed="true"/>
        <one:T><![CDATA[Starred done bullet]]></one:T>
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

    def test_notebook_scoped_xml(self):
        """When root is a Notebook element (ID-scoped hierarchy call)."""
        xml = f"""<?xml version="1.0"?>
<one:Notebook xmlns:one="{NS}" name="Work Notes" ID="nb-001">
  <one:Section name="Meeting Notes" ID="sec-001">
    <one:Page ID="page-001" name="Standup" pageLevel="0"/>
  </one:Section>
</one:Notebook>"""
        notebooks = parse_notebooks(xml)
        assert len(notebooks) == 1
        assert notebooks[0].name == "Work Notes"
        assert notebooks[0].id == "nb-001"
        assert len(notebooks[0].sections) == 1
        assert notebooks[0].sections[0].name == "Meeting Notes"
        assert len(notebooks[0].sections[0].pages) == 1

    def test_section_scoped_xml(self):
        """When root is a Section element (section-ID-scoped hierarchy call)."""
        xml = f"""<?xml version="1.0"?>
<one:Section xmlns:one="{NS}" name="Meeting Notes" ID="sec-001">
  <one:Page ID="page-001" name="Standup" pageLevel="0"/>
  <one:Page ID="page-002" name="Sprint Review" pageLevel="0"/>
</one:Section>"""
        notebooks = parse_notebooks(xml)
        assert len(notebooks) == 1
        nb = notebooks[0]
        assert len(nb.sections) == 1
        sec = nb.sections[0]
        assert sec.id == "sec-001"
        assert sec.name == "Meeting Notes"
        assert len(sec.pages) == 2
        assert sec.pages[0].name == "Standup"


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

    def test_paragraph_style_not_rendered_as_heading(self):
        """When QuickStyleDef maps index 2 to 'p', it must NOT become a heading."""
        md, _ = parse_page_to_markdown(REAL_STYLE_XML)
        assert "## Date Heading" in md
        # These are qsi=2 which the page defines as "p" — must be plain text
        assert "Normal topic text" in md
        assert "### Normal topic text" not in md
        assert "Another plain paragraph" in md
        assert "### Another plain paragraph" not in md

    def test_bullets_under_paragraph(self):
        """Bullet items nested under a paragraph topic should render as list items."""
        md, _ = parse_page_to_markdown(REAL_STYLE_XML)
        assert "- A bullet detail" in md
        assert "  - Nested bullet" in md

    def test_list_prefix_wins_over_heading(self):
        """If an OE has both a heading style index AND a List child,
        the list prefix should take precedence."""
        ns = NS
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{ns}" ID="p" name="ListVsHeading">
  <one:QuickStyleDef index="1" name="h1"/>
  <one:Outline>
    <one:OEChildren>
      <one:OE quickStyleIndex="1">
        <one:List><one:Bullet/></one:List>
        <one:T><![CDATA[Should be bullet not heading]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        assert "- Should be bullet not heading" in md
        assert "## Should be bullet" not in md


class TestTags:
    def test_unchecked_todo(self):
        md, _ = parse_page_to_markdown(TAGS_XML)
        assert "[ ] Unchecked task" in md

    def test_checked_todo(self):
        md, _ = parse_page_to_markdown(TAGS_XML)
        assert "[x] Completed task" in md

    def test_important_star(self):
        md, _ = parse_page_to_markdown(TAGS_XML)
        assert "⭐ Important item" in md

    def test_question_mark(self):
        md, _ = parse_page_to_markdown(TAGS_XML)
        assert "❓ A question" in md

    def test_multiple_tags_stacked(self):
        md, _ = parse_page_to_markdown(TAGS_XML)
        assert "[ ] ⭐ Important unchecked task" in md

    def test_tags_with_bullet_list(self):
        md, _ = parse_page_to_markdown(TAGS_XML)
        assert "- [x] ⭐ Starred done bullet" in md

    def test_tags_inside_table_cells(self):
        ns = NS
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{ns}" ID="p" name="TableTags">
  <one:TagDef index="0" name="To Do" type="0" symbol="3"/>
  <one:TagDef index="1" name="Important" type="1" symbol="13"/>
  <one:QuickStyleDef index="0" name="p"/>
  <one:Outline>
    <one:OEChildren>
      <one:OE>
        <one:Table>
          <one:Row>
            <one:Cell><one:OEChildren>
              <one:OE><one:T><![CDATA[Category]]></one:T></one:OE>
            </one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren>
              <one:OE><one:T><![CDATA[Tasks]]></one:T></one:OE>
            </one:OEChildren></one:Cell>
          </one:Row>
          <one:Row>
            <one:Cell><one:OEChildren>
              <one:OE><one:T><![CDATA[Sprint 1]]></one:T></one:OE>
            </one:OEChildren></one:Cell>
            <one:Cell><one:OEChildren>
              <one:OE>
                <one:Tag index="0" completed="false"/>
                <one:T><![CDATA[Fix the bug]]></one:T>
              </one:OE>
              <one:OE>
                <one:Tag index="0" completed="true"/>
                <one:T><![CDATA[Write tests]]></one:T>
              </one:OE>
              <one:OE>
                <one:Tag index="1" completed="true"/>
                <one:T><![CDATA[Deploy to prod]]></one:T>
              </one:OE>
            </one:OEChildren></one:Cell>
          </one:Row>
        </one:Table>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""
        md, _ = parse_page_to_markdown(xml)
        # Tags should appear in table cells
        assert "[ ] Fix the bug" in md
        assert "[x] Write tests" in md
        assert "⭐ Deploy to prod" in md
        # Multiple OEs in one cell should be separated by <br>
        assert "<br>" in md


# ── Tests for extract_tagged_items ───────────────────────────────────


from onenote_lib.xml_parser import extract_tagged_items, TaggedItem


# XML with creationDate on tags and surrounding context lines
TAGGED_EXTRACT_XML = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="page-ext" name="Extract Page">
  <one:TagDef index="0" name="To Do" type="0" symbol="3"/>
  <one:TagDef index="1" name="Important" type="1" symbol="13"/>
  <one:TagDef index="2" name="Question" type="2" symbol="26"/>
  <one:QuickStyleDef index="0" name="PageTitle"/>
  <one:QuickStyleDef index="1" name="p"/>
  <one:Outline>
    <one:OEChildren>
      <one:OE quickStyleIndex="1">
        <one:T><![CDATA[Context line before]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:Tag index="0" completed="false" creationDate="2026-04-20T10:00:00Z"/>
        <one:T><![CDATA[Open task]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:T><![CDATA[Context line after]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:Tag index="0" completed="true" creationDate="2026-04-10T08:00:00Z"/>
        <one:T><![CDATA[Done task]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:Tag index="1" completed="true" creationDate="2026-04-22T12:00:00Z"/>
        <one:T><![CDATA[Star item]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:Tag index="2" completed="true" creationDate="2026-04-15T09:00:00Z"/>
        <one:T><![CDATA[Question item]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:Tag index="0" completed="false" creationDate="2026-04-18T07:00:00Z"/>
        <one:Tag index="1" completed="true" creationDate="2026-04-18T07:00:00Z"/>
        <one:T><![CDATA[Multi-tagged item]]></one:T>
      </one:OE>
      <one:OE quickStyleIndex="1">
        <one:T><![CDATA[Untagged line at end]]></one:T>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""


class TestExtractTaggedItems:
    def test_extracts_all_tagged(self):
        items, lines = extract_tagged_items(TAGGED_EXTRACT_XML)
        assert len(items) == 5
        texts = [i.text for i in items]
        assert "Open task" in texts
        assert "Done task" in texts
        assert "Star item" in texts
        assert "Question item" in texts
        assert "Multi-tagged item" in texts

    def test_untagged_lines_not_included(self):
        items, _ = extract_tagged_items(TAGGED_EXTRACT_XML)
        texts = [i.text for i in items]
        assert "Context line before" not in texts
        assert "Context line after" not in texts
        assert "Untagged line at end" not in texts

    def test_tag_types_classified(self):
        items, _ = extract_tagged_items(TAGGED_EXTRACT_XML)
        by_text = {i.text: i for i in items}
        assert by_text["Open task"].tags[0]["type"] == "todo"
        assert by_text["Star item"].tags[0]["type"] == "important"
        assert by_text["Question item"].tags[0]["type"] == "question"

    def test_completion_state(self):
        items, _ = extract_tagged_items(TAGGED_EXTRACT_XML)
        by_text = {i.text: i for i in items}
        assert by_text["Open task"].tags[0]["completed"] is False
        assert by_text["Done task"].tags[0]["completed"] is True

    def test_creation_date_preserved(self):
        items, _ = extract_tagged_items(TAGGED_EXTRACT_XML)
        by_text = {i.text: i for i in items}
        assert by_text["Open task"].tags[0]["created"] == "2026-04-20T10:00:00Z"

    def test_multi_tags(self):
        items, _ = extract_tagged_items(TAGGED_EXTRACT_XML)
        by_text = {i.text: i for i in items}
        multi = by_text["Multi-tagged item"]
        assert len(multi.tags) == 2
        types = {t["type"] for t in multi.tags}
        assert types == {"todo", "important"}

    def test_all_lines_include_everything(self):
        _, lines = extract_tagged_items(TAGGED_EXTRACT_XML)
        text = "\n".join(lines)
        assert "Context line before" in text
        assert "Context line after" in text
        assert "Open task" in text
        assert "Untagged line at end" in text

    def test_line_index_valid(self):
        items, lines = extract_tagged_items(TAGGED_EXTRACT_XML)
        for item in items:
            assert 0 <= item.line_index < len(lines)
            assert item.text in lines[item.line_index]

    def test_context_window(self):
        items, lines = extract_tagged_items(TAGGED_EXTRACT_XML)
        # Open task should be at index 1 (after "Context line before" at 0)
        open_task = [i for i in items if i.text == "Open task"][0]
        # Line before should be "Context line before"
        assert "Context line before" in lines[open_task.line_index - 1]
        # Line after should be "Context line after"
        assert "Context line after" in lines[open_task.line_index + 1]

    def test_tags_in_table_extracted(self):
        """Tags inside table cells are extracted."""
        ns = NS
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{ns}" ID="p" name="TableExtract">
  <one:TagDef index="0" name="To Do" type="0" symbol="3"/>
  <one:QuickStyleDef index="0" name="p"/>
  <one:Outline>
    <one:OEChildren>
      <one:OE>
        <one:Table>
          <one:Row>
            <one:Cell><one:OEChildren>
              <one:OE><one:T><![CDATA[Header]]></one:T></one:OE>
            </one:OEChildren></one:Cell>
          </one:Row>
          <one:Row>
            <one:Cell><one:OEChildren>
              <one:OE>
                <one:Tag index="0" completed="false" creationDate="2026-04-20T10:00:00Z"/>
                <one:T><![CDATA[Table task]]></one:T>
              </one:OE>
            </one:OEChildren></one:Cell>
          </one:Row>
        </one:Table>
      </one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""
        items, _ = extract_tagged_items(xml)
        texts = [i.text for i in items]
        assert "Table task" in texts

    def test_no_tags_returns_empty(self):
        xml = f"""<?xml version="1.0"?>
<one:Page xmlns:one="{NS}" ID="p" name="NoTags">
  <one:Outline>
    <one:OEChildren>
      <one:OE><one:T><![CDATA[Just text]]></one:T></one:OE>
    </one:OEChildren>
  </one:Outline>
</one:Page>"""
        items, lines = extract_tagged_items(xml)
        assert len(items) == 0
        assert len(lines) > 0
