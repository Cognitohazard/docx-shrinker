"""Tests for pure functions in docx_shrinker.core.

Each test targets a distinct code path, branch, or edge case.
No two tests exercise the same logic with interchangeable data.
"""

import zipfile

import pytest
from docx_shrinker.core import (
    extract_vml_dimensions,
    object_to_drawing,
    next_doc_pr_id,
    ensure_namespaces,
    strip_bookmarks,
    strip_comment_refs,
    strip_revisions,
    _find_page_border,
    _border_clip_rect,
    _vsdx_page_context,
    _tmp_vsdx_base,
    _render_pdf_to_image,
)


# ---------------------------------------------------------------------------
# extract_vml_dimensions
# ---------------------------------------------------------------------------

class TestExtractVmlDimensions:
    def test_empty_string_returns_fallback(self):
        assert extract_vml_dimensions("") == (3048000, 2286000)

    def test_pt_conversion(self):
        xml = '<v:shape style="width:100pt;height:50pt"/>'
        assert extract_vml_dimensions(xml) == (100 * 12700, 50 * 12700)

    def test_in_conversion(self):
        xml = '<v:shape style="width:2in;height:1in"/>'
        assert extract_vml_dimensions(xml) == (2 * 914400, 1 * 914400)

    def test_mixed_units_pt_and_in(self):
        xml = '<v:shape style="width:1in;height:72pt"/>'
        w, h = extract_vml_dimensions(xml)
        assert w == 914400
        assert h == 72 * 12700

    def test_width_only_needs_both_to_match(self):
        """Requires both width and height in the same style attr."""
        assert extract_vml_dimensions('<v:shape style="width:100pt"/>') == (3048000, 2286000)

    def test_skips_non_dimension_styles_picks_first_match(self):
        """The 'miter' style lacks width/height; the function should skip it
        and use the first style that has both."""
        xml = (
            '<v:shape style="miter:10"/>'
            '<v:shape style="width:200pt;height:100pt"/>'
            '<v:shape style="width:300pt;height:150pt"/>'
        )
        assert extract_vml_dimensions(xml) == (200 * 12700, 100 * 12700)

    def test_fractional_pt_truncates_to_int(self):
        xml = '<v:shape style="width:1.5pt;height:2.5pt"/>'
        assert extract_vml_dimensions(xml) == (int(1.5 * 12700), int(2.5 * 12700))

    @pytest.mark.parametrize("bad_val", ["1.2.3", "1..2", ".", "1.2."])
    def test_malformed_decimal_returns_fallback_not_crash(self, bad_val):
        """The old regex matched invalid float strings like '1.2.3',
        crashing with ValueError. Now the tighter numeric pattern simply
        doesn't match, so fallback dimensions are returned."""
        xml = f'<v:shape style="width:{bad_val}pt;height:{bad_val}pt"/>'
        assert extract_vml_dimensions(xml) == (3048000, 2286000)


# ---------------------------------------------------------------------------
# object_to_drawing
# ---------------------------------------------------------------------------

class TestObjectToDrawing:
    def test_no_imagedata_returns_input_unchanged(self):
        xml = "<w:object>no image here</w:object>"
        assert object_to_drawing(xml, 1) == xml

    def test_converts_vml_to_drawingml_with_correct_dimensions(self):
        xml = (
            '<w:object>'
            '<v:shape style="width:100pt;height:50pt">'
            '<v:imagedata r:id="rId5"/>'
            '</v:shape>'
            '</w:object>'
        )
        result = object_to_drawing(xml, 42)
        assert result.startswith("<w:drawing>")
        assert result.endswith("</w:drawing>")
        # rId wired through
        assert 'r:embed="rId5"' in result
        # doc_pr_id appears in docPr and cNvPr
        assert result.count('id="42"') == 2
        # Dimensions match extract_vml_dimensions and appear in both extent and a:ext
        assert result.count('cx="1270000"') == 2
        assert result.count('cy="635000"') == 2


# ---------------------------------------------------------------------------
# next_doc_pr_id
# ---------------------------------------------------------------------------

class TestNextDocPrId:
    def test_empty_returns_one(self):
        assert next_doc_pr_id("") == 1

    def test_single_docpr(self):
        assert next_doc_pr_id('<wp:docPr id="5" name="Pic"/>') == 6

    def test_matches_cnvpr_too(self):
        """The regex alternation (?:docPr|cNvPr) must match both tag names."""
        assert next_doc_pr_id('<pic:cNvPr id="7" name="Pic"/>') == 8

    def test_multiple_ids_returns_max_plus_one(self):
        xml = '<wp:docPr id="3"/><pic:cNvPr id="10"/><wp:docPr id="1"/>'
        assert next_doc_pr_id(xml) == 11

    def test_ignores_non_docpr_ids(self):
        """An id= on an unrelated tag should not affect the counter."""
        xml = '<w:other id="100"/><wp:docPr id="5"/>'
        assert next_doc_pr_id(xml) == 6


# ---------------------------------------------------------------------------
# ensure_namespaces
# ---------------------------------------------------------------------------

class TestEnsureNamespaces:
    def test_no_w_document_tag_returns_unchanged(self):
        xml = "<w:body>text</w:body>"
        assert ensure_namespaces(xml) == xml

    def test_adds_both_namespaces_when_missing(self):
        xml = '<w:document xmlns:w="http://example.com">body</w:document>'
        result = ensure_namespaces(xml)
        assert 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' in result
        assert 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' in result
        assert "body</w:document>" in result  # content preserved

    def test_does_not_duplicate_existing_namespace(self):
        xml = '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        result = ensure_namespaces(xml)
        assert result.count('xmlns:r=') == 1
        assert 'xmlns:wp=' in result

    def test_already_has_both_is_idempotent(self):
        xml = (
            '<w:document '
            'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        )
        assert ensure_namespaces(xml) == xml

    def test_namespace_in_attr_value_not_confused_with_declaration(self):
        """If 'xmlns:wp' appears inside another attribute's VALUE (not as an
        attribute name), the function must still add the real declaration.
        The old `if attr not in tag` check was fooled by this."""
        xml = '<w:document xmlns:w="http://xmlns:wp/test">'
        result = ensure_namespaces(xml)
        assert 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' in result


# ---------------------------------------------------------------------------
# strip_bookmarks
# ---------------------------------------------------------------------------

class TestStripBookmarks:
    def test_empty_string(self):
        assert strip_bookmarks("") == ("", 0)

    def test_removes_goback_and_its_end_tag(self):
        doc = (
            'before'
            '<w:bookmarkStart w:id="1" w:name="_GoBack"/>'
            'middle'
            '<w:bookmarkEnd w:id="1"/>'
            'after'
        )
        result, count = strip_bookmarks(doc)
        assert count == 1
        assert result == "beforemiddleafter"

    def test_removes_empty_name_bookmark(self):
        doc = (
            '<w:bookmarkStart w:id="2" w:name=""/>'
            'text'
            '<w:bookmarkEnd w:id="2"/>'
        )
        result, count = strip_bookmarks(doc)
        assert count == 1
        assert result == "text"

    def test_keeps_user_defined_bookmarks(self):
        doc = (
            '<w:bookmarkStart w:id="3" w:name="Chapter1"/>'
            'text'
            '<w:bookmarkEnd w:id="3"/>'
        )
        result, count = strip_bookmarks(doc)
        assert count == 0
        assert result == doc

    def test_reversed_attribute_order_still_matches(self):
        """Attributes may appear in any order in XML."""
        doc = (
            '<w:bookmarkStart w:name="_GoBack" w:id="1"/>'
            '<w:bookmarkEnd w:id="1"/>'
        )
        _, count = strip_bookmarks(doc)
        assert count == 1

    def test_two_goback_bookmarks_with_different_ids(self):
        """Multiple _GoBack bookmarks (e.g. from merging docs) should all go."""
        doc = (
            '<w:bookmarkStart w:id="1" w:name="_GoBack"/>'
            '<w:bookmarkEnd w:id="1"/>'
            'text'
            '<w:bookmarkStart w:id="5" w:name="_GoBack"/>'
            '<w:bookmarkEnd w:id="5"/>'
        )
        result, count = strip_bookmarks(doc)
        assert count == 2
        assert result == "text"


# ---------------------------------------------------------------------------
# strip_comment_refs
# ---------------------------------------------------------------------------

class TestStripCommentRefs:
    def test_strips_all_three_comment_tag_types(self):
        doc = (
            '<w:commentRangeStart w:id="1"/>'
            'text'
            '<w:commentRangeEnd w:id="1"/>'
            '<w:r><w:rPr><w:commentReference w:id="1"/></w:rPr></w:r>'
        )
        result = strip_comment_refs(doc)
        assert "commentRange" not in result
        assert "commentReference" not in result
        assert "<w:r><w:rPr></w:rPr></w:r>" in result
        assert "text" in result

    def test_leaves_non_comment_xml_untouched(self):
        doc = "<w:r><w:t>hello</w:t></w:r>"
        assert strip_comment_refs(doc) == doc


# ---------------------------------------------------------------------------
# strip_revisions
# ---------------------------------------------------------------------------

class TestStripRevisions:
    def test_simple_del_removed(self):
        doc = 'before<w:del w:id="1"><w:r><w:t>gone</w:t></w:r></w:del>after'
        result, counts = strip_revisions(doc, [])
        assert result == "beforeafter"
        assert counts['deletions'] == 1

    def test_simple_ins_unwrapped_content_kept(self):
        doc = 'before<w:ins w:id="1"><w:r><w:t>kept</w:t></w:r></w:ins>after'
        result, counts = strip_revisions(doc, [])
        assert result == "before<w:r><w:t>kept</w:t></w:r>after"
        assert counts['insertions'] == 1

    def test_unbalanced_del_warns_and_preserves(self):
        doc = '<w:del w:id="1"><w:r>text</w:r>'
        warnings = []
        result, counts = strip_revisions(doc, warnings)
        assert any("Mismatched w:del" in w for w in warnings)
        assert "<w:del" in result  # preserved to avoid content loss

    def test_rsid_attributes_stripped(self):
        doc = '<w:r w:rsidR="00A1234" w:rsidRPr="00B5678"><w:t>text</w:t></w:r>'
        result, _ = strip_revisions(doc, [])
        assert result == "<w:r><w:t>text</w:t></w:r>"

    def test_rpr_change_stripped(self):
        doc = '<w:rPr><w:rPrChange w:id="1"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr>'
        result, counts = strip_revisions(doc, [])
        assert result == "<w:rPr></w:rPr>"
        assert counts['property_changes'] == 1

    def test_del_wrapping_ins_removes_everything(self):
        """When a deletion wraps an insertion, all content should be removed."""
        doc = 'before<w:del w:id="1"><w:ins w:id="2"><w:r>text</w:r></w:ins></w:del>after'
        result, _ = strip_revisions(doc, [])
        assert result == "beforeafter"

    # --- Nesting bugs (the reason for the iterative-innermost-first fix) ---

    def test_nested_del_fully_removed(self):
        """Paragraph-level <w:del> containing run-level <w:del>.
        The old non-greedy .*? matched from outer open to inner close,
        leaking content and leaving an orphaned </w:del>."""
        doc = (
            'before'
            '<w:del w:id="1">'
            '<w:p>'
            '<w:del w:id="2"><w:r><w:t>inner</w:t></w:r></w:del>'
            '<w:r><w:t>outer</w:t></w:r>'
            '</w:p>'
            '</w:del>'
            'after'
        )
        result, _ = strip_revisions(doc, [])
        assert result == "beforeafter"

    def test_nested_ins_fully_unwrapped(self):
        """Paragraph-level <w:ins> containing run-level <w:ins>.
        All ins wrappers must be removed; all content must survive."""
        doc = (
            'before'
            '<w:ins w:id="1">'
            '<w:p>'
            '<w:ins w:id="2"><w:r><w:t>inner</w:t></w:r></w:ins>'
            '<w:r><w:t>outer</w:t></w:r>'
            '</w:p>'
            '</w:ins>'
            'after'
        )
        result, _ = strip_revisions(doc, [])
        assert "<w:ins" not in result
        assert "</w:ins>" not in result
        expected_content = (
            'before<w:p><w:r><w:t>inner</w:t></w:r>'
            '<w:r><w:t>outer</w:t></w:r></w:p>after'
        )
        assert result == expected_content

    def test_three_level_nested_del(self):
        """Verifies the while-loop terminates correctly for 3 nesting levels."""
        doc = (
            '<w:del w:id="1">'
            'L1'
            '<w:del w:id="2">'
            'L2'
            '<w:del w:id="3">L3</w:del>'
            '</w:del>'
            '</w:del>'
        )
        result, _ = strip_revisions(doc, [])
        assert result == ""


# ---------------------------------------------------------------------------
# _find_page_border / _border_clip_rect
# ---------------------------------------------------------------------------

def _make_pdf_page_with_border(width=100, height=100, stroke_width=0.5):
    """Helper: create a single-page PDF with a stroked rectangle at the edges."""
    import fitz
    doc = fitz.open()
    page = doc.new_page(width=width, height=height)
    shape = page.new_shape()
    shape.draw_rect(page.rect)
    shape.finish(color=(0, 0, 0), width=stroke_width)
    shape.commit()
    return page, doc


class TestBorderClipRect:
    def test_no_drawings_returns_full_rect(self):
        """A blank page should return the full page rect."""
        import fitz
        doc = fitz.open()
        page = doc.new_page(width=100, height=100)
        clip = _border_clip_rect(page)
        assert clip == page.rect
        doc.close()

    def test_detects_border_and_clips(self):
        """A stroked rectangle at page edges should produce a smaller clip."""
        page, doc = _make_pdf_page_with_border(stroke_width=0.75)
        clip = _border_clip_rect(page)
        assert clip.x0 > page.rect.x0
        assert clip.y0 > page.rect.y0
        assert clip.x1 < page.rect.x1
        assert clip.y1 < page.rect.y1
        doc.close()

    def test_thick_stroke_ignored(self):
        """Strokes wider than 2pt are not page frame borders."""
        page, doc = _make_pdf_page_with_border(stroke_width=3.0)
        clip = _border_clip_rect(page)
        assert clip == page.rect
        doc.close()

    def test_interior_rect_ignored(self):
        """A rectangle not at the page edges should not be detected."""
        import fitz
        doc = fitz.open()
        page = doc.new_page(width=100, height=100)
        shape = page.new_shape()
        shape.draw_rect(fitz.Rect(10, 10, 90, 90))
        shape.finish(color=(0, 0, 0), width=0.5)
        shape.commit()
        clip = _border_clip_rect(page)
        assert clip == page.rect
        doc.close()


# ---------------------------------------------------------------------------
# _vsdx_page_context
# ---------------------------------------------------------------------------

def _make_vsdx(tmp_path, pages_xml, windows_xml=None,
               pages_rels_xml=None, page_rels=None):
    """Build a minimal .vsdx ZIP for testing.

    page_rels: optional dict {page_filename: rels_xml} for per-page rels.
    """
    p = tmp_path / 't.vsdx'
    with zipfile.ZipFile(p, 'w') as zf:
        zf.writestr('visio/pages/pages.xml', pages_xml)
        if windows_xml is not None:
            zf.writestr('visio/windows.xml', windows_xml)
        if pages_rels_xml is not None:
            zf.writestr('visio/pages/_rels/pages.xml.rels', pages_rels_xml)
        if page_rels:
            for name, xml in page_rels.items():
                zf.writestr(f'visio/pages/_rels/{name}.rels', xml)
    return str(p)


class TestVsdxPageContext:
    def test_missing_pages_xml_returns_fallback(self, tmp_path):
        p = tmp_path / 'bad.vsdx'
        with zipfile.ZipFile(p, 'w') as zf:
            zf.writestr('other.xml', '<x/>')
        assert _vsdx_page_context(str(p)) == (0, None)

    def test_no_windows_xml_returns_fallback(self, tmp_path):
        p = _make_vsdx(tmp_path,
                       '<Pages><Page ID="0"/><Page ID="1"/></Pages>')
        assert _vsdx_page_context(p) == (0, None)

    def test_active_middle_page_without_rels(self, tmp_path):
        """Active page can be located, but no rels → index but no filter."""
        p = _make_vsdx(
            tmp_path,
            '<Pages><Page ID="0"/><Page ID="5"/><Page ID="9"/></Pages>',
            windows_xml='<Windows><Window WindowType="Drawing" Page="5"/></Windows>',
        )
        assert _vsdx_page_context(p) == (1, None)

    def test_background_pages_excluded_from_index(self, tmp_path):
        """Background pages are not exported by visPrintAll, so they must not
        count toward the page index."""
        p = _make_vsdx(
            tmp_path,
            '<Pages>'
            '<Page ID="0" Background="1"/>'
            '<Page ID="5"/>'
            '<Page ID="9"/>'
            '</Pages>',
            windows_xml='<Windows><Window WindowType="Drawing" Page="9"/></Windows>',
        )
        # ID=9 is the 2nd foreground page => index 1
        assert _vsdx_page_context(p) == (1, None)

    def test_active_id_not_in_foreground_falls_back(self, tmp_path):
        p = _make_vsdx(
            tmp_path,
            '<Pages><Page ID="0"/><Page ID="5"/></Pages>',
            windows_xml='<Windows><Window WindowType="Drawing" Page="99"/></Windows>',
        )
        assert _vsdx_page_context(p) == (0, None)

    def test_non_drawing_window_ignored(self, tmp_path):
        p = _make_vsdx(
            tmp_path,
            '<Pages><Page ID="0"/><Page ID="5"/></Pages>',
            windows_xml='<Windows>'
                        '<Window WindowType="ShapeSheet" Page="5"/>'
                        '</Windows>',
        )
        assert _vsdx_page_context(p) == (0, None)

    def test_corrupt_zip_returns_fallback(self, tmp_path):
        p = tmp_path / 'corrupt.vsdx'
        p.write_bytes(b'not a zip')
        assert _vsdx_page_context(str(p)) == (0, None)

    def test_returns_active_page_media_basenames(self, tmp_path):
        p = _make_vsdx(
            tmp_path,
            '<Pages>'
            '<Page ID="0"><Rel r:id="rId1"/></Page>'
            '<Page ID="5"><Rel r:id="rId2"/></Page>'
            '</Pages>',
            windows_xml='<Windows><Window WindowType="Drawing" Page="5"/></Windows>',
            pages_rels_xml='<Relationships>'
                           '<Relationship Id="rId1" Target="page1.xml"/>'
                           '<Relationship Id="rId2" Target="page2.xml"/>'
                           '</Relationships>',
            page_rels={'page2.xml':
                '<Relationships>'
                '<Relationship Id="rId1" Target="../media/image2.png"/>'
                '<Relationship Id="rId2" Target="../media/image3.jpg"/>'
                '</Relationships>'},
        )
        assert _vsdx_page_context(p) == (1, {'image2.png', 'image3.jpg'})

    def test_active_page_with_no_rels_returns_empty_set(self, tmp_path):
        """An active page whose rels file is missing has zero source images —
        callers can short-circuit."""
        p = _make_vsdx(
            tmp_path,
            '<Pages><Page ID="0"><Rel r:id="rId1"/></Page></Pages>',
            windows_xml='<Windows><Window WindowType="Drawing" Page="0"/></Windows>',
            pages_rels_xml='<Relationships>'
                           '<Relationship Id="rId1" Target="page1.xml"/>'
                           '</Relationships>',
        )
        assert _vsdx_page_context(p) == (0, set())


# ---------------------------------------------------------------------------
# _tmp_vsdx_base
# ---------------------------------------------------------------------------

class TestTmpVsdxBase:
    def test_includes_basename(self):
        stem = _tmp_vsdx_base('/tmp/foo/Microsoft_Visio_Drawing.vsdx')
        assert stem.startswith('Microsoft_Visio_Drawing_')

    def test_different_paths_same_basename_get_distinct_stems(self):
        """The same filename in different directories must not collide —
        otherwise their rendered PDFs/PNGs overwrite each other in tmp."""
        a = _tmp_vsdx_base('/tmp/dirA/drawing.vsdx')
        b = _tmp_vsdx_base('/tmp/dirB/drawing.vsdx')
        assert a != b

    def test_same_path_deterministic(self):
        """Same path must produce the same stem across calls so interactive
        re-render can locate the PDF the initial conversion wrote."""
        path = '/tmp/x/drawing.vsdx'
        assert _tmp_vsdx_base(path) == _tmp_vsdx_base(path)


# ---------------------------------------------------------------------------
# _render_pdf_to_image page_index parameter
# ---------------------------------------------------------------------------

class TestRenderPdfPageIndex:
    def _make_multipage_pdf(self, tmp_path):
        import fitz
        pdf_path = str(tmp_path / 't.pdf')
        doc = fitz.open()
        doc.new_page(width=100, height=100)
        doc.new_page(width=200, height=150)
        doc.new_page(width=400, height=300)
        doc.save(pdf_path)
        doc.close()
        return pdf_path

    def test_renders_requested_page(self, tmp_path):
        import fitz
        pdf_path = self._make_multipage_pdf(tmp_path)
        img_path = str(tmp_path / 'out.png')
        _render_pdf_to_image(pdf_path, img_path, fmt='png', dpi=72,
                             page_index=1)
        pix = fitz.Pixmap(img_path)
        # Page 1 is 200x150 at 72 dpi (scale=1) and no border to clip
        assert pix.width == 200
        assert pix.height == 150

    def test_default_page_index_renders_first_page(self, tmp_path):
        import fitz
        pdf_path = self._make_multipage_pdf(tmp_path)
        img_path = str(tmp_path / 'out.png')
        _render_pdf_to_image(pdf_path, img_path, fmt='png', dpi=72)
        pix = fitz.Pixmap(img_path)
        assert pix.width == 100
        assert pix.height == 100

    def test_out_of_range_page_index_falls_back_to_first(self, tmp_path):
        """A stale or mismatched index must not crash — fall back to page 0."""
        import fitz
        pdf_path = self._make_multipage_pdf(tmp_path)
        img_path = str(tmp_path / 'out.png')
        _render_pdf_to_image(pdf_path, img_path, fmt='png', dpi=72,
                             page_index=99)
        pix = fitz.Pixmap(img_path)
        assert pix.width == 100
        assert pix.height == 100
