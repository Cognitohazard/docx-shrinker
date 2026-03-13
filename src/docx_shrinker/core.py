"""Core functionality for shrinking and sanitizing Word (.docx) documents."""

import hashlib
import zipfile
import os
import re
import shutil
import tempfile

# Patterns for parts removed during cleanup (used in content types and rels)
_CLEANUP_PATTERNS = [
    r'vbaProject\.bin', r'comments\.xml', r'commentsExtended\.xml',
    r'commentsIds\.xml', r'thumbnail\.\w+', r'vbaData\.xml',
    r'printerSettings/', r'activeX/', r'customXml/', r'custom\.xml',
    r'embeddings/[^"]*\.vsdx',
]


def extract_vml_dimensions(obj_xml):
    """Extract width/height in EMU from a VML <w:object> block.
    Searches all style attributes for one containing both 'width:' and 'height:'
    (skipping unrelated styles like 'miter' on stroke elements)."""
    styles = re.findall(r'style="([^"]*)"', obj_xml)
    width_emu = 3048000   # fallback 3.2 inches
    height_emu = 2286000  # fallback 2.4 inches

    for style in styles:
        w_m = re.search(r'width:([\d.]+)(pt|in)', style)
        h_m = re.search(r'height:([\d.]+)(pt|in)', style)
        if w_m and h_m:
            w_val = float(w_m.group(1))
            h_val = float(h_m.group(1))
            width_emu = int(w_val * 12700) if w_m.group(2) == 'pt' else int(w_val * 914400)
            height_emu = int(h_val * 12700) if h_m.group(2) == 'pt' else int(h_val * 914400)
            break

    return width_emu, height_emu


def object_to_drawing(obj_xml, doc_pr_id):
    """Convert a VML <w:object> block to a DrawingML <w:drawing> block.
    doc_pr_id must be unique across the document."""
    img_match = re.search(r'<v:imagedata\s[^>]*r:id="(rId\d+)"', obj_xml)
    if not img_match:
        return obj_xml

    img_rid = img_match.group(1)
    cx, cy = extract_vml_dimensions(obj_xml)

    return (
        f'<w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:docPr id="{doc_pr_id}" name="Picture {doc_pr_id}"/>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{doc_pr_id}" name="Picture {doc_pr_id}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{img_rid}"/>'
        f'<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic>'
        f'</wp:inline></w:drawing>'
    )


def next_doc_pr_id(doc_xml):
    """Find the highest existing docPr/cNvPr id in the document and return max + 1."""
    ids = (int(m) for m in re.findall(r'(?:docPr|cNvPr)\b[^>]*\bid="(\d+)"', doc_xml))
    return max(ids, default=0) + 1


def ensure_namespaces(doc_xml):
    """Ensure the root <w:document> element declares wp: and r: namespaces
    needed by generated DrawingML blocks."""
    ns = {
        'xmlns:wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    m = re.search(r'(<w:document\b[^>]*)(>)', doc_xml)
    if not m:
        return doc_xml
    tag = m.group(1)
    for attr, uri in ns.items():
        if attr not in tag:
            tag += f' {attr}="{uri}"'
    return doc_xml[:m.start()] + tag + m.group(2) + doc_xml[m.end():]


def compress_media_images(media_dir, max_width=2000, quality=85, skip_files=None):
    """Re-encode oversized raster images in word/media/ to reduce file size.
    Resizes images wider than max_width and re-compresses JPGs.
    skip_files: set of filenames to skip (e.g. freshly converted Visio images).
    Returns list of (filename, old_kb, new_kb) for images that were shrunk."""
    import fitz

    results = []
    if not os.path.isdir(media_dir):
        return results
    skip = skip_files or set()

    for fname in os.listdir(media_dir):
        if fname in skip:
            continue
        ext = os.path.splitext(fname)[1].lower()
        if ext not in ('.png', '.jpg', '.jpeg'):
            continue
        path = os.path.join(media_dir, fname)
        old_size = os.path.getsize(path)

        try:
            pix = fitz.Pixmap(path)
        except Exception:
            continue

        # Drop alpha for JPG output
        if pix.alpha:
            pix = fitz.Pixmap(fitz.csRGB, pix)

        resized = False
        if pix.width > max_width:
            scale = max_width / pix.width
            new_w = max_width
            new_h = int(pix.height * scale)
            # Resize via PDF re-render
            tmp_pdf = fitz.open()
            page = tmp_pdf.new_page(width=pix.width, height=pix.height)
            page.insert_image(page.rect, pixmap=pix)
            mat = fitz.Matrix(new_w / pix.width, new_h / pix.height)
            pix = tmp_pdf[0].get_pixmap(matrix=mat, alpha=False)
            tmp_pdf.close()
            resized = True

        # Re-compress: save to temp first, only overwrite if actually smaller
        if ext in ('.jpg', '.jpeg'):
            tmp_path = path + '.recompress' + ext
            pix.save(tmp_path, jpg_quality=quality)
            new_size = os.path.getsize(tmp_path)
            saved_pct = (old_size - new_size) / old_size * 100 if old_size else 0
            if new_size < old_size and saved_pct > 5:
                os.replace(tmp_path, path)
                results.append((fname, old_size // 1024, new_size // 1024))
            else:
                os.remove(tmp_path)
        else:
            # For PNG, only save if we resized (PNG is lossless, rewriting won't help)
            if resized:
                pix.save(path)
                new_size = os.path.getsize(path)
                if new_size < old_size:
                    results.append((fname, old_size // 1024, new_size // 1024))
            else:
                continue

    return results


def dedup_media(media_dir, unpack_dir):
    """Deduplicate identical files in word/media/ by hash.
    Rewrites all .rels files to point duplicates to a single canonical file.
    Returns number of duplicates removed."""
    if not os.path.isdir(media_dir):
        return 0

    # Hash all media files
    hash_to_files = {}
    for fname in os.listdir(media_dir):
        path = os.path.join(media_dir, fname)
        if not os.path.isfile(path):
            continue
        h = hashlib.md5()
        with open(path, 'rb') as fh:
            for chunk in iter(lambda: fh.read(65536), b''):
                h.update(chunk)
        h = h.hexdigest()
        hash_to_files.setdefault(h, []).append(fname)

    # Find duplicates and build rename map
    renames = {}  # {duplicate_fname: canonical_fname}
    removed = 0
    for fnames in hash_to_files.values():
        if len(fnames) < 2:
            continue
        canonical = fnames[0]
        for dup in fnames[1:]:
            renames[dup] = canonical
            os.remove(os.path.join(media_dir, dup))
            removed += 1

    if not renames:
        return 0

    # Update all .rels files to point to canonical
    for rels_dir, _, files in os.walk(unpack_dir):
        for fname in files:
            if not fname.endswith('.rels'):
                continue
            rels_path = os.path.join(rels_dir, fname)
            with open(rels_path, 'r', encoding='utf-8') as f:
                content = f.read()
            changed = False
            for dup, canon in renames.items():
                if dup in content:
                    content = content.replace(dup, canon)
                    changed = True
            if changed:
                with open(rels_path, 'w', encoding='utf-8') as f:
                    f.write(content)

    return removed


def strip_bookmarks(doc):
    """Strip auto-generated bookmarks (_GoBack, empty name) from document XML string.
    Returns (modified_doc, count)."""
    to_remove = []
    for pattern in [
        r'<w:bookmarkStart\b[^>]*w:name="_GoBack"[^>]*/>',
        r'<w:bookmarkStart\b[^>]*w:name=""[^>]*/>',
    ]:
        for m in re.finditer(pattern, doc):
            id_m = re.search(r'w:id="(\d+)"', m.group(0))
            if id_m:
                to_remove.append((m.group(0), id_m.group(1)))

    for start_tag, bid in to_remove:
        doc = doc.replace(start_tag, '', 1)
        doc = re.sub(rf'<w:bookmarkEnd\b[^>]*w:id="{bid}"[^>]*/>', '', doc)
    return doc, len(to_remove)


def _render_pdf_to_image(pdf_path, img_path, fmt='jpg', dpi=300, quality=95,
                         max_width=0):
    """Render the first page of a PDF to an image file via PyMuPDF.
    Insets by 1pt to crop the Visio page frame border.
    Returns True on success."""
    import fitz

    pdf_doc = fitz.open(pdf_path)
    page = pdf_doc[0]
    scale = dpi / 72

    # Crop border: inset by ~1pt (the Visio page frame)
    inset = 1.0  # points
    rect = page.rect
    clip = fitz.Rect(rect.x0 + inset, rect.y0 + inset,
                     rect.x1 - inset, rect.y1 - inset)

    # Cap scale so output width doesn't exceed max_width
    if max_width > 0:
        content_width_pt = clip.x1 - clip.x0
        full_width_px = content_width_pt * scale
        if full_width_px > max_width:
            scale = max_width / content_width_pt

    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, alpha=False, clip=clip)
    pdf_doc.close()

    if fmt == 'jpg':
        pix.save(img_path, jpg_quality=quality)
    else:
        pix.save(img_path)
    return True


_VISIO_OPEN_FLAGS = 0x8 | 0x2  # visOpenDontList | visOpenRO


def _export_vsdx_to_pdf(visio, vsdx_path, pdf_path):
    """Export a single .vsdx to PDF using an already-running Visio instance."""
    doc = visio.Documents.OpenEx(os.path.abspath(vsdx_path), _VISIO_OPEN_FLAGS)
    doc.ExportAsFixedFormat(1, os.path.abspath(pdf_path), 1, 0)  # PDF, Print, All
    doc.Close()


def _get_visio(warnings):
    """Launch Visio COM and return the application object, or None."""
    try:
        import win32com.client
    except ImportError:
        warnings.append('pywin32 not installed — skipping Visio conversion')
        return None
    try:
        visio = win32com.client.Dispatch('Visio.Application')
    except Exception:
        warnings.append('Visio not available — keeping EMF previews')
        return None
    visio.Visible = False
    visio.AlertResponse = 7  # suppress dialogs (answer "No")
    return visio


def convert_vsdx_via_visio(vsdx_paths, out_dir, warnings, fmt='jpg', dpi=300,
                           quality=95, max_width=2000, keep_pdfs=False):
    """Convert .vsdx files via Visio COM (vsdx->PDF) then PyMuPDF (PDF->image).

    Returns dict: {vsdx_path: image_path} for successful conversions."""
    results = {}
    visio = _get_visio(warnings)
    if visio is None:
        return results

    try:
        for vsdx_path in vsdx_paths:
            base = os.path.splitext(os.path.basename(vsdx_path))[0]
            pdf_path = os.path.join(out_dir, f'{base}.pdf')
            img_path = os.path.join(out_dir, f'{base}.{fmt}')
            try:
                _export_vsdx_to_pdf(visio, vsdx_path, pdf_path)

                if os.path.exists(pdf_path):
                    _render_pdf_to_image(pdf_path, img_path, fmt=fmt,
                                         dpi=dpi, quality=quality,
                                         max_width=max_width)

                    if not keep_pdfs:
                        try:
                            os.remove(pdf_path)
                        except OSError:
                            pass

                if os.path.exists(img_path):
                    results[vsdx_path] = img_path
            except Exception as e:
                warnings.append(f'Visio failed on {base}.vsdx: {e}')
                for p in [pdf_path, img_path]:
                    if os.path.exists(p):
                        try:
                            os.remove(p)
                        except OSError:
                            pass
    finally:
        try:
            visio.Quit()
        except Exception:
            pass

    return results


def _strip_xml_tags(path, tags):
    """Remove specified XML tags from a file."""
    if not os.path.exists(path):
        return
    with open(path, 'r', encoding='utf-8') as f:
        xml = f.read()
    for tag in tags:
        xml = re.sub(rf'<{re.escape(tag)}[^>]*>.*?</{re.escape(tag)}>', '', xml, flags=re.DOTALL)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(xml)


def sanitize_core_props(unpack_dir):
    """Strip personal info from docProps/core.xml."""
    _strip_xml_tags(os.path.join(unpack_dir, 'docProps', 'core.xml'),
                    ['dc:creator', 'cp:lastModifiedBy', 'cp:lastPrinted',
                     'cp:revision', 'dc:subject', 'cp:keywords',
                     'cp:category', 'cp:contentStatus'])


def sanitize_app_props(unpack_dir):
    """Strip sensitive fields from docProps/app.xml."""
    _strip_xml_tags(os.path.join(unpack_dir, 'docProps', 'app.xml'),
                    ['Company', 'Manager', 'HyperlinkBase'])


def remove_comment_files(unpack_dir):
    """Remove comments.xml, commentsExtended.xml, commentsIds.xml. Returns count removed."""
    count = 0
    for name in ['comments.xml', 'commentsExtended.xml', 'commentsIds.xml']:
        p = os.path.join(unpack_dir, 'word', name)
        if os.path.exists(p):
            os.remove(p)
            count += 1
    return count


def strip_comment_refs(doc):
    """Strip comment range/reference tags from document XML string."""
    doc = re.sub(r'<w:commentRangeStart\b[^>]*/>', '', doc)
    doc = re.sub(r'<w:commentRangeEnd\b[^>]*/>', '', doc)
    doc = re.sub(r'<w:commentReference\b[^>]*/>', '', doc)
    return doc


def strip_revisions(doc, warnings):
    """Accept all tracked changes by stripping revision markup from document XML string."""
    balanced = True
    for tag in ['w:del', 'w:ins']:
        opens = len(re.findall(rf'<{tag}\b', doc))
        closes = len(re.findall(rf'</{tag}>', doc))
        if opens != closes:
            warnings.append(f'Mismatched {tag} tags ({opens} open, {closes} close) — '
                            f'skipping {tag} removal to avoid content loss')
            balanced = False

    if balanced:
        doc = re.sub(r'<w:del\b[^>]*>.*?</w:del>', '', doc, flags=re.DOTALL)
        doc = re.sub(r'<w:ins\b[^>]*>(.*?)</w:ins>', r'\1', doc, flags=re.DOTALL)

    # Always safe to strip property-change blocks and rsid attributes
    for tag in ['rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange',
                'tblGridChange', 'tcPrChange', 'trPrChange']:
        doc = re.sub(rf'<w:{tag}\b[^>]*>.*?</w:{tag}>', '', doc, flags=re.DOTALL)
    doc = re.sub(r'\s+w:rsid\w*="[^"]*"', '', doc)
    return doc


def remove_garbage_parts(unpack_dir):
    """Remove thumbnail, VBA macros, printer settings, ActiveX, custom XML data."""
    removed = []

    for ext in ['jpeg', 'jpg', 'png', 'emf', 'wmf']:
        p = os.path.join(unpack_dir, 'docProps', f'thumbnail.{ext}')
        if os.path.exists(p):
            os.remove(p)
            removed.append(f'thumbnail.{ext}')

    vba_path = os.path.join(unpack_dir, 'word', 'vbaProject.bin')
    if os.path.exists(vba_path):
        os.remove(vba_path)
        removed.append('vbaProject.bin')
    vba_data = os.path.join(unpack_dir, 'word', 'vbaData.xml')
    if os.path.exists(vba_data):
        os.remove(vba_data)
        removed.append('vbaData.xml')

    printer_dir = os.path.join(unpack_dir, 'word', 'printerSettings')
    if os.path.exists(printer_dir):
        shutil.rmtree(printer_dir)
        removed.append('printerSettings/')

    activex_dir = os.path.join(unpack_dir, 'word', 'activeX')
    if os.path.exists(activex_dir):
        shutil.rmtree(activex_dir)
        removed.append('activeX/')

    custom_xml_dir = os.path.join(unpack_dir, 'customXml')
    if os.path.exists(custom_xml_dir):
        shutil.rmtree(custom_xml_dir)
        removed.append('customXml/')

    custom_props = os.path.join(unpack_dir, 'docProps', 'custom.xml')
    if os.path.exists(custom_props):
        os.remove(custom_props)
        removed.append('custom.xml')

    return removed


def clean_content_types(unpack_dir):
    """Remove Content_Types entries that reference deleted parts."""
    ct_path = os.path.join(unpack_dir, '[Content_Types].xml')
    if not os.path.exists(ct_path):
        return
    with open(ct_path, 'r', encoding='utf-8') as f:
        ct = f.read()
    for pattern in _CLEANUP_PATTERNS:
        ct = re.sub(rf'<Override[^>]*{pattern}[^>]*/>', '', ct)
    ct = re.sub(r'<Default[^>]*Extension="bin"[^>]*vbaProject[^>]*/>', '', ct)
    with open(ct_path, 'w', encoding='utf-8') as f:
        f.write(ct)


def clean_relationships(unpack_dir):
    """Remove .rels entries that reference deleted parts."""
    for rels_dir, _, files in os.walk(unpack_dir):
        for fname in files:
            if not fname.endswith('.rels'):
                continue
            rels_path = os.path.join(rels_dir, fname)
            with open(rels_path, 'r', encoding='utf-8') as f:
                rels = f.read()
            changed = False
            for pattern in _CLEANUP_PATTERNS:
                for tag_re in [rf'<Relationship\b[^>]*Target="[^"]*{pattern}[^"]*"[^>]*/>',
                               rf'<Relationship\b[^>]*Target="[^"]*{pattern}[^"]*"[^>]*>.*?</Relationship>']:
                    rels, n = re.subn(tag_re, '', rels, flags=re.DOTALL)
                    if n:
                        changed = True
            if changed:
                with open(rels_path, 'w', encoding='utf-8') as f:
                    f.write(rels)


def _interactive_reconvert(media_dir, emf_to_vsdx, conversions,
                           tmp_dir, fmt, dpi, warnings, max_width=0):
    """Show the top 5 largest converted images and let the user re-convert
    selected ones from the original .vsdx at a different quality."""

    # Build list of (filename, size, vsdx_path) for converted images
    items = []
    for emf_name, vsdx_path in emf_to_vsdx.items():
        if vsdx_path not in conversions:
            continue
        emf_base = os.path.splitext(emf_name)[0]
        img_file = f'{emf_base}.{fmt}'
        img_path = os.path.join(media_dir, img_file)
        if os.path.exists(img_path):
            items.append((img_file, os.path.getsize(img_path), vsdx_path))

    if not items:
        return

    items.sort(key=lambda x: x[1], reverse=True)
    top5 = items[:5]

    print(f'\n  Top {len(top5)} largest converted images:')
    for i, (fname, size, _) in enumerate(top5, 1):
        print(f'    {i}. {fname} ({size // 1024} KB)')

    print(f'\n  Enter numbers to re-convert (e.g. "1 3 5"), new quality (e.g. "1,3 q=80"),')
    print(f'  or press Enter to skip: ', end='', flush=True)

    try:
        user_input = input().strip()
    except (EOFError, KeyboardInterrupt):
        return
    if not user_input:
        return

    # Parse quality override
    new_quality = 85  # default re-conversion quality (lower than initial)
    if 'q=' in user_input:
        q_match = re.search(r'q=(\d+)', user_input)
        if q_match:
            new_quality = max(1, min(100, int(q_match.group(1))))
        user_input = re.sub(r'q=\d+', '', user_input).strip()

    # Parse selected indices
    selected = set()
    for part in re.split(r'[,\s]+', user_input):
        if part.isdigit():
            idx = int(part)
            if 1 <= idx <= len(top5):
                selected.add(idx - 1)

    if not selected:
        return

    # Check which items need PDF re-export from Visio
    needs_reexport = []
    for idx in sorted(selected):
        fname, _, vsdx_path = top5[idx]
        base = os.path.splitext(os.path.basename(vsdx_path))[0]
        pdf_path = os.path.join(tmp_dir, f'{base}.pdf')
        if not os.path.exists(pdf_path) and os.path.exists(vsdx_path):
            needs_reexport.append((idx, vsdx_path, pdf_path))

    # Re-export missing PDFs via Visio (one session for all)
    if needs_reexport:
        visio = _get_visio(warnings)
        if visio is not None:
            try:
                for _, vsdx_path, pdf_path in needs_reexport:
                    try:
                        _export_vsdx_to_pdf(visio, vsdx_path, pdf_path)
                    except Exception as e:
                        base = os.path.splitext(os.path.basename(vsdx_path))[0]
                        warnings.append(f'Could not re-export {base}: {e}')
            finally:
                try:
                    visio.Quit()
                except Exception:
                    pass

    print(f'  Re-converting {len(selected)} image(s) at quality={new_quality}...')
    for idx in sorted(selected):
        fname, _, vsdx_path = top5[idx]
        img_path = os.path.join(media_dir, fname)
        base = os.path.splitext(os.path.basename(vsdx_path))[0]
        pdf_path = os.path.join(tmp_dir, f'{base}.pdf')

        if not os.path.exists(pdf_path):
            warnings.append(f'PDF not available for {base}, skipping')
            continue

        try:
            old_size = os.path.getsize(img_path)
            _render_pdf_to_image(pdf_path, img_path, fmt=fmt, dpi=dpi,
                                 quality=new_quality, max_width=max_width)
            new_size = os.path.getsize(img_path)
            print(f'    {fname}: {old_size // 1024} KB -> {new_size // 1024} KB')
        except Exception as e:
            warnings.append(f'Re-conversion failed for {fname}: {e}')


def shrink_docx(src_path, dst_path, fmt='jpg', dpi=300, quality=95, max_width=2000,
                interactive=False):
    """Shrink and sanitize a Word document.

    Returns a dict with:
        original_size_mb, new_size_mb, reduction_mb, reduction_percent,
        output_path, visio_converted (list), visio_removed (int),
        images_compressed (list of (name, old_kb, new_kb)),
        duplicates_removed (int), comments_removed (int),
        bookmarks_removed (int), garbage_removed (list),
        warnings (list of str)
    """
    result = {
        'original_size_mb': 0,
        'new_size_mb': 0,
        'reduction_mb': 0,
        'reduction_percent': 0,
        'output_path': dst_path,
        'visio_converted': [],
        'visio_removed': 0,
        'images_compressed': [],
        'duplicates_removed': 0,
        'comments_removed': 0,
        'bookmarks_removed': 0,
        'garbage_removed': [],
        'warnings': [],
    }
    warnings = result['warnings']

    with tempfile.TemporaryDirectory() as work:
        unpack_dir = os.path.join(work, 'unpacked')
        tmp_dir = os.path.join(work, 'tmp')
        os.makedirs(tmp_dir)

        # Unpack
        with zipfile.ZipFile(src_path) as zf:
            zf.extractall(unpack_dir)

        media_dir = os.path.join(unpack_dir, 'word', 'media')
        embed_dir = os.path.join(unpack_dir, 'word', 'embeddings')

        # --- 1. Convert .vsdx images via Visio, then remove .vsdx embeddings ---
        rels_path = os.path.join(unpack_dir, 'word', '_rels', 'document.xml.rels')
        if os.path.exists(rels_path):
            with open(rels_path, 'r', encoding='utf-8') as f:
                rels_xml = f.read()
        else:
            rels_xml = ''

        doc_path = os.path.join(unpack_dir, 'word', 'document.xml')
        with open(doc_path, 'r', encoding='utf-8') as f:
            doc = f.read()

        # Build rId -> Target mapping from rels
        rid_to_target = {}
        for m in re.finditer(r'<Relationship\b[^>]*\bId="(rId\d+)"[^>]*\bTarget="([^"]*)"', rels_xml):
            rid_to_target[m.group(1)] = m.group(2)

        # Find each <w:object> and extract OLE rId (-> .vsdx) and image rId (-> .emf)
        emf_to_vsdx = {}  # {emf_filename: vsdx_full_path}
        for obj_m in re.finditer(r'<w:object\b[^>]*>.*?</w:object>', doc, flags=re.DOTALL):
            obj_xml = obj_m.group(0)
            ole_match = re.search(r'<o:OLEObject\b[^>]*r:id="(rId\d+)"', obj_xml)
            img_match = re.search(r'<v:imagedata\s[^>]*r:id="(rId\d+)"', obj_xml)
            if not ole_match or not img_match:
                continue
            ole_target = rid_to_target.get(ole_match.group(1), '')
            img_target = rid_to_target.get(img_match.group(1), '')
            if ole_target.endswith('.vsdx') and img_target:
                emf_name = os.path.basename(img_target)
                vsdx_path = os.path.join(unpack_dir, 'word', ole_target.replace('/', os.sep))
                emf_to_vsdx[emf_name] = vsdx_path

        # Convert via Visio COM (batch — opens Visio once for all files)
        vsdx_paths = [p for p in emf_to_vsdx.values() if os.path.exists(p)]
        conversions = convert_vsdx_via_visio(vsdx_paths, tmp_dir, warnings,
                                              fmt=fmt, dpi=dpi, quality=quality,
                                              max_width=max_width,
                                              keep_pdfs=interactive)

        # Place converted images and update refs
        converted = []  # list of emf_basename_no_ext
        for emf_name, vsdx_path in emf_to_vsdx.items():
            img_path = conversions.get(vsdx_path)
            if img_path is None:
                continue

            emf_base = os.path.splitext(emf_name)[0]
            dest = os.path.join(media_dir, f'{emf_base}.{fmt}')
            shutil.copy2(img_path, dest)

            # Remove the old EMF
            emf_path = os.path.join(media_dir, emf_name)
            if os.path.exists(emf_path):
                os.remove(emf_path)

            converted.append(emf_base)
            size_kb = os.path.getsize(dest) // 1024
            result['visio_converted'].append((emf_base, size_kb))

        if emf_to_vsdx and not converted:
            warnings.append(f'Kept {len(emf_to_vsdx)} EMF preview(s) (Visio unavailable)')

        # Delete all .vsdx files (whether conversion succeeded or not)
        vsdx_removed = 0
        if os.path.isdir(embed_dir):
            remaining = []
            for f in os.listdir(embed_dir):
                if f.endswith('.vsdx'):
                    os.remove(os.path.join(embed_dir, f))
                    vsdx_removed += 1
                else:
                    remaining.append(f)
            if not remaining:
                os.rmdir(embed_dir)
        result['visio_removed'] = vsdx_removed

        # --- 2. Convert OLE objects to DrawingML ---
        _id_counter = [next_doc_pr_id(doc)]

        def _replace_object(m):
            r = object_to_drawing(m.group(0), _id_counter[0])
            _id_counter[0] += 1
            return r

        doc = re.sub(
            r'<w:object\b[^>]*>.*?</w:object>',
            _replace_object,
            doc, flags=re.DOTALL
        )
        doc = ensure_namespaces(doc)

        # --- 2b. Strip comments, revisions, bookmarks from document XML (in-memory) ---
        doc = strip_comment_refs(doc)
        doc = strip_revisions(doc, warnings)
        doc, bm_removed = strip_bookmarks(doc)
        result['bookmarks_removed'] = bm_removed

        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(doc)

        # --- 3. Update relationships ---
        if rels_xml:
            rels = rels_xml

            # Remove Visio embedding relationships
            rels = re.sub(r'<Relationship[^>]*Target="embeddings/[^"]*\.vsdx"[^/]*/>', '', rels)

            # Update converted image refs: .emf -> new format
            for emf_base in converted:
                rels = rels.replace(f'media/{emf_base}.emf"', f'media/{emf_base}.{fmt}"')

            with open(rels_path, 'w', encoding='utf-8') as f:
                f.write(rels)

        # --- 4. Update Content_Types ---
        ct_path = os.path.join(unpack_dir, '[Content_Types].xml')
        with open(ct_path, 'r', encoding='utf-8') as f:
            ct = f.read()

        ct = re.sub(r'<Override[^>]*\.vsdx[^>]*/>', '', ct)
        # Remove stale Default entries for formats no longer present
        has_emf = os.path.isdir(media_dir) and any(f.endswith('.emf') for f in os.listdir(media_dir))
        if not has_emf:
            ct = re.sub(r'<Default[^>]*Extension="emf"[^>]*/>', '', ct)
        has_vsdx = os.path.isdir(embed_dir) and any(f.endswith('.vsdx') for f in os.listdir(embed_dir))
        if not has_vsdx:
            ct = re.sub(r'<Default[^>]*Extension="vsdx"[^>]*/>', '', ct)
        _CONTENT_TYPE_MAP = {'jpg': 'image/jpeg', 'png': 'image/png'}
        if converted and fmt in _CONTENT_TYPE_MAP:
            ext_attr = f'Extension="{fmt}"'
            if ext_attr not in ct:
                mime = _CONTENT_TYPE_MAP[fmt]
                ct = ct.replace('</Types>',
                                f'<Default Extension="{fmt}" ContentType="{mime}"/></Types>')

        with open(ct_path, 'w', encoding='utf-8') as f:
            f.write(ct)

        # --- 5. Compress/resize oversized raster images ---
        skip_files = {f'{b}.{fmt}' for b in converted}
        compressed = compress_media_images(media_dir, max_width=max_width,
                                           quality=quality, skip_files=skip_files)
        result['images_compressed'] = compressed

        # --- 6. Deduplicate identical media files ---
        result['duplicates_removed'] = dedup_media(media_dir, unpack_dir)

        # --- 7. Remove personal info and sensitive data ---
        sanitize_core_props(unpack_dir)
        sanitize_app_props(unpack_dir)

        result['comments_removed'] = remove_comment_files(unpack_dir)

        # --- 9. Remove garbage parts ---
        result['garbage_removed'] = remove_garbage_parts(unpack_dir)

        # --- 10. Clean up references to deleted parts ---
        clean_content_types(unpack_dir)
        clean_relationships(unpack_dir)

        # --- 11. Interactive: show top 5 largest images, offer re-conversion ---
        if interactive and emf_to_vsdx:
            _interactive_reconvert(media_dir, emf_to_vsdx, conversions,
                                   tmp_dir, fmt, dpi, warnings, max_width=max_width)

        # --- 12. Repack (write to temp file first for atomicity) ---
        tmp_output = dst_path + '.tmp'
        try:
            with zipfile.ZipFile(tmp_output, 'w', zipfile.ZIP_DEFLATED) as zout:
                for root, dirs, files in os.walk(unpack_dir):
                    for f in sorted(files):
                        full = os.path.join(root, f)
                        arc = os.path.relpath(full, unpack_dir)
                        zout.write(full, arc)

            # Validate output before finalizing
            with zipfile.ZipFile(tmp_output, 'r') as zcheck:
                bad = zcheck.testzip()
                if bad:
                    raise RuntimeError(f'Output ZIP is corrupt: {bad}')
                if '[Content_Types].xml' not in zcheck.namelist():
                    raise RuntimeError('Output ZIP missing [Content_Types].xml')

            os.replace(tmp_output, dst_path)
        except Exception:
            if os.path.exists(tmp_output):
                os.remove(tmp_output)
            raise

    orig_size = os.path.getsize(src_path)
    final_size = os.path.getsize(dst_path)
    reduction = orig_size - final_size
    result['original_size_mb'] = round(orig_size / 1024 / 1024, 2)
    result['new_size_mb'] = round(final_size / 1024 / 1024, 2)
    result['reduction_mb'] = round(reduction / 1024 / 1024, 2)
    result['reduction_percent'] = round(reduction / orig_size * 100, 1) if orig_size else 0

    return result
