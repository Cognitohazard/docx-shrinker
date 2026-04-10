# docx-shrinker

Shrink and sanitize Word (.docx) documents. Converts embedded Visio diagrams to raster images, compresses oversized media, deduplicates files, and strips metadata, comments, tracked changes, macros, and other cruft.

## What it does

1. **Convert Visio embeddings** — `.vsdx` → PDF (via Visio COM) → JPG/PNG (via PyMuPDF). Falls back to keeping the EMF preview when Visio is unavailable.
2. **Convert OLE objects** — Replaces legacy VML `<w:object>` blocks with modern DrawingML `<w:drawing>` inline pictures.
3. **Compress images** — Resizes raster images exceeding a pixel width threshold and re-compresses JPGs.
4. **Deduplicate media** — Identifies identical files by hash and rewrites relationships to point to a single copy.
5. **Strip personal info** — Removes author, last modified by, company, manager, keywords, and other document properties.
6. **Remove comments and tracked changes** — Deletes comment files and accepts all revisions inline.
7. **Strip bookmarks** — Removes auto-generated bookmarks (`_GoBack`, empty).
8. **Remove garbage parts** — Thumbnail, VBA macros, printer settings, ActiveX controls, custom XML data.
9. **Clean up** — Updates `[Content_Types].xml` and `.rels` files to reflect removed parts.
10. **Validate output** — Checks ZIP integrity and presence of `[Content_Types].xml` before finalizing.

## Requirements

- **Python 3.10+**
- **PyMuPDF** (`pymupdf`) — image compression and PDF-to-image rendering
- **pywin32** — Visio COM automation (Windows only; Visio conversion is skipped if unavailable)
- **Microsoft Visio** (optional) — required only for converting embedded `.vsdx` to high-quality images

## Installation

```
pip install docx-shrinker
```

Or with [uv](https://docs.astral.sh/uv/):

```
uv tool install docx-shrinker
```

## Usage

### Command line

```
docx-shrinker report.docx
```

This produces `report (shrunk).docx` in the same directory.

Specify an output path:

```
docx-shrinker report.docx output.docx
```

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `--format {jpg,png}` | `jpg` | Image format for converted Visio figures |
| `--dpi N` | `300` | Rasterization DPI for Visio conversion |
| `--quality N` | `95` | JPG quality (1–100). Ignored for PNG. |
| `--max-width N` | `2000` | Max pixel width for raster images. `0` to disable. |
| `-i, --interactive` | off | After conversion, show top 5 largest images and offer to re-convert at different quality |
| `--version` | | Show version and exit |

### Examples

Convert Visio figures to PNG at 150 DPI:

```
docx-shrinker report.docx --format png --dpi 150
```

Aggressive compression (lower quality, smaller max width):

```
docx-shrinker report.docx --quality 80 --max-width 1200
```

Interactive mode to fine-tune large images:

```
docx-shrinker report.docx -i
```

### Python API

```python
from docx_shrinker import shrink_docx

result = shrink_docx("input.docx", "output.docx", fmt="jpg", dpi=300, quality=95)

print(f"{result['original_size_mb']} MB -> {result['new_size_mb']} MB")
print(f"Reduction: {result['reduction_percent']}%")
```

The `result` dict contains:

| Key | Type | Description |
|-----|------|-------------|
| `original_size_mb` | `float` | Original file size |
| `new_size_mb` | `float` | Output file size |
| `reduction_mb` | `float` | Size saved |
| `reduction_percent` | `float` | Percentage reduction |
| `output_path` | `str` | Path to the output file |
| `visio_converted` | `list` | `(name, size_kb)` tuples for each converted Visio diagram |
| `visio_removed` | `int` | Number of `.vsdx` embeddings removed |
| `images_compressed` | `list` | `(filename, old_kb, new_kb)` tuples |
| `duplicates_removed` | `int` | Number of duplicate media files removed |
| `comments_removed` | `int` | Number of comment files removed |
| `bookmarks_removed` | `int` | Number of bookmarks removed |
| `garbage_removed` | `list` | Names of removed garbage parts |
| `warnings` | `list` | Warning messages (e.g., Visio unavailable) |

## How it works

A `.docx` file is a ZIP archive containing XML and media files. docx-shrinker extracts the archive into a temp directory, applies all transformations in-place, then repacks it into a new ZIP. The original file is never modified.

Visio diagrams embedded as OLE objects include both the full `.vsdx` source and a low-resolution EMF preview image. docx-shrinker replaces these with a high-quality raster render and strips the heavy `.vsdx` originals — often the single biggest source of bloat.

## Technical reference

### Processing pipeline

`shrink_docx()` unpacks the `.docx` ZIP into a temp directory, applies all transformations in-place, then repacks into a new ZIP. The original file is never modified.

#### 1. Visio conversion (`.vsdx` → PDF → image)

Embedded Visio diagrams are OLE objects containing both the full `.vsdx` source and a low-resolution EMF preview. The conversion pipeline:

1. **`_export_vsdx_to_pdf`** — Opens each `.vsdx` via Visio COM (`ExportAsFixedFormat`) and exports to PDF. This two-step path exists because Visio's direct raster export (`Page.Export` to PNG/BMP) produces extremely low-quality output. The PDF intermediate preserves full vector fidelity.

2. **`_restore_pdf_images`** — Visio's PDF export **downscales and JPEG-compresses** any raster images embedded in `.vsdx` files, even if the originals are lossless PNGs (e.g., a 1590x633 PNG becomes a 668x266 JPEG with visible chroma artifacts). There is no COM setting to control this. The fix: extract the original images from the `.vsdx` ZIP, match them to degraded PDF images by aspect ratio, and replace them with `page.replace_image()`. The corrected PDF is saved and reopened before rasterization. A critical detail: the PDF image transform matrix often has a **negative Y scale** (`matrix.d < 0`, since PDF origin is bottom-left), meaning the image data is stored vertically flipped. When replacing, the original must be flipped to match. Not all images need flipping — some transforms have positive Y scale.

3. **`_border_clip_rect`** — Visio always draws a 0.75pt black stroked rectangle at the page edges of every exported PDF. This border is **not centered** on the page boundary — left/top are nearly flush while bottom/right overshoot outward by 0.02–0.12pt. The code detects this rectangle via `page.get_drawings()` (typically drawing #0), computes per-side inset from the actual stroke overshoot, and clips the render rect inward to exclude both the stroke and its anti-alias fringe.

   Earlier approaches that failed:
   - **Fixed uniform inset** (0.5pt or 1pt) — either clipped content or left borders on some sides due to the asymmetric overshoot.
   - **Pixel-level detection** after rasterization — fragile; anti-aliased gray pixels don't pass a simple threshold, and results varied per image.
   - **White rectangle overlay** — the overlay's own edges get anti-aliased, replacing one faint border with another.

4. **`_render_pdf_to_image`** — Rasterizes the first page of the corrected PDF to PNG/JPG via PyMuPDF, using the computed clip rect. Respects DPI, quality, and max-width settings.

#### 2. OLE/VML to DrawingML conversion

- **`extract_vml_dimensions`** — Parses width/height in EMU from `<w:object>` style attributes (handles both `pt` and `in` units).
- **`object_to_drawing`** — Rewrites legacy VML `<w:object>` blocks as modern DrawingML `<w:drawing>` inline pictures, preserving the image relationship ID and dimensions.
- **`next_doc_pr_id`** — Scans the document XML for the highest existing `docPr`/`cNvPr` id to generate unique IDs for new drawing elements.
- **`ensure_namespaces`** — Adds `wp:` and `r:` namespace declarations to the root `<w:document>` if missing, which is required for DrawingML elements.

#### 3. Image compression

- **`compress_media_images`** — Re-encodes oversized raster images in `word/media/`. PNGs with high estimated compression potential are converted to JPEG. Existing JPEGs are re-saved at the target quality. Images exceeding `max_width` are downscaled. Skips images that would grow larger after re-encoding.

#### 4. Media deduplication

- **`dedup_media`** — Identifies identical files in `word/media/` by MD5 hash. Keeps one canonical copy and rewrites all `.rels` references to point to it, removing the duplicates.

#### 5. Metadata and markup stripping

- **`sanitize_core_props`** / **`sanitize_app_props`** — Strip personal info (author, last modified by, company, manager, keywords, etc.) from `docProps/core.xml` and `docProps/app.xml`.
- **`remove_comment_files`** — Deletes `comments.xml`, `commentsExtended.xml`, and `commentsIds.xml`.
- **`strip_comment_refs`** — Removes comment range/reference XML tags (`commentRangeStart`, `commentRangeEnd`, `commentReference`) from document XML.
- **`strip_revisions`** — Accepts all tracked changes inline: unwraps `<w:ins>` content, removes `<w:del>` blocks and their content, strips revision property tags (`rPrChange`, `pPrChange`, `sectPrChange`, `tblPrChange`).
- **`strip_bookmarks`** — Removes auto-generated bookmarks (`_GoBack`, empty-name).

#### 6. Garbage part removal

- **`remove_garbage_parts`** — Deletes thumbnail, VBA macros (`vbaProject.bin`, `vbaData.xml`), printer settings, ActiveX controls, custom XML data, and the `.vsdx` embeddings themselves (after conversion).

#### 7. Cleanup and validation

- **`clean_content_types`** — Removes `[Content_Types].xml` entries referencing deleted parts.
- **`clean_relationships`** — Removes `.rels` entries referencing deleted parts across all relationship files.
- **`_strip_xml_tags`** / **`_strip_nested_tag`** — Low-level helpers for removing XML tags, handling arbitrarily nested structures by processing innermost matches first.
- Output ZIP integrity is validated (well-formed ZIP, `[Content_Types].xml` present) before finalizing.

#### 8. Interactive mode

- **`_interactive_reconvert`** — When `-i` is passed, presents the top 5 largest converted images and offers to re-convert selected ones at a different quality/DPI setting.

## License

MIT
