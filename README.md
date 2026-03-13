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

## License

MIT
