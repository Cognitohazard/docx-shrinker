"""Command-line interface for docx-shrinker."""

import argparse
import os
import sys
from importlib.metadata import version, PackageNotFoundError

from .core import shrink_docx

try:
    __version__ = version("docx-shrinker")
except PackageNotFoundError:
    from . import __version__


_TAG_LABELS = {
    'dc:creator': 'Author',
    'cp:lastModifiedBy': 'Last modified by',
    'cp:lastPrinted': 'Last printed',
    'cp:revision': 'Revision',
    'dc:subject': 'Subject',
    'cp:keywords': 'Keywords',
    'cp:category': 'Category',
    'cp:contentStatus': 'Content status',
    'Company': 'Company',
    'Manager': 'Manager',
    'HyperlinkBase': 'Hyperlink base',
}


def _print_result(result):
    """Print a clean summary of what was done."""
    lines = []

    # Visio conversions
    for name, size_kb in result['visio_converted']:
        lines.append(f'  Visio converted:    {name} -> {size_kb} KB')
    if result['visio_removed']:
        lines.append(f'  Visio removed:      {result["visio_removed"]} embedding(s)')

    # Image compression
    for fname, old_kb, new_kb in result['images_compressed']:
        lines.append(f'  Compressed:         {fname}: {old_kb} KB -> {new_kb} KB')

    # Deduplication
    if result['duplicates_removed']:
        lines.append(f'  Deduplicated:       {result["duplicates_removed"]} file(s)')
    personal = result.get('personal_info_stripped', {})
    if personal:
        for tag, value in personal.items():
            label = _TAG_LABELS.get(tag, tag)
            # Truncate long values
            display = value if len(value) <= 60 else value[:57] + '...'
            lines.append(f'  Stripped:           {label}: {display}')
    else:
        lines.append(f'  Personal info:      (none found)')

    # Tracked changes
    revisions = result.get('revisions_stripped', {})
    rev_parts = []
    if revisions.get('deletions'):
        rev_parts.append(f'{revisions["deletions"]} deletion(s)')
    if revisions.get('insertions'):
        rev_parts.append(f'{revisions["insertions"]} insertion(s)')
    if revisions.get('property_changes'):
        rev_parts.append(f'{revisions["property_changes"]} property change(s)')
    if rev_parts:
        lines.append(f'  Tracked changes:    {", ".join(rev_parts)} accepted')
    else:
        lines.append(f'  Tracked changes:    (none found)')

    if result['comments_removed']:
        lines.append(f'  Comments:           {result["comments_removed"]} file(s) removed')
    if result['bookmarks_removed']:
        lines.append(f'  Bookmarks:          {result["bookmarks_removed"]} removed')

    # Garbage parts
    if result['garbage_removed']:
        lines.append(f'  Garbage removed:    {", ".join(result["garbage_removed"])}')

    # Size summary
    lines.append('')
    lines.append(f'  {result["original_size_mb"]} MB -> {result["new_size_mb"]} MB '
                 f'(-{result["reduction_mb"]} MB, {result["reduction_percent"]}%)')

    print('\n'.join(lines))


def _print_warnings(warnings):
    """Print any warnings that occurred."""
    for w in warnings:
        print(f'  WARNING: {w}', file=sys.stderr)


def main() -> int:
    """Main entry point for the CLI."""
    parser = argparse.ArgumentParser(
        prog='docx-shrinker',
        description='Shrink and sanitize a Word document.',
        epilog='''steps performed:
  1. Convert embedded Visio .vsdx -> PDF (via Visio COM) -> JPG/PNG (via PyMuPDF)
     Falls back to keeping EMF previews when Visio is unavailable.
  2. Convert OLE/VML objects to DrawingML inline pictures
  3. Compress/resize oversized raster images (--max-width)
  4. Deduplicate identical media files
  5. Strip personal info (author, company, manager, etc.)
  6. Remove comments, tracked changes, and revision history
  7. Strip internal bookmarks (_GoBack, etc.)
  8. Remove thumbnail, VBA macros, printer settings, ActiveX, custom XML
  9. Clean up relationships and content types
  10. Validate output ZIP integrity''',
        formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('input', help='Input .docx file')
    parser.add_argument('output', nargs='?', default=None,
                        help='Output .docx file (default: "input (shrunk).docx")')
    parser.add_argument('--format', choices=['jpg', 'png'], default='jpg',
                        help='Image format for converted Visio figures (default: jpg)')
    parser.add_argument('--dpi', type=int, default=300,
                        help='Rasterization DPI (default: 300)')
    parser.add_argument('--quality', type=int, default=95,
                        help='JPG quality 1-100 (default: 95). PNG is always lossless.')
    parser.add_argument('--max-width', type=int, default=2000,
                        help='Max pixel width for non-Visio images (default: 2000). '
                             '0 to disable resizing.')
    parser.add_argument('-i', '--interactive', action='store_true',
                        help='After conversion, show top 5 largest images and '
                             'offer to re-convert at different quality.')
    parser.add_argument('--version', action='version', version=f'%(prog)s {__version__}')

    args = parser.parse_args()

    src = args.input
    if args.output:
        dst = args.output
    else:
        base, ext = os.path.splitext(src)
        dst = f'{base} (shrunk){ext}'

    try:
        print(f'Shrinking: {src}')
        result = shrink_docx(src, dst, fmt=args.format, dpi=args.dpi,
                             quality=args.quality, max_width=args.max_width,
                             interactive=args.interactive)
        _print_warnings(result['warnings'])
        _print_result(result)
        print(f'Saved: {dst}')
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
