#!/usr/bin/env python3
"""Fast visual index for photo PDFs. No OCR, no remote calls, no source changes."""
import argparse
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import fitz


def render_index(pdf_path: Path, out_dir: Path, columns: int = 4, rows: int = 4,
                 thumb_width: int = 420) -> list[Path]:
    if columns < 1 or rows < 1 or thumb_width < 120:
        raise ValueError('Invalid contact sheet dimensions')
    out_dir.mkdir(parents=True, exist_ok=True)
    names = []
    with fitz.open(pdf_path) as doc:
        count = len(doc)
        cell_h = int(thumb_width * 1.38) + 35
        pad = 14
        per_sheet = columns * rows
        font_path = Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
        font = ImageFont.truetype(str(font_path), 20) if font_path.exists() else ImageFont.load_default()
        for first in range(0, count, per_sheet):
            last = min(first + per_sheet, count)
            sheet = Image.new('RGB', (columns * (thumb_width + pad) + pad,
                                      rows * (cell_h + pad) + pad), 'white')
            draw = ImageDraw.Draw(sheet)
            for i in range(first, last):
                page = doc[i]
                scale = thumb_width / page.rect.width
                pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), colorspace=fitz.csRGB, alpha=False)
                image = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
                if image.height > cell_h - 35:
                    image.thumbnail((thumb_width, cell_h - 35), Image.Resampling.LANCZOS)
                j = i - first
                x = pad + (j % columns) * (thumb_width + pad)
                y = pad + (j // columns) * (cell_h + pad)
                sheet.paste(image, (x + (thumb_width - image.width)//2, y + 28))
                draw.text((x + 6, y + 6), f'Лист {i + 1:02d}', fill='black', font=font)
                draw.rectangle((x, y, x + thumb_width, y + cell_h), outline='#888', width=2)
            filename = out_dir / f'contact_{first//per_sheet+1:02d}.jpg'
            sheet.save(filename, quality=84, optimize=True)
            names.append(filename)
    return names


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf', type=Path)
    parser.add_argument('--out-dir', type=Path, required=True)
    parser.add_argument('--columns', type=int, default=4)
    parser.add_argument('--rows', type=int, default=4)
    parser.add_argument('--thumb-width', type=int, default=420)
    args = parser.parse_args()
    for path in render_index(args.pdf, args.out_dir, args.columns, args.rows, args.thumb_width):
        print(path)


if __name__ == '__main__':
    main()
