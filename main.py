"""
Standalone usage example for ppt-generator.

Run directly:
    python main.py chapter01.pdf --title "Chapter 1: Introduction" --book "My Book"
"""

import sys
import shutil
import tempfile
import logging
from pathlib import Path

# Add package to path for standalone usage
sys.path.insert(0, str(Path(__file__).parent))

from ppt_generator import generate_ppt

try:
    from pdf_image_extractor import extract_images, extract_vector_diagrams
    HAS_IMAGE_EXTRACTOR = True
except ImportError:
    HAS_IMAGE_EXTRACTOR = False

logger = logging.getLogger(__name__)


def cleanup_old_outputs(output_path: Path):
    """Remove the old output file and any leftover temp image directories."""
    # Remove old output PPTX
    if output_path.exists():
        output_path.unlink()
        print(f"🗑️  Removed old output: {output_path}")

    # Remove leftover temp image directories from previous runs
    temp_dir = Path(tempfile.gettempdir())
    cleaned = 0
    for d in temp_dir.glob("ppt_images_*"):
        if d.is_dir():
            shutil.rmtree(d, ignore_errors=True)
            cleaned += 1
    if cleaned:
        print(f"🗑️  Cleaned {cleaned} old temp image director{'ies' if cleaned != 1 else 'y'}")


def extract_text_and_images(pdf_path):
    """
    Extract text and images (raster + vector diagrams) from PDF.

    Returns:
        tuple: (text_content, images_dir) - extracted text and path to images directory
    """
    pdf_path = Path(pdf_path)
    images_dir = None

    if HAS_IMAGE_EXTRACTOR:
        try:
            temp_dir = tempfile.mkdtemp(prefix="ppt_images_")
            print(f"📸 Extracting images to {temp_dir}...")

            # Extract raster images
            result = extract_images(str(pdf_path), output_dir=temp_dir)
            raster_count = len(result.images)
            print(f"   Raster: found {result.total_found}, kept {result.kept}")

            # Extract vector diagrams (flowcharts, matrices, etc.)
            diagram_count = 0
            try:
                diagrams = extract_vector_diagrams(pdf_path, output_dir=temp_dir, dpi=200)
                diagram_count = len(diagrams)
                for d in diagrams:
                    print(f"   Diagram: page {d['page']}, {d['name']} ({d['width']}x{d['height']})")
            except Exception as e:
                print(f"   ⚠️  Diagram extraction failed: {e}")

            total = raster_count + diagram_count
            print(f"   Total content images: {total} ({raster_count} raster + {diagram_count} diagrams)")

            if total > 0:
                images_dir = temp_dir
        except Exception as e:
            print(f"⚠️  Image extraction failed: {e}")
    
    # Extract text from PDF
    print(f"📄 Extracting text from {pdf_path}...")
    try:
        import fitz
        doc = fitz.open(pdf_path)
        text_content = []
        for page_num, page in enumerate(doc):
            text = page.get_text()
            text_content.append(text)
            print(f"   Page {page_num + 1}/{len(doc)}")
        doc.close()
        text = "\n".join(text_content)
        print(f"   Extracted {len(text)} characters")
        return text, images_dir
    except Exception as e:
        print(f"⚠️  Text extraction failed: {e}")
        return "", images_dir


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Generate PPT from PDF chapter")
    parser.add_argument("pdf", help="Path to PDF file")
    parser.add_argument("-o", "--output", default="output.pptx", help="Output PPTX path")
    parser.add_argument("-t", "--title", default="", help="Chapter title")
    parser.add_argument("-b", "--book", default="", help="Book title")
    parser.add_argument("-n", "--number", type=int, default=0, help="Chapter number")
    parser.add_argument("-i", "--images", default=None, help="Images directory")
    parser.add_argument("--extract", action="store_true", help="Extract text and images before generation")
    args = parser.parse_args()

    if not Path(args.pdf).exists():
        print(f"Error: File not found: {args.pdf}")
        sys.exit(1)

    # Clean up old outputs before starting
    cleanup_old_outputs(Path(args.output))

    # Extract text and images if requested
    extracted_images_dir = args.images
    if args.extract:
        text, extracted_images_dir = extract_text_and_images(args.pdf)
        print(f"\n✅ Extraction complete. Starting PPT generation...\n")

    result = generate_ppt(
        args.pdf,
        output_path=args.output,
        chapter_title=args.title,
        book_title=args.book,
        chapter_number=args.number,
        images_dir=extracted_images_dir,
    )

    print(f"✅ Generated {result['slide_count']} slides → {args.output}")
    print(f"   Pages: {result['pages_processed']}, Sections: {result['sections_detected']}, "
          f"Images: {result['image_count']}, LLM: {'Yes' if result.get('used_llm') else 'No'}")


if __name__ == "__main__":
    main()
