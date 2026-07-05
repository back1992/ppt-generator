"""
CLI entry point for ppt-generate command.

Usage:
    ppt-generate <pdf_path> --output slides.pptx --title "Chapter 1"
"""

import argparse
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(
        description="Generate a PowerPoint presentation from a PDF chapter.",
    )
    parser.add_argument(
        "pdf_path",
        help="Path to the PDF file",
    )
    parser.add_argument(
        "-o", "--output",
        default="output.pptx",
        help="Output PPTX file path (default: output.pptx)",
    )
    parser.add_argument(
        "-t", "--title",
        default="",
        help="Chapter title",
    )
    parser.add_argument(
        "-b", "--book-title",
        default="",
        help="Book title (optional)",
    )
    parser.add_argument(
        "-n", "--chapter-number",
        type=int,
        default=0,
        help="Chapter number (optional)",
    )
    parser.add_argument(
        "-i", "--images-dir",
        default=None,
        help="Directory containing images to include (optional)",
    )
    parser.add_argument(
        "--llm-api-key",
        default="",
        help="API key for LLM service (defaults to DASHSCOPE_API_KEY env var)",
    )
    parser.add_argument(
        "--llm-base-url",
        default="",
        help="Base URL for LLM API (defaults to DashScope)",
    )
    parser.add_argument(
        "--llm-model",
        default="",
        help="LLM model name (defaults to DASHSCOPE_MODEL env var)",
    )

    args = parser.parse_args()

    pdf_path = Path(args.pdf_path)
    if not pdf_path.exists():
        print(f"Error: PDF file not found: {pdf_path}", file=sys.stderr)
        sys.exit(1)

    output_path = Path(args.output)
    images_dir = Path(args.images_dir) if args.images_dir else None

    from .generator import PPTGenerator

    generator = PPTGenerator()

    print(f"Generating PPT from: {pdf_path}")
    print(f"Output: {output_path}")
    if args.title:
        print(f"Chapter: {args.title}")
    if args.book_title:
        print(f"Book: {args.book_title}")

    result = generator.generate_from_chapter(
        pdf_path=pdf_path,
        output_path=output_path,
        chapter_title=args.title,
        book_title=args.book_title,
        chapter_number=args.chapter_number,
        images_dir=images_dir,
        llm_api_key=args.llm_api_key,
        llm_base_url=args.llm_base_url,
        llm_model=args.llm_model,
    )

    print(f"\nDone! Generated {result['slide_count']} slides.")
    print(f"  Pages processed: {result['pages_processed']}")
    print(f"  Sections detected: {result['sections_detected']}")
    print(f"  Images included: {result['image_count']}")
    print(f"  LLM used: {'Yes' if result.get('used_llm') else 'No (heuristic fallback)'}")
    print(f"  Saved to: {output_path}")


if __name__ == "__main__":
    main()
