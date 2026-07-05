"""
ppt_generator - Generate professional PowerPoint presentations from PDF chapter content.

Provides intelligent text cleaning, section detection, key point extraction,
LLM-powered slide design, and varied slide types (overview, detail, definition,
quote, summary, comparison).

Quick start:
    from ppt_generator import generate_ppt

    result = generate_ppt(
        "chapter.pdf",
        output_path="slides.pptx",
        chapter_title="Chapter 1",
        book_title="My Book",
    )
    print(f"Generated {result['slide_count']} slides")
"""

__version__ = "1.0.0"

from .generator import PPTGenerator
from .builder import PPTBuilder
from .text_analysis import (
    TextCleaner,
    SectionDetector,
    KeyPointExtractor,
    Section,
    CleanPage,
)
from .theme import COLORS, SLIDE_W, SLIDE_H, MARGIN, CONTENT_W

__all__ = [
    # Convenience function
    "generate_ppt",
    # Classes
    "PPTGenerator",
    "PPTBuilder",
    "TextCleaner",
    "SectionDetector",
    "KeyPointExtractor",
    # Dataclasses
    "Section",
    "CleanPage",
    # Theme constants
    "COLORS",
    "SLIDE_W",
    "SLIDE_H",
    "MARGIN",
    "CONTENT_W",
]


def generate_ppt(
    pdf_path,
    output_path="output.pptx",
    chapter_title="",
    book_title="",
    chapter_number=0,
    images_dir=None,
    llm_api_key="",
    llm_base_url="",
    llm_model="",
):
    """
    Convenience function to generate a PPT from a PDF chapter.

    Args:
        pdf_path: Path to the PDF file (str or Path).
        output_path: Path where the PPTX will be saved (str or Path).
        chapter_title: Title of the chapter.
        book_title: Title of the book (optional).
        chapter_number: Chapter number (optional).
        images_dir: Directory containing images to include (str, Path, or None).
        llm_api_key: API key for LLM service (defaults to DASHSCOPE_API_KEY env var).
        llm_base_url: Base URL for LLM API (defaults to DashScope).
        llm_model: LLM model name (defaults to DASHSCOPE_MODEL env var).

    Returns:
        dict with keys: slide_count, image_count, pages_processed,
        sections_detected, used_llm
    """
    from pathlib import Path

    pdf_path = Path(pdf_path)
    output_path = Path(output_path)
    if images_dir is not None:
        images_dir = Path(images_dir)

    generator = PPTGenerator()
    return generator.generate_from_chapter(
        pdf_path=pdf_path,
        output_path=output_path,
        chapter_title=chapter_title,
        book_title=book_title,
        chapter_number=chapter_number,
        images_dir=images_dir,
        llm_api_key=llm_api_key,
        llm_base_url=llm_base_url,
        llm_model=llm_model,
    )
