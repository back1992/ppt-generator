"""
PPT Generator - main orchestrator for presentation generation.

Combines text analysis, LLM content analysis, and slide building
into a single `generate_from_chapter()` call.
"""

import re
import logging
from pathlib import Path
from typing import Optional

import fitz

from .text_analysis import TextCleaner, SectionDetector, KeyPointExtractor, Section
from .builder import PPTBuilder
from .llm_service import get_llm_analyzer

logger = logging.getLogger(__name__)


class PPTGenerator:
    """Generate professional presentations from PDF chapter content."""

    @staticmethod
    def _map_images_to_sections(
        image_files: list[Path],
        sections: list[Section],
    ) -> dict[int, list[Path]]:
        """
        Map images to sections based on their source page numbers.

        Images with filenames like 'page_5_img_1.png' are mapped to the
        section that covers page 5. Returns a dict of {section_index: [images]}.

        Args:
            image_files: List of image file paths.
            sections: List of Section objects with page ranges.

        Returns:
            Dict mapping section index to list of image paths.
        """
        result: dict[int, list[Path]] = {}

        for img_file in image_files:
            match = re.match(r'page_(\d+)_img_(\d+)', img_file.stem)
            if not match:
                continue
            img_page = int(match.group(1))

            # Find which section contains this page
            for sec_idx, sec in enumerate(sections):
                if sec.pages and sec.pages[0] <= img_page <= sec.pages[-1]:
                    result.setdefault(sec_idx, []).append(img_file)
                    break

        return result

    def _generate_with_llm(
        self,
        chapter_title: str,
        book_title: str,
        sections: list[Section],
        images_dir: Optional[Path] = None,
        llm_api_key: str = "",
        llm_base_url: str = "",
        llm_model: str = "",
    ) -> Optional[dict]:
        """Attempt to generate slides using LLM analysis. Returns slide data or None."""
        analyzer = get_llm_analyzer(
            api_key=llm_api_key,
            base_url=llm_base_url,
            model=llm_model,
        )
        if not analyzer:
            return None

        # Prepare section data for LLM
        section_data = []
        for sec in sections:
            if sec.title and sec.text:
                section_data.append({
                    "title": sec.title,
                    "content": sec.text[:1500],  # Limit per section
                    "pages": f"{sec.pages[0]}-{sec.pages[-1]}" if sec.pages else "",
                })

        if not section_data:
            return None

        try:
            result = analyzer.analyze_chapter(
                chapter_title=chapter_title,
                book_title=book_title,
                sections=section_data,
                max_slides=20,
            )

            if result.get("error"):
                return None

            return result

        except Exception as e:
            logger.error(f"LLM analysis failed: {e}")
            return None

    def _build_slides_from_llm(
        self,
        llm_result: dict,
        chapter_title: str,
        book_title: str,
        chapter_number: int,
        output_path: Path,
        images_dir: Optional[Path] = None,
    ) -> dict:
        """Build PPTX from LLM-structured slide data."""
        builder = PPTBuilder()
        slide_count = 0

        slides = llm_result.get("slides", [])

        for slide_data in slides:
            slide_type = slide_data.get("type", "content")

            if slide_type == "title":
                builder.slide_title(
                    slide_data.get("title", chapter_title),
                    slide_data.get("subtitle", book_title),
                    chapter_number,
                    0,  # page count unknown
                )
                slide_count += 1

            elif slide_type == "outline":
                if slide_data.get("items"):
                    from dataclasses import dataclass

                    @dataclass
                    class FakeSection:
                        title: str
                        level: int = 1

                    fake_sections = [FakeSection(title=item) for item in slide_data["items"][:12]]
                    builder.slide_outline(fake_sections)
                    slide_count += 1

            elif slide_type == "content":
                builder.slide_content(
                    slide_data.get("title", "Content"),
                    slide_data.get("points", []),
                    source_pages=slide_data.get("source_pages", ""),
                )
                slide_count += 1

            elif slide_type == "quote":
                builder.slide_quote(
                    slide_data.get("quote", ""),
                    attribution=slide_data.get("attribution", chapter_title),
                    context=slide_data.get("context", ""),
                )
                slide_count += 1

            elif slide_type == "definition":
                builder.slide_definition(
                    slide_data.get("term", ""),
                    slide_data.get("definition", ""),
                )
                slide_count += 1

            elif slide_type == "comparison":
                builder.slide_two_column(
                    slide_data.get("title", "Comparison"),
                    slide_data.get("left_title", "Left"),
                    slide_data.get("left_points", []),
                    slide_data.get("right_title", "Right"),
                    slide_data.get("right_points", []),
                )
                slide_count += 1

            elif slide_type == "summary":
                builder.slide_summary(
                    chapter_title,
                    slide_data.get("items", []),
                    {"slides": len(slides)},
                )
                slide_count += 1

        # Add image slides — pair with content slides where possible
        image_count = 0
        if images_dir and images_dir.exists():
            image_files = sorted(
                f for f in images_dir.glob("*")
                if f.suffix.lower() in ('.png', '.jpg', '.jpeg', '.gif', '.bmp')
            )

            # Map images to LLM content slides by page number
            placed_images = set()
            for img_file in image_files:
                match = re.match(r'page_(\d+)', img_file.stem)
                if not match:
                    continue
                img_page = int(match.group(1))
                img_label = img_file.stem.replace('_', ' ').title()

                # Try to find a content slide from LLM that mentions this page range
                # Insert image slide after the matching content slide
                if image_count < 6:  # reasonable limit
                    builder.slide_image(
                        img_file,
                        caption=img_label,
                        source=f"From: {chapter_title}",
                    )
                    slide_count += 1
                    image_count += 1
                    placed_images.add(str(img_file))

        builder.prs.save(str(output_path))

        return {
            "slide_count": slide_count,
            "image_count": image_count,
            "pages_processed": 0,
            "sections_detected": 0,
            "used_llm": True,
        }

    def generate_from_chapter(
        self,
        pdf_path: Path,
        output_path: Path,
        chapter_title: str,
        book_title: str = "",
        chapter_number: int = 0,
        images_dir: Optional[Path] = None,
        llm_api_key: str = "",
        llm_base_url: str = "",
        llm_model: str = "",
    ) -> dict:
        """
        Generate a professional PowerPoint presentation from a PDF chapter.

        Args:
            pdf_path: Path to the PDF file.
            output_path: Path where the PPTX will be saved.
            chapter_title: Title of the chapter.
            book_title: Title of the book (optional).
            chapter_number: Chapter number (optional).
            images_dir: Directory containing images to include (optional).
            llm_api_key: API key for LLM service (defaults to DASHSCOPE_API_KEY env var).
            llm_base_url: Base URL for LLM API (defaults to DashScope).
            llm_model: LLM model name (defaults to DASHSCOPE_MODEL env var).

        Returns:
            dict with keys: slide_count, image_count, pages_processed,
            sections_detected, used_llm
        """
        # Sanitize inputs - strip control characters
        chapter_title = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', chapter_title).strip()
        book_title = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', book_title).strip()

        # 1. Extract and clean text
        doc = fitz.open(str(pdf_path))
        cleaner = TextCleaner()
        raw_pages = []

        for i in range(doc.page_count):
            page = doc[i]
            text = page.get_text()
            if text.strip():
                raw_pages.append(cleaner.clean_page(
                    text, i + 1,
                    book_title=book_title,
                    chapter_title=chapter_title,
                ))

        doc.close()

        # 2. Merge paragraphs within each page
        for page in raw_pages:
            page.lines = cleaner.merge_paragraphs(page.lines)

        # 3. Detect sections
        detector = SectionDetector()
        sections = detector.detect(raw_pages, chapter_title)

        # 4. Extract text per section and find key points
        extractor = KeyPointExtractor(book_title=book_title)
        page_map = {p.page_num: p for p in raw_pages}

        for sec in sections:
            all_paras = []
            for pn in sec.pages:
                if pn in page_map:
                    all_paras.extend(page_map[pn].lines)
            sec.text = '\n'.join(all_paras)
            sec.key_points = extractor.extract(all_paras, max_points=5)

        # 5. Try LLM-based generation first
        llm_result = self._generate_with_llm(
            chapter_title, book_title, sections, images_dir,
            llm_api_key=llm_api_key,
            llm_base_url=llm_base_url,
            llm_model=llm_model,
        )

        if llm_result:
            result = self._build_slides_from_llm(
                llm_result, chapter_title, book_title,
                chapter_number, output_path, images_dir,
            )
            return result

        # Collect image files for heuristic path
        image_files = []
        if images_dir and images_dir.exists():
            image_files = sorted(images_dir.glob("*"))

        # Fallback: Build slides with heuristic approach
        builder = PPTBuilder()
        slide_count = 0

        # Map images to sections by page number
        section_images = self._map_images_to_sections(image_files, sections)

        # Title slide
        builder.slide_title(chapter_title, book_title, chapter_number, len(raw_pages))
        slide_count += 1

        # Outline slide
        if len(sections) > 2:
            builder.slide_outline(sections)
            slide_count += 1

        # Section slides - limit extras to keep total ~20 slides
        total_sections = len(sections)

        # Pick top 4 sections (by key_points count) for extra slides
        ranked = sorted(enumerate(sections),
                        key=lambda x: len(x[1].key_points), reverse=True)
        top_extra_indices = {idx for idx, _ in ranked[:4]}
        added_extras = 0
        seen_titles = set()
        placed_images = set()  # track which images have been placed
        image_count = 0

        for sec_idx, sec in enumerate(sections):
            # Section divider for major topic changes (every 3-4 sections)
            if sec_idx > 0 and sec_idx % 3 == 0 and total_sections > 5 and slide_count < 22:
                divider_title = sec.title[:30] if sec.title else f"Section {sec_idx + 1}"
                builder.slide_section_divider(divider_title, sec_idx + 1, total_sections)
                slide_count += 1

            # Content slide with key points (always add this)
            if sec.key_points:
                page_range = f"pp. {sec.pages[0]}–{sec.pages[-1]}" if sec.pages else ""
                slide_title = sec.title if sec.title else f"Section {sec_idx + 1}"
                # Clean long titles: truncate at natural break points
                if len(slide_title) > 30:
                    for sep in ['关于', '是指', '所谓', '这一', '从', '在']:
                        idx = slide_title.find(sep)
                        if 4 < idx < 25:
                            slide_title = slide_title[:idx]
                            break
                    else:
                        for sep in [': ', ' - ', ' — ', ', ', ' (', ' and ']:
                            idx = slide_title.find(sep)
                            if 4 < idx < 25:
                                slide_title = slide_title[:idx]
                                break
                        else:
                            slide_title = slide_title[:27] + '...'

                # Check if this section has an image to pair with
                sec_imgs = section_images.get(sec_idx, [])
                if sec_imgs and slide_count < 28:
                    # Use combined content + image slide
                    img_file = sec_imgs[0]
                    placed_images.add(str(img_file))
                    img_match = re.match(r'page_(\d+)_img_(\d+)', img_file.stem)
                    img_caption = f"Fig {img_match.group(2)}" if img_match else img_file.stem
                    builder.slide_content_with_image(
                        slide_title,
                        sec.key_points,
                        img_file,
                        source_pages=page_range,
                        image_caption=img_caption,
                    )
                    slide_count += 1
                    image_count += 1
                else:
                    builder.slide_content(
                        slide_title,
                        sec.key_points,
                        source_pages=page_range,
                    )
                    slide_count += 1

                # Add extras only for top sections and if under budget
                if sec_idx in top_extra_indices and added_extras < 5 and slide_count < 25:
                    if sec.title not in seen_titles:
                        seen_titles.add(sec.title)
                        for point in sec.key_points:
                            # Quote detection: CJK + Latin quotation marks
                            if re.search(r'[""「」""]', point) and 40 < len(point) < 150:
                                builder.slide_quote(
                                    point,
                                    attribution=chapter_title,
                                    context=f"pp. {sec.pages[0]}–{sec.pages[-1]}" if sec.pages else "",
                                )
                                slide_count += 1
                                added_extras += 1
                                break
                        else:
                            # Try definition: CJK + English keywords
                            for para in sec.text.split('\n'):
                                if re.search(r'(所谓|是指|定义为)', para) and 40 < len(para) < 180:
                                    builder.slide_definition(
                                        sec.title or f"Section {sec_idx + 1}",
                                        para,
                                    )
                                    slide_count += 1
                                    added_extras += 1
                                    break
                                if re.search(r'\b(is defined as|refers to|can be defined as)\b', para, re.IGNORECASE) and 40 < len(para) < 180:
                                    builder.slide_definition(
                                        sec.title or f"Section {sec_idx + 1}",
                                        para,
                                    )
                                    slide_count += 1
                                    added_extras += 1
                                    break

            elif sec.text.strip():
                # Section with text but no good key points
                paras = [p for p in sec.text.split('\n') if len(p.strip()) > 30][:4]
                if paras:
                    builder.slide_content(
                        sec.title or f"Section {sec_idx + 1}",
                        paras,
                    )
                    slide_count += 1

            # Comparison slide: detect contrast keywords in title
            COMPARISON_PATTERNS = r'(差异|比较|对比|vs\.?|versus|comparison|contrast|differences?\s+between)'
            if (re.search(COMPARISON_PATTERNS, sec.title or '', re.IGNORECASE) and slide_count < 26
                    and sec_idx not in seen_titles):
                seen_titles.add(sec_idx)
                # Split key points into two groups for the two columns
                half = max(2, len(sec.key_points) // 2)
                left_pts = sec.key_points[:half]
                right_pts = sec.key_points[half:]
                # Fallback: split text paragraphs if key_points insufficient
                if len(left_pts) < 2 or len(right_pts) < 2:
                    all_paras = [p.strip()[:150] for p in sec.text.split('\n') if len(p.strip()) > 30]
                    half = max(2, len(all_paras) // 2)
                    left_pts = all_paras[:half][:3]
                    right_pts = all_paras[half:][:3]
                # Deduplicate between columns
                right_pts = [p for p in right_pts if p not in left_pts]
                if left_pts and right_pts:
                    builder.slide_two_column(
                        sec.title[:30] if len(sec.title) > 30 else sec.title,
                        "Perspective A",
                        left_pts[:3],
                        "Perspective B",
                        right_pts[:3],
                    )
                    slide_count += 1

        # Remaining image slides (images not placed with sections)
        remaining_images = [
            f for f in image_files
            if str(f) not in placed_images
            and f.suffix.lower() in ('.png', '.jpg', '.jpeg', '.gif', '.bmp')
        ]
        for img_file in remaining_images:
            match = re.match(r'page_(\d+)', img_file.stem)
            label = f"Page {match.group(1)}" if match else img_file.stem
            # Clean up label: page_6_diagram → "Page 6 Diagram"
            label = img_file.stem.replace('_', ' ').title()
            builder.slide_image(
                img_file,
                caption=label,
                source=f"From: {chapter_title}",
            )
            slide_count += 1
            image_count += 1

        # Summary slide
        builder.slide_summary(
            chapter_title,
            [s.title for s in sections if s.title],
            {
                "pages": len(raw_pages),
                "sections": total_sections,
                "slides": slide_count + 1,
                "images": image_count,
            },
        )
        slide_count += 1

        # Save
        builder.prs.save(str(output_path))

        return {
            "slide_count": slide_count,
            "image_count": image_count,
            "pages_processed": len(raw_pages),
            "sections_detected": total_sections,
            "used_llm": False,
        }
