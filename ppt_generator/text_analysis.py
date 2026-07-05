"""
Text analysis utilities for PPT generation.

Handles PDF text cleaning, section detection, and key point extraction.
Supports both CJK (Chinese/Japanese/Korean) and Latin-script content.
"""

import re
from dataclasses import dataclass, field


def _is_cjk(text: str) -> bool:
    """Detect if text is predominantly CJK characters."""
    cjk_count = sum(1 for c in text if '\u4e00' <= c <= '\u9fff' or '\u3040' <= c <= '\u30ff')
    latin_count = sum(1 for c in text if c.isascii() and c.isalpha())
    return cjk_count > latin_count


@dataclass
class Section:
    title: str
    pages: list[int] = field(default_factory=list)
    text: str = ""
    key_points: list[str] = field(default_factory=list)
    level: int = 1  # 1=section, 2=sub-section


@dataclass
class CleanPage:
    page_num: int
    lines: list[str]
    raw_text: str


class TextCleaner:
    """Remove PDF artifacts: headers, footers, page numbers, footnote refs."""

    # Generic running header patterns (language-agnostic)
    HEADER_PATTERNS = [
        r'^\d+$',  # standalone page numbers
        r'^\d+\s*$',
    ]

    # CJK chapter/section patterns
    CJK_CHAPTER_PATTERN = re.compile(r'^第.+章\s')

    def clean_page(self, text: str, page_num: int, book_title: str = "", chapter_title: str = "") -> CleanPage:
        """Clean a single page's text, removing headers/footers/noise."""
        raw_lines = text.split('\n')
        clean_lines = []

        for line in raw_lines:
            line = line.strip()
            if not line:
                continue
            # Skip standalone page numbers
            if re.match(r'^\d{1,4}$', line):
                continue
            # Skip standalone footnote markers
            if re.match(r'^[a-z]$', line):
                continue
            # Skip running headers: exact match with book/chapter title
            if book_title and line == book_title:
                continue
            if chapter_title and line == chapter_title:
                continue
            # Skip generic chapter-like headers (第X章 or Chapter N)
            if self.CJK_CHAPTER_PATTERN.match(line) and len(line) < 20:
                continue
            if re.match(r'^Chapter\s+\d+\s*$', line, re.IGNORECASE) and len(line) < 30:
                continue
            # Clean inline footnote refs (CJK punctuation)
            line = re.sub(r'([。！？）"】])\s*[a-z]\s*$', r'\1', line)
            # Clean inline footnote refs (Latin punctuation)
            line = re.sub(r'([.!?)"\]])\s*[a-z]\s*$', r'\1', line)
            line = re.sub(r'\s+[a-z]$', '', line)
            # Clean all control chars
            line = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', line)
            line = line.strip()
            if line:
                clean_lines.append(line)

        return CleanPage(page_num=page_num, lines=clean_lines, raw_text=text)

    def merge_paragraphs(self, lines: list[str]) -> list[str]:
        """Merge broken lines into complete paragraphs, keeping section headers separate."""
        paragraphs = []
        current = []

        def _is_header(line):
            # CJK patterns
            if re.match(r'^第[一二三四五六七八九十]+节', line):
                return True
            if re.match(r'^[一二三四五六七八九十]+、', line) and len(line) < 30:
                return True
            # English patterns
            if re.match(r'^\d+\.\s', line) and len(line) < 60:
                return True
            if re.match(r'^\d+\.\d+\s', line) and len(line) < 60:
                return True
            if re.match(r'^Section\s+\d+', line, re.IGNORECASE) and len(line) < 60:
                return True
            if re.match(r'^Chapter\s+\d+', line, re.IGNORECASE) and len(line) < 60:
                return True
            # Bullet/list markers
            if re.match(r'^[•●○■□▸▹►»]\s', line) and len(line) < 80:
                return True
            return False

        # Sentence-ending punctuation (both CJK and Latin)
        SENTENCE_ENDS = ('。', '！', '？', '"', '）', '】', '；', '.', '!', '?', ')', ']', ';')
        SENTENCE_MIDS = ('。', '，', '；', '、', '.', ',', ';')

        for line in lines:
            if not current:
                current = [line]
            elif _is_header(line):
                paragraphs.append(''.join(current))
                current = [line]
            elif _is_header(current[0]) and len(current) == 1:
                paragraphs.append(''.join(current))
                current = [line]
            elif line and line[0] in ' \t\u3000':
                paragraphs.append(''.join(current))
                current = [line]
            elif current[-1].endswith(SENTENCE_ENDS) and len(current[-1]) < 60:
                paragraphs.append(''.join(current))
                current = [line]
            elif len(line) < 15 and not current[-1].endswith(SENTENCE_MIDS):
                paragraphs.append(''.join(current))
                current = [line]
            else:
                current.append(line)

        if current:
            paragraphs.append(''.join(current))

        return paragraphs


class SectionDetector:
    """Detect logical sections within chapter text."""

    # Section patterns: (regex, level)
    # Supports both CJK and English numbering
    SECTION_PATTERNS = [
        # CJK: 第X节 title
        (r'^第[一二三四五六七八九十]+节\s+(.{3,25})$', 1),
        (r'^第[一二三四五六七八九十]+节\s*$', 1),
        # CJK: 一、title
        (r'^[一二三四五六七八九十]+、(.{3,20})$', 2),
        # English: 1. title, 1.1 title
        (r'^\d+\.\s+(.{3,40})$', 2),
        (r'^\d+\.\d+\s+(.{3,40})$', 2),
        # English: Section N title
        (r'^Section\s+\d+[:.\s]+(.{3,40})$', 2),
    ]

    SUB_SECTION_PATTERNS = [
        r'^[一二三四五六七八九十]+、(.{3,25})$',
        r'^\d+\.\s+(.{3,25})$',
        r'^\d+\.\d+\s+(.{3,25})$',
        r'^第[一二三四五六七八九十]+节\s+(.{3,25})',
        r'^Section\s+\d+[:.\s]+(.{3,25})$',
    ]

    # Prefix patterns to strip from detected titles
    TITLE_PREFIX_PATTERNS = [
        r'^第[一二三四五六七八九十]+节\s*',
        r'^[一二三四五六七八九十]+、\s*',
        r'^\d+\.\s*',
        r'^\d+\.\d+\s*',
        r'^（[一二三四五六七八九十\d]+）\s*',
        r'^Section\s+\d+[:.\s]*',
    ]

    def _clean_title(self, title: str) -> str:
        """Strip numbering prefix from a section title."""
        for prefix in self.TITLE_PREFIX_PATTERNS:
            cleaned = re.sub(prefix, '', title, flags=re.IGNORECASE).strip()
            if cleaned and cleaned != title:
                return cleaned
        return title

    def detect(self, pages: list[CleanPage], chapter_title: str) -> list[Section]:
        """Detect sections from cleaned pages."""
        sections = []
        current_section = Section(title=chapter_title, level=0)

        for page in pages:
            for line in page.lines:
                detected = False
                for pattern, level in self.SECTION_PATTERNS:
                    match = re.match(pattern, line, re.IGNORECASE)
                    if match:
                        if current_section.pages:
                            sections.append(current_section)

                        title = match.group(0).strip()
                        title = self._clean_title(title)
                        if not title or len(title) < 3:
                            title = line.strip()
                        if len(title) > 40:
                            title = title[:37] + '...'

                        current_section = Section(
                            title=title,
                            level=level,
                            pages=[page.page_num],
                        )
                        detected = True
                        break

                if not detected:
                    if not current_section.pages or current_section.pages[-1] != page.page_num:
                        current_section.pages.append(page.page_num)

        if current_section.pages:
            sections.append(current_section)

        return self._balance_sections(sections, pages)

    def _split_large_sections(self, sections, pages):
        """Split sections with > 10 pages into sub-sections."""
        page_map = {p.page_num: p for p in pages}
        result = []

        for sec in sections:
            if len(sec.pages) <= 10:
                result.append(sec)
                continue

            sub_splits = []
            for pn in sec.pages:
                if pn in page_map:
                    page = page_map[pn]
                    for line in page.lines[:10]:
                        for pattern in self.SUB_SECTION_PATTERNS:
                            match = re.match(pattern, line, re.IGNORECASE)
                            if match:
                                title = match.group(0).strip()
                                title = self._clean_title(title)
                                if len(title) > 40:
                                    title = title[:37] + '...'
                                sub_splits.append((pn, title))
                                break

            if len(sub_splits) < 2:
                result.append(sec)
                continue

            for i, (start_page, title) in enumerate(sub_splits):
                if i + 1 < len(sub_splits):
                    end_page = sub_splits[i + 1][0]
                    sub_pages = [pn for pn in sec.pages if start_page <= pn < end_page]
                else:
                    sub_pages = [pn for pn in sec.pages if pn >= start_page]

                if sub_pages:
                    result.append(Section(
                        title=title,
                        pages=sub_pages,
                        level=2,
                    ))

            first_split = sub_splits[0][0]
            pre_pages = [pn for pn in sec.pages if pn < first_split]
            if pre_pages:
                result.insert(len(result) - len(sub_splits),
                              Section(title=sec.title, pages=pre_pages, level=sec.level))

        return result

    def _balance_sections(self, sections: list[Section], pages: list[CleanPage]) -> list[Section]:
        """Ensure 3-10 sections with reasonable content."""
        if not sections:
            chunk_size = max(4, len(pages) // 7)
            result = []
            for i in range(0, len(pages), chunk_size):
                chunk = pages[i:i + chunk_size]
                title = self._infer_title(chunk)
                result.append(Section(
                    title=title,
                    pages=[p.page_num for p in chunk],
                    level=2,
                ))
            return result

        sections = self._split_large_sections(sections, pages)

        while len(sections) > 10:
            min_size = float('inf')
            min_idx = 0
            for i in range(len(sections) - 1):
                combined = len(sections[i].pages) + len(sections[i+1].pages)
                if combined < min_size:
                    min_size = combined
                    min_idx = i
            sections[min_idx].pages.extend(sections[min_idx + 1].pages)
            sections.pop(min_idx + 1)

        if len(sections) < 3 and len(pages) > 8:
            return self._split_evenly(sections, pages)

        return sections

    def _split_evenly(self, sections: list[Section], pages: list[CleanPage]) -> list[Section]:
        """Split pages into ~5-8 sections."""
        target = min(8, max(5, len(pages) // 5))
        chunk_size = max(3, len(pages) // target)
        result = []
        for i in range(0, len(pages), chunk_size):
            chunk = pages[i:i + chunk_size]
            title = self._infer_title(chunk) or f"Section {len(result) + 1}"
            result.append(Section(
                title=title,
                pages=[p.page_num for p in chunk],
                level=2,
            ))
        return result

    def _infer_title(self, pages: list[CleanPage]) -> str:
        """Infer a title from the first substantive line of page group."""
        for page in pages:
            for line in page.lines:
                line = line.strip()
                if len(line) < 8 or len(line) > 60:
                    continue
                if re.match(r'^\d+$', line):
                    continue
                if self.CJK_CHAPTER_PATTERN.match(line) if hasattr(self, 'CJK_CHAPTER_PATTERN') else False:
                    continue
                if re.match(r'^Chapter\s+\d+', line, re.IGNORECASE):
                    continue
                return line
        return ""

    # Class-level pattern for _infer_title
    CJK_CHAPTER_PATTERN = re.compile(r'^第.+章\s')


class KeyPointExtractor:
    """Extract key sentences from a section's text."""

    # Generic patterns to skip (running headers, chapter markers)
    SKIP_PATTERNS = [
        r'^Chapter\s+\d+',
        r'^第.+章',
    ]

    def __init__(self, book_title: str = ""):
        """
        Args:
            book_title: Book title to filter out as running header.
        """
        self.book_title = book_title
        if book_title:
            self.SKIP_PATTERNS = self.SKIP_PATTERNS + [re.escape(book_title)]

    def extract(self, paragraphs: list[str], max_points: int = 5) -> list[str]:
        """Pick the most informative sentences."""
        if not paragraphs:
            return []

        scored = []
        for para in paragraphs:
            para = para.strip()
            if len(para) < 25:
                continue
            # Skip fragments that start mid-sentence (CJK and Latin punctuation)
            if para[0] in '，。、；：！？）】"，.!?)]"' and not para[0] in '""':
                continue
            # Skip footnote-like text
            if para.startswith(('a ', 'a ', 'b ', 'b ')):
                continue
            # Skip running headers
            skip = False
            for pat in self.SKIP_PATTERNS:
                if re.match(pat, para, re.IGNORECASE):
                    skip = True
                    break
            if skip:
                continue

            score = self._score_sentence(para)
            scored.append((score, para))

        # Sort by score descending
        scored.sort(key=lambda x: x[0], reverse=True)

        # Pick top N, maintaining original order
        selected = sorted(scored[:max_points], key=lambda x: paragraphs.index(x[1]) if x[1] in paragraphs else 999)

        result = []
        seen = set()
        for _, para in selected:
            para_key = para[:50]
            if para_key in seen:
                continue
            seen.add(para_key)

            # Truncate very long paragraphs at natural break points
            if len(para) > 180:
                cut = self._find_cut_point(para, 180)
                para = para[:cut + 1] + "..."

            # Final skip check
            skip = False
            for pat in self.SKIP_PATTERNS:
                if re.match(pat, para, re.IGNORECASE):
                    skip = True
                    break
            if skip:
                continue
            result.append(para)

        return result

    def _find_cut_point(self, text: str, max_len: int) -> int:
        """Find a natural cut point for truncation (CJK and Latin punctuation)."""
        # Try CJK punctuation first
        for punct in ['，', '。', '；']:
            cut = text[:max_len].rfind(punct)
            if cut >= 80:
                return cut
        # Try Latin punctuation
        for punct in [',', '.', ';']:
            cut = text[:max_len].rfind(punct)
            if cut >= 80:
                return cut
        return max_len - 3

    def _score_sentence(self, text: str) -> float:
        """Score a sentence by informativeness (language-agnostic)."""
        score = 0.0

        # Length: prefer medium-length sentences
        length = len(text)
        if 40 < length < 150:
            score += 3.0
        elif 20 < length < 200:
            score += 1.5

        # Named entities (names in parentheses, dates)
        if re.search(r'[（(][A-Z][a-z]+\s+[A-Z]', text):
            score += 2.0
        if re.search(r'\d{4}\s*年?', text):
            score += 1.5

        # Book titles (CJK: 《》, English: italic markers)
        if re.search(r'《[^》]+》', text):
            score += 2.0

        # Definition keywords (CJK and English)
        if re.search(r'(是指|所谓|意味着|定义为|即)', text):
            score += 3.0
        if re.search(r'\b(is defined as|refers to|means that|can be defined)\b', text, re.IGNORECASE):
            score += 3.0

        # Contrast/comparison (CJK and English)
        if re.search(r'(而|但|然而|相反|与之相对|区别在于)', text):
            score += 1.5
        if re.search(r'\b(however|although|in contrast|on the other hand|whereas|differs from)\b', text, re.IGNORECASE):
            score += 1.5

        # Quoted text
        if re.search(r'[""「」""]', text):
            score += 1.0

        # Penalty for chapter headers
        for pat in self.SKIP_PATTERNS:
            if re.match(pat, text, re.IGNORECASE):
                score -= 10.0

        return score
