# ppt-generator

Generate professional PowerPoint presentations from PDF chapter content with intelligent text analysis and LLM-powered slide design.

## Features

- **Intelligent text cleaning** — removes PDF headers, footers, page numbers, footnote references
- **Section detection** — automatically detects chapter sections (Chinese patterns: 第X节, 一、, 1.)
- **Key point extraction** — scores and ranks sentences by informativeness
- **LLM-powered design** — optionally uses Qwen (DashScope) to design slide layouts and content
- **Varied slide types** — title, outline, content, quote, definition, comparison, image, summary
- **Heuristic fallback** — works without LLM using rule-based slide generation

## Installation

```bash
# Basic (no LLM)
pip install -e .

# With LLM support
pip install -e ".[llm]"

# Development
pip install -e ".[dev]"
```

## Quick Start

### Python API

```python
from ppt_generator import generate_ppt

result = generate_ppt(
    "chapter.pdf",
    output_path="slides.pptx",
    chapter_title="第一章 传播学概述",
    book_title="传播学引论",
    chapter_number=1,
    images_dir="./images",  # optional
)

print(f"Generated {result['slide_count']} slides")
# result keys: slide_count, image_count, pages_processed, sections_detected, used_llm
```

### CLI

```bash
ppt-generate chapter.pdf \
    --output slides.pptx \
    --title "第一章 传播学概述" \
    --book-title "传播学引论" \
    --chapter-number 1 \
    --images-dir ./images
```

### Standalone Script

```bash
python main.py chapter.pdf \
    -t "第一章 传播学概述" \
    -b "传播学引论" \
    -o slides.pptx
```

## LLM Configuration

For AI-powered slide design, set the DashScope API key:

```bash
export DASHSCOPE_API_KEY="your-api-key"

# Optional overrides
export DASHSCOPE_BASE_URL="https://dashscope.aliyuncs.com/compatible-mode/v1"
export DASHSCOPE_MODEL="qwen-plus"
```

Or pass directly in code:

```python
result = generate_ppt(
    "chapter.pdf",
    output_path="slides.pptx",
    chapter_title="Chapter 1",
    llm_api_key="your-api-key",
    llm_base_url="https://your-llm-api.com/v1",
    llm_model="your-model",
)
```

When LLM is unavailable, the generator falls back to heuristic-based slide generation.

## Advanced Usage

### Fine-grained control

```python
from ppt_generator import (
    PPTGenerator, PPTBuilder,
    TextCleaner, SectionDetector, KeyPointExtractor,
)

# 1. Clean text
cleaner = TextCleaner()
page = cleaner.clean_page(raw_text, page_num=1, book_title="My Book")

# 2. Detect sections
detector = SectionDetector()
sections = detector.detect(pages, chapter_title="Chapter 1")

# 3. Extract key points
extractor = KeyPointExtractor()
points = extractor.extract(paragraphs, max_points=5)

# 4. Build slides manually
builder = PPTBuilder()
builder.slide_title("Chapter 1", "My Book", 1, 20)
builder.slide_content("Key Concept", ["Point 1", "Point 2"])
builder.slide_quote("Important quote...", attribution="Author")
builder.prs.save("output.pptx")
```

## Package Structure

```
ppt_generator/
├── __init__.py          # Public API exports
├── generator.py         # PPTGenerator (main orchestrator)
├── builder.py           # PPTBuilder (slide construction)
├── text_analysis.py     # TextCleaner, SectionDetector, KeyPointExtractor
├── llm_service.py       # LLMContentAnalyzer (DashScope/Qwen)
├── theme.py             # Colors, dimensions, constants
└── cli.py               # CLI: ppt-generate
```

## Dependencies

| Core | Optional (LLM) |
|------|---------------|
| `python-pptx>=0.6.23` | `openai>=1.0.0` |
| `PyMuPDF>=1.24.0` | `httpx>=0.24.0` |

## License

MIT
