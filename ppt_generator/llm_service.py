"""
LLM Service for PPT Content Analysis

Uses Qwen via DashScope (OpenAI-compatible API) to analyze chapter text
and produce structured slide content.
"""

import os
import json
import logging
from typing import Optional

logger = logging.getLogger(__name__)


class LLMContentAnalyzer:
    """Analyze chapter content using Qwen LLM to produce structured slide data."""

    def __init__(self, api_key: str = "", base_url: str = "", model: str = ""):
        """
        Initialize LLM analyzer.

        Args:
            api_key: API key for the LLM service. Defaults to DASHSCOPE_API_KEY env var.
            base_url: API base URL. Defaults to DashScope compatible endpoint.
            model: Model name. Defaults to DASHSCOPE_MODEL env var or 'qwen-plus'.
        """
        import httpx
        from openai import OpenAI

        self.api_key = api_key or os.getenv("DASHSCOPE_API_KEY", "")
        if not self.api_key:
            raise ValueError("API key not configured. Set DASHSCOPE_API_KEY or pass api_key.")

        self.base_url = base_url or os.getenv(
            "DASHSCOPE_BASE_URL",
            "https://dashscope.aliyuncs.com/compatible-mode/v1"
        )
        self.model = model or os.getenv("DASHSCOPE_MODEL", "qwen-plus")

        # Temporarily clear proxy env vars (DashScope is a domestic China API, no proxy needed)
        saved_proxy = {}
        for key in ['ALL_PROXY', 'all_proxy', 'HTTPS_PROXY', 'https_proxy',
                     'HTTP_PROXY', 'http_proxy']:
            if key in os.environ:
                saved_proxy[key] = os.environ.pop(key)

        try:
            http_client = httpx.Client(
                base_url=self.base_url,
                timeout=60.0,
                verify=False,
                proxy=None,
            )

            self.client = OpenAI(
                api_key=self.api_key,
                base_url=self.base_url,
                default_headers={"Content-Type": "application/json"},
                http_client=http_client,
            )
        finally:
            # Restore proxy env vars
            os.environ.update(saved_proxy)

    @staticmethod
    def _is_cjk(text: str) -> bool:
        """Detect if text is predominantly CJK characters."""
        cjk = sum(1 for c in text if '\u4e00' <= c <= '\u9fff' or '\u3040' <= c <= '\u30ff')
        latin = sum(1 for c in text if c.isascii() and c.isalpha())
        return cjk > latin

    def _build_prompt(self, chapter_title, book_title, sections, max_slides, is_cjk):
        """Build the prompt in the appropriate language."""
        sections_text = "\n\n".join([
            f"[{s['title']}]\n{s['content'][:800]}"
            for s in sections[:15]
        ])

        if is_cjk:
            return f"""你是一位专业的教育内容设计师，擅长将学术章节转化为高质量的演示文稿。

请分析以下章节内容，设计一份结构清晰、视觉丰富的PPT大纲。

章节信息：
- 书名：{book_title}
- 章节：{chapter_title}
- 目标幻灯片数：{max_slides}

章节内容：
{sections_text}

请按以下JSON格式返回设计建议：

{{
  "theme": "academic_blue",
  "slides": [
    {{"type": "title", "title": "章节标题", "subtitle": "书名"}},
    {{"type": "outline", "title": "本章概览", "items": ["要点1", "要点2"]}},
    {{"type": "content", "title": "小节标题", "points": ["要点1", "要点2"], "source_pages": "pp. 1-3"}},
    {{"type": "quote", "quote": "重要引用句", "attribution": "来源", "context": "上下文说明"}},
    {{"type": "definition", "term": "术语", "definition": "定义说明"}},
    {{"type": "comparison", "title": "对比标题", "left_title": "左侧", "left_points": ["要点"], "right_title": "右侧", "right_points": ["要点"]}},
    {{"type": "summary", "title": "本章小结", "items": ["要点1", "要点2"]}}
  ]
}}

要求：
1. 保持学术严谨性，提炼核心概念
2. 每个section最多1个content slide
3. 选择最重要的2-3个概念做quote或definition slide
4. 如有对比概念，使用comparison slide
5. 控制总幻灯片数在{max_slides}以内
6. 只返回JSON，不要其他说明文字
"""
        else:
            return f"""You are a professional educational content designer skilled at transforming academic chapters into high-quality presentations.

Please analyze the following chapter content and design a well-structured, visually rich PPT outline.

Chapter information:
- Book: {book_title}
- Chapter: {chapter_title}
- Target slide count: {max_slides}

Chapter content:
{sections_text}

Return your design in the following JSON format:

{{
  "theme": "academic_blue",
  "slides": [
    {{"type": "title", "title": "Chapter title", "subtitle": "Book name"}},
    {{"type": "outline", "title": "Chapter Overview", "items": ["Point 1", "Point 2"]}},
    {{"type": "content", "title": "Section title", "points": ["Point 1", "Point 2"], "source_pages": "pp. 1-3"}},
    {{"type": "quote", "quote": "Important quote", "attribution": "Source", "context": "Context"}},
    {{"type": "definition", "term": "Term", "definition": "Definition text"}},
    {{"type": "comparison", "title": "Comparison title", "left_title": "Left side", "left_points": ["Point"], "right_title": "Right side", "right_points": ["Point"]}},
    {{"type": "summary", "title": "Chapter Summary", "items": ["Point 1", "Point 2"]}}
  ]
}}

Requirements:
1. Maintain academic rigor, distill core concepts
2. At most 1 content slide per section
3. Select the 2-3 most important concepts for quote or definition slides
4. Use comparison slides when contrasting concepts exist
5. Keep total slides within {max_slides}
6. Return only valid JSON, no other text
"""

    def analyze_chapter(
        self,
        chapter_title: str,
        book_title: str,
        sections: list[dict],
        max_slides: int = 20,
    ) -> dict:
        """
        Analyze chapter content and produce structured slide data.

        Auto-detects content language (CJK vs Latin) and uses the appropriate
        prompt template.

        Args:
            chapter_title: Title of the chapter
            book_title: Title of the book
            sections: List of sections with title and content
            max_slides: Target number of slides

        Returns:
            Structured JSON with slide layouts and content
        """
        # Detect language from chapter content
        sample_text = chapter_title + " " + " ".join(s.get("content", "")[:200] for s in sections[:5])
        is_cjk = self._is_cjk(sample_text)

        prompt = self._build_prompt(chapter_title, book_title, sections, max_slides, is_cjk)

        system_msg = "你是PPT设计专家，只返回有效的JSON格式。" if is_cjk \
            else "You are a PPT design expert. Return only valid JSON."

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_msg},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=4000,
                response_format={"type": "json_object"}
            )

            content = response.choices[0].message.content
            result = json.loads(content)

            logger.info(f"LLM analysis complete: {len(result.get('slides', []))} slides designed")
            return result

        except Exception as e:
            logger.error(f"LLM analysis failed: {e}")
            return {"error": str(e), "slides": []}

    def is_available(self) -> bool:
        """Check if LLM service is configured and available."""
        return bool(self.api_key)


# ─── Module-level singleton ────────────────────────────────────────────────

_llm_analyzer: Optional[LLMContentAnalyzer] = None


def get_llm_analyzer(
    api_key: str = "",
    base_url: str = "",
    model: str = "",
) -> Optional[LLMContentAnalyzer]:
    """
    Get or create the LLM analyzer singleton.

    Args:
        api_key: API key (defaults to DASHSCOPE_API_KEY env var)
        base_url: API base URL (defaults to DashScope)
        model: Model name (defaults to DASHSCOPE_MODEL env var)

    Returns:
        LLMContentAnalyzer instance, or None if not configured.
    """
    global _llm_analyzer
    if _llm_analyzer is None:
        try:
            _llm_analyzer = LLMContentAnalyzer(
                api_key=api_key, base_url=base_url, model=model
            )
        except (ValueError, ImportError):
            return None
    return _llm_analyzer


def reset_llm_analyzer():
    """Reset the singleton (useful for testing)."""
    global _llm_analyzer
    _llm_analyzer = None
