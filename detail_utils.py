"""CRM 富文本转换为适合桌面客户端展示的纯文本。"""

from __future__ import annotations

import re
from html.parser import HTMLParser


_IMAGE_MARKER_RE = re.compile(r"\ue000(\d+)\ue001")


class _RichTextParser(HTMLParser):
    BLOCK_TAGS = {
        "blockquote",
        "div",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "hr",
        "ol",
        "p",
        "pre",
        "table",
        "tr",
        "ul",
    }

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.image_sources: list[str] = []

    def _newline(self) -> None:
        if not self.parts or not self.parts[-1].endswith("\n"):
            self.parts.append("\n")

    def handle_starttag(
        self,
        tag: str,
        attrs: list[tuple[str, str | None]],
    ) -> None:
        tag = tag.lower()
        if tag in self.BLOCK_TAGS or tag == "br":
            self._newline()
        if tag == "li":
            self._newline()
            self.parts.append("• ")
        if tag == "img":
            src = str(dict(attrs).get("src") or "").strip()
            if src:
                index = len(self.image_sources)
                self.image_sources.append(src)
                self.parts.append(f"\ue000{index}\ue001")
            else:
                self.parts.append("[图片地址缺失]")

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in self.BLOCK_TAGS or tag.lower() == "li":
            self._newline()

    def handle_data(self, data: str) -> None:
        self.parts.append(data.replace("\xa0", " "))


def _normalize_text(value: str) -> str:
    lines: list[str] = []
    blank = False
    for raw_line in value.splitlines():
        line = re.sub(r"[ \t\r\f\v]+", " ", raw_line).strip()
        if line:
            lines.append(line)
            blank = False
        elif lines and not blank:
            lines.append("")
            blank = True
    return "\n".join(lines).strip()


def rich_content_segments(value: str) -> list[tuple[str, str]]:
    """按富文本中的原始位置返回 text/image 片段。"""
    if not value:
        return [("text", "（未填写）")]
    parser = _RichTextParser()
    parser.feed(value)
    parser.close()
    normalized = _normalize_text("".join(parser.parts))
    if not normalized:
        return [("text", "（未填写）")]

    segments: list[tuple[str, str]] = []
    cursor = 0
    for match in _IMAGE_MARKER_RE.finditer(normalized):
        before = normalized[cursor : match.start()]
        if before:
            segments.append(("text", before))
        index = int(match.group(1))
        if 0 <= index < len(parser.image_sources):
            segments.append(("image", parser.image_sources[index]))
        cursor = match.end()
    after = normalized[cursor:]
    if after:
        segments.append(("text", after))
    return segments or [("text", "（未填写）")]


def extract_image_sources(value: str) -> list[str]:
    """提取富文本图片地址并保持原始顺序。"""
    return [
        content
        for kind, content in rich_content_segments(value)
        if kind == "image"
    ]


def rich_text_to_plain(value: str) -> str:
    """将富文本转换为复制用纯文本，图片保留占位标记。"""
    parts = [
        content if kind == "text" else "[图片]"
        for kind, content in rich_content_segments(value)
    ]
    return _normalize_text("".join(parts)) or "（未填写）"
