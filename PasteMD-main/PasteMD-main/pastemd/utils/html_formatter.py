"""Utilities for cleaning and formatting HTML fragments before conversion."""

from __future__ import annotations

import re
from typing import Dict, Optional

try:
    from bs4 import BeautifulSoup, NavigableString  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    BeautifulSoup = None  # type: ignore
    NavigableString = None  # type: ignore


def clean_html_content(html: str, options: Optional[Dict[str, object]] = None) -> str:
    """
    清理 HTML 内容，移除不可用元素，并按配置应用格式化规则。

    Args:
        html: 原始 HTML 内容。
        options: 可选格式化配置，如 ``{"strikethrough_to_del": True}``。

    Returns:
        清理后的 HTML 内容。
    """

    options = options or {}
    enable_strike_conversion = bool(options.get("strikethrough_to_del", True))

    if BeautifulSoup is not None:
        try:
            soup = BeautifulSoup(html, "lxml")

            # 删除所有 <svg> 标签
            for svg in soup.find_all("svg"):
                svg.decompose()

            # 删除 src 指向 .svg 的 <img> 标签
            for img in soup.find_all("img", src=True):
                if img["src"].lower().endswith(".svg"):
                    img.decompose()

            if enable_strike_conversion:
                _convert_strikethrough_to_del(soup)

            return f"<!DOCTYPE html>\n<meta charset='utf-8'>\n{str(soup)}"
        except Exception:
            # 解析失败时回退到正则处理
            pass

    # 如果没有 BeautifulSoup 或解析失败，使用简单的正则清理
    html = re.sub(r"<svg[^>]*>.*?</svg>", "", html, flags=re.DOTALL | re.IGNORECASE)
    html = re.sub(r'<img[^>]*src=["\'][^"\']*\.svg["\'][^>]*>', "", html, flags=re.IGNORECASE)
    if enable_strike_conversion:
        html = re.sub(r'~~([^~]+?)~~', r'<del>\1</del>', html)
    return html


def _convert_strikethrough_to_del(soup) -> None:
    """
    在 BeautifulSoup 解析树中查找文本节点，将 ``~~text~~`` 替换为 ``<del>text</del>``。

    Args:
        soup: BeautifulSoup 对象，会被原地修改。
    """

    if NavigableString is None:
        return

    # 递归处理所有文本节点
    for element in soup.find_all(text=True):
        if isinstance(element, NavigableString):
            if "~~" not in element:
                continue
            pattern = r'~~([^~]+?)~~'
            if not re.search(pattern, element):
                continue

            new_content = []
            last_end = 0
            for match in re.finditer(pattern, element):
                if match.start() > last_end:
                    new_content.append(element[last_end:match.start()])

                del_tag = soup.new_tag("del")
                del_tag.string = match.group(1)
                new_content.append(del_tag)
                last_end = match.end()

            if last_end < len(element):
                new_content.append(element[last_end:])

            parent = element.parent
            if not parent:
                continue
            index = parent.contents.index(element)
            element.extract()
            for i, item in enumerate(new_content):
                if isinstance(item, str):
                    parent.insert(index + i, NavigableString(item))
                else:
                    parent.insert(index + i, item)
