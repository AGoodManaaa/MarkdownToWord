# -*- coding: utf-8 -*-

import os
import re
from typing import Any, Dict, List, Optional

import requests
from PIL import Image


_IMAGE_PATTERN = re.compile(r"!\[[^\]]*\]\(([^\)]+)\)")
_TABLE_SEPARATOR_PATTERN = re.compile(
    r"^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$"
)


def _is_url(s: str) -> bool:
    return s.startswith(("http://", "https://"))


def _clean_image_target(raw: str) -> str:
    s = (raw or "").strip()
    if not s:
        return s
    if s.startswith("<") and s.endswith(">"):
        s = s[1:-1].strip()
    if s.startswith("data:"):
        return s
    if " " in s and not _is_url(s):
        s = s.split(" ", 1)[0].strip()
    return s


def _looks_like_table_row(line: str) -> bool:
    s = (line or "").strip()
    if not s:
        return False
    if s.startswith("|") and s.count("|") >= 2:
        return True
    if s.endswith("|") and "|" in s:
        return True
    return False


def _table_col_count(line: str) -> int:
    s = (line or "").strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    parts = [p.strip() for p in s.split("|")]
    parts = [p for p in parts if p != ""]
    return len(parts)


def _is_table_separator(line: str) -> bool:
    return bool(_TABLE_SEPARATOR_PATTERN.match((line or "").strip()))


def _check_remote_image(url: str, timeout_sec: float = 3.0) -> bool:
    try:
        resp = requests.head(url, allow_redirects=True, timeout=timeout_sec)
        if resp.status_code >= 400:
            return False
        return True
    except Exception:
        return False


def _try_get_local_image_stats(path: str) -> Dict[str, Any]:
    stats: Dict[str, Any] = {}
    try:
        stats['size_bytes'] = int(os.path.getsize(path))
    except Exception:
        pass
    try:
        with Image.open(path) as im:
            w, h = im.size
            stats['width_px'] = int(w)
            stats['height_px'] = int(h)
            stats['pixels'] = int(w) * int(h)
    except Exception:
        pass
    return stats


def run_preflight(
    markdown_text: str,
    base_dir: Optional[str] = None,
    options: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    issues: List[Dict[str, Any]] = []
    text = markdown_text or ""

    opt = options or {}
    check_remote_images = bool(opt.get('check_remote_images', False))
    max_images_warn = int(opt.get('max_images_warn', 30))
    large_image_mb = float(opt.get('large_image_mb', 10))
    large_image_pixels = int(opt.get('large_image_pixels', 16000000))

    lines = text.split("\n")

    code_fence_lines: List[int] = []
    for i, line in enumerate(lines, start=1):
        if line.strip().startswith("```"):
            code_fence_lines.append(i)
    if len(code_fence_lines) % 2 == 1:
        issues.append(
            {
                "level": "error",
                "line": code_fence_lines[-1],
                "message": "代码块未闭合（``` 数量为奇数）",
                "hint": "检查是否漏写了结尾的 ```",
            }
        )

    math_fence_lines: List[int] = []
    for i, line in enumerate(lines, start=1):
        if line.strip().startswith("$$"):
            math_fence_lines.append(i)
    if len(math_fence_lines) % 2 == 1:
        issues.append(
            {
                "level": "error",
                "line": math_fence_lines[-1],
                "message": "块级公式未闭合（$$ 数量为奇数）",
                "hint": "检查是否漏写了结尾的 $$",
            }
        )

    if base_dir is None:
        base_dir = os.getcwd()

    # 表格检查：缺少分隔行、列数不一致
    in_code = False
    idx = 0
    while idx < len(lines):
        line = lines[idx]
        if (line or "").strip().startswith("```"):
            in_code = not in_code
            idx += 1
            continue
        if in_code:
            idx += 1
            continue

        if _looks_like_table_row(line) and idx + 1 < len(lines) and _looks_like_table_row(lines[idx + 1]):
            header_cols = _table_col_count(lines[idx])
            sep_line_no = idx + 2
            if not _is_table_separator(lines[idx + 1]):
                issues.append(
                    {
                        "level": "warn",
                        "line": sep_line_no,
                        "message": "表格缺少分隔行（|---|---|）或分隔行格式不规范",
                        "hint": "表头下一行需要类似 | --- | --- | 的分隔行",
                    }
                )
            else:
                sep_cols = _table_col_count(lines[idx + 1])
                if header_cols and sep_cols and header_cols != sep_cols:
                    issues.append(
                        {
                            "level": "warn",
                            "line": sep_line_no,
                            "message": f"表格列数不一致：表头 {header_cols} 列，分隔行 {sep_cols} 列",
                            "hint": "检查每行 | 的数量是否一致",
                        }
                    )

            j = idx + 2
            while j < len(lines) and _looks_like_table_row(lines[j]) and (lines[j] or "").strip() != "":
                row_cols = _table_col_count(lines[j])
                if header_cols and row_cols and row_cols != header_cols:
                    issues.append(
                        {
                            "level": "warn",
                            "line": j + 1,
                            "message": f"表格列数不一致：表头 {header_cols} 列，该行 {row_cols} 列",
                            "hint": "检查该行是否少了/多了一个 | 或单元格",
                        }
                    )
                j += 1

            idx = j
            continue

        idx += 1

    image_count = 0
    for i, line in enumerate(lines, start=1):
        for m in _IMAGE_PATTERN.finditer(line):
            target_raw = m.group(1)
            target = _clean_image_target(target_raw)
            if not target or target.startswith("data:"):
                continue
            if _is_url(target):
                image_count += 1
                if check_remote_images and not _check_remote_image(target):
                    issues.append(
                        {
                            "level": "warn",
                            "line": i,
                            "message": f"网络图片可能不可访问: {target}",
                            "hint": "建议改为本地图片或检查网络/权限",
                        }
                    )
                continue

            image_count += 1
            if os.path.isabs(target):
                exists = os.path.exists(target)
                local_path = target
            else:
                local_path = os.path.join(base_dir, target)
                exists = os.path.exists(local_path)

            if not exists:
                issues.append(
                    {
                        "level": "error",
                        "line": i,
                        "message": f"图片不存在: {target}",
                        "hint": "检查图片路径，或把图片放到 Markdown 同目录",
                    }
                )
            else:
                stats = _try_get_local_image_stats(local_path)
                try:
                    size_mb = float(stats.get('size_bytes', 0)) / (1024 * 1024)
                    if size_mb and size_mb >= large_image_mb:
                        issues.append(
                            {
                                "level": "warn",
                                "line": i,
                                "message": f"图片文件较大（约 {size_mb:.1f}MB）: {target}",
                                "hint": "可能导致导出变慢或内存占用高，可考虑压缩图片",
                            }
                        )
                except Exception:
                    pass
                try:
                    pixels = int(stats.get('pixels', 0))
                    if pixels and pixels >= large_image_pixels:
                        issues.append(
                            {
                                "level": "warn",
                                "line": i,
                                "message": f"图片分辨率较高（约 {pixels} 像素）: {target}",
                                "hint": "可能导致导出变慢，可考虑缩放图片",
                            }
                        )
                except Exception:
                    pass

    if image_count >= max_images_warn:
        issues.append(
            {
                "level": "warn",
                "line": None,
                "message": f"图片数量较多（共 {image_count} 张）",
                "hint": "图片过多可能导致导出较慢或文件较大",
            }
        )

    return issues
