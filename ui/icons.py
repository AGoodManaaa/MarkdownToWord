# -*- coding: utf-8 -*-
"""图标和表情管理模块 - 提供丰富可爱的 UI 图标"""

from dataclasses import dataclass
from typing import Dict, Optional


@dataclass
class IconSet:
    """图标集合"""
    # 文件操作
    file_new: str = "📄"
    file_open: str = "📂"
    file_save: str = "💾"
    file_export: str = "📤"
    file_import: str = "📥"
    file_pdf: str = "📑"
    file_word: str = "📝"
    file_html: str = "🌐"
    file_image: str = "🖼️"
    
    # 编辑操作
    edit_undo: str = "↩️"
    edit_redo: str = "↪️"
    edit_cut: str = "✂️"
    edit_copy: str = "📋"
    edit_paste: str = "📌"
    edit_delete: str = "🗑️"
    edit_select_all: str = "☑️"
    edit_find: str = "🔍"
    edit_replace: str = "🔄"
    edit_format: str = "✨"
    
    # 视图操作
    view_preview: str = "👁️"
    view_sidebar: str = "📊"
    view_minimap: str = "🗺️"
    view_fullscreen: str = "🖥️"
    view_split_h: str = "⬜"
    view_split_v: str = "⬛"
    view_zoom_in: str = "🔎"
    view_zoom_out: str = "🔍"
    view_focus: str = "🎯"
    view_reading: str = "📖"
    
    # 工具
    tool_settings: str = "⚙️"
    tool_theme: str = "🎨"
    tool_help: str = "❓"
    tool_info: str = "ℹ️"
    tool_warning: str = "⚠️"
    tool_error: str = "❌"
    tool_success: str = "✅"
    tool_loading: str = "⏳"
    tool_refresh: str = "🔄"
    
    # 格式化
    format_bold: str = "𝐁"
    format_italic: str = "𝐼"
    format_underline: str = "U̲"
    format_strike: str = "S̶"
    format_code: str = "⟨⟩"
    format_quote: str = "❝"
    format_link: str = "🔗"
    format_list_ul: str = "•"
    format_list_ol: str = "1."
    format_table: str = "▦"
    format_heading: str = "H"
    format_hr: str = "―"
    
    # 插入
    insert_image: str = "🖼️"
    insert_link: str = "🔗"
    insert_table: str = "📊"
    insert_code: str = "💻"
    insert_math: str = "∑"
    insert_emoji: str = "😊"
    insert_date: str = "📅"
    insert_toc: str = "📑"
    
    # 协作
    collab_user: str = "👤"
    collab_users: str = "👥"
    collab_chat: str = "💬"
    collab_share: str = "🔗"
    collab_invite: str = "📨"
    collab_online: str = "🟢"
    collab_offline: str = "🔴"
    collab_away: str = "🟡"
    collab_meeting: str = "🎥"
    collab_lock: str = "🔒"
    collab_unlock: str = "🔓"
    
    # AI 功能
    ai_assistant: str = "🤖"
    ai_write: str = "✍️"
    ai_translate: str = "🌍"
    ai_summary: str = "📝"
    ai_polish: str = "💎"
    ai_magic: str = "✨"
    
    # 状态
    status_saved: str = "💾"
    status_modified: str = "📝"
    status_syncing: str = "🔄"
    status_synced: str = "☁️"
    status_error: str = "⚠️"
    status_connected: str = "🌐"
    status_disconnected: str = "📵"
    
    # 导航
    nav_home: str = "🏠"
    nav_back: str = "⬅️"
    nav_forward: str = "➡️"
    nav_up: str = "⬆️"
    nav_down: str = "⬇️"
    nav_folder: str = "📁"
    nav_file: str = "📄"
    nav_search: str = "🔍"
    
    # 其他
    misc_star: str = "⭐"
    misc_heart: str = "❤️"
    misc_fire: str = "🔥"
    misc_sparkle: str = "✨"
    misc_rocket: str = "🚀"
    misc_lightning: str = "⚡"
    misc_clock: str = "🕐"
    misc_calendar: str = "📅"
    misc_pin: str = "📌"
    misc_tag: str = "🏷️"
    misc_bookmark: str = "🔖"
    misc_bell: str = "🔔"
    misc_gift: str = "🎁"
    misc_trophy: str = "🏆"
    misc_medal: str = "🏅"
    misc_crown: str = "👑"
    misc_gem: str = "💎"
    misc_rainbow: str = "🌈"
    misc_sun: str = "☀️"
    misc_moon: str = "🌙"
    misc_cloud: str = "☁️"
    misc_umbrella: str = "☂️"
    misc_snowflake: str = "❄️"
    misc_leaf: str = "🍃"
    misc_flower: str = "🌸"
    misc_tree: str = "🌳"
    misc_cat: str = "🐱"
    misc_dog: str = "🐶"
    misc_panda: str = "🐼"
    misc_unicorn: str = "🦄"
    misc_butterfly: str = "🦋"


# 可爱风格图标集
CUTE_ICONS = IconSet(
    # 文件操作 - 更可爱的版本
    file_new="📃",
    file_open="📂",
    file_save="💾",
    file_export="📤",
    file_import="📥",
    file_pdf="📕",
    file_word="📘",
    file_html="🌐",
    file_image="🖼️",
    
    # 编辑操作
    edit_undo="↩️",
    edit_redo="↪️",
    edit_cut="✂️",
    edit_copy="📋",
    edit_paste="📌",
    edit_delete="🗑️",
    edit_select_all="☑️",
    edit_find="🔎",
    edit_replace="🔄",
    edit_format="✨",
    
    # 视图操作
    view_preview="👀",
    view_sidebar="📊",
    view_minimap="🗺️",
    view_fullscreen="🖥️",
    view_split_h="◫",
    view_split_v="◧",
    view_zoom_in="🔎",
    view_zoom_out="🔍",
    view_focus="🎯",
    view_reading="📖",
    
    # 工具
    tool_settings="⚙️",
    tool_theme="🎨",
    tool_help="💡",
    tool_info="ℹ️",
    tool_warning="⚠️",
    tool_error="❌",
    tool_success="✅",
    tool_loading="⏳",
    tool_refresh="🔄",
    
    # AI 功能
    ai_assistant="🤖",
    ai_write="✍️",
    ai_translate="🌍",
    ai_summary="📝",
    ai_polish="💎",
    ai_magic="🪄",
    
    # 协作
    collab_user="👤",
    collab_users="👥",
    collab_chat="💬",
    collab_share="🔗",
    collab_invite="💌",
    collab_online="🟢",
    collab_offline="🔴",
    collab_away="🟡",
    collab_meeting="📹",
    collab_lock="🔐",
    collab_unlock="🔓",
    
    # 状态
    status_saved="✅",
    status_modified="📝",
    status_syncing="🔄",
    status_synced="☁️",
    status_error="⚠️",
    status_connected="🌐",
    status_disconnected="📵",
    
    # 其他可爱图标
    misc_star="⭐",
    misc_heart="💖",
    misc_fire="🔥",
    misc_sparkle="✨",
    misc_rocket="🚀",
    misc_lightning="⚡",
    misc_clock="⏰",
    misc_calendar="📅",
    misc_pin="📍",
    misc_tag="🏷️",
    misc_bookmark="🔖",
    misc_bell="🔔",
    misc_gift="🎁",
    misc_trophy="🏆",
    misc_medal="🎖️",
    misc_crown="👑",
    misc_gem="💎",
    misc_rainbow="🌈",
    misc_sun="🌞",
    misc_moon="🌛",
    misc_cloud="☁️",
    misc_umbrella="🌂",
    misc_snowflake="❄️",
    misc_leaf="🍀",
    misc_flower="🌷",
    misc_tree="🌲",
    misc_cat="😺",
    misc_dog="🐕",
    misc_panda="🐼",
    misc_unicorn="🦄",
    misc_butterfly="🦋",
)


# 工具栏图标映射 - 使用更可爱的图标
TOOLBAR_ICONS = {
    "open": "📂",
    "save": "💾",
    "format": "✨",
    "search": "🔎",
    "preview": "👀",
    "export": "📤",
    "pdf": "📕",
    "ocr": "📷",
    "ai": "🤖",
    "batch": "📦",
    "chart": "📊",
    "mindmap": "🧠",
    "bibliography": "📚",
    "version": "🔄",
    "link": "🔗",
    "database": "📚",
    "collab": "👥",
    "insert": "➕",
    "sidebar": "☰",
    "font_minus": "A⁻",
    "font_plus": "A⁺",
    "theme_light": "🌞",
    "theme_dark": "🌛",
    "theme_editor": "🎨",
    "focus": "🎯",
    "reading": "📖",
    "minimap": "🗺️",
    "split": "◫",
    "fullscreen": "🖥️",
    "settings": "⚙️",
    "help": "💡",
    "undo": "↩️",
    "redo": "↪️",
    "copy": "📋",
    "paste": "📌",
    "cut": "✂️",
    "delete": "🗑️",
}


# 状态栏图标
STATUS_ICONS = {
    "ready": "✅",
    "saving": "💾",
    "saved": "✅",
    "modified": "📝",
    "error": "⚠️",
    "loading": "⏳",
    "connected": "🌐",
    "disconnected": "📵",
    "syncing": "🔄",
    "synced": "☁️",
    "meeting": "📹",
    "users": "👥",
}


# 表情符号选择器 - 常用表情
EMOJI_PICKER = {
    "笑脸": ["😀", "😃", "😄", "😁", "😆", "😅", "🤣", "😂", "🙂", "😊", "😇", "🥰", "😍", "🤩", "😘", "😗", "😚", "😙", "🥲", "😋"],
    "手势": ["👍", "👎", "👌", "✌️", "🤞", "🤟", "🤘", "🤙", "👈", "👉", "👆", "👇", "☝️", "✋", "🤚", "🖐️", "🖖", "👋", "🤝", "👏"],
    "动物": ["🐱", "🐶", "🐭", "🐹", "🐰", "🦊", "🐻", "🐼", "🐨", "🐯", "🦁", "🐮", "🐷", "🐸", "🐵", "🐔", "🐧", "🐦", "🦄", "🦋"],
    "自然": ["🌸", "🌷", "🌹", "🌺", "🌻", "🌼", "🌱", "🌲", "🌳", "🌴", "🌵", "🍀", "🍁", "🍂", "🍃", "🌈", "☀️", "🌙", "⭐", "✨"],
    "食物": ["🍎", "🍐", "🍊", "🍋", "🍌", "🍉", "🍇", "🍓", "🫐", "🍒", "🍑", "🥭", "🍍", "🥥", "🥝", "🍅", "🥑", "🍔", "🍕", "🍰"],
    "物品": ["💎", "💍", "👑", "🎀", "🎁", "🎈", "🎉", "🎊", "🏆", "🥇", "🎯", "🎮", "🎨", "🎭", "🎪", "📱", "💻", "⌚", "📷", "🔮"],
    "符号": ["❤️", "🧡", "💛", "💚", "💙", "💜", "🖤", "🤍", "💖", "💝", "💘", "💗", "💓", "💞", "💕", "❣️", "💔", "🔥", "💯", "✅"],
    "天气": ["☀️", "🌤️", "⛅", "🌥️", "☁️", "🌦️", "🌧️", "⛈️", "🌩️", "🌨️", "❄️", "🌬️", "💨", "🌪️", "🌫️", "🌈", "☔", "⚡", "🌙", "⭐"],
}


# 应用状态消息图标
MESSAGE_ICONS = {
    "success": "✅",
    "error": "❌",
    "warning": "⚠️",
    "info": "ℹ️",
    "loading": "⏳",
    "saved": "💾",
    "copied": "📋",
    "exported": "📤",
    "imported": "📥",
    "deleted": "🗑️",
    "created": "✨",
    "updated": "🔄",
    "connected": "🌐",
    "disconnected": "📵",
    "welcome": "👋",
    "goodbye": "👋",
    "tip": "💡",
    "hint": "💭",
    "question": "❓",
    "answer": "💬",
}


def get_icon(name: str, style: str = "cute") -> str:
    """获取图标
    
    Args:
        name: 图标名称
        style: 图标风格 ("cute" 或 "default")
    
    Returns:
        图标字符串
    """
    icons = CUTE_ICONS if style == "cute" else IconSet()
    return getattr(icons, name, "❓")


def get_toolbar_icon(name: str) -> str:
    """获取工具栏图标"""
    return TOOLBAR_ICONS.get(name, "❓")


def get_status_icon(name: str) -> str:
    """获取状态栏图标"""
    return STATUS_ICONS.get(name, "")


def get_message_icon(msg_type: str) -> str:
    """获取消息图标"""
    return MESSAGE_ICONS.get(msg_type, "")


def get_emoji_categories() -> Dict[str, list]:
    """获取表情分类"""
    return EMOJI_PICKER


# 全局图标实例
icons = CUTE_ICONS
