# -*- coding: utf-8 -*-

from .tooltips import TooltipManager
from .header_style import HeaderStyler
from .command_palette import CommandPalette
from .insert_templates import InsertTemplatesFeature
from .status_bar import StatusBarFeature
from .editor_context_menu import EditorContextMenuFeature
from .help_dialog import HelpDialogFeature
from .auto_save import AutoSaveFeature
from .file_ops import FileOpsFeature
from .theme_feature import ThemeFeature
from .preview_sync import PreviewSyncFeature
from .window_geometry import WindowGeometryFeature
from .pdf_export import PDFExportFeature
from .preview_zoom import PreviewZoomFeature
from .editor_zoom import EditorZoomFeature
from .tab_manager import TabManagerFeature, TabData
from .statistics_detail import StatisticsDetailFeature, DocumentStatistics

# 新功能模块
from .batch_convert import BatchConvertFeature
from .word_to_markdown import WordToMarkdownFeature, WordToMarkdownConverter
from .template_manager import TemplateManager
from .ai_assistant import AIAssistantFeature
from .diagram_support import DiagramFeature, MermaidRenderer, PlantUMLRenderer
from .table_editor import TableEditorFeature
from .advanced_statistics import AdvancedStatisticsFeature
from .advanced_search import AdvancedSearchFeature
from .document_security import DocumentSecurityFeature
from .quick_tools import QuickToolsFeature
from .plugin_system import PluginManager, Plugin
from .undo_redo import UndoRedoFeature, UndoRedoManager, EnhancedTextWidget

# Phase 1 新增功能
from .focus_reading_mode import FocusModeFeature, ReadingModeFeature
from .toc_generator import TOCGeneratorFeature
from .watermark import WatermarkFeature
from .theme_editor import ThemeEditorFeature
from .template_selector import TemplateSelectorFeature
from .header_footer_editor import HeaderFooterFeature

# Phase 3 新增功能
from .document_stats import DocumentStatsFeature
from .global_search_replace import GlobalSearchReplaceFeature
from .link_checker import LinkCheckerFeature
from .snippet_library import SnippetLibraryFeature
from .batch_export import BatchExportFeature
from .chart_editor import ChartEditorFeature
from .mindmap import MindmapFeature
from .bibliography import BibliographyFeature
from .version_control import VersionControlFeature
from .autocomplete import AutocompleteFeature
