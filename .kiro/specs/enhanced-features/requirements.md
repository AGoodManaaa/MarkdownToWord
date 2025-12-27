# Requirements Document

## Introduction

本文档定义了 MarkdowntoWord 应用程序的四项增强功能需求：PDF 导出、预览区缩放、多标签页编辑和字数统计详情。这些功能旨在提升用户体验，使应用程序更加专业和实用。

## Glossary

- **MarkdowntoWord**: 本项目的 Markdown 转 Word 转换器应用程序
- **PDF**: Portable Document Format，便携式文档格式
- **预览区**: 应用程序右侧显示 Markdown 渲染效果的区域
- **标签页**: 允许用户在同一窗口中切换多个文档的 UI 组件
- **字数统计**: 对文档内容进行字符、单词、段落等维度的计数分析

## Requirements

### Requirement 1: PDF 导出功能

**User Story:** As a user, I want to export my Markdown content directly to PDF format, so that I can share documents without requiring Word software.

#### Acceptance Criteria

1. WHEN a user clicks the PDF export button THEN the system SHALL display a file save dialog with PDF as the default format
2. WHEN a user confirms the PDF export THEN the system SHALL convert the Markdown content to PDF and save it to the specified path
3. WHEN the PDF export completes successfully THEN the system SHALL display a success notification with the file path
4. IF the PDF export fails THEN the system SHALL display an error message describing the failure reason
5. WHEN exporting to PDF THEN the system SHALL preserve all formatting including headings, tables, code blocks, and images
6. WHEN exporting to PDF THEN the system SHALL apply the same page size settings (A4/Letter) as Word export

### Requirement 2: 预览区缩放功能

**User Story:** As a user, I want to adjust the preview area zoom level, so that I can view the rendered content at a comfortable size.

#### Acceptance Criteria

1. WHEN the preview panel is visible THEN the system SHALL display zoom control buttons (zoom in, zoom out, reset)
2. WHEN a user clicks the zoom in button THEN the system SHALL increase the preview scale by 10% up to a maximum of 150%
3. WHEN a user clicks the zoom out button THEN the system SHALL decrease the preview scale by 10% down to a minimum of 50%
4. WHEN a user clicks the reset button THEN the system SHALL restore the preview scale to 100%
5. WHEN the zoom level changes THEN the system SHALL immediately update the preview display
6. WHEN the application restarts THEN the system SHALL restore the last used zoom level from saved configuration

### Requirement 3: 多标签页编辑功能

**User Story:** As a user, I want to edit multiple Markdown files in tabs, so that I can work on several documents simultaneously without opening multiple application windows.

#### Acceptance Criteria

1. WHEN the application starts THEN the system SHALL display a tab bar above the editor area
2. WHEN a user opens a new file THEN the system SHALL create a new tab for that file
3. WHEN a user clicks on a tab THEN the system SHALL switch to display that tab's content in the editor and preview
4. WHEN a user clicks the close button on a tab THEN the system SHALL close that tab after checking for unsaved changes
5. WHEN a tab has unsaved changes THEN the system SHALL display a visual indicator (asterisk) on the tab title
6. WHEN a user attempts to close a tab with unsaved changes THEN the system SHALL prompt the user to save, discard, or cancel
7. WHEN all tabs are closed THEN the system SHALL create a new empty tab automatically
8. WHEN a user drags a tab THEN the system SHALL allow reordering tabs by drag and drop
9. WHEN a user right-clicks on a tab THEN the system SHALL display a context menu with options (Close, Close Others, Close All)

### Requirement 4: 字数统计详情功能

**User Story:** As a user, I want to see detailed statistics about my document, so that I can track my writing progress and estimate reading time.

#### Acceptance Criteria

1. WHEN the editor contains text THEN the system SHALL display character count in the status bar
2. WHEN the editor contains text THEN the system SHALL display word count in the status bar
3. WHEN the editor contains text THEN the system SHALL display paragraph count in the status bar
4. WHEN the editor contains text THEN the system SHALL display estimated reading time in the status bar
5. WHEN a user clicks on the statistics area THEN the system SHALL display a detailed statistics popup
6. WHEN the detailed statistics popup is shown THEN the system SHALL display: total characters, characters without spaces, Chinese characters, English words, paragraphs, lines, and estimated reading time
7. WHEN the editor content changes THEN the system SHALL update all statistics within 500 milliseconds
8. WHEN calculating reading time THEN the system SHALL use 300 Chinese characters per minute or 200 English words per minute as the reading speed
