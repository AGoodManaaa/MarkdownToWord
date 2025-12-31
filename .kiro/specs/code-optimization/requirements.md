# Requirements Document

## Introduction

本文档定义了 MarkdowntoWord 项目的代码优化需求，主要针对两个方面：
1. **代码结构优化** - 重构 `gui.py` 文件，拆分 `App` 类职责，消除重复代码
2. **性能优化** - 改进实时预览性能和图片异步下载

## Glossary

- **App**: 主应用窗口类，位于 `gui.py`，当前承担了过多职责
- **Feature**: 功能模块，位于 `ui/features/` 目录下的独立功能类
- **防抖 (Debounce)**: 延迟执行机制，避免频繁触发操作
- **增量渲染**: 只渲染变化的部分，而非全量重新渲染
- **线程池**: 管理多个工作线程的机制，用于异步执行任务

## Requirements

### Requirement 1: 消除重复初始化

**User Story:** 作为开发者，我希望代码中没有重复的初始化语句，以避免资源浪费和潜在的 bug。

#### Acceptance Criteria

1. WHEN the App class initializes THEN the system SHALL create each Feature instance exactly once
2. WHEN reviewing gui.py THEN the system SHALL contain no duplicate feature initialization statements
3. WHEN the application starts THEN the system SHALL log a warning if any feature is initialized more than once

### Requirement 2: 拆分 App 类的 UI 布局职责

**User Story:** 作为开发者，我希望 UI 布局代码从 App 类中分离出来，以提高代码可维护性。

#### Acceptance Criteria

1. WHEN the App class is loaded THEN the system SHALL delegate header creation to a dedicated HeaderBuilder class
2. WHEN the App class is loaded THEN the system SHALL delegate main content creation to a dedicated MainContentBuilder class
3. WHEN the App class is loaded THEN the system SHALL delegate status bar creation to a dedicated module
4. WHEN UI builders are used THEN the system SHALL maintain the same visual appearance as before refactoring

### Requirement 3: 拆分 App 类的快捷键绑定职责

**User Story:** 作为开发者，我希望快捷键绑定逻辑集中管理，便于维护和扩展。

#### Acceptance Criteria

1. WHEN the App initializes THEN the system SHALL delegate keyboard shortcut binding to a KeyboardShortcutsManager class
2. WHEN a new shortcut is added THEN the system SHALL require modification only in the KeyboardShortcutsManager
3. WHEN shortcuts are bound THEN the system SHALL support easy customization through configuration

### Requirement 4: 优化实时预览性能

**User Story:** 作为用户，我希望在编辑大文档时预览仍然流畅，不会出现明显卡顿。

#### Acceptance Criteria

1. WHEN the user edits a document larger than 10KB THEN the system SHALL render preview within 200ms
2. WHEN the user types rapidly THEN the system SHALL use debounce mechanism with configurable delay
3. WHEN only a small portion of document changes THEN the system SHALL perform incremental update instead of full re-render
4. WHEN the preview is updating THEN the system SHALL not block the editor input

### Requirement 5: 异步图片下载

**User Story:** 作为用户，我希望在文档包含网络图片时，下载过程不会阻塞界面操作。

#### Acceptance Criteria

1. WHEN a network image is encountered THEN the system SHALL download it asynchronously using a thread pool
2. WHILE an image is downloading THEN the system SHALL display a placeholder in the preview
3. WHEN an image download completes THEN the system SHALL update the preview without blocking the UI
4. WHEN multiple images need downloading THEN the system SHALL limit concurrent downloads to prevent resource exhaustion
5. IF an image download fails THEN the system SHALL display an error placeholder and allow retry

### Requirement 6: 图片缓存机制

**User Story:** 作为用户，我希望已下载的网络图片被缓存，避免重复下载。

#### Acceptance Criteria

1. WHEN a network image is downloaded THEN the system SHALL cache it locally
2. WHEN the same image URL is requested again THEN the system SHALL use the cached version
3. WHEN the cache exceeds a configurable size limit THEN the system SHALL remove oldest entries
4. WHEN the application restarts THEN the system SHALL preserve the image cache

