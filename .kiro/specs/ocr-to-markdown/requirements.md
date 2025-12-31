# Requirements Document

## Introduction

本文档定义了 OCR 图片转 Markdown 功能的需求。该功能允许用户通过截图、拍照或导入图片，自动识别其中的文字、表格和公式，并转换为 Markdown 格式。

## Glossary

- **OCR**: Optical Character Recognition，光学字符识别技术
- **LaTeX**: 数学公式排版系统
- **Tesseract**: 开源 OCR 引擎
- **PaddleOCR**: 百度开源的 OCR 工具包，支持中英文识别
- **表格识别**: 从图片中识别表格结构并提取数据
- **公式识别**: 从图片中识别数学公式并转换为 LaTeX

## Requirements

### Requirement 1: 图片导入和截图

**User Story:** 作为用户，我希望能够通过多种方式导入图片进行 OCR 识别，以便快速将图片内容转为文字。

#### Acceptance Criteria

1. WHEN the user clicks the OCR button THEN the system SHALL display options for image import (file, clipboard, screenshot)
2. WHEN the user selects a local image file THEN the system SHALL load and display the image preview
3. WHEN the user pastes from clipboard THEN the system SHALL detect and import image data
4. WHEN the user initiates screenshot mode THEN the system SHALL allow region selection and capture
5. WHEN an image is imported THEN the system SHALL validate the image format (PNG, JPG, BMP, WebP)
6. IF the image file is corrupted or unsupported THEN the system SHALL display an error message

### Requirement 2: 文字识别 (OCR)

**User Story:** 作为用户，我希望系统能准确识别图片中的文字，包括中英文混合内容。

#### Acceptance Criteria

1. WHEN an image is loaded THEN the system SHALL perform OCR recognition automatically
2. WHEN recognizing text THEN the system SHALL support Chinese, English, and mixed content
3. WHEN recognition completes THEN the system SHALL display the recognized text with confidence scores
4. WHEN the text contains paragraphs THEN the system SHALL preserve paragraph structure
5. WHEN the text contains lists THEN the system SHALL detect and format as Markdown lists
6. IF recognition confidence is below 70% THEN the system SHALL highlight uncertain regions for user review

### Requirement 3: 表格识别

**User Story:** 作为用户，我希望系统能识别图片中的表格并自动生成 Markdown 表格格式。

#### Acceptance Criteria

1. WHEN an image contains a table THEN the system SHALL detect table boundaries
2. WHEN a table is detected THEN the system SHALL extract cell contents and structure
3. WHEN table extraction completes THEN the system SHALL generate valid Markdown table syntax
4. WHEN table cells contain merged cells THEN the system SHALL handle them appropriately
5. WHEN the table has headers THEN the system SHALL identify and format the header row
6. IF table structure is ambiguous THEN the system SHALL provide manual adjustment options

### Requirement 4: 公式识别

**User Story:** 作为用户，我希望系统能识别图片中的数学公式并转换为 LaTeX 格式。

#### Acceptance Criteria

1. WHEN an image contains mathematical formulas THEN the system SHALL detect formula regions
2. WHEN a formula is detected THEN the system SHALL convert it to LaTeX syntax
3. WHEN conversion completes THEN the system SHALL wrap formulas in appropriate Markdown syntax ($...$ or $$...$$)
4. WHEN the formula is inline THEN the system SHALL use single dollar signs
5. WHEN the formula is block-level THEN the system SHALL use double dollar signs
6. IF formula recognition fails THEN the system SHALL preserve the original image as fallback

### Requirement 5: 批量处理

**User Story:** 作为用户，我希望能够批量导入多张图片进行 OCR 识别，提高工作效率。

#### Acceptance Criteria

1. WHEN the user selects multiple images THEN the system SHALL queue them for batch processing
2. WHILE batch processing is running THEN the system SHALL display progress for each image
3. WHEN batch processing completes THEN the system SHALL combine results into a single document
4. WHEN processing multiple images THEN the system SHALL maintain image order in output
5. WHEN an image fails processing THEN the system SHALL continue with remaining images and report errors
6. IF the user cancels batch processing THEN the system SHALL stop and preserve completed results

### Requirement 6: 结果编辑和导出

**User Story:** 作为用户，我希望能够编辑 OCR 识别结果并将其插入到当前文档或导出。

#### Acceptance Criteria

1. WHEN OCR completes THEN the system SHALL display results in an editable preview
2. WHEN the user edits results THEN the system SHALL update the Markdown output in real-time
3. WHEN the user clicks insert THEN the system SHALL insert the Markdown at cursor position
4. WHEN the user clicks copy THEN the system SHALL copy the Markdown to clipboard
5. WHEN the user clicks export THEN the system SHALL save the Markdown to a file
6. WHEN inserting results THEN the system SHALL preserve the original image as a reference comment
