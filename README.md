# Markdown to Word Converter

一个聚焦于 `Markdown -> Word` 高保真导出的桌面转换器。

## 当前产品方向

项目正在从“大而全编辑器”收敛为“专业转换器”：

- 默认主界面聚焦编辑、预览、模板和导出
- 高级能力改为按需加载，避免首屏干扰
- 优先保证导出质量、性能和稳定性

## 核心能力

- Markdown 实时编辑与预览
- 导出为 Word
- 导出为 PDF / HTML
- 批量导出
- 模板、页眉页脚、水印
- 图片、表格、代码块、数学公式支持

## 启动

```bash
pip install -r requirements.txt
python gui.py
```

## 项目结构

```text
gui.py              # 主应用窗口
converter.py        # Markdown -> Word 核心转换逻辑
parser.py           # Markdown 解析
handlers.py         # 图片 / 表格 / 数学公式等处理器
ui/                 # 桌面界面
tests/              # 测试
auto_dev_agent/     # 独立实验子系统
```

## 当前重构重点

1. 聚焦转换主链路
2. 减少首屏功能噪音
3. 引入核心 / 扩展能力分层加载
4. 补齐转换回归测试

## 说明

- 当前仓库仍包含部分实验功能模块
- 主版本默认不强调 AI、协作、数据库、OCR 等高级入口
- 后续会继续拆分 `gui.py` 和 `converter.py`
