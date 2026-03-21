# 产品重构蓝图

## 目标

把当前项目收敛为一个专业的 `Markdown / Word` 转换器。

优先级：

1. 高保真导出
2. 流畅编辑与预览
3. 低心智负担
4. 大文档稳定性

## 默认主链路

- 打开或粘贴 Markdown
- 实时预览
- 选择模板
- 导出 Word / PDF / HTML
- 查看导出历史

## 默认保留在主入口的能力

- 打开文件
- 保存文件
- Markdown 规范化
- 搜索
- 预览切换
- 分屏
- 导出 Word
- 导出 PDF
- 导出 HTML
- 批量导出

## 默认降级为按需加载的能力

- AI 助手
- OCR
- 图表
- 导图
- 文献
- 版本历史
- 链接检查
- 文档库
- 协作
- 快捷键面板
- 文件夹视图

## 代码结构目标

```text
core/
  parser/
  converter/
  diagnostics/

ui/
  shell/
  editor/
  preview/
  export/

services/
  files/
  templates/
  cache/
```

## 当前已落地的第一轮收敛

- 修复 README 冲突残留
- 修复命令执行编码导致的测试失败
- 修复 AI 配置重复装饰器
- 引入 `product_mode` 与 `show_advanced_toolbar`
- 把一批非核心功能改为按需初始化
- 默认工具栏聚焦转换主链路

## 下一轮建议

1. 拆分 `gui.py`
2. 拆分 `converter.py`
3. 建立 `tests/regression`
4. 为核心样例建立导出基线
