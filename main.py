# -*- coding: utf-8 -*-
"""
Markdown转Word转换器 - 命令行入口
支持LaTeX风格、表格、公式、图片等完整转换
"""

import argparse
import os
import sys
from converter import MarkdownToWordConverter


def main():
    parser = argparse.ArgumentParser(
        description='Markdown转Word转换器 (LaTeX风格)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python main.py input.md output.docx
  python main.py input.md                    # 输出为 input.docx
  python main.py -t "# 标题\\n正文内容" output.docx
        """
    )
    
    parser.add_argument('input', nargs='?', help='输入Markdown文件路径')
    parser.add_argument('output', nargs='?', help='输出Word文件路径（可选）')
    parser.add_argument('-t', '--text', help='直接输入Markdown文本（替代文件输入）')
    parser.add_argument('-b', '--base-dir', help='图片基准目录（用于相对路径解析）')
    
    args = parser.parse_args()
    
    # 验证参数
    if not args.input and not args.text:
        parser.print_help()
        print("\n错误: 请提供输入文件或使用 -t 参数提供文本")
        sys.exit(1)
    
    # 确定输出路径
    if args.output:
        output_path = args.output
    elif args.input:
        # 默认输出同名.docx文件
        base_name = os.path.splitext(args.input)[0]
        output_path = base_name + '.docx'
    else:
        output_path = 'output.docx'
    
    # 确定基准目录
    if args.base_dir:
        base_dir = args.base_dir
    elif args.input:
        base_dir = os.path.dirname(os.path.abspath(args.input))
    else:
        base_dir = os.getcwd()
    
    # 创建转换器
    converter = MarkdownToWordConverter(base_dir=base_dir)
    
    # 执行转换
    if args.text:
        # 从文本转换
        text = args.text.replace('\\n', '\n')  # 处理转义换行
        converter.convert_text(text)
        converter.save(output_path)
        print(f"转换完成: {output_path}")
    else:
        # 从文件转换
        if not os.path.exists(args.input):
            print(f"错误: 文件不存在 - {args.input}")
            sys.exit(1)
        
        success = converter.convert_file(args.input, output_path)
        if not success:
            sys.exit(1)


def convert_markdown_to_word(markdown_text: str, output_path: str, base_dir: str = None) -> bool:
    """
    便捷函数：直接调用转换
    
    Args:
        markdown_text: Markdown文本内容
        output_path: 输出Word文件路径
        base_dir: 图片基准目录
        
    Returns:
        是否成功
    """
    try:
        converter = MarkdownToWordConverter(base_dir=base_dir)
        converter.convert_text(markdown_text)
        converter.save(output_path)
        return True
    except Exception as e:
        print(f"转换失败: {e}")
        return False


def convert_file(input_path: str, output_path: str = None) -> bool:
    """
    便捷函数：转换文件
    
    Args:
        input_path: Markdown文件路径
        output_path: 输出Word文件路径（可选，默认同名.docx）
        
    Returns:
        是否成功
    """
    if output_path is None:
        base_name = os.path.splitext(input_path)[0]
        output_path = base_name + '.docx'
    
    converter = MarkdownToWordConverter()
    return converter.convert_file(input_path, output_path)


# 示例用法
if __name__ == '__main__':
    main()
