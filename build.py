import PyInstaller.__main__
import os
import shutil

def build():
    # 确定应用图标路径
    icon_path = "app.ico"
    if not os.path.exists(icon_path):
        print("Warning: app.ico not found, using default icon")
        icon_path = None

    # 构建参数
    args = [
        'gui.py',  # 入口文件
        '--name=MarkdownToWord',  # 输出文件名
        '--noconfirm',  # 不确认覆盖
        '--windowed',  # 窗口模式（无控制台）
        '--clean',  # 清理临时文件
        '--onefile',  # 单文件模式 (分发更方便，通过QQ/微信发送只需一个文件)
        # '--onedir', # 目录模式 (启动快，排错容易)
        
        # 隐藏导入 (根据项目依赖添加)
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=babel.numbers',
        '--hidden-import=win32timezone',
        '--hidden-import=docx',
        '--hidden-import=customtkinter',
        '--hidden-import=openai',
        
        # 数据文件
        '--add-data=snippets.json;.',
        
    ]

    # 添加图标
    if icon_path:
        args.append(f'--icon={icon_path}')
        args.append(f'--add-data={icon_path};.')

    # 收集 customtkinter 数据文件
    import customtkinter
    ctk_path = os.path.dirname(customtkinter.__file__)
    args.append(f'--add-data={ctk_path};customtkinter')
    
    # 添加 templates 目录 (如果存在)
    if os.path.exists('templates'):
        args.append('--add-data=templates;templates')
    
    # 添加 assets/icons 目录 (SVG 图标)
    if os.path.exists('assets'):
        args.append('--add-data=assets;assets')
    
    # 添加 ui/themes 目录 (主题文件)
    if os.path.exists('ui/themes'):
        args.append('--add-data=ui/themes;ui/themes')

    print("开始构建...")
    PyInstaller.__main__.run(args)
    print("构建完成！请查看 dist/MarkdownToWord 目录")

if __name__ == '__main__':
    build()
