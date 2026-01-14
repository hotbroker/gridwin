import PyInstaller.__main__
import os
import shutil

def build_exe():
    print("开始打包 GridWin...")
    
    # 清理旧的构建文件
    for folder in ['build', 'dist']:
        if os.path.exists(folder):
            shutil.rmtree(folder)
            
    # 打包参数
    params = [
        'main.py',              # 主入口
        '--name=GridWin',       # 生成的 exe 名称
        '--noconsole',          # 无控制台窗口
        '--onefile',            # 打包成单个文件
        '--add-data=image.jpg;.', # 将图标打包进 exe (Windows 使用分号分隔)
        '--icon=image.jpg',     # 设置 exe 文件本身的图标
        '--clean',              # 清理缓存
    ]
    
    PyInstaller.__main__.run(params)
    print("\n打包完成！EXE 文件位于 dist 目录。")

if __name__ == "__main__":
    build_exe()
