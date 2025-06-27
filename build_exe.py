#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
文件筛选工具打包脚本
此脚本用于将文件筛选工具打包为可执行文件(exe)
"""

import os
import sys
import subprocess
import shutil

def check_dependencies():
    """检查并安装必要的依赖项"""
    dependencies = ["pyinstaller", "pillow"]
    
    for dep in dependencies:
        try:
            __import__(dep.replace("pillow", "PIL"))
            print(f"{dep} 已安装")
        except ImportError:
            print(f"{dep} 未安装，正在安装...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", dep])
                print(f"{dep} 安装成功")
            except Exception as e:
                print(f"安装 {dep} 失败: {e}")
                if dep == "pyinstaller":
                    return False
    return True

def create_icon():
    """创建应用图标"""
    print("正在生成应用图标...")
    try:
        # 使用create_icon.py脚本生成图标
        if os.path.exists("create_icon.py"):
            # 导入图标创建模块
            sys.path.append(os.getcwd())
            import create_icon
            icon_path = create_icon.save_icon()
            print(f"图标已生成: {icon_path}")
            return icon_path
        else:
            print("警告: 未找到create_icon.py脚本，将使用默认图标")
            return None
    except Exception as e:
        print(f"生成图标时出错: {e}")
        return None

def build_exe(icon_path=None):
    """构建可执行文件"""
    print("开始打包应用程序...")
    
    # 打包命令
    cmd = [
        sys.executable, 
        "-m", "PyInstaller",
        "--name=文件筛选工具",
        "--windowed",  # 使用GUI模式，不显示控制台
        "--onefile",   # 打包为单个exe文件
        "--add-data=requirements.txt;.",  # 添加额外文件
    ]
    
    # 如果有图标，添加图标参数
    if icon_path and os.path.exists(icon_path):
        cmd.append(f"--icon={icon_path}")
    else:
        print("警告: 未指定图标或图标文件不存在，将使用默认图标")
    
    # 添加主程序文件
    cmd.append("file_selector.py")
    
    try:
        subprocess.check_call(cmd)
        print("打包完成!")
        
        # 显示打包后的文件位置
        dist_path = os.path.abspath(os.path.join(os.getcwd(), "dist", "文件筛选工具.exe"))
        if os.path.exists(dist_path):
            print(f"可执行文件已生成: {dist_path}")
        else:
            print("警告: 未找到生成的可执行文件")
    except Exception as e:
        print(f"打包过程中出错: {e}")

def clean_build_files():
    """清理打包过程中生成的临时文件"""
    print("清理临时文件...")
    
    # 要删除的目录
    dirs_to_remove = ["build", "__pycache__", "文件筛选工具.spec"]
    
    for item in dirs_to_remove:
        try:
            if os.path.isdir(item):
                shutil.rmtree(item)
            elif os.path.exists(item):
                os.remove(item)
        except Exception as e:
            print(f"清理 {item} 时出错: {e}")

def main():
    """主函数"""
    print("=" * 50)
    print("文件筛选工具打包脚本")
    print("=" * 50)
    
    # 检查依赖项
    if not check_dependencies():
        print("错误: 无法安装必要的依赖项，打包过程终止")
        return
    
    # 创建图标
    icon_path = create_icon()
    
    # 构建exe
    build_exe(icon_path)
    
    # 清理临时文件
    clean_build_files()
    
    print("=" * 50)
    print("打包过程完成")
    print("=" * 50)

if __name__ == "__main__":
    main()
    input("按Enter键退出...") 