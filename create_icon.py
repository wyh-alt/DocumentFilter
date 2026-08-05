#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
生成文件处理工具的图标文件
"""

import os
import re
import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QPainter, QPixmap, QColor, QIcon, QFont
from PyQt6.QtCore import Qt, QPoint, QSize

# 全局持有 QApplication 实例，避免局部引用被 GC 销毁导致后续绘制崩溃
_app = None


def _ensure_app():
    global _app
    if _app is None:
        _app = QApplication(sys.argv)
    return _app

def create_app_icon():
    """创建应用图标"""
    # 创建一个64x64的图标
    icon_size = 64
    pixmap = QPixmap(icon_size, icon_size)
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

    # 绘制文件夹图标背景
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor(65, 105, 225))  # 蓝色
    painter.drawRoundedRect(4, 12, 56, 44, 5, 5)

    # 绘制文件夹顶部
    painter.setBrush(QColor(100, 149, 237))  # 浅蓝色
    painter.drawRoundedRect(4, 4, 36, 16, 3, 3)

    # 绘制筛选/过滤图标
    painter.setBrush(QColor(255, 255, 255))  # 白色

    # 绘制漏斗形状
    points = [
        (24, 20),  # 顶部中心
        (40, 20),  # 顶部右侧
        (34, 35),  # 中部右侧
        (34, 50),  # 底部右侧
        (30, 50),  # 底部左侧
        (30, 35),  # 中部左侧
        (18, 20)   # 顶部左侧
    ]

    # 将点列表转换为Qt可用的多边形
    polygon = [QPoint(x, y) for x, y in points]

    # 绘制漏斗
    painter.drawPolygon(polygon)

    # 绘制文件图标
    painter.setBrush(QColor(255, 255, 255))  # 白色
    painter.drawRect(42, 28, 12, 16)  # 文件主体
    painter.drawRect(44, 32, 8, 2)    # 文件行1
    painter.drawRect(44, 36, 8, 2)    # 文件行2
    painter.drawRect(44, 40, 8, 2)    # 文件行3

    painter.end()
    return pixmap

def save_splash():
    """生成 PyInstaller 启动页图片。

    该图片由 PyInstaller 的 --splash 参数嵌入 exe，在双击 exe 后立即
    显示（Python 解释器启动之前），覆盖 onefile 解压与依赖加载阶段，
    直到程序调用 pyi_splash.close() 才关闭。
    """
    _ensure_app()

    width, height = 640, 400
    pixmap = QPixmap(width, height)
    pixmap.fill(QColor(24, 26, 38))  # 深色背景

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    painter.setPen(Qt.PenStyle.NoPen)
    # 顶部品牌色条
    painter.setBrush(QColor(65, 105, 225))
    painter.drawRect(0, 0, width, 8)
    # 中央应用图标（文件夹 + 漏斗，与主窗口图标风格一致）
    icon_size = 144
    icon_x, icon_y = (width - icon_size) // 2, 46
    painter.setBrush(QColor(65, 105, 225))
    painter.drawRoundedRect(icon_x + 4, icon_y + 20, icon_size - 8, icon_size - 22, 10, 10)
    painter.setBrush(QColor(100, 149, 237))
    painter.drawRoundedRect(icon_x + 4, icon_y, int((icon_size - 8) * 0.62), 36, 6, 6)
    painter.setBrush(QColor(255, 255, 255))
    funnel_points = [
        (icon_x + 28, icon_y + 48), (icon_x + 60, icon_y + 48),
        (icon_x + 68, icon_y + 80), (icon_x + 66, icon_y + 104),
        (icon_x + 58, icon_y + 104), (icon_x + 56, icon_y + 80),
        (icon_x + 20, icon_y + 48),
    ]
    painter.drawPolygon([QPoint(x, y) for x, y in funnel_points])
    # 标题
    title_font = QFont("Microsoft YaHei", 34)
    title_font.setWeight(QFont.Weight.Bold)
    painter.setFont(title_font)
    painter.setPen(QColor(255, 255, 255))
    painter.drawText(0, 210, width, 56, Qt.AlignmentFlag.AlignCenter, "文件处理工具")
    # 版本号（从主程序提取 APP_VERSION，保持单一来源）
    version = "1.0"
    try:
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "file_selector.py"),
                  encoding="utf-8") as f:
            m = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', f.read())
            if m:
                version = m.group(1)
    except Exception:
        pass
    version_font = QFont("Microsoft YaHei", 15)
    painter.setFont(version_font)
    painter.setPen(QColor(160, 170, 200))
    painter.drawText(0, 272, width, 36, Qt.AlignmentFlag.AlignCenter, f"v{version}")
    # 底部静态提示：exe 解压阶段（Python 未运行）显示此文字；
    # 依赖加载阶段由 pyi_splash.update_text() 的动态文字覆盖此区域。
    tip_font = QFont("Microsoft YaHei", 11)
    painter.setFont(tip_font)
    painter.setPen(QColor(120, 130, 160))
    painter.drawText(0, 340, width, 30, Qt.AlignmentFlag.AlignCenter, "正在启动，请稍候...")
    painter.end()

    splash_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "splash.png")
    pixmap.save(splash_path, "PNG")
    print(f"启动页已生成: {splash_path}")
    return splash_path


def save_icon():
    """保存图标为文件"""
    _ensure_app()

    pixmap = create_app_icon()
    
    # 保存PNG格式的图标
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "file_selector_icon.png")
    pixmap.save(icon_path, "PNG")
    print(f"图标已保存为PNG格式: {icon_path}")
    
    # 尝试使用PIL库保存为.ico格式
    try:
        from PIL import Image
        ico_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "file_selector_icon.ico")
        # 将QPixmap转换为PIL Image
        temp_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp_icon.png")
        pixmap.save(temp_path, "PNG")
        img = Image.open(temp_path)
        
        # 保存为ICO格式
        sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128)]
        img.save(ico_path, format="ICO", sizes=sizes)
        print(f"图标已保存为ICO格式: {ico_path}")
        
        # 删除临时文件
        os.remove(temp_path)
        
        return ico_path
    except ImportError:
        print("PIL库未安装，无法保存为.ico格式")
        print("您可以使用在线工具将PNG转换为ICO格式: https://convertio.co/png-ico/")
    except Exception as e:
        print(f"保存为.ico格式时出错: {e}")
    
    return icon_path

if __name__ == "__main__":
    try:
        icon_path = save_icon()
        print("\n您可以使用此图标路径更新打包脚本:")
        print(f'--icon="{icon_path}"')
    except Exception as e:
        print(f"生成图标时出错: {e}")
    
    input("\n按Enter键退出...") 