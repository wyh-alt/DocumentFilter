#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
生成文件筛选工具的图标文件
"""

import os
import sys
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QPainter, QPixmap, QColor, QIcon
from PyQt6.QtCore import Qt, QPoint, QSize

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

def save_icon():
    """保存图标为文件"""
    app = QApplication(sys.argv)
    
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