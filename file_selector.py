#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import re
import shutil
import concurrent.futures
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                           QListWidget, QListWidgetItem, QPushButton, QFileDialog, QLabel,
                           QCheckBox, QMessageBox, QProgressBar, QComboBox, QGroupBox,
                           QLineEdit, QSplitter, QFrame, QTextEdit, QTableWidget, QTableWidgetItem, QHeaderView,
                           QStyledItemDelegate)
from PyQt6.QtCore import Qt, QMimeData, QUrl, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QDrag, QDropEvent, QIcon, QFont, QColor, QTextDocument, QAbstractTextDocumentLayout, QPainter, QPixmap

from file_matcher import FileMatcher

# 创建应用图标
def create_app_icon():
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
    from PyQt6.QtCore import QPoint
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
    return QIcon(pixmap)

# 自定义代理类用于高亮显示单元格中的关键字
class HTMLDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.highlights = {}  # 存储每个单元格的高亮关键字
        
    def add_highlight(self, row, column, key):
        """添加单元格高亮信息"""
        self.highlights[(row, column)] = key
        
    def paint(self, painter, option, index):
        row, column = index.row(), index.column()
        if (row, column) in self.highlights:
            key = self.highlights[(row, column)]
            text = index.data(Qt.ItemDataRole.DisplayRole)
            if text and key in text:
                # 创建HTML文档
                doc = QTextDocument()
                highlighted_text = text.replace(key, f'<span style="color:green;">{key}</span>')
                doc.setHtml(highlighted_text)
                
                # 保存画笔状态
                painter.save()
                
                # 设置绘制区域
                painter.translate(option.rect.topLeft())
                clip_rect = option.rect.translated(-option.rect.topLeft())
                painter.setClipRect(clip_rect)
                
                # 绘制文档
                doc.documentLayout().draw(painter, QAbstractTextDocumentLayout.PaintContext())
                
                # 恢复画笔状态
                painter.restore()
                return
        super().paint(painter, option, index)

# 自定义复选框容器，用于在表格中显示复选框
class CheckBoxWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(3, 1, 3, 1)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.checkbox = QCheckBox()
        layout.addWidget(self.checkbox)
        
    def isChecked(self):
        return self.checkbox.isChecked()
    
    def setChecked(self, checked):
        self.checkbox.setChecked(checked)

class MatchWorker(QThread):
    result_ready = pyqtSignal(list, int)
    error = pyqtSignal(str)
    progress = pyqtSignal(int, int)
    file_counts_ready = pyqtSignal(int, int)  # 新增信号：源目录和目标目录文件数量
    def __init__(self, match_method_index, source_dir, target_dir, match_basis, format_match, text_input, text_match_basis=0, use_multithreading=True):
        super().__init__()
        self.match_method_index = match_method_index
        self.source_dir = source_dir
        self.target_dir = target_dir
        self.match_basis = match_basis
        self.format_match = format_match
        self.text_input = text_input
        self.text_match_basis = text_match_basis
        self.use_multithreading = use_multithreading
        # 设置线程数为CPU核心数
        self.thread_count = max(4, os.cpu_count() or 4)
        self._is_cancelled = False
    
    def cancel(self):
        """取消匹配操作"""
        self._is_cancelled = True
    
    def _load_files(self, directory):
        """延迟加载文件列表"""
        if not directory or not os.path.isdir(directory):
            return []
        
        # 使用生成器表达式，避免一次性加载所有文件到内存
        return [os.path.join(directory, f) for f in os.listdir(directory) 
                if os.path.isfile(os.path.join(directory, f))]
    
    def _split_list(self, items, num_chunks):
        """将列表分割成多个块，用于并行处理"""
        if not items:
            return []
        
        chunk_size = max(1, len(items) // num_chunks)
        return [items[i:i + chunk_size] for i in range(0, len(items), chunk_size)]
    
    def _process_chunk_by_name(self, source_files, target_dict, target_dict_no_ext):
        """处理一批源文件（用于多线程处理）"""
        result = []
        for s in source_files:
            s_name = os.path.basename(s)
            
            if self.format_match:  # 匹配包括扩展名
                if s_name in target_dict:
                    t = target_dict[s_name]
                    t_name = os.path.basename(t)
                    result.append((s, t, s_name, t_name, s_name))
            else:  # 不匹配扩展名
                s_name_no_ext = os.path.splitext(s_name)[0]
                if s_name_no_ext in target_dict_no_ext:
                    for t in target_dict_no_ext[s_name_no_ext]:
                        t_name = os.path.basename(t)
                        result.append((s, t, s_name, t_name, s_name_no_ext))
        return result
    
    def _process_chunk_by_id(self, target_files, s_dict, id_pattern):
        """处理一批目标文件（用于ID匹配的多线程处理）"""
        result = []
        for t in target_files:
            t_name = os.path.basename(t)
            match = id_pattern.match(t_name)
            if match:
                tid = match.group(1)
                if tid in s_dict:
                    for s in s_dict[tid]:
                        s_name = os.path.basename(s)
                        result.append((s, t, s_name, t_name, tid))
        return result
    
    def _process_chunk_by_text(self, target_files, keywords):
        """处理一批目标文件（用于文本匹配的多线程处理）"""
        result = []
        for t in target_files:
            t_name = os.path.basename(t)
            
            for kw in keywords:
                if kw in t_name:
                    result.append((None, t, '', t_name, kw))
                    break
        return result
    
    def run(self):
        try:
            # 首先统计文件数量并发送信号
            source_count = 0
            target_count = 0
            
            # 使用更高效的方式统计文件数量
            if self.source_dir and os.path.isdir(self.source_dir):
                source_count = sum(1 for f in os.listdir(self.source_dir) 
                                 if os.path.isfile(os.path.join(self.source_dir, f)))
            
            if self.target_dir and os.path.isdir(self.target_dir):
                target_count = sum(1 for f in os.listdir(self.target_dir) 
                                 if os.path.isfile(os.path.join(self.target_dir, f)))
            
            # 发送文件数量信号
            self.file_counts_ready.emit(source_count, target_count)
            
            # 检查是否已取消
            if self._is_cancelled:
                return
            
            matched_pairs = []
            if self.match_method_index == 0:
                if not self.source_dir or not self.target_dir:
                    self.error.emit("请先选择源目录和目标目录")
                    return
                
                # 优化1: 延迟加载文件列表
                self.progress.emit(0, 100)
                source_files = self._load_files(self.source_dir)
                if self._is_cancelled:
                    return
                
                self.progress.emit(10, 100)
                target_files = self._load_files(self.target_dir)
                if self._is_cancelled:
                    return
                
                self.progress.emit(20, 100)
                
                # 优化2: 预处理目标文件，创建查找索引
                target_dict = {}
                target_dict_no_ext = {}
                
                # 使用批处理方式构建索引
                batch_size = 1000  # 每批处理的文件数量
                for i in range(0, len(target_files), batch_size):
                    batch = target_files[i:i+batch_size]
                    for t in batch:
                        t_name = os.path.basename(t)
                        target_dict[t_name] = t
                        t_name_no_ext = os.path.splitext(t_name)[0]
                        if t_name_no_ext not in target_dict_no_ext:
                            target_dict_no_ext[t_name_no_ext] = []
                        target_dict_no_ext[t_name_no_ext].append(t)
                    
                    # 更新进度并检查是否取消
                    progress = 20 + int((i + len(batch)) / len(target_files) * 20)
                    self.progress.emit(progress, 100)
                    if self._is_cancelled:
                        return
                
                total = len(source_files)
                count = 0
                
                if self.match_basis == 0:  # 完整文件名
                    # 检查是否使用多线程处理
                    if self.use_multithreading and len(source_files) > 100:
                        # 分割源文件列表为多个块
                        chunks = self._split_list(source_files, self.thread_count)
                        
                        # 使用线程池并行处理
                        with concurrent.futures.ThreadPoolExecutor(max_workers=self.thread_count) as executor:
                            # 提交任务
                            future_to_chunk = {
                                executor.submit(self._process_chunk_by_name, chunk, target_dict, target_dict_no_ext): i 
                                for i, chunk in enumerate(chunks)
                            }
                            
                            # 处理结果
                            completed = 0
                            for future in concurrent.futures.as_completed(future_to_chunk):
                                chunk_result = future.result()
                                matched_pairs.extend(chunk_result)
                                completed += 1
                                # 更新进度
                                progress = int((completed / len(chunks)) * total)
                                self.progress.emit(progress, total)
                    else:
                        # 单线程处理
                        for s in source_files:
                            s_name = os.path.basename(s)
                            
                            if self.format_match:  # 匹配包括扩展名
                                if s_name in target_dict:
                                    t = target_dict[s_name]
                                    t_name = os.path.basename(t)
                                    matched_pairs.append((s, t, s_name, t_name, s_name))
                            else:  # 不匹配扩展名
                                s_name_no_ext = os.path.splitext(s_name)[0]
                                if s_name_no_ext in target_dict_no_ext:
                                    for t in target_dict_no_ext[s_name_no_ext]:
                                        t_name = os.path.basename(t)
                                        matched_pairs.append((s, t, s_name, t_name, s_name_no_ext))
                            
                            count += 1
                            if count % 20 == 0 or count == total:
                                self.progress.emit(count, total)
                
                else:  # ID前缀
                    # 优化3: 提前编译正则表达式
                    id_pattern = re.compile(r'([a-zA-Z0-9]+)')
                    
                    # 创建ID索引
                    s_dict = {}
                    for s in source_files:
                        s_name = os.path.basename(s)
                        match = id_pattern.match(s_name)
                        if match:
                            sid = match.group(1)
                            if sid not in s_dict:
                                s_dict[sid] = []
                            s_dict[sid].append(s)
                    
                    # 使用ID索引匹配目标文件
                    total = len(target_files)
                    
                    # 检查是否使用多线程处理
                    if self.use_multithreading and len(target_files) > 100:
                        # 分割目标文件列表为多个块
                        chunks = self._split_list(target_files, self.thread_count)
                        
                        # 使用线程池并行处理
                        with concurrent.futures.ThreadPoolExecutor(max_workers=self.thread_count) as executor:
                            # 提交任务
                            future_to_chunk = {
                                executor.submit(self._process_chunk_by_id, chunk, s_dict, id_pattern): i 
                                for i, chunk in enumerate(chunks)
                            }
                            
                            # 处理结果
                            completed = 0
                            for future in concurrent.futures.as_completed(future_to_chunk):
                                chunk_result = future.result()
                                matched_pairs.extend(chunk_result)
                                completed += 1
                                # 更新进度
                                progress = int((completed / len(chunks)) * total)
                                self.progress.emit(progress, total)
                    else:
                        # 单线程处理
                        count = 0
                        for t in target_files:
                            t_name = os.path.basename(t)
                            match = id_pattern.match(t_name)
                            if match:
                                tid = match.group(1)
                                if tid in s_dict:
                                    for s in s_dict[tid]:
                                        s_name = os.path.basename(s)
                                        matched_pairs.append((s, t, s_name, t_name, tid))
                            
                            count += 1
                            if count % 20 == 0 or count == total:
                                self.progress.emit(count, total)
            
            elif self.match_method_index == 1:
                if not self.target_dir:
                    self.error.emit("请先选择目标目录")
                    return
                
                filter_text = self.text_input.strip()
                if not filter_text:
                    self.error.emit("请输入检索文本")
                    return
                
                # 获取匹配依据选项
                text_match_basis = self.text_match_basis
                
                # 处理多行或分号分隔的关键字
                keywords = [kw.strip() for kw in filter_text.replace(';', '\n').split('\n') if kw.strip()]
                
                # 优化4: 创建关键字集合，提高查找速度
                keywords_set = set(keywords)
                
                # 获取所有目标文件
                target_files = [os.path.join(self.target_dir, f) for f in os.listdir(self.target_dir) if os.path.isfile(os.path.join(self.target_dir, f))]
                
                # 优化5: 根据匹配模式预处理文件名
                if text_match_basis == 0:  # 完全匹配
                    # 创建不含扩展名的文件名字典
                    name_dict = {}
                    for t in target_files:
                        t_name = os.path.basename(t)
                        t_name_no_ext = os.path.splitext(t_name)[0]
                        name_dict[t_name_no_ext] = (t, t_name)
                    
                    # 直接查找匹配项
                    for kw in keywords:
                        if kw in name_dict:
                            t, t_name = name_dict[kw]
                            matched_pairs.append((None, t, '', t_name, kw))
                
                else:  # 检索匹配
                    total = len(target_files)
                    
                    # 检查是否使用多线程处理
                    if self.use_multithreading and len(target_files) > 100:
                        # 分割目标文件列表为多个块
                        chunks = self._split_list(target_files, self.thread_count)
                        
                        # 使用线程池并行处理
                        with concurrent.futures.ThreadPoolExecutor(max_workers=self.thread_count) as executor:
                            # 提交任务
                            future_to_chunk = {
                                executor.submit(self._process_chunk_by_text, chunk, keywords): i 
                                for i, chunk in enumerate(chunks)
                            }
                            
                            # 处理结果
                            completed = 0
                            for future in concurrent.futures.as_completed(future_to_chunk):
                                chunk_result = future.result()
                                matched_pairs.extend(chunk_result)
                                completed += 1
                                # 更新进度
                                progress = int((completed / len(chunks)) * total)
                                self.progress.emit(progress, total)
                    else:
                        # 单线程处理
                        count = 0
                        for t in target_files:
                            t_name = os.path.basename(t)
                            
                            # 优化6: 一次检查所有关键字
                            for kw in keywords:
                                if kw in t_name:
                                    matched_pairs.append((None, t, '', t_name, kw))
                                    break
                            
                            count += 1
                            if count % 20 == 0 or count == total:
                                self.progress.emit(count, total)
            
            self.result_ready.emit(matched_pairs, len(matched_pairs))
        except Exception as e:
            self.error.emit(str(e))

class FileSelector(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.source_dir = ""
        self.target_dir = ""
        self.output_dir = ""
        self.matched_files = []
        self._search_anim_timer = None
        self._search_anim_step = 0
        self.processing = False
        
        # 设置应用图标
        self.setWindowIcon(create_app_icon())

    def init_ui(self):
        self.setWindowTitle("文件筛选工具")
        self.setGeometry(300, 300, 1000, 800)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 进度条和取消按钮的水平布局
        progress_layout = QHBoxLayout()
        
        # 添加进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        progress_layout.addWidget(self.progress_bar, 1)  # 1是拉伸因子，让进度条占据大部分空间
        
        # 添加取消按钮
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.cancel_operation)
        self.cancel_btn.setVisible(False)
        progress_layout.addWidget(self.cancel_btn)
        
        main_layout.addLayout(progress_layout)
        
        splitter = QSplitter(Qt.Orientation.Horizontal)
        # 左侧设置面板
        left_panel = QWidget()
        left_panel.setMinimumWidth(320)
        left_panel.setMaximumWidth(340)
        left_layout = QVBoxLayout(left_panel)
        # 目录选择区域
        dir_group = QGroupBox("目录设置")
        dir_layout = QVBoxLayout()
        # 源目录选择
        source_layout = QHBoxLayout()
        source_layout.addWidget(QLabel("源目录:"))
        self.source_label = QLabel("未选择")
        self.source_label.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.source_label.setFixedWidth(300)
        self.source_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextBrowserInteraction)
        self.source_label.setWordWrap(False)
        self.source_label.setToolTip("未选择")
        self.source_label.setAcceptDrops(True)
        self.source_label.dragEnterEvent = self.drag_enter_event
        self.source_label.dropEvent = lambda event: self.drop_event(event, "source")
        source_btn = QPushButton("浏览...")
        source_btn.clicked.connect(lambda: self.select_directory("source"))
        source_layout.addWidget(self.source_label)
        source_layout.addWidget(source_btn)
        # 目标目录选择
        target_layout = QHBoxLayout()
        target_layout.addWidget(QLabel("目标目录:"))
        self.target_label = QLabel("未选择")
        self.target_label.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.target_label.setFixedWidth(300)
        self.target_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextBrowserInteraction)
        self.target_label.setWordWrap(False)
        self.target_label.setToolTip("未选择")
        self.target_label.setAcceptDrops(True)
        self.target_label.dragEnterEvent = self.drag_enter_event
        self.target_label.dropEvent = lambda event: self.drop_event(event, "target")
        target_btn = QPushButton("浏览...")
        target_btn.clicked.connect(lambda: self.select_directory("target"))
        target_layout.addWidget(self.target_label)
        target_layout.addWidget(target_btn)
        # 输出目录选择
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出目录:"))
        self.output_label = QLabel("未选择")
        self.output_label.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.output_label.setFixedWidth(300)
        self.output_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextBrowserInteraction)
        self.output_label.setWordWrap(False)
        self.output_label.setToolTip("未选择")
        self.output_label.setAcceptDrops(True)
        self.output_label.dragEnterEvent = self.drag_enter_event
        self.output_label.dropEvent = lambda event: self.drop_event(event, "output")
        output_btn = QPushButton("浏览...")
        output_btn.clicked.connect(lambda: self.select_directory("output"))
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(output_btn)
        dir_layout.addLayout(source_layout)
        dir_layout.addLayout(target_layout)
        dir_layout.addLayout(output_layout)
        dir_group.setLayout(dir_layout)
        left_layout.addWidget(dir_group)
        # 匹配设置区域
        match_group = QGroupBox("匹配设置")
        match_layout = QVBoxLayout()
        match_layout.addWidget(QLabel("匹配方式:"))
        self.match_method = QComboBox()
        self.match_method.addItems([
            "根据源目录匹配提取",
            "根据文本匹配提取"
        ])
        self.match_method.currentIndexChanged.connect(self.on_match_method_changed)
        match_layout.addWidget(self.match_method)
        # 匹配选项容器
        self.match_options_widget = QWidget()
        self.match_options_layout = QVBoxLayout(self.match_options_widget)
        # 源目录匹配提取设置
        self.dir_match_layout = QVBoxLayout()
        # 匹配依据
        match_basis_row = QHBoxLayout()
        match_basis_row.addWidget(QLabel("匹配依据:"))
        self.match_basis_combo = QComboBox()
        self.match_basis_combo.addItems(["完整文件名", "ID前缀"])
        match_basis_row.addWidget(self.match_basis_combo)
        self.dir_match_layout.addLayout(match_basis_row)
        # 文件格式匹配
        format_match_row = QHBoxLayout()
        format_match_row.addWidget(QLabel("文件格式匹配:"))
        self.format_match_combo = QComboBox()
        self.format_match_combo.addItems(["是", "否"])
        format_match_row.addWidget(self.format_match_combo)
        self.dir_match_layout.addLayout(format_match_row)
        # 文本匹配提取设置
        self.text_match_layout = QVBoxLayout()
        label = QLabel("检索文本（支持多行或分号分隔）:")
        self.text_match_layout.addWidget(label)
        self.text_match_input = QTextEdit()
        self.text_match_input.setPlaceholderText("可输入多个文件名或关键字，支持换行或分号分隔")
        self.text_match_input.setFixedHeight(120)
        self.text_match_layout.addWidget(self.text_match_input)
        
        # 添加文本匹配依据选项
        text_match_basis_row = QHBoxLayout()
        text_match_basis_row.addWidget(QLabel("匹配依据:"))
        self.text_match_basis_combo = QComboBox()
        self.text_match_basis_combo.addItems(["完全匹配", "检索匹配"])
        text_match_basis_row.addWidget(self.text_match_basis_combo)
        self.text_match_layout.addLayout(text_match_basis_row)
        
        # 添加所有布局到选项容器
        self.match_options_layout.addLayout(self.dir_match_layout)
        self.match_options_layout.addLayout(self.text_match_layout)
        match_layout.addWidget(self.match_options_widget)
        # 匹配按钮
        self.match_btn = QPushButton("查找匹配文件")
        self.match_btn.clicked.connect(self.match_files)
        match_layout.addWidget(self.match_btn)
        match_group.setLayout(match_layout)
        left_layout.addWidget(match_group)
        # 匹配文件统计和操作区域（移到左侧）
        op_group = QGroupBox("操作")
        op_layout = QVBoxLayout()
        # 新增源/目标目录文件数量统计
        self.source_file_count_label = QLabel("源目录文件数: 0")
        self.target_file_count_label = QLabel("目标目录文件数: 0")
        op_layout.addWidget(self.source_file_count_label)
        op_layout.addWidget(self.target_file_count_label)
        self.match_count_label = QLabel("匹配文件总数: 0")
        op_layout.addWidget(self.match_count_label)
        self.status_label = QLabel("准备就绪")
        op_layout.addWidget(self.status_label)
        
        # 选择操作按钮
        select_layout = QHBoxLayout()
        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.clicked.connect(self.select_all_files)
        self.deselect_all_btn = QPushButton("取消全选")
        self.deselect_all_btn.clicked.connect(self.deselect_all_files)
        self.invert_select_btn = QPushButton("反选")
        self.invert_select_btn.clicked.connect(self.invert_selection)
        select_layout.addWidget(self.select_all_btn)
        select_layout.addWidget(self.deselect_all_btn)
        select_layout.addWidget(self.invert_select_btn)
        op_layout.addLayout(select_layout)
        
        # 文件操作按钮
        self.copy_btn = QPushButton("复制选中文件")
        self.copy_btn.clicked.connect(lambda: self.process_files("copy"))
        op_layout.addWidget(self.copy_btn)
        self.move_btn = QPushButton("移动选中文件")
        self.move_btn.clicked.connect(lambda: self.process_files("move"))
        op_layout.addWidget(self.move_btn)
        self.delete_btn = QPushButton("删除选中文件")
        self.delete_btn.clicked.connect(lambda: self.process_files("delete"))
        op_layout.addWidget(self.delete_btn)
        op_group.setLayout(op_layout)
        left_layout.addWidget(op_group)
        left_layout.addStretch()
        splitter.addWidget(left_panel)
        # 右侧文件匹配结果列表（左右排布）
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        # 匹配结果表格
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(3)  # 复选框、源文件、目标文件
        self.result_table.setHorizontalHeaderLabels(["选择", "源文件", "目标文件"])
        self.result_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.result_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.result_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.result_table.setSelectionMode(QTableWidget.SelectionMode.ExtendedSelection)
        self.result_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        # 设置表格样式，取消选中行的高亮
        self.result_table.setStyleSheet("""
        QTableWidget {
            outline: none;
        }
        QTableWidget::item:selected {
            background-color: transparent;
            color: black;
        }
        """)
        # 创建并设置代理
        self.html_delegate = HTMLDelegate(self.result_table)
        self.result_table.setItemDelegate(self.html_delegate)
        right_layout.addWidget(self.result_table)
        splitter.addWidget(right_panel)
        splitter.setSizes([350, 850])
        main_layout.addWidget(splitter)
        self.setAcceptDrops(True)
        self.on_match_method_changed(0)

    def on_match_method_changed(self, index):
        # 隐藏所有设置项
        def hide_all_widgets(layout):
            for i in range(layout.count()):
                item = layout.itemAt(i)
                if item.widget():
                    item.widget().hide()
                elif item.layout():
                    hide_all_widgets(item.layout())
        hide_all_widgets(self.match_options_layout)
        self.match_options_widget.hide()
        if index == 0:  # 根据源目录匹配提取
            self.match_options_widget.show()
            for i in range(self.dir_match_layout.count()):
                item = self.dir_match_layout.itemAt(i)
                if item.widget():
                    item.widget().show()
                elif item.layout():
                    for j in range(item.layout().count()):
                        subitem = item.layout().itemAt(j)
                        if subitem and subitem.widget():
                            subitem.widget().show()
        elif index == 1:  # 根据文本匹配提取
            self.match_options_widget.show()
            for i in range(self.text_match_layout.count()):
                item = self.text_match_layout.itemAt(i)
                if item.widget():
                    item.widget().show()
                elif item.layout():
                    for j in range(item.layout().count()):
                        subitem = item.layout().itemAt(j)
                        if subitem and subitem.widget():
                            subitem.widget().show()

    def drag_enter_event(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def drop_event(self, event, target_type):
        if event.mimeData().hasUrls():
            url = event.mimeData().urls()[0]
            file_path = url.toLocalFile()
            
            if os.path.isdir(file_path):
                if target_type == "source":
                    self.source_dir = file_path
                    self.source_label.setText(file_path)
                    self.source_label.setToolTip(file_path)
                elif target_type == "target":
                    self.target_dir = file_path
                    self.target_label.setText(file_path)
                    self.target_label.setToolTip(file_path)
                elif target_type == "output":
                    self.output_dir = file_path
                    self.output_label.setText(file_path)
                    self.output_label.setToolTip(file_path)

    def select_directory(self, dir_type):
        directory = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if directory:
            if dir_type == "source":
                self.source_dir = directory
                self.source_label.setText(directory)
                self.source_label.setToolTip(directory)
            elif dir_type == "target":
                self.target_dir = directory
                self.target_label.setText(directory)
                self.target_label.setToolTip(directory)
            elif dir_type == "output":
                self.output_dir = directory
                self.output_label.setText(directory)
                self.output_label.setToolTip(directory)

    def update_source_file_list(self):
        """更新源文件列表"""
        self.source_file_list.clear()
        if not self.source_dir:
            self.source_count_label.setText("源文件总数: 0")
            return
            
        try:
            source_files = [f for f in os.listdir(self.source_dir) if os.path.isfile(os.path.join(self.source_dir, f))]
            for file_name in source_files:
                item = QListWidgetItem(file_name)
                self.source_file_list.addItem(item)
            
            self.source_count_label.setText(f"源文件总数: {len(source_files)}")
        except Exception as e:
            self.source_count_label.setText(f"源文件总数: 0 (错误: {str(e)})")

    def match_files(self):
        # 如果有正在运行的匹配任务，先取消它
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.cancel()
            self.worker.wait()
        
        self.result_table.setRowCount(0)  # 清空表格
        self.matched_files = []
        match_method_index = self.match_method.currentIndex()
        self.match_btn.setEnabled(False)
        
        # 显示进度条和取消按钮
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(100)
        self.cancel_btn.setVisible(True)
        self.cancel_btn.setEnabled(True)
        
        self._start_search_anim()
        source_dir = self.source_dir if hasattr(self, 'source_dir') else ''
        target_dir = self.target_dir if hasattr(self, 'target_dir') else ''
        match_basis = self.match_basis_combo.currentIndex() if hasattr(self, 'match_basis_combo') else 0
        format_match = self.format_match_combo.currentIndex() == 0 if hasattr(self, 'format_match_combo') else True
        text_input = self.text_match_input.toPlainText() if hasattr(self, 'text_match_input') else ''
        text_match_basis = self.text_match_basis_combo.currentIndex() if hasattr(self, 'text_match_basis_combo') else 0
        
        # 使用多线程处理大型目录
        use_multithreading = True  # 默认启用多线程
        
        # 创建并启动工作线程
        self.worker = MatchWorker(
            match_method_index, 
            source_dir, 
            target_dir, 
            match_basis, 
            format_match, 
            text_input, 
            text_match_basis, 
            use_multithreading
        )
        
        # 连接信号
        self.worker.result_ready.connect(self.on_match_result)
        self.worker.error.connect(self.on_match_error)
        self.worker.progress.connect(self.on_match_progress)
        self.worker.file_counts_ready.connect(self.on_file_counts_ready)
        
        # 启动工作线程
        QApplication.processEvents()  # 在启动线程前处理所有待处理的事件
        self.worker.start()
    
    def on_match_progress(self, value, total):
        """更新匹配进度"""
        progress_percent = int(value / total * 100) if total > 0 else 0
        self.progress_bar.setValue(progress_percent)
        self.status_label.setText(f"正在查找匹配文件... {progress_percent}%")

    def _start_search_anim(self):
        self._search_anim_step = 0
        if not self._search_anim_timer:
            self._search_anim_timer = QTimer(self)
            self._search_anim_timer.timeout.connect(self._update_search_anim)
        self._search_anim_timer.start(400)
        self.status_label.setText("正在查找匹配文件，请稍候…")

    def _update_search_anim(self):
        dots = ["…", "……", "………", "……", "…"]
        self.status_label.setText(f"正在查找匹配文件，请稍候{dots[self._search_anim_step % len(dots)]}")
        self._search_anim_step += 1

    def _stop_search_anim(self, msg=None):
        if self._search_anim_timer:
            self._search_anim_timer.stop()
        if msg:
            self.status_label.setText(msg)

    def on_match_result(self, matched_pairs, matched_count):
        self.result_table.setRowCount(0)  # 清空表格
        self.matched_files = []
        self._stop_search_anim(f"找到 {matched_count} 个匹配文件")
        self.match_btn.setEnabled(True)
        
        # 隐藏进度条和取消按钮
        self.progress_bar.setVisible(False)
        self.cancel_btn.setVisible(False)
        
        # 清除旧的高亮信息
        self.html_delegate.highlights.clear()
        
        # 设置表格行数
        self.result_table.setRowCount(matched_count)
        
        # 填充表格
        for row, (s, t, s_name, t_name, key) in enumerate(matched_pairs):
            # 复选框
            checkbox_widget = CheckBoxWidget()
            checkbox_widget.setChecked(True)
            self.result_table.setCellWidget(row, 0, checkbox_widget)
            
            # 源文件列
            if s:
                source_item = QTableWidgetItem(s_name)
                self.result_table.setItem(row, 1, source_item)
                # 如果有匹配关键字，设置单元格部分文本颜色
                if key and key in s_name:
                    self._highlight_cell_text(row, 1, key)
            else:
                self.result_table.setItem(row, 1, QTableWidgetItem(""))
            
            # 目标文件列
            target_item = QTableWidgetItem(t_name)
            self.result_table.setItem(row, 2, target_item)
            # 如果有匹配关键字，设置单元格部分文本颜色
            if key and key in t_name:
                self._highlight_cell_text(row, 2, key)
            
            # 存储文件数据
            self.matched_files.append((s, t, s_name, t_name, key))
        
        self.match_count_label.setText(f"匹配文件总数: {matched_count}")
        if matched_count == 0:
            QMessageBox.information(self, "提示", "未找到匹配的文件")

    def on_match_error(self, msg):
        self._stop_search_anim("查找失败")
        self.match_btn.setEnabled(True)
        
        # 隐藏进度条和取消按钮
        self.progress_bar.setVisible(False)
        self.cancel_btn.setVisible(False)
        
        QMessageBox.critical(self, "错误", f"匹配文件时出错: {msg}")

    def _highlight_cell_text(self, row, column, key):
        """使用自定义文本渲染器高亮单元格中的关键字"""
        # 为特定单元格添加高亮信息
        self.html_delegate.add_highlight(row, column, key)

    def select_all_files(self):
        for row in range(self.result_table.rowCount()):
            checkbox_widget = self.result_table.cellWidget(row, 0)
            if checkbox_widget:
                checkbox_widget.setChecked(True)

    def deselect_all_files(self):
        for row in range(self.result_table.rowCount()):
            checkbox_widget = self.result_table.cellWidget(row, 0)
            if checkbox_widget:
                checkbox_widget.setChecked(False)

    def invert_selection(self):
        for row in range(self.result_table.rowCount()):
            checkbox_widget = self.result_table.cellWidget(row, 0)
            if checkbox_widget:
                checkbox_widget.setChecked(not checkbox_widget.isChecked())

    def process_files(self, operation):
        if not self.result_table.rowCount():
            QMessageBox.warning(self, "警告", "没有可处理的文件")
            return
        if operation != "delete" and not hasattr(self, 'output_dir'):
            QMessageBox.warning(self, "警告", "请先选择输出目录")
            return
        if operation != "delete" and not self.output_dir:
            QMessageBox.warning(self, "警告", "请先选择输出目录")
            return
        
        # 获取选中的文件
        selected_files = []
        for row in range(self.result_table.rowCount()):
            checkbox_widget = self.result_table.cellWidget(row, 0)
            if checkbox_widget and checkbox_widget.isChecked():
                selected_files.append(self.matched_files[row])
        
        if not selected_files:
            QMessageBox.warning(self, "警告", "请先选择要处理的文件")
            return
        
        operation_desc = {"copy": "复制", "move": "移动", "delete": "删除"}
        reply = QMessageBox.question(self, "确认操作", 
                                     f"确定要{operation_desc[operation]}选中的{len(selected_files)}个文件吗？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        
        # 显示进度条和取消按钮
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(selected_files))
        self.progress_bar.setValue(0)
        self.cancel_btn.setVisible(True)
        self.cancel_btn.setEnabled(True)
        
        # 创建文件处理线程
        self.processing = True
        self._is_cancelled = False
        
        processed = 0
        failed = 0
        
        # 创建一个定时器来检查取消状态
        cancel_check_timer = QTimer()
        cancel_check_timer.timeout.connect(lambda: setattr(self, '_is_cancelled', True) if not self.cancel_btn.isEnabled() else None)
        cancel_check_timer.start(100)  # 每100毫秒检查一次
        
        try:
            for i, (s, t, s_name, t_name, key) in enumerate(selected_files):
                if self._is_cancelled:
                    self.status_label.setText("操作已取消")
                    break
                
                try:
                    src_path = t if s is None else t  # 只处理目标文件
                    filename = t_name
                    if operation == "copy":
                        dest_path = os.path.join(self.output_dir, filename)
                        shutil.copy2(src_path, dest_path)
                    elif operation == "move":
                        dest_path = os.path.join(self.output_dir, filename)
                        shutil.move(src_path, dest_path)
                    elif operation == "delete":
                        os.remove(src_path)
                    processed += 1
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"处理文件 {filename} 时出错: {str(e)}")
                    failed += 1
                finally:
                    # 更新进度条
                    self.progress_bar.setValue(processed + failed)
                    self.status_label.setText(f"正在{operation_desc[operation]}文件... {processed + failed}/{len(selected_files)}")
                    QApplication.processEvents()
        finally:
            # 停止定时器
            cancel_check_timer.stop()
            
            # 隐藏进度条和取消按钮
            self.progress_bar.setVisible(False)
            self.cancel_btn.setVisible(False)
            self.processing = False
            
            if self._is_cancelled:
                self.status_label.setText(f"操作已取消: 成功 {processed} 个文件, 失败 {failed} 个文件")
                QMessageBox.information(self, "已取消", f"操作已取消: 成功{operation_desc[operation]}{processed}个文件，失败{failed}个文件")
            else:
                self.status_label.setText(f"{operation_desc[operation]}完成: 成功 {processed} 个文件, 失败 {failed} 个文件")
                QMessageBox.information(self, "完成", f"操作完成: 成功{operation_desc[operation]}{processed}个文件，失败{failed}个文件")
            
            if operation in ["move", "delete"] and processed > 0:
                self.match_files()

    def on_file_counts_ready(self, source_count, target_count):
        """更新文件数量统计"""
        self.source_file_count_label.setText(f"源目录文件数: {source_count}")
        self.target_file_count_label.setText(f"目标目录文件数: {target_count}")

    def cancel_operation(self):
        """取消当前操作"""
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.cancel()
            self.status_label.setText("正在取消操作...")
            self.cancel_btn.setEnabled(False)

    def closeEvent(self, event):
        """应用关闭时的处理"""
        # 取消所有正在运行的线程
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.cancel()
            self.worker.wait(1000)  # 等待最多1秒
            if self.worker.isRunning():
                self.worker.terminate()  # 强制终止
        
        # 停止动画计时器
        if self._search_anim_timer and self._search_anim_timer.isActive():
            self._search_anim_timer.stop()
        
        # 接受关闭事件
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # 使用Fusion样式，看起来更现代
    
    # 设置应用图标
    app_icon = create_app_icon()
    app.setWindowIcon(app_icon)
    
    window = FileSelector()
    window.show()
    sys.exit(app.exec()) 