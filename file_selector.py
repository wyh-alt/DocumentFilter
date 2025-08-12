#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import re
import shutil
import concurrent.futures
import tempfile
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                           QListWidget, QListWidgetItem, QPushButton, QFileDialog, QLabel,
                           QCheckBox, QMessageBox, QProgressBar, QComboBox, QGroupBox,
                           QLineEdit, QSplitter, QFrame, QTextEdit, QTableWidget, QTableWidgetItem, QHeaderView,
                           QStyledItemDelegate, QTabWidget, QRadioButton, QButtonGroup, QSizePolicy)
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
    unmatched_ready = pyqtSignal(list, list)  # 新增信号：未匹配的源文件和检索文本
    def __init__(self, match_method_index, source_dir, target_dir, match_basis, format_match, text_input, text_match_basis=0, expand_search=False, use_multithreading=True):
        super().__init__()
        self.match_method_index = match_method_index
        self.source_dir = source_dir
        self.target_dir = target_dir
        self.match_basis = match_basis
        self.format_match = format_match
        self.text_input = text_input
        self.text_match_basis = text_match_basis
        self.expand_search = expand_search
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
                    result.append((None, t, kw, t_name, kw))
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
            unmatched_source_files = []  # 未匹配的源文件
            unmatched_keywords = []      # 未匹配的检索文本
            
            if self.match_method_index == 0:  # 根据源目录匹配提取
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
                matched_source_files = set()  # 记录已匹配的源文件
                
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
                                # 记录已匹配的源文件
                                for s, t, s_name, t_name, key in chunk_result:
                                    if s:
                                        matched_source_files.add(s)
                                completed += 1
                                # 更新进度
                                progress = int((completed / len(chunks)) * total)
                                self.progress.emit(progress, total)
                    else:
                        # 单线程处理
                        for s in source_files:
                            s_name = os.path.basename(s)
                            matched = False
                            
                            if self.format_match:  # 匹配包括扩展名
                                if s_name in target_dict:
                                    t = target_dict[s_name]
                                    t_name = os.path.basename(t)
                                    matched_pairs.append((s, t, s_name, t_name, s_name))
                                    matched = True
                            else:  # 不匹配扩展名
                                s_name_no_ext = os.path.splitext(s_name)[0]
                                if s_name_no_ext in target_dict_no_ext:
                                    for t in target_dict_no_ext[s_name_no_ext]:
                                        t_name = os.path.basename(t)
                                        matched_pairs.append((s, t, s_name, t_name, s_name_no_ext))
                                        matched = True
                            
                            if matched:
                                matched_source_files.add(s)
                            
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
                                # 记录已匹配的源文件
                                for s, t, s_name, t_name, key in chunk_result:
                                    if s:
                                        matched_source_files.add(s)
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
                                        matched_source_files.add(s)
                            
                            count += 1
                            if count % 20 == 0 or count == total:
                                self.progress.emit(count, total)
                
                # 找出未匹配的源文件
                unmatched_source_files = [s for s in source_files if s not in matched_source_files]
            
            elif self.match_method_index == 1:  # 根据文本匹配提取
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
                    matched_keywords = set()
                    for kw in keywords:
                        if kw in name_dict:
                            t, t_name = name_dict[kw]
                            matched_pairs.append((None, t, '', t_name, kw))
                            matched_keywords.add(kw)
                    
                    # 找出未匹配的关键字
                    unmatched_keywords = [kw for kw in keywords if kw not in matched_keywords]
                
                else:  # 检索匹配
                    total = len(target_files)
                    matched_keywords = set()
                    
                    if not self.expand_search:  # 精确匹配模式
                        # 为每个关键字找到最佳匹配的文件
                        keyword_best_matches = {}  # 存储每个关键字的最佳匹配
                        
                        for t in target_files:
                            t_name = os.path.basename(t)
                            t_name_no_ext = os.path.splitext(t_name)[0]
                            
                            for kw in keywords:
                                if kw in t_name_no_ext:
                                    # 计算匹配度（精确匹配优先）
                                    if t_name_no_ext == kw:
                                        # 完全匹配，最高优先级
                                        match_score = 100
                                    elif t_name_no_ext.startswith(kw + '-') or t_name_no_ext.startswith(kw + '_'):
                                        # 以关键字开头，次高优先级
                                        match_score = 80
                                    elif t_name_no_ext.endswith('-' + kw) or t_name_no_ext.endswith('_' + kw):
                                        # 以关键字结尾，中等优先级
                                        match_score = 60
                                    else:
                                        # 包含关键字，最低优先级
                                        match_score = 40
                                    
                                    # 如果这个关键字还没有匹配，或者当前匹配度更高，则更新
                                    if kw not in keyword_best_matches or match_score > keyword_best_matches[kw][2]:
                                        keyword_best_matches[kw] = (t, t_name, match_score)
                        
                        # 将最佳匹配添加到结果中
                        for kw, (t, t_name, score) in keyword_best_matches.items():
                            matched_pairs.append((None, t, kw, t_name, kw))
                            matched_keywords.add(kw)
                    
                    else:  # 扩大匹配模式
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
                                    # 记录已匹配的关键字
                                    for s, t, s_name, t_name, key in chunk_result:
                                        matched_keywords.add(key)
                                    completed += 1
                                    # 更新进度
                                    progress = int((completed / len(chunks)) * total)
                                    self.progress.emit(progress, total)
                        else:
                            # 单线程处理
                            count = 0
                            for t in target_files:
                                t_name = os.path.basename(t)
                                
                                # 检查所有关键字
                                for kw in keywords:
                                    if kw in t_name:
                                        matched_pairs.append((None, t, kw, t_name, kw))
                                        matched_keywords.add(kw)
                                
                                count += 1
                                if count % 20 == 0 or count == total:
                                    self.progress.emit(count, total)
                    
                    # 找出未匹配的关键字
                    unmatched_keywords = [kw for kw in keywords if kw not in matched_keywords]
            
            # 发送未匹配文件信号
            self.unmatched_ready.emit(unmatched_source_files, unmatched_keywords)
            
            self.result_ready.emit(matched_pairs, len(matched_pairs))
        except Exception as e:
            self.error.emit(str(e))

class RenameWorker(QThread):
    """批量重命名工作线程"""
    progress = pyqtSignal(int, int)
    result_ready = pyqtSignal(list, int)
    error = pyqtSignal(str)
    file_count_ready = pyqtSignal(int)
    
    def __init__(self, directory, rename_type, prefix="", suffix="", 
                 delete_text="", replace_from="", replace_to="", 
                 start_number=1, number_format="{:03d}"):
        super().__init__()
        self.directory = directory
        self.rename_type = rename_type
        self.prefix = prefix
        self.suffix = suffix
        self.delete_text = delete_text
        self.replace_from = replace_from
        self.replace_to = replace_to
        self.start_number = start_number
        self.number_format = number_format
        self._is_cancelled = False
    
    def cancel(self):
        """取消重命名操作"""
        self._is_cancelled = True
    
    def run(self):
        try:
            # 获取目录中的所有文件
            if not self.directory or not os.path.isdir(self.directory):
                self.error.emit("请选择有效的目录")
                return
            
            files = [f for f in os.listdir(self.directory) if os.path.isfile(os.path.join(self.directory, f))]
            
            # 发送文件数量信号
            self.file_count_ready.emit(len(files))
            
            if not files:
                self.error.emit("目录中没有文件")
                return
            
            # 检查是否已取消
            if self._is_cancelled:
                return
            
            # 预览重命名结果
            results = []
            for i, filename in enumerate(files):
                if self._is_cancelled:
                    return
                
                name, ext = os.path.splitext(filename)
                new_name = name
                
                # 根据不同的重命名类型处理
                if self.rename_type == 0:  # 添加前缀
                    new_name = self.prefix + name
                elif self.rename_type == 1:  # 添加后缀
                    new_name = name + self.suffix
                elif self.rename_type == 2:  # 删除指定字段
                    if self.delete_text:
                        new_name = name.replace(self.delete_text, "")
                elif self.rename_type == 3:  # 替换指定字段
                    if self.replace_from:
                        new_name = name.replace(self.replace_from, self.replace_to)
                elif self.rename_type == 4:  # 按升序编号
                    number = self.start_number + i
                    formatted_number = self.number_format.format(number)
                    new_name = formatted_number + "_" + name
                
                new_filename = new_name + ext
                results.append((filename, new_filename))
                
                # 更新进度
                self.progress.emit(i + 1, len(files))
            
            # 发送结果
            self.result_ready.emit(results, len(results))
            
        except Exception as e:
            self.error.emit(str(e))

class FileSelector(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.source_dir = ""
        self.target_dir = ""
        self.output_dir = ""
        self.rename_dir = ""  # 新增重命名目录
        self.matched_files = []
        self._search_anim_timer = None
        self._search_anim_step = 0
        self.processing = False
        self.rename_results = []  # 存储重命名结果
        self.rename_history = []  # 存储重命名历史，用于撤回功能
        
        # Excel相关属性
        self.excel_file_path = None  # Excel文件路径
        self.excel_data = None  # Excel数据
        self.temp_excel_dir = None  # 临时Excel文件目录
        
        # 设置应用图标
        self.setWindowIcon(create_app_icon())

    def init_ui(self):
        self.setWindowTitle("文件处理工具 v2.2        *使用遇到bug或有功能建议请及时联系王永皓")
        self.setGeometry(300, 300, 1000, 800)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 创建选项卡控件
        self.tab_widget = QTabWidget()
        
        # 批量文件提取选项卡
        file_selector_tab = QWidget()
        file_selector_layout = QVBoxLayout(file_selector_tab)
        
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
        
        file_selector_layout.addLayout(progress_layout)
        
        splitter = QSplitter(Qt.Orientation.Horizontal)
        # 左侧设置面板
        left_panel = QWidget()
        left_panel.setMinimumWidth(320)
        left_panel.setMaximumWidth(340)
        left_layout = QVBoxLayout(left_panel)
        # 目录选择区域
        dir_group = QGroupBox("目录设置")
        dir_layout = QVBoxLayout()
        # 目标目录选择
        target_layout = QHBoxLayout()
        target_layout.addWidget(QLabel("目标目录:"))
        self.target_label = QLabel("未选择")
        self.target_label.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.target_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.target_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextBrowserInteraction)
        self.target_label.setWordWrap(False)
        self.target_label.setToolTip("未选择")
        self.target_label.setAcceptDrops(True)
        self.target_label.dragEnterEvent = self.drag_enter_event
        def drop_event_target(event):
            self.drop_event(event, "target")
        self.target_label.dropEvent = drop_event_target
        target_btn = QPushButton("浏览...")
        target_btn.clicked.connect(lambda: self.select_directory("target"))
        target_layout.addWidget(self.target_label, 1)
        target_layout.addWidget(target_btn, 0)
        # 输出目录选择
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出目录:"))
        self.output_label = QLabel("未选择")
        self.output_label.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.output_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.output_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextBrowserInteraction)
        self.output_label.setWordWrap(False)
        self.output_label.setToolTip("未选择")
        self.output_label.setAcceptDrops(True)
        self.output_label.dragEnterEvent = self.drag_enter_event
        def drop_event_output(event):
            self.drop_event(event, "output")
        self.output_label.dropEvent = drop_event_output
        output_btn = QPushButton("浏览...")
        output_btn.clicked.connect(lambda: self.select_directory("output"))
        output_layout.addWidget(self.output_label, 1)
        output_layout.addWidget(output_btn, 0)
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
        # 源目录选择
        source_layout = QHBoxLayout()
        source_layout.addWidget(QLabel("源目录:"))
        self.source_label = QLabel("未选择")
        self.source_label.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.source_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.source_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextBrowserInteraction)
        self.source_label.setWordWrap(False)
        self.source_label.setToolTip("未选择")
        self.source_label.setAcceptDrops(True)
        self.source_label.dragEnterEvent = self.drag_enter_event
        self.source_label.dropEvent = lambda event: self.drop_event(event, "source")
        source_btn = QPushButton("浏览...")
        source_btn.clicked.connect(lambda: self.select_directory("source"))
        source_layout.addWidget(self.source_label, 1)
        source_layout.addWidget(source_btn, 0)
        self.dir_match_layout.addLayout(source_layout)
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
        self.text_match_basis_combo.currentIndexChanged.connect(self.on_text_match_basis_changed)
        text_match_basis_row.addWidget(self.text_match_basis_combo)
        self.text_match_layout.addLayout(text_match_basis_row)
        
        # 添加扩大文件名检索范围选项
        expand_search_row = QHBoxLayout()
        expand_search_row.addWidget(QLabel("扩大文件名检索范围:"))
        self.expand_search_checkbox = QCheckBox()
        self.expand_search_checkbox.setChecked(False)  # 默认取消勾选
        expand_search_row.addWidget(self.expand_search_checkbox)
        
        # 添加功能说明标签
        info_label = QLabel("ⓘ")
        info_label.setToolTip("当未勾选时，根据用户检索内容进行文件一对一匹配，当多个文件包含用户提供的文本时，仅保留匹配度最高的文件。\n例如检索文本'123'，仅匹配'123-原唱'，忽略'1234-原唱'以及'11234-原唱'，避免重复匹配。\n当勾选该选项时，将对目录内包含检索文本的所有文件进行匹配。")
        expand_search_row.addWidget(info_label)
        
        self.text_match_layout.addLayout(expand_search_row)
        
        # 初始状态：隐藏扩大检索范围选项
        self.expand_search_checkbox.hide()
        expand_search_row.itemAt(0).widget().hide()  # 隐藏标签
        expand_search_row.itemAt(2).widget().hide()  # 隐藏说明标签
        
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
        file_selector_layout.addWidget(splitter)
        
        # 批量重命名选项卡
        rename_tab = QWidget()
        rename_layout = QVBoxLayout(rename_tab)
        
        # 重命名目录选择
        rename_dir_group = QGroupBox("目录设置")
        rename_dir_layout = QVBoxLayout()
        
        rename_dir_row = QHBoxLayout()
        rename_dir_row.addWidget(QLabel("重命名目录:"))
        self.rename_dir_label = QLabel("未选择")
        self.rename_dir_label.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.rename_dir_label.setFixedWidth(300)
        self.rename_dir_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse | Qt.TextInteractionFlag.TextBrowserInteraction)
        self.rename_dir_label.setWordWrap(False)
        self.rename_dir_label.setToolTip("未选择")
        self.rename_dir_label.setAcceptDrops(True)
        self.rename_dir_label.dragEnterEvent = self.drag_enter_event
        self.rename_dir_label.dropEvent = lambda event: self.drop_event(event, "rename")
        
        rename_dir_btn = QPushButton("浏览...")
        rename_dir_btn.clicked.connect(lambda: self.select_directory("rename"))
        
        rename_dir_row.addWidget(self.rename_dir_label)
        rename_dir_row.addWidget(rename_dir_btn)
        rename_dir_layout.addLayout(rename_dir_row)
        
        # 显示文件数量
        self.rename_file_count_label = QLabel("目录文件数: 0")
        rename_dir_layout.addWidget(self.rename_file_count_label)
        
        rename_dir_group.setLayout(rename_dir_layout)
        rename_layout.addWidget(rename_dir_group)
        
        # 重命名设置
        rename_settings_group = QGroupBox("重命名设置")
        rename_settings_layout = QVBoxLayout()
        
        # 重命名类型选择
        rename_type_layout = QVBoxLayout()
        rename_type_layout.addWidget(QLabel("重命名类型:"))
        
        # 使用单选按钮组
        self.rename_type_group = QButtonGroup(self)
        
        # 添加前缀选项
        prefix_radio = QRadioButton("添加前缀")
        prefix_radio.setChecked(True)  # 默认选中
        self.rename_type_group.addButton(prefix_radio, 0)
        rename_type_layout.addWidget(prefix_radio)
        
        self.prefix_input = QLineEdit()
        self.prefix_input.setPlaceholderText("请输入前缀")
        rename_type_layout.addWidget(self.prefix_input)
        
        # 添加后缀选项
        suffix_radio = QRadioButton("添加后缀")
        self.rename_type_group.addButton(suffix_radio, 1)
        rename_type_layout.addWidget(suffix_radio)
        
        self.suffix_input = QLineEdit()
        self.suffix_input.setPlaceholderText("请输入后缀")
        self.suffix_input.setEnabled(False)
        rename_type_layout.addWidget(self.suffix_input)
        
        # 删除指定字段选项
        delete_radio = QRadioButton("删除指定字段")
        self.rename_type_group.addButton(delete_radio, 2)
        rename_type_layout.addWidget(delete_radio)
        
        self.delete_input = QLineEdit()
        self.delete_input.setPlaceholderText("请输入要删除的字段")
        self.delete_input.setEnabled(False)
        rename_type_layout.addWidget(self.delete_input)
        
        # 替换指定字段选项
        replace_radio = QRadioButton("替换指定字段")
        self.rename_type_group.addButton(replace_radio, 3)
        rename_type_layout.addWidget(replace_radio)
        
        replace_layout = QHBoxLayout()
        self.replace_from_input = QLineEdit()
        self.replace_from_input.setPlaceholderText("原字段")
        self.replace_from_input.setEnabled(False)
        replace_layout.addWidget(self.replace_from_input)
        
        replace_layout.addWidget(QLabel("→"))
        
        self.replace_to_input = QLineEdit()
        self.replace_to_input.setPlaceholderText("新字段")
        self.replace_to_input.setEnabled(False)
        replace_layout.addWidget(self.replace_to_input)
        rename_type_layout.addLayout(replace_layout)
        
        # 按升序编号选项
        number_radio = QRadioButton("按升序编号")
        self.rename_type_group.addButton(number_radio, 4)
        rename_type_layout.addWidget(number_radio)
        
        number_layout = QHBoxLayout()
        number_layout.addWidget(QLabel("起始编号:"))
        self.start_number_input = QLineEdit("1")
        self.start_number_input.setEnabled(False)
        number_layout.addWidget(self.start_number_input)
        
        number_layout.addWidget(QLabel("格式:"))
        self.number_format_combo = QComboBox()
        self.number_format_combo.addItems(["001", "01", "1"])
        self.number_format_combo.setEnabled(False)
        number_layout.addWidget(self.number_format_combo)
        rename_type_layout.addLayout(number_layout)
        
        # Excel替换命名选项
        excel_radio = QRadioButton("Excel替换命名")
        self.rename_type_group.addButton(excel_radio, 5)
        rename_type_layout.addWidget(excel_radio)
        
        excel_layout = QHBoxLayout()
        self.open_excel_btn = QPushButton("打开表格")
        self.open_excel_btn.setEnabled(False)
        self.open_excel_btn.clicked.connect(self.open_excel_file)
        excel_layout.addWidget(self.open_excel_btn)
        
        self.excel_status_label = QLabel("请先选择重命名目录")
        excel_layout.addWidget(self.excel_status_label)
        rename_type_layout.addLayout(excel_layout)
        
        # 连接单选按钮信号
        self.rename_type_group.buttonClicked.connect(self.on_rename_type_changed)
        
        rename_settings_layout.addLayout(rename_type_layout)
        rename_settings_group.setLayout(rename_settings_layout)
        rename_layout.addWidget(rename_settings_group)
        
        # 预览和执行按钮
        rename_buttons_layout = QHBoxLayout()
        self.preview_rename_btn = QPushButton("预览重命名")
        self.preview_rename_btn.clicked.connect(self.preview_rename)
        rename_buttons_layout.addWidget(self.preview_rename_btn)
        
        self.execute_rename_btn = QPushButton("执行重命名")
        self.execute_rename_btn.clicked.connect(self.execute_rename)
        self.execute_rename_btn.setEnabled(False)
        rename_buttons_layout.addWidget(self.execute_rename_btn)
        
        self.undo_rename_btn = QPushButton("撤回重命名")
        self.undo_rename_btn.clicked.connect(self.undo_rename)
        self.undo_rename_btn.setEnabled(False)
        rename_buttons_layout.addWidget(self.undo_rename_btn)
        
        rename_layout.addLayout(rename_buttons_layout)
        
        # 重命名状态标签
        self.rename_status_label = QLabel("准备就绪")
        rename_layout.addWidget(self.rename_status_label)
        
        # 重命名预览表格
        self.rename_preview_table = QTableWidget()
        self.rename_preview_table.setColumnCount(2)
        self.rename_preview_table.setHorizontalHeaderLabels(["原文件名", "新文件名"])
        self.rename_preview_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.rename_preview_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        rename_layout.addWidget(self.rename_preview_table)
        
        # 添加选项卡
        self.tab_widget.addTab(file_selector_tab, "批量文件提取")
        self.tab_widget.addTab(rename_tab, "批量重命名")
        
        main_layout.addWidget(self.tab_widget)
        self.setAcceptDrops(True)
        self.on_match_method_changed(0)

    def on_rename_type_changed(self, button):
        """当重命名类型改变时更新UI"""
        # 禁用所有输入框
        self.prefix_input.setEnabled(False)
        self.suffix_input.setEnabled(False)
        self.delete_input.setEnabled(False)
        self.replace_from_input.setEnabled(False)
        self.replace_to_input.setEnabled(False)
        self.start_number_input.setEnabled(False)
        self.number_format_combo.setEnabled(False)
        self.open_excel_btn.setEnabled(False)
        
        # 根据选择的类型启用相应的输入框
        type_id = self.rename_type_group.id(button)
        if type_id == 0:  # 添加前缀
            self.prefix_input.setEnabled(True)
        elif type_id == 1:  # 添加后缀
            self.suffix_input.setEnabled(True)
        elif type_id == 2:  # 删除指定字段
            self.delete_input.setEnabled(True)
        elif type_id == 3:  # 替换指定字段
            self.replace_from_input.setEnabled(True)
            self.replace_to_input.setEnabled(True)
        elif type_id == 4:  # 按升序编号
            self.start_number_input.setEnabled(True)
            self.number_format_combo.setEnabled(True)
        elif type_id == 5: # Excel替换命名
            self.open_excel_btn.setEnabled(True)
            self.excel_status_label.setText("请先选择重命名目录")
            # 如果已经选择了重命名目录，自动生成Excel文件
            if self.rename_dir and os.path.isdir(self.rename_dir):
                try:
                    files = [f for f in os.listdir(self.rename_dir) if os.path.isfile(os.path.join(self.rename_dir, f))]
                    self.generate_excel_file(files)
                except Exception as e:
                    self.excel_status_label.setText(f"生成Excel文件失败: {str(e)}")

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
                elif target_type == "rename":
                    self.rename_dir = file_path
                    self.rename_dir_label.setText(file_path)
                    self.rename_dir_label.setToolTip(file_path)
                    # 清除之前的Excel缓存
                    self.clear_excel_cache()
                    self.update_rename_file_count()

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
            elif dir_type == "rename":
                self.rename_dir = directory
                self.rename_dir_label.setText(directory)
                self.rename_dir_label.setToolTip(directory)
                # 清除之前的Excel缓存
                self.clear_excel_cache()
                self.update_rename_file_count()

    def update_rename_file_count(self):
        """更新重命名目录的文件数量"""
        if not self.rename_dir or not os.path.isdir(self.rename_dir):
            self.rename_file_count_label.setText("目录文件数: 0")
            return
        
        try:
            files = [f for f in os.listdir(self.rename_dir) if os.path.isfile(os.path.join(self.rename_dir, f))]
            self.rename_file_count_label.setText(f"目录文件数: {len(files)}")
            
            # 如果选择了Excel替换命名，自动生成Excel文件
            if self.rename_type_group.checkedId() == 5:
                self.generate_excel_file(files)
        except Exception as e:
            self.rename_file_count_label.setText(f"目录文件数: 0 (错误: {str(e)})")

    def generate_excel_file(self, files):
        """生成Excel文件"""
        try:
            # 如果已经有Excel文件且存在，先检查是否可用
            if self.excel_file_path and os.path.exists(self.excel_file_path):
                try:
                    # 尝试读取文件，检查是否被占用
                    test_df = pd.read_excel(self.excel_file_path, engine='openpyxl')
                    # 如果文件可以正常读取，检查文件数量是否匹配
                    if not test_df.empty and len(test_df) == len(files):
                        # 确保数据类型正确
                        test_df['源文件名'] = test_df['源文件名'].astype(str)
                        test_df['重命名'] = test_df['重命名'].astype(str)
                        self.excel_status_label.setText(f"Excel文件已存在: {len(files)}个文件")
                        self.excel_data = test_df
                        return
                except:
                    # 如果文件被占用或无法读取，继续生成新文件
                    pass
            
            # 创建临时目录
            if not self.temp_excel_dir:
                self.temp_excel_dir = tempfile.mkdtemp(prefix="rename_excel_")
            
            # 创建Excel文件路径
            self.excel_file_path = os.path.join(self.temp_excel_dir, "rename_files.xlsx")
            
            # 创建DataFrame
            df = pd.DataFrame({
                '源文件名': [str(f) for f in files],
                '重命名': [str(f) for f in files]  # 默认重命名与源文件名相同
            })
            
            # 保存到Excel文件
            df.to_excel(self.excel_file_path, index=False, engine='openpyxl')
            
            # 更新状态
            self.excel_status_label.setText(f"Excel文件已生成: {len(files)}个文件")
            self.excel_data = df
            
        except Exception as e:
            self.excel_status_label.setText(f"生成Excel文件失败: {str(e)}")

    def open_excel_file(self):
        """打开Excel文件"""
        # 如果没有Excel文件，先尝试生成
        if not self.excel_file_path or not os.path.exists(self.excel_file_path):
            if not self.rename_dir or not os.path.isdir(self.rename_dir):
                QMessageBox.warning(self, "警告", "请先选择重命名目录以生成Excel文件")
                return
            
            try:
                files = [f for f in os.listdir(self.rename_dir) if os.path.isfile(os.path.join(self.rename_dir, f))]
                self.generate_excel_file(files)
            except Exception as e:
                QMessageBox.warning(self, "错误", f"生成Excel文件失败: {str(e)}")
                return
        
        # 检查Excel文件是否被占用
        try:
            test_df = pd.read_excel(self.excel_file_path, engine='openpyxl')
            # 验证文件格式
            if test_df.empty:
                QMessageBox.warning(self, "错误", "Excel文件为空")
                return
            if '源文件名' not in test_df.columns or '重命名' not in test_df.columns:
                QMessageBox.warning(self, "错误", "Excel文件格式不正确，缺少必要的列")
                return
        except Exception as e:
            QMessageBox.warning(self, "错误", f"Excel文件被占用或损坏，请关闭Excel程序后重试: {str(e)}")
            return
        
        try:
            # 使用系统默认程序打开Excel文件
            if sys.platform == "win32":
                os.startfile(self.excel_file_path)
            elif sys.platform == "darwin":  # macOS
                os.system(f"open '{self.excel_file_path}'")
            else:  # Linux
                os.system(f"xdg-open '{self.excel_file_path}'")
            
            # 显示使用说明
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("使用说明")
            msg_box.setText("Excel文件已打开，请按以下步骤操作：\n\n"
                "1. 在'重命名'列中编辑每个文件的新名称\n"
                "2. 如果新名称不包含文件扩展名，将自动保持原文件扩展名\n"
                "3. 保存Excel文件并关闭\n"
                "4. 返回程序点击'预览重命名'查看效果\n"
                "5. 确认无误后点击'执行重命名'\n\n"
                "注意：请勿修改'源文件名'列，只编辑'重命名'列")
            msg_box.setIcon(QMessageBox.Icon.Information)
            msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
            # 设置弹窗置顶
            msg_box.setWindowFlags(msg_box.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
            msg_box.exec()
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"无法打开Excel文件: {str(e)}")

    def load_excel_data(self):
        """加载Excel数据"""
        if not self.excel_file_path or not os.path.exists(self.excel_file_path):
            return None
        
        try:
            df = pd.read_excel(self.excel_file_path, engine='openpyxl')
            
            # 验证DataFrame结构
            if df.empty:
                QMessageBox.warning(self, "错误", "Excel文件为空")
                return None
            
            # 检查必要的列是否存在
            required_columns = ['源文件名', '重命名']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                QMessageBox.warning(self, "错误", f"Excel文件缺少必要的列: {', '.join(missing_columns)}")
                return None
            
            # 确保数据类型正确
            df['源文件名'] = df['源文件名'].astype(str)
            df['重命名'] = df['重命名'].astype(str)
            
            # 处理NaN值
            df['源文件名'] = df['源文件名'].fillna('')
            df['重命名'] = df['重命名'].fillna('')
            
            return df
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"无法读取Excel文件: {str(e)}")
            print(f"Excel文件读取详细错误: {str(e)}")
            import traceback
            traceback.print_exc()
            return None

    def clear_excel_cache(self):
        """清除Excel文件缓存"""
        try:
            if self.temp_excel_dir and os.path.exists(self.temp_excel_dir):
                # 尝试删除Excel文件
                if self.excel_file_path and os.path.exists(self.excel_file_path):
                    try:
                        os.remove(self.excel_file_path)
                    except PermissionError:
                        # 如果文件被占用，等待一下再试
                        import time
                        time.sleep(1)
                        try:
                            os.remove(self.excel_file_path)
                        except:
                            pass  # 如果还是无法删除，就跳过
                    except:
                        pass  # 其他错误也跳过
                
                # 尝试删除临时目录
                try:
                    shutil.rmtree(self.temp_excel_dir)
                except PermissionError:
                    # 如果目录被占用，等待一下再试
                    import time
                    time.sleep(1)
                    try:
                        shutil.rmtree(self.temp_excel_dir)
                    except:
                        pass  # 如果还是无法删除，就跳过
                except:
                    pass  # 其他错误也跳过
                
                self.temp_excel_dir = None
                self.excel_file_path = None
                self.excel_data = None
        except Exception as e:
            print(f"清除Excel缓存失败: {str(e)}")
            # 即使清除失败，也要重置变量
            self.temp_excel_dir = None
            self.excel_file_path = None
            self.excel_data = None

    def preview_rename(self):
        """预览重命名结果"""
        if not self.rename_dir:
            QMessageBox.warning(self, "警告", "请先选择重命名目录")
            return
        
        # 获取重命名类型和参数
        rename_type = self.rename_type_group.checkedId()
        
        # 获取输入参数
        prefix = self.prefix_input.text()
        suffix = self.suffix_input.text()
        delete_text = self.delete_input.text()
        replace_from = self.replace_from_input.text()
        replace_to = self.replace_to_input.text()
        
        # 获取编号参数
        try:
            start_number = int(self.start_number_input.text())
        except ValueError:
            start_number = 1
        
        # 获取编号格式
        number_format_index = self.number_format_combo.currentIndex()
        number_format = "{:03d}" if number_format_index == 0 else "{:02d}" if number_format_index == 1 else "{:d}"
        
        # 验证输入
        if rename_type == 0 and not prefix:
            QMessageBox.warning(self, "警告", "请输入前缀")
            return
        elif rename_type == 1 and not suffix:
            QMessageBox.warning(self, "警告", "请输入后缀")
            return
        elif rename_type == 2 and not delete_text:
            QMessageBox.warning(self, "警告", "请输入要删除的字段")
            return
        elif rename_type == 3 and not replace_from:
            QMessageBox.warning(self, "警告", "请输入要替换的原字段")
            return
        elif rename_type == 5:  # Excel替换命名
            # 加载Excel数据
            df = self.load_excel_data()
            if df is None:
                QMessageBox.warning(self, "警告", "无法读取Excel文件，请确保文件已保存")
                return
            
            # 直接处理Excel数据
            self.process_excel_rename(df)
            return
        
        # 显示进度条和取消按钮
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.cancel_btn.setVisible(True)
        self.cancel_btn.setEnabled(True)
        
        # 禁用预览按钮
        self.preview_rename_btn.setEnabled(False)
        self.execute_rename_btn.setEnabled(False)
        
        # 创建并启动工作线程
        self.rename_worker = RenameWorker(
            self.rename_dir,
            rename_type,
            prefix,
            suffix,
            delete_text,
            replace_from,
            replace_to,
            start_number,
            number_format
        )
        
        # 连接信号
        self.rename_worker.progress.connect(self.on_rename_progress)
        self.rename_worker.result_ready.connect(self.on_rename_preview_result)
        self.rename_worker.error.connect(self.on_rename_error)
        self.rename_worker.file_count_ready.connect(self.on_rename_file_count_ready)
        
        # 更新状态
        self.rename_status_label.setText("正在生成预览...")
        
        # 启动工作线程
        self.rename_worker.start()

    def process_excel_rename(self, df):
        """处理Excel重命名数据"""
        try:
            results = []
            files = [f for f in os.listdir(self.rename_dir) if os.path.isfile(os.path.join(self.rename_dir, f))]
            
            # 创建文件名到新名称的映射
            rename_map = {}
            for index, row in df.iterrows():
                try:
                    old_name = str(row['源文件名']) if pd.notna(row['源文件名']) else ""
                    new_name = str(row['重命名']) if pd.notna(row['重命名']) else ""
                    
                    if new_name and new_name != old_name:
                        # 检查新名称是否包含文件扩展名
                        if '.' not in new_name:
                            # 如果没有扩展名，保持原文件的扩展名
                            old_name_without_ext, ext = os.path.splitext(old_name)
                            new_name = new_name + ext
                        rename_map[old_name] = new_name
                except Exception as e:
                    print(f"处理第{index}行数据时出错: {str(e)}")
                    continue
            
            # 生成重命名结果
            for file in files:
                if file in rename_map:
                    new_name = rename_map[file]
                    results.append((file, new_name))
                else:
                    results.append((file, file))  # 没有重命名的保持原名
            
            # 显示结果
            self.on_rename_preview_result(results, len(files))
            
        except Exception as e:
            QMessageBox.warning(self, "错误", f"处理Excel数据失败: {str(e)}")
            print(f"Excel数据处理详细错误: {str(e)}")
            import traceback
            traceback.print_exc()

    def on_rename_progress(self, value, total):
        """更新重命名进度"""
        progress_percent = int(value / total * 100) if total > 0 else 0
        self.progress_bar.setValue(progress_percent)

    def on_rename_file_count_ready(self, count):
        """更新重命名文件数量"""
        self.rename_file_count_label.setText(f"目录文件数: {count}")

    def on_rename_preview_result(self, results, count):
        """显示重命名预览结果"""
        # 存储结果
        self.rename_results = results
        
        # 清空表格
        self.rename_preview_table.setRowCount(0)
        
        # 设置表格行数
        self.rename_preview_table.setRowCount(count)
        
        # 填充表格
        for row, (old_name, new_name) in enumerate(results):
            self.rename_preview_table.setItem(row, 0, QTableWidgetItem(old_name))
            self.rename_preview_table.setItem(row, 1, QTableWidgetItem(new_name))
        
        # 更新状态
        self.rename_status_label.setText(f"预览生成完成，共 {count} 个文件")
        
        # 隐藏进度条和取消按钮
        self.progress_bar.setVisible(False)
        self.cancel_btn.setVisible(False)
        
        # 启用按钮
        self.preview_rename_btn.setEnabled(True)
        self.execute_rename_btn.setEnabled(True)
        
        # 检查是否有重名文件
        new_names = [new_name for _, new_name in results]
        if len(new_names) != len(set(new_names)):
            QMessageBox.warning(self, "警告", "检测到重命名后存在重复的文件名，请修改重命名规则")
            self.execute_rename_btn.setEnabled(False)

    def on_rename_error(self, msg):
        """处理重命名错误"""
        self.rename_status_label.setText(f"错误: {msg}")
        
        # 隐藏进度条和取消按钮
        self.progress_bar.setVisible(False)
        self.cancel_btn.setVisible(False)
        
        # 启用预览按钮
        self.preview_rename_btn.setEnabled(True)
        
        QMessageBox.critical(self, "错误", f"重命名时出错: {msg}")

    def execute_rename(self):
        """执行重命名操作"""
        if not self.rename_results:
            QMessageBox.warning(self, "警告", "请先预览重命名结果")
            return
        
        reply = QMessageBox.question(self, "确认操作", 
                                     f"确定要重命名 {len(self.rename_results)} 个文件吗？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        
        # 显示进度条和取消按钮
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(len(self.rename_results))
        self.cancel_btn.setVisible(True)
        self.cancel_btn.setEnabled(True)
        
        # 禁用按钮
        self.preview_rename_btn.setEnabled(False)
        self.execute_rename_btn.setEnabled(False)
        
        # 更新状态
        self.rename_status_label.setText("正在重命名文件...")
        
        # 执行重命名
        success_count = 0
        error_count = 0
        rename_operations = []  # 记录重命名操作，用于撤回
        
        try:
            for i, (old_name, new_name) in enumerate(self.rename_results):
                # 检查取消状态
                QApplication.processEvents()
                if not self.cancel_btn.isEnabled():
                    self.rename_status_label.setText("重命名已取消")
                    break
                
                try:
                    old_path = os.path.join(self.rename_dir, old_name)
                    new_path = os.path.join(self.rename_dir, new_name)
                    
                    # 检查目标文件是否已存在
                    if os.path.exists(new_path) and old_path != new_path:
                        # 如果目标文件已存在，使用临时文件名
                        temp_path = new_path + ".temp"
                        os.rename(old_path, temp_path)
                        os.rename(temp_path, new_path)
                    else:
                        os.rename(old_path, new_path)
                    
                    # 记录重命名操作
                    rename_operations.append((new_name, old_name))
                    success_count += 1
                except Exception as e:
                    error_count += 1
                    print(f"重命名文件 {old_name} 时出错: {str(e)}")
                
                # 更新进度
                self.progress_bar.setValue(i + 1)
                self.rename_status_label.setText(f"正在重命名文件... {i+1}/{len(self.rename_results)}")
        
        finally:
            # 隐藏进度条和取消按钮
            self.progress_bar.setVisible(False)
            self.cancel_btn.setVisible(False)
            
            # 启用按钮
            self.preview_rename_btn.setEnabled(True)
            
            # 更新状态
            self.rename_status_label.setText(f"重命名完成: 成功 {success_count} 个，失败 {error_count} 个")
            
            # 更新文件列表
            self.update_rename_file_count()
            
            # 清空结果
            self.rename_results = []
            
            # 如果有成功的重命名，保存到历史记录并启用撤回按钮
            if success_count > 0:
                self.rename_history.append(rename_operations)
                self.undo_rename_btn.setEnabled(True)
                # 重新预览
                self.preview_rename()
            else:
                self.execute_rename_btn.setEnabled(False)
            
            QMessageBox.information(self, "完成", f"重命名完成: 成功 {success_count} 个，失败 {error_count} 个")

    def undo_rename(self):
        """撤回重命名操作"""
        if not self.rename_history:
            QMessageBox.warning(self, "警告", "没有可撤回的重命名操作")
            return
        
        reply = QMessageBox.question(self, "确认撤回", 
                                     f"确定要撤回最近一次的重命名操作吗？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply != QMessageBox.StandardButton.Yes:
            return
        
        # 获取最近一次的重命名操作
        last_operations = self.rename_history.pop()
        
        # 显示进度条和取消按钮
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(len(last_operations))
        self.cancel_btn.setVisible(True)
        self.cancel_btn.setEnabled(True)
        
        # 禁用按钮
        self.preview_rename_btn.setEnabled(False)
        self.execute_rename_btn.setEnabled(False)
        self.undo_rename_btn.setEnabled(False)
        
        # 更新状态
        self.rename_status_label.setText("正在撤回重命名操作...")
        
        # 执行撤回操作
        success_count = 0
        error_count = 0
        
        try:
            for i, (current_name, original_name) in enumerate(last_operations):
                # 检查取消状态
                QApplication.processEvents()
                if not self.cancel_btn.isEnabled():
                    self.rename_status_label.setText("撤回操作已取消")
                    break
                
                try:
                    current_path = os.path.join(self.rename_dir, current_name)
                    original_path = os.path.join(self.rename_dir, original_name)
                    
                    # 检查目标文件是否已存在
                    if os.path.exists(original_path) and current_path != original_path:
                        # 如果目标文件已存在，使用临时文件名
                        temp_path = original_path + ".temp"
                        os.rename(current_path, temp_path)
                        os.rename(temp_path, original_path)
                    else:
                        os.rename(current_path, original_path)
                    
                    success_count += 1
                except Exception as e:
                    error_count += 1
                    print(f"撤回重命名文件 {current_name} 时出错: {str(e)}")
                
                # 更新进度
                self.progress_bar.setValue(i + 1)
                self.rename_status_label.setText(f"正在撤回重命名操作... {i+1}/{len(last_operations)}")
        
        finally:
            # 隐藏进度条和取消按钮
            self.progress_bar.setVisible(False)
            self.cancel_btn.setVisible(False)
            
            # 启用按钮
            self.preview_rename_btn.setEnabled(True)
            
            # 如果没有更多历史记录，禁用撤回按钮
            if not self.rename_history:
                self.undo_rename_btn.setEnabled(False)
            
            # 更新状态
            self.rename_status_label.setText(f"撤回完成: 成功 {success_count} 个，失败 {error_count} 个")
            
            # 更新文件列表
            self.update_rename_file_count()
            
            # 如果有成功的撤回，重新预览
            if success_count > 0:
                self.preview_rename()
            
            QMessageBox.information(self, "完成", f"撤回完成: 成功 {success_count} 个，失败 {error_count} 个")

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
            # 设置表格列标题为源文件
            self.result_table.setHorizontalHeaderLabels(["选择", "源文件", "目标文件"])
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
            # 设置表格列标题为检索文本
            self.result_table.setHorizontalHeaderLabels(["选择", "检索文本", "目标文件"])
            # 检查文本匹配依据，决定是否显示扩大检索范围选项
            self.on_text_match_basis_changed(self.text_match_basis_combo.currentIndex())

    def on_text_match_basis_changed(self, index):
        """根据文本匹配依据的变化，更新扩大检索范围选项的可见性"""
        # 只有当选择"检索匹配"时才显示扩大检索范围选项
        if index == 1:  # 检索匹配
            self.expand_search_checkbox.show()
            # 显示标签和说明
            expand_search_row = self.text_match_layout.itemAt(3)  # 扩大检索范围选项的布局
            if expand_search_row and expand_search_row.layout():
                expand_search_row.layout().itemAt(0).widget().show()  # 显示标签
                expand_search_row.layout().itemAt(2).widget().show()  # 显示说明标签
        else:  # 完全匹配
            self.expand_search_checkbox.hide()
            # 隐藏标签和说明
            expand_search_row = self.text_match_layout.itemAt(3)  # 扩大检索范围选项的布局
            if expand_search_row and expand_search_row.layout():
                expand_search_row.layout().itemAt(0).widget().hide()  # 隐藏标签
                expand_search_row.layout().itemAt(2).widget().hide()  # 隐藏说明标签
            # 重置勾选状态
            self.expand_search_checkbox.setChecked(False)

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
            self.expand_search_checkbox.isChecked(), 
            use_multithreading
        )
        
        # 连接信号
        self.worker.result_ready.connect(self.on_match_result)
        self.worker.error.connect(self.on_match_error)
        self.worker.progress.connect(self.on_match_progress)
        self.worker.file_counts_ready.connect(self.on_file_counts_ready)
        self.worker.unmatched_ready.connect(self.on_unmatched_ready)
        
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
        
        # 计算总行数（包括未匹配的文件）
        total_rows = matched_count
        if hasattr(self, 'unmatched_source_files'):
            total_rows += len(self.unmatched_source_files)
        if hasattr(self, 'unmatched_keywords'):
            total_rows += len(self.unmatched_keywords)
        
        # 设置表格行数
        self.result_table.setRowCount(total_rows)
        
        current_row = 0
        
        # 首先显示未匹配的文件（红色字体）
        if hasattr(self, 'unmatched_source_files') and self.unmatched_source_files:
            for unmatched_file in self.unmatched_source_files:
                # 复选框
                checkbox_widget = CheckBoxWidget()
                checkbox_widget.setChecked(False)  # 未匹配的文件默认不选中
                self.result_table.setCellWidget(current_row, 0, checkbox_widget)
                
                # 源文件列（红色字体）
                source_item = QTableWidgetItem(os.path.basename(unmatched_file))
                source_item.setForeground(QColor(255, 0, 0))  # 红色字体
                self.result_table.setItem(current_row, 1, source_item)
                
                # 目标文件列（空）
                self.result_table.setItem(current_row, 2, QTableWidgetItem(""))
                
                # 存储文件数据（未匹配的源文件）
                self.matched_files.append((unmatched_file, None, os.path.basename(unmatched_file), "", ""))
                
                current_row += 1
        
        # 显示未匹配的关键字（红色字体）
        if hasattr(self, 'unmatched_keywords') and self.unmatched_keywords:
            for unmatched_keyword in self.unmatched_keywords:
                # 复选框
                checkbox_widget = CheckBoxWidget()
                checkbox_widget.setChecked(False)  # 未匹配的关键字默认不选中
                self.result_table.setCellWidget(current_row, 0, checkbox_widget)
                
                # 检索文本列（红色字体）
                keyword_item = QTableWidgetItem(unmatched_keyword)
                keyword_item.setForeground(QColor(255, 0, 0))  # 红色字体
                self.result_table.setItem(current_row, 1, keyword_item)
                
                # 目标文件列（空）
                self.result_table.setItem(current_row, 2, QTableWidgetItem(""))
                
                # 存储文件数据（未匹配的关键字）
                self.matched_files.append((None, None, unmatched_keyword, "", ""))
                
                current_row += 1
        
        # 然后显示匹配的文件
        for s, t, s_name, t_name, key in matched_pairs:
            # 复选框
            checkbox_widget = CheckBoxWidget()
            checkbox_widget.setChecked(True)
            self.result_table.setCellWidget(current_row, 0, checkbox_widget)
            
            # 源文件列或检索文本列
            if s:
                # 源目录匹配模式：显示源文件名
                source_item = QTableWidgetItem(s_name)
                self.result_table.setItem(current_row, 1, source_item)
                # 如果有匹配关键字，设置单元格部分文本颜色
                if key and key in s_name:
                    self._highlight_cell_text(current_row, 1, key)
            else:
                # 文本匹配模式：显示检索文本
                source_item = QTableWidgetItem(s_name if s_name else key)
                self.result_table.setItem(current_row, 1, source_item)
                # 如果有匹配关键字，设置单元格部分文本颜色
                if key and key in (s_name if s_name else key):
                    self._highlight_cell_text(current_row, 1, key)
            
            # 目标文件列
            target_item = QTableWidgetItem(t_name)
            self.result_table.setItem(current_row, 2, target_item)
            # 如果有匹配关键字，设置单元格部分文本颜色
            if key and key in t_name:
                self._highlight_cell_text(current_row, 2, key)
            
            # 存储文件数据
            self.matched_files.append((s, t, s_name, t_name, key))
            
            current_row += 1
        
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
                    # 检查是否为未匹配的文件（没有目标文件路径）
                    if t is None:
                        if operation == "delete":
                            # 对于未匹配的源文件，可以删除源文件
                            if s:
                                os.remove(s)
                                processed += 1
                            else:
                                # 对于未匹配的关键字，无法执行操作
                                failed += 1
                        else:
                            # 复制和移动操作需要目标文件，未匹配的文件无法操作
                            failed += 1
                        continue
                    
                    src_path = t  # 只处理目标文件
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
                    QMessageBox.critical(self, "错误", f"处理文件 {filename if 'filename' in locals() else s_name} 时出错: {str(e)}")
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
        
        # 温和地清除Excel缓存（不强制删除被占用的文件）
        try:
            if self.temp_excel_dir and os.path.exists(self.temp_excel_dir):
                # 只重置变量，不强制删除文件
                self.temp_excel_dir = None
                self.excel_file_path = None
                self.excel_data = None
        except:
            pass
        
        # 接受关闭事件
        event.accept()

    def on_unmatched_ready(self, unmatched_source_files, unmatched_keywords):
        """处理未匹配的文件和关键字"""
        self.unmatched_source_files = unmatched_source_files
        self.unmatched_keywords = unmatched_keywords

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # 使用Fusion样式，看起来更现代
    
    # 设置应用图标
    app_icon = create_app_icon()
    app.setWindowIcon(app_icon)
    
    window = FileSelector()
    window.show()
    sys.exit(app.exec()) 