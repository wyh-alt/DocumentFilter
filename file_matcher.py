#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import re
from difflib import SequenceMatcher

class FileMatcher:
    """文件匹配器类，提供多种文件匹配策略"""
    
    @staticmethod
    def match_by_number_prefix(source_files, target_files, min_similarity=0.5):
        """基于数字前缀匹配文件
        例如：12345-原唱.mp3 匹配 12345-伴奏.mp3
        """
        matched_pairs = []
        
        # 创建目标文件的查找索引
        target_dict = {}
        for target_file in target_files:
            # 提取数字前缀
            match = re.search(r'^(\d+)', os.path.basename(target_file))
            if match:
                prefix = match.group(1)
                if prefix not in target_dict:
                    target_dict[prefix] = []
                target_dict[prefix].append(target_file)
        
        # 查找匹配的文件
        for source_file in source_files:
            match = re.search(r'^(\d+)', os.path.basename(source_file))
            if match:
                prefix = match.group(1)
                if prefix in target_dict:
                    for target_file in target_dict[prefix]:
                        matched_pairs.append((source_file, target_file))
        
        return matched_pairs
    
    @staticmethod
    def match_by_similarity(source_files, target_files, min_similarity=0.5):
        """基于文件名相似度匹配文件"""
        matched_pairs = []
        
        for source_file in source_files:
            source_name = os.path.splitext(os.path.basename(source_file))[0]
            best_match = None
            best_score = min_similarity
            
            for target_file in target_files:
                target_name = os.path.splitext(os.path.basename(target_file))[0]
                
                # 计算相似度
                similarity = SequenceMatcher(None, source_name, target_name).ratio()
                
                if similarity > best_score:
                    best_score = similarity
                    best_match = target_file
            
            if best_match:
                matched_pairs.append((source_file, best_match))
        
        return matched_pairs
    
    @staticmethod
    def match_by_pattern_replacement(source_files, target_files, pattern_pairs):
        """基于模式替换匹配文件
        例如：将"原唱"替换为"伴奏"进行匹配
        """
        matched_pairs = []
        
        # 创建目标文件查找索引
        target_dict = {os.path.basename(f): f for f in target_files}
        
        for source_file in source_files:
            source_basename = os.path.basename(source_file)
            
            for old_pattern, new_pattern in pattern_pairs:
                if old_pattern in source_basename:
                    potential_target = source_basename.replace(old_pattern, new_pattern)
                    if potential_target in target_dict:
                        matched_pairs.append((source_file, target_dict[potential_target]))
        
        return matched_pairs
    
    @staticmethod
    def match_by_custom_regex(source_files, target_files, source_pattern, target_pattern):
        """使用自定义正则表达式匹配文件"""
        matched_pairs = []
        
        source_regex = re.compile(source_pattern)
        target_dict = {}
        
        # 为目标文件创建索引
        for target_file in target_files:
            target_basename = os.path.basename(target_file)
            match = re.search(target_pattern, target_basename)
            if match and match.groups():
                key = match.group(1)  # 使用第一个捕获组作为键
                target_dict[key] = target_file
        
        # 查找匹配的源文件
        for source_file in source_files:
            source_basename = os.path.basename(source_file)
            match = re.search(source_pattern, source_basename)
            if match and match.groups():
                key = match.group(1)  # 使用第一个捕获组作为键
                if key in target_dict:
                    matched_pairs.append((source_file, target_dict[key]))
        
        return matched_pairs 