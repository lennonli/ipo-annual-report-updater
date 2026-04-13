#!/usr/bin/env python3
"""
sync_files.py — 底稿文件同步工具

从权威源目录向多个目标目录"点对点"同步文件。
支持：清除旧件后复制新件、仅复制新件、仅清除旧件。

用法：
    python3 sync_files.py --source <源目录> --targets <目标目录1>,<目标目录2> [--clear-old]
    python3 sync_files.py --clear-only <目标目录>
"""

import os
import sys
import shutil
import argparse


def clear_directory(target_dir, preserve_patterns=None):
    """
    清空目录中的所有文件和子目录。
    
    Args:
        target_dir: 目标目录
        preserve_patterns: 需要保留的文件名模式列表（如 ['笔录', '.docx']）
    """
    if not os.path.exists(target_dir):
        print(f'警告：目录不存在 {target_dir}')
        return 0
    
    removed = 0
    for item in os.listdir(target_dir):
        if item == '.DS_Store':
            continue
        
        # 检查是否需要保留
        if preserve_patterns:
            should_preserve = any(p in item for p in preserve_patterns)
            if should_preserve:
                print(f'  保留: {item}')
                continue
        
        item_path = os.path.join(target_dir, item)
        if os.path.isfile(item_path):
            os.remove(item_path)
            removed += 1
        elif os.path.isdir(item_path):
            shutil.rmtree(item_path)
            removed += 1
    
    return removed


def sync_directory(source_dir, target_dir, clear_old=True, preserve_patterns=None):
    """
    将源目录中的文件同步到目标目录。
    
    Args:
        source_dir: 源目录（权威数据来源）
        target_dir: 目标目录
        clear_old: 是否先清除旧件
        preserve_patterns: 清除时需要保留的文件名模式
    
    Returns:
        tuple: (清除数量, 复制数量)
    """
    if not os.path.exists(source_dir):
        print(f'错误：源目录不存在 {source_dir}')
        return (0, 0)
    
    if not os.path.exists(target_dir):
        print(f'错误：目标目录不存在 {target_dir}')
        return (0, 0)
    
    removed = 0
    copied = 0
    
    if clear_old:
        removed = clear_directory(target_dir, preserve_patterns)
    
    for f in os.listdir(source_dir):
        if f == '.DS_Store':
            continue
        src = os.path.join(source_dir, f)
        if os.path.isfile(src):
            shutil.copy2(src, target_dir)
            copied += 1
    
    return (removed, copied)


def batch_rename_dirs(base_path, old_text, new_text):
    """
    递归重命名目录中包含指定文本的子文件夹。
    
    Args:
        base_path: 根目录
        old_text: 需要替换的文本
        new_text: 替换为的新文本
    
    Returns:
        list: 被重命名的目录列表
    """
    renamed = []
    for root, dirs, files in os.walk(base_path):
        for d in dirs:
            if old_text in d:
                old_path = os.path.join(root, d)
                new_name = d.replace(old_text, new_text)
                new_path = os.path.join(root, new_name)
                if not os.path.exists(new_path):
                    os.rename(old_path, new_path)
                    renamed.append((d, new_name))
                else:
                    print(f'跳过（目标已存在）: {new_name}')
    return renamed


def main():
    parser = argparse.ArgumentParser(description='底稿文件同步工具')
    subparsers = parser.add_subparsers(dest='command')
    
    # 同步命令
    sync_parser = subparsers.add_parser('sync', help='从源目录同步到目标目录')
    sync_parser.add_argument('--source', required=True, help='源目录')
    sync_parser.add_argument('--targets', required=True, help='目标目录（逗号分隔）')
    sync_parser.add_argument('--clear-old', action='store_true', default=True, help='先清除旧件')
    sync_parser.add_argument('--preserve', help='需要保留的文件名模式（逗号分隔）')
    
    # 清理命令
    clear_parser = subparsers.add_parser('clear', help='清空目标目录')
    clear_parser.add_argument('--targets', required=True, help='目标目录（逗号分隔）')
    
    # 重命名命令
    rename_parser = subparsers.add_parser('rename', help='批量重命名子目录')
    rename_parser.add_argument('--base', required=True, help='根目录')
    rename_parser.add_argument('--old', required=True, help='旧文本')
    rename_parser.add_argument('--new', required=True, help='新文本')
    
    args = parser.parse_args()
    
    if args.command == 'sync':
        targets = [t.strip() for t in args.targets.split(',')]
        preserve = [p.strip() for p in args.preserve.split(',')] if args.preserve else None
        
        for target in targets:
            removed, copied = sync_directory(args.source, target, args.clear_old, preserve)
            print(f'✅ {os.path.basename(target)}: 清除{removed}项, 复制{copied}项')
    
    elif args.command == 'clear':
        targets = [t.strip() for t in args.targets.split(',')]
        for target in targets:
            removed = clear_directory(target)
            print(f'✅ {os.path.basename(target)}: 清除{removed}项')
    
    elif args.command == 'rename':
        renamed = batch_rename_dirs(args.base, args.old, args.new)
        for old_name, new_name in renamed:
            print(f'✅ {old_name} → {new_name}')
        print(f'\n共重命名 {len(renamed)} 个目录')
    
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
