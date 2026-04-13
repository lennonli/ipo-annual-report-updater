#!/usr/bin/env python3
"""
scan_old_files.py — IPO底稿旧件穿透扫描工具

扫描指定的底稿根目录，识别并分类残留的旧期文件。
输出结构化的审计报告（Markdown格式）。

用法：
    python3 scan_old_files.py <底稿根目录> [--keywords 2024,2023,申报] [--exclude-backup]
"""

import os
import sys
import time
import re
import argparse
from collections import defaultdict


# 误报白名单正则（匹配到的文件不算旧件）
FALSE_POSITIVE_PATTERNS = [
    r'202\d{10}',                     # 专利申请流水号
    r'日常申报',                       # 当期日常申报明细
    r'\d{4}年\d+月\d+日.*受理',       # 历史固定时点
]


def is_false_positive(filename):
    """检查文件名是否属于误报"""
    for pattern in FALSE_POSITIVE_PATTERNS:
        if re.search(pattern, filename):
            return True
    return False


def scan_directory(base_dir, keywords, cutoff_year=2025, exclude_backup=True):
    """
    扫描目录，返回分类后的旧件列表。
    
    Args:
        base_dir: 底稿根目录
        keywords: 用于文件名匹配的关键字列表
        cutoff_year: 文件修改时间早于此年份的视为旧件
        exclude_backup: 是否排除备份目录
    
    Returns:
        dict: {
            'critical': [...],      # 关键过时件
            'historical': [...],    # 历史实证件
            'false_positive': [...] # 误报排除
        }
    """
    results = {
        'critical': [],
        'historical': [],
        'false_positive': []
    }
    
    for root, dirs, files in os.walk(base_dir):
        # 排除备份目录
        if exclude_backup and '_备份_' in root:
            continue
        
        for f in files:
            if f == '.DS_Store' or f.startswith('~') or f.startswith('.'):
                continue
            
            full_path = os.path.join(root, f)
            rel_path = os.path.relpath(full_path, base_dir)
            
            try:
                stat = os.stat(full_path)
                mtime_year = time.localtime(stat.st_mtime).tm_year
            except:
                continue
            
            is_old = False
            reasons = []
            
            # 检查修改时间
            if mtime_year < cutoff_year:
                is_old = True
                reasons.append(f'修改时间: {mtime_year}年')
            
            # 检查文件名关键字
            matched_keywords = [k for k in keywords if k in f]
            if matched_keywords:
                is_old = True
                reasons.append(f'关键字匹配: {", ".join(matched_keywords)}')
            
            if is_old:
                entry = {
                    'path': rel_path,
                    'name': f,
                    'year': mtime_year,
                    'size': stat.st_size,
                    'reasons': reasons,
                    'dir': os.path.relpath(root, base_dir)
                }
                
                if is_false_positive(f):
                    results['false_positive'].append(entry)
                elif mtime_year < cutoff_year:
                    results['critical'].append(entry)
                else:
                    results['historical'].append(entry)
    
    return results


def generate_report(base_dir, results):
    """生成 Markdown 格式的审计报告"""
    lines = []
    lines.append(f'# 底稿旧件审计报告\n')
    lines.append(f'**扫描目录**: `{base_dir}`')
    lines.append(f'**扫描时间**: {time.strftime("%Y-%m-%d %H:%M:%S")}')
    lines.append(f'**发现旧件**: 关键{len(results["critical"])}项 | '
                 f'历史{len(results["historical"])}项 | '
                 f'误报排除{len(results["false_positive"])}项\n')
    
    # 关键过时件
    lines.append('## 1. 关键过时件（建议立即更新）\n')
    if results['critical']:
        by_dir = defaultdict(list)
        for item in results['critical']:
            by_dir[item['dir']].append(item)
        for d, items in sorted(by_dir.items()):
            lines.append(f'### 📁 {d}\n')
            for item in items:
                lines.append(f'- `{item["name"]}` ({", ".join(item["reasons"])})')
            lines.append('')
    else:
        lines.append('无关键过时件。\n')
    
    # 历史实证件
    lines.append('## 2. 历史实证件（建议保留/确认）\n')
    if results['historical']:
        by_dir = defaultdict(list)
        for item in results['historical']:
            by_dir[item['dir']].append(item)
        for d, items in sorted(by_dir.items()):
            lines.append(f'### 📁 {d}\n')
            for item in items[:5]:
                lines.append(f'- `{item["name"]}` ({", ".join(item["reasons"])})')
            if len(items) > 5:
                lines.append(f'- ... 共 {len(items)} 项')
            lines.append('')
    else:
        lines.append('无历史实证件。\n')
    
    # 误报排除
    lines.append('## 3. 误报排除\n')
    if results['false_positive']:
        lines.append(f'共 {len(results["false_positive"])} 项被识别为误报并排除。\n')
    else:
        lines.append('无误报。\n')
    
    return '\n'.join(lines)


def main():
    parser = argparse.ArgumentParser(description='IPO底稿旧件穿透扫描工具')
    parser.add_argument('base_dir', help='底稿根目录路径')
    parser.add_argument('--keywords', default='2024,2023,申报,IPO',
                        help='用于文件名匹配的关键字，逗号分隔')
    parser.add_argument('--cutoff-year', type=int, default=2025,
                        help='早于此年份的文件视为旧件（默认2025）')
    parser.add_argument('--exclude-backup', action='store_true', default=True,
                        help='排除备份目录（默认是）')
    parser.add_argument('--output', help='输出报告文件路径（默认打印到终端）')
    
    args = parser.parse_args()
    keywords = [k.strip() for k in args.keywords.split(',')]
    
    if not os.path.exists(args.base_dir):
        print(f'错误：目录不存在 {args.base_dir}')
        sys.exit(1)
    
    results = scan_directory(args.base_dir, keywords, args.cutoff_year, args.exclude_backup)
    report = generate_report(args.base_dir, results)
    
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f'报告已保存: {args.output}')
    else:
        print(report)


if __name__ == '__main__':
    main()
