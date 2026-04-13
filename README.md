# ipo-annual-report-updater

IPO/北交所上市项目年度报告更新底稿迁移与管理工具。

## 安装

```bash
npx skills add lennonli/ipo-annual-report-updater
```

# IPO 年度报告更新底稿迁移工具

## 概述

本 skill 用于 IPO/北交所上市项目中，当年度报告更新时，系统性地将全套底稿材料从旧报告期（如2024年半年度）迁移至新报告期（如2025年度）。这是一个高度结构化的工作流程，涉及数百个文件夹和上千份文件的精准操作。

## 核心工作流程

### Phase 1: 现状梳理（必做）

在进行任何修改之前，必须先做全景扫描，输出结构化的审计报告。

```
步骤：
1. 扫描目标底稿根目录，递归列出所有含文件的子目录
2. 对每个文件，通过以下维度判断"新旧"：
   - 文件修改时间戳（os.stat().st_mtime）
   - 文件名关键字匹配（如：2024、2023、申报、IPO、半年度等）
3. 按章节分组汇总，输出审计报告
4. 将旧件分类为：
   a. 关键过时件（需立即更新） — 如标注了旧年度的确认函
   b. 历史实证件（建议保留） — 如专利流水号、跨期合同
   c. 当期件误报（排除）     — 如"日常申报明细表"实为当期数据
```

**关键脚本模板**（见 `scripts/scan_old_files.py`）

> [!IMPORTANT]
> 必须排除备份目录（含 `_备份_` 的路径）和系统文件（`.DS_Store`）。
> 仅靠文件名关键字或时间戳判定是不够的——必须结合两者，并对误报进行人工复核。

### Phase 2: 文件同步与替换

核心操作模式：**从权威源目录向多个目标目录"点对点"同步**。

```
典型的源→目标映射关系：

申报底稿(2-年报更新申报底稿)
  ├── 6-1-1 股东名册          → 3-股东核查/1-1, 4-1, 5-4, 6-1
  ├── 6-1-2 调查问卷/穿透表   → 3-股东核查/6-2, 4-离职人员/1-2, 1-3
  ├── 9-2-2 关联方访谈签署版  → 1-问询反馈/3-3-5
  ├── 9-2-3 关联交易凭证      → 1-问询反馈/3-6-3
  └── 23-1  网络核查截图       → 4-离职人员/1-8
```

**同步操作的标准流程**：
```python
# 伪代码
def sync_files(source_dir, target_dir, delete_old=True):
    if delete_old:
        clear_all_files(target_dir)  # 删除旧件
    copy_all_files(source_dir, target_dir)  # 复制新件
```

> [!WARNING]
> **安全第一**：所有删除操作前，确认用户已有备份（如 `_备份_20260413` 目录）。
> 绝不主动创建或删除备份目录，除非用户明确指示。

### Phase 3: 目录状态标注

为已完成审核的目录添加状态标识，这是一个极好的审计实践。

**命名规范**（来自实战经验）：
```
目录状态标注格式：
  【无更新】    → 已确认该章节无需更新，旧件已清理
  【参见XXX】   → 底稿证据在其他章节，标明交叉引用路径

示例：
  3-2-4【无更新】取得了发那特和发行人关于工艺流程的说明...
  12-3-2【参见2025年度报告新增底稿14-1-4】发行人关于更换主办...
```

**批量重命名脚本模板**：
```python
import os

def batch_rename_dirs(base_path, old_text, new_text):
    """递归重命名目录中包含指定文本的子文件夹"""
    for root, dirs, files in os.walk(base_path):
        for d in dirs:
            if old_text in d:
                old_path = os.path.join(root, d)
                new_name = d.replace(old_text, new_text)
                new_path = os.path.join(root, new_name)
                if not os.path.exists(new_path):
                    os.rename(old_path, new_path)
                    print(f'Renamed: {d} -> {new_name}')
```

### Phase 4: DOCX 内容批量修改（核心难点）

这是整个工作流中技术含量最高的环节。Word文档(.docx)本质上是ZIP包，内含XML文件。

> [!CAUTION]
> **血泪教训**：不要用 `zipfile` 直接操作 XML 文本替换！
> Word 会将一个完整的中文句子拆分成多个 `<w:t>` 节点（XML runs），
> 导致搜索"补充法律意见书（二）"时在 XML 层面根本匹配不到，因为它可能被拆成了：
> `<w:t>补充法律意见书</w:t><w:t>（</w:t><w:t>二</w:t><w:t>）</w:t>`
>
> **必须使用 `python-docx` 库**，它会将同一段落内的所有 runs 合并为完整的 `.text` 属性。

**正确的替换方法**：
```python
import docx

def robust_docx_replace(filepath, replacements):
    """
    使用 python-docx 对 DOCX 文件进行鲁棒的文本替换。
    同时处理段落和表格中的文本。
    
    Args:
        filepath: DOCX 文件路径
        replacements: dict，{旧文本: 新文本}
    """
    doc = docx.Document(filepath)
    modified = False
    
    def process_paragraphs(paragraphs):
        nonlocal modified
        for para in paragraphs:
            text = para.text
            original = text
            for old, new in replacements.items():
                if old in text:
                    text = text.replace(old, new)
            if text != original:
                # 保存首个 run 的格式属性
                rPr_xml = None
                if para.runs:
                    for r in para.runs:
                        if len(r.text.strip()) > 0 and r._r.rPr is not None:
                            rPr_xml = r._r.rPr
                            break
                # 替换文本（会清除所有 runs，合并为单个 run）
                para.text = text
                # 重新应用格式
                if rPr_xml is not None and para.runs:
                    para.runs[0]._r.append(rPr_xml)
                modified = True
    
    # 处理正文段落
    process_paragraphs(doc.paragraphs)
    
    # 处理表格中的段落（极易遗漏！）
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs)
    
    if modified:
        doc.save(filepath)
        return True
    return False
```

> [!IMPORTANT]
> **多轮替换策略**（实战经验总结）：
>
> 第一轮替换很容易遗漏细微变体。必须分多轮执行，每轮使用不同的关键字集合：
>
> **第一轮（显式版本号）**：
> - `补充法律意见书（二）` → `（三）`
> - `2024年上半年度新增底稿` → `2025年度报告新增底稿`
> - `（2025年半年度更新版本）` → `（2025年度更新版本）`
>
> **第二轮（隐式上下文）**：
> - `加审期间` → `本次年度更新期间`
> - `2024年半年报底稿` → `2025年报更新申报底稿`
> - `2025年半年报更新` → `2025年度报告更新`
>
> **第三轮（表格内用语）**：
> - `加审期内` → `2025年度`（常见于底稿目录表格的项目描述中）
>
> **每轮结束后必须用校验脚本扫描确认**。

**校验脚本模板**：
```python
def verify_docx_clean(base_dir, forbidden_terms):
    """验证目录下所有 DOCX 文件不含禁止词"""
    import docx, os
    issues = []
    for f in os.listdir(base_dir):
        if f.endswith('.docx') and not f.startswith(('~', '.')):
            doc = docx.Document(os.path.join(base_dir, f))
            for p in doc.paragraphs:
                for term in forbidden_terms:
                    if term in p.text:
                        issues.append((f, 'Paragraph', p.text[:80]))
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            for term in forbidden_terms:
                                if term in p.text:
                                    issues.append((f, 'Table', p.text[:80]))
    return issues
```

### Phase 5: 残留审计与收尾

最后一步：用全量审计脚本做收官扫描。

```python
# 审计维度
audit_keywords = ['2024', '2023', '申报', 'IPO', '半年度', '半年报', '加审']

# 需要排除的误报模式
false_positive_patterns = [
    r'202\d{10}',       # 专利申请流水号（如 202410752436X）
    r'日常申报',        # "日常申报明细表"是当期数据
    r'（一）|（二）',   # 作为列举序号而非版本号（如"（二）查验事项"）
    r'\d{4}年\d+月\d+日.*受理',  # 历史固定时间点（如"2024年12月26日受理"）
]
```

## 典型的 IPO 底稿目录结构

```
项目根目录/
├── 1-问询反馈底稿/           # 交易所问询函的逐项回复证据
│   ├── 3-关于同业竞争/
│   ├── 12-其他问题/
│   └── ...
├── 2-年报更新申报底稿/       # 【权威源】最新年度的全套申报材料
│   ├── 6-发起人、股东及实际控制人/
│   ├── 9-关联交易及同业竞争/
│   ├── 14-发行人股东大会.../
│   ├── 23-附卷/
│   └── ...
├── 3-股东信息披露专项核查/    # 股东穿透核查、名册、承诺函
├── 4-证监会离职人员核查/      # 证监系统离职人员网络比对
├── 5-内核文件/                # 律所内核委员会审批文件
│   ├── 1-说明.docx
│   ├── 2-律师工作情况.docx
│   ├── 3-1 查验计划.docx
│   ├── 3-2 底稿目录.docx
│   ├── 4-1 股东信息查验计划.docx
│   ├── 4-2 股东信息底稿目录.docx
│   ├── 5-1 离职人员查验计划.docx
│   └── 5-2 离职人员底稿目录.docx
└── _备份_YYYYMMDD/            # 操作前的完整备份
```

## 操作安全规范

1. **备份确认**：执行任何删除操作前，先确认项目根目录下存在 `_备份_` 目录。
2. **逐步推进**：每完成一个章节的更新，立即运行审计脚本验证，不要一次性处理所有章节。
3. **交叉引用一致性**：当更新 `1-问询反馈底稿` 中的证据时，必须确保其来源（通常是 `2-年报更新申报底稿`）中的对应文件确实存在。
4. **目录重命名安全**：重命名目录前检查目标路径是否已存在，防止覆盖。
5. **DOCX 修改后校验**：每次修改 DOCX 后，用校验脚本确认无遗漏；建议用 Word 打开抽检排版是否完好。

## 常见陷阱与教训

| 陷阱 | 表现 | 解决方案 |
|------|------|----------|
| XML runs 拆分 | zipfile 直替无效，长中文句子匹配不到 | 必须用 python-docx |
| 表格内文本遗漏 | 段落已清洁，但表格单元格里的旧文本未改 | 必须遍历 `doc.tables` |
| 隐式旧文本 | "加审期间"、"半年报底稿" 等不含年份数字 | 多轮替换+禁止词校验 |
| 临时文件干扰 | `~$file.docx` 或 `.~file.docx` 被误识别 | 跳过 `~` 和 `.` 开头的文件 |
| 目录重复 | 同一编号的文件夹在根级和子级各有一套 | 用 `os.walk` 全量遍历，两套都处理 |
| 误报干扰 | 专利流水号含"2024"被标为旧件 | 建立误报白名单 |
| 版本号 vs 序号 | "（二）查验事项" 不应被改为"（三）" | 仅替换明确的版本标识 |

## 依赖

- Python 3.x（内置 `os`, `shutil`, `zipfile`, `time`）
- `python-docx`（`pip install python-docx`）— DOCX 内容修改必需

## License

MIT

## 📱 关注作者

如果这个项目对你有帮助，欢迎关注我获取更多技术分享：

- **X (Twitter)**: [@vista8](https://x.com/vista8)
- **微信公众号「向阳乔木推荐看」**:

<p align="center">
  <img src="https://github.com/joeseesun/terminal-boost/raw/main/assets/wechat-qr.jpg?raw=true" alt="向阳乔木推荐看公众号二维码" width="300">
</p>
