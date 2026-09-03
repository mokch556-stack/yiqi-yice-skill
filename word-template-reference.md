# 一企一册 Word 文档模板参考

本文档包含生成"一企一册"实践学期学习手册所需的全部 Word (docx) 表格模板和代码模式。  
**所有个人信息和企业信息已脱敏（用 xxx 替代）。**

---

## 1. 技术栈

- **语言**: Python
- **库**: `zipfile` (标准库，直接操作 docx ZIP)
- **不使用**: `python-docx`（模板表格需直接操作 XML）

## 2. 目录结构

```
一企一册项目/
├── 【模板】电商专业实践手册_大三.docx    # 原始空模板（内含17个预置表格）
├── _tables_xml/                          # 从模板提取的表格 XML
│   ├── table_0.xml ~ table_16.xml        # 17个表格
│   └── styles_tbl.xml                    # 表格样式定义
├── gen_脚本.py                           # 生成脚本
└── 企业名_专业实践学期学习手册.docx       # 输出文档
```

## 3. 表格索引（table_0 ~ table_16）

| 索引 | 用途 | 说明 |
|------|------|------|
| 0 | 教学安排表 | 学程/项目/考核内容/周次 |
| 1 | 成绩总览表 | 实习成绩汇总 |
| 2 | 工作内容表 | 各学程工作内容描述 |
| 3 | 任务书/论文计划表 | 项目任务+论文计划 |
| 4 | 实习周记表 | 7天×3列（日期\|岗位\|日小结） |
| 5 | 企业指导教师评价表 | 企业导师评价 |
| 6 | (预留) | |
| 7 | (预留) | |
| 8 | (预留) | |
| 9 | (预留) | |
| 10 | (预留) | |
| 11 | 通用能力实践报告 | 通用能力评价 |
| 12 | 实习评价表 | 学校/企业综合评价 |
| 13 | 成绩核定表(学程1) | 第一阶段成绩 |
| 14 | 成绩核定表(学程2) | 第二阶段成绩 + 工作态度/绩效 |
| 15 | 成绩核定表(学程3) | 第三阶段成绩 + 销售/绩效 |
| 16 | KPI考核总表 | 绩效考核总分汇总 |

## 4. 核心代码模式

### 4.1 加载模板表格

```python
import zipfile, os, io

def load_tbl(idx):
    with open(f'_tables_xml/table_{idx}.xml', 'r', encoding='utf-8') as f:
        return f.read()

# 加载关键表格
T = {
    'teach': load_tbl(0),     # 教学安排
    'thesis': load_tbl(3),    # 论文计划
    'weekly': load_tbl(4),    # 周记
    'gen_ability': load_tbl(11),  # 通用能力
    'eval': load_tbl(12),     # 评价表
    'score1': load_tbl(13),   # 成绩核定1
    'score2': load_tbl(14),   # 成绩核定2
    'score3': load_tbl(15),   # 成绩核定3
    'kpi': load_tbl(16),      # KPI总表
}
```

### 4.2 自定义表格生成（无底色）

```python
def make_table(col_widths, headers, data_rows):
    """
    生成自定义表格
    col_widths: [w1, w2, ...] 各列宽度(DXA)，总和=9026
    headers: [h1, h2, ...] 表头文字
    data_rows: [[row1_col1, row1_col2, ...], ...] 数据行
    """
    xml = ('<w:tbl><w:tblPr>'
           '<w:tblStyle w:val="15"/>'
           '<w:tblW w:w="9026" w:type="dxa"/>'
           '<w:jc w:val="center"/>'
           '<w:tblBorders>'
           '<w:top w:val="single" w:color="auto" w:sz="4" w:space="0"/>'
           '<w:left w:val="single" w:color="auto" w:sz="4" w:space="0"/>'
           '<w:bottom w:val="single" w:color="auto" w:sz="4" w:space="0"/>'
           '<w:right w:val="single" w:color="auto" w:sz="4" w:space="0"/>'
           '<w:insideH w:val="single" w:color="auto" w:sz="4" w:space="0"/>'
           '<w:insideV w:val="single" w:color="auto" w:sz="4" w:space="0"/>'
           '</w:tblBorders>'
           '<w:tblLayout w:type="fixed"/>'
           '</w:tblPr>'
           '<w:tblGrid>')
    for w in col_widths:
        xml += f'<w:gridCol w:w="{w}"/>'
    xml += '</w:tblGrid>'

    # 表头
    if headers:
        xml += '<w:tr><w:trPr><w:jc w:val="center"/></w:trPr>'
        for i, h in enumerate(headers):
            xml += (f'<w:tc><w:tcPr><w:tcW w:w="{col_widths[i]}" w:type="dxa"/>'
                    f'<w:vAlign w:val="center"/></w:tcPr>'
                    f'<w:p><w:pPr><w:jc w:val="center"/>'
                    f'<w:spacing w:before="0" w:after="0" w:line="320"/></w:pPr>'
                    f'<w:r><w:rPr>'
                    f'<w:rFonts w:ascii="仿宋_GB2312" w:hAnsi="仿宋_GB2312" w:eastAsia="仿宋_GB2312"/>'
                    f'<w:b/><w:sz w:val="24"/><w:szCs w:val="24"/>'
                    f'</w:rPr><w:t>{h}</w:t></w:r></w:p></w:tc>')
        xml += '</w:tr>'

    # 数据行
    for row in data_rows:
        xml += '<w:tr><w:trPr><w:jc w:val="center"/></w:trPr>'
        for i, ct in enumerate(row):
            ct = str(ct)
            xml += (f'<w:tc><w:tcPr><w:tcW w:w="{col_widths[i]}" w:type="dxa"/>'
                    f'<w:vAlign w:val="center"/></w:tcPr>'
                    f'<w:p><w:pPr><w:jc w:val="center"/>'
                    f'<w:spacing w:before="0" w:after="0" w:line="320"/></w:pPr>'
                    f'<w:r><w:rPr>'
                    f'<w:rFonts w:ascii="仿宋_GB2312" w:hAnsi="仿宋_GB2312" w:eastAsia="仿宋_GB2312"/>'
                    f'<w:sz w:val="24"/><w:szCs w:val="24"/>'
                    f'</w:rPr><w:t xml:space="preserve">{ct}</w:t></w:r></w:p></w:tc>')
        xml += '</w:tr>'
    return xml + '</w:tbl>'
```

### 4.3 信息展示表（双列表格）

```python
def make_info_table(items):
    """items = [(label, value), ...] 生成双列信息表"""
    xml = ('<w:tbl><w:tblPr>...<w:tblGrid>'
           '<w:gridCol w:w="2200"/><w:gridCol w:w="6826"/>'
           '</w:tblGrid>')
    for label, value in items:
        xml += (f'<w:tr>...'
                f'<w:tc><w:tcW w:w="2200"...><w:t>{label}</w:t></w:tc>'
                f'<w:tc><w:tcW w:w="6826"...><w:t xml:space="preserve">{value}</w:t></w:tc>'
                f'</w:tr>')
    return xml + '</w:tbl>'
```

### 4.4 文字段落函数

```python
# 大标题（居中/左对齐）
def hp(text, level=1):
    """level=0: 封面标题44pt, level=1: 大标题36pt, level=2: 中标题32pt, level=3: 小标题30pt"""
    sz = {0:44, 1:36, 2:32, 3:30}.get(level, 28)
    al = 'center' if level == 0 else 'left'
    return (f'<w:p><w:pPr><w:jc w:val="{al}"/>'
            f'<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/></w:pPr>'
            f'<w:r><w:rPr>'
            f'<w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体"/>'
            f'<w:b/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
            f'</w:rPr><w:t>{text}</w:t></w:r></w:p>')

# 正文段落
def pa(text, indent=True, center=False):
    """首行缩进2字符（480DXA），仿宋12pt，行距固定值360"""
    ind = '<w:ind w:firstLine="480"/>' if indent else ''
    al = 'center' if center else 'left'
    return (f'<w:p><w:pPr><w:jc w:val="{al}"/>'
            f'<w:spacing w:before="60" w:after="60" w:line="360" w:lineRule="auto"/>{ind}</w:pPr>'
            f'<w:r><w:rPr>'
            f'<w:rFonts w:ascii="仿宋_GB2312" w:hAnsi="仿宋_GB2312" w:eastAsia="仿宋_GB2312"/>'
            f'<w:sz w:val="24"/><w:szCs w:val="24"/>'
            f'</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')

# 居中小标题（表格标题）
def pt(text):
    return (f'<w:p><w:pPr><w:jc w:val="center"/>'
            f'<w:spacing w:before="200" w:after="120" w:line="360" w:lineRule="auto"/></w:pPr>'
            f'<w:r><w:rPr>'
            f'<w:rFonts w:ascii="黑体" w:hAnsi="黑体" w:eastAsia="黑体"/>'
            f'<w:b/><w:sz w:val="24"/><w:szCs w:val="24"/>'
            f'</w:rPr><w:t>{text}</w:t></w:r></w:p>')

# 分页符
def bk():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

# 小标题（大纲级别2，带书签）
def sub_heading(text):
    return (f'<w:p><w:pPr><w:spacing w:before="120" w:after="60" w:line="360" w:lineRule="auto"/>'
            f'<w:outlineLvl w:val="2"/></w:pPr>'
            f'<w:bookmarkStart w:id="{bid}" w:name="_Toc{bid}"/>'
            f'<w:r><w:rPr>...黑体...<w:b/><w:sz w:val="24"/></w:rPr><w:t>{text}</w:t></w:r>'
            f'<w:bookmarkEnd w:id="{bid}"/></w:p>')
```

### 4.5 带书签和大纲级别的标题

```python
bookmark_id = [100]

def heading_with_bookmark(text, level=2):
    bid = bookmark_id[0]; bookmark_id[0] += 1
    sz = {0:44, 1:36, 2:32, 3:30}.get(level, 28)
    olvl = {0:'0', 1:'0', 2:'1', 3:'2'}.get(level, '2')
    al = 'center' if level == 0 else 'left'
    return (f'<w:p><w:pPr><w:jc w:val="{al}"/>'
            f'<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>'
            f'<w:outlineLvl w:val="{olvl}"/></w:pPr>'
            f'<w:bookmarkStart w:id="{bid}" w:name="_Toc{bid}"/>'
            f'<w:r><w:rPr>...黑体...<w:b/></w:rPr><w:t>{text}</w:t></w:r>'
            f'<w:bookmarkEnd w:id="{bid}"/></w:p>')
```

## 5. 文档结构

### 5.1 完整组装流程

```python
# 1. 收集所有段落
c = []

# 2. 封面
c.append('<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="3000"/></w:pPr></w:p>')
c.append(hp('电子商务专业实践教学手册', 0))
c.append(pa('（适用于2024级xxx岗位）', center=True))
c.append(pa('实践企业：xxx', center=True))
c.append('<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="800"/></w:pPr></w:p>')
c.append(pa('主    编：xxx', center=True))
c.append(pa('班    级：________________________', center=True))
c.append(pa('姓    名：________________________', center=True))
c.append(pa('学    号：________________________', center=True))
c.append(pa('实习单位（公章）：________________', center=True))
c.append(pa('xxx学院  xxx专业', center=True))
c.append(pa('2026年6月', center=True))

# 3. 目录（手写带页码）
c.append(bk())
c.append(heading_with_bookmark('目  录', 0))
toc_items = [('第一部分  通用部分', 1), ...]
for text, page in toc_items:
    dots = '.' * max(2, 50 - len(text))
    c.append(pa(f'{text}{dots}{page}'))

# 4. 第一部分：通用部分
c.append(bk())
c.append(heading_with_bookmark('第一部分  通用部分', 1))
# ... 4节内容

# 5. 第二部分：岗前培训
c.append(bk())
c.append(heading_with_bookmark('第二部分  岗前培训', 1))

# 6. 第三部分：专业部分
c.append(bk())
c.append(heading_with_bookmark('第三部分  专业部分', 1))
# ... 4节内容

# 7. 周记（N周）
for wn in weeks:
    c.append(pt_page(f'实习周记（{wn}）'))
    c.append(weekly_ns)

# 8. 打包写入docx
body = '\n'.join(c)
doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            ...>
<w:body>{body}
  <w:sectPr>
    <w:pgSz w:w="11906" w:h="16838"/>
    <w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440" .../>
  </w:sectPr>
</w:body></w:document>'''

# 从模板复制所有文件，替换 document.xml
with zipfile.ZipFile(tmpl_path, 'r') as z:
    files = {item.filename: z.read(item.filename) for item in z.infolist()}

files['word/document.xml'] = doc_xml.encode('utf-8')

# 写入输出
buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
    for name, data in files.items():
        zout.writestr(name, data)
with open('企业名_专业实践学期学习手册.docx', 'wb') as f:
    f.write(buf.getvalue())
```

### 5.2 分页策略

```
保留的分页符位置：
├── 封面后（目录前）
├── 目录后（第一部分前）
├── 第一部分后（第二部分前）
├── 第二部分后（第三部分前）
├── 每个周记前（独占一页）
├── 论文计划表前
├── 通用能力报告前
├── 评价表前
├── 成绩核定表前
├── 每个KPI考核表前
```

### 5.3 表格跨页控制

```python
# 周记表格加 cantSplit 防跨页
weekly_ns = T['weekly'].replace('<w:tblPr>', '<w:tblPr><w:cantSplit w:val="true"/>')

# 考核表不加（让它们自然分页）
```

## 6. 字体规范

| 用途 | 字体 | 字号 | 格式 |
|------|------|------|------|
| 封面标题 | 黑体 | 44pt(88号) | 加粗居中 |
| 大标题(第一部分) | 黑体 | 36pt(72号) | 加粗 |
| 中标题(一、二、三、四) | 黑体 | 32pt(64号) | 加粗 |
| 小标题(（一）（二）等) | 黑体 | 24pt(48号) | 加粗 |
| 正文 | 仿宋_GB2312 | 12pt(24号) | 首行缩进2字符 |
| 表格内容 | 仿宋_GB2312 | 12pt(24号) | 居中 |
| 表格表头 | 仿宋_GB2312 | 12pt(24号) | 加粗居中 |
| 目录 | 仿宋_GB2312 | 12pt(24号) | |
| 安全培训/沟通记录等 | 黑体 | 12pt(24号) | 居中加粗 |

## 7. 页面设置

```
A4纸: w:w="11906" w:h="16838" (DXA)
页边距: 上下1440(2cm)，左右1440(2cm)
内容宽度: 11906 - 1440*2 = 9026 DXA
```

## 8. 考核表内容模板（脱敏）

### 8.1 score2（学程1阶段）- 工作绩效部分

```python
# 通用考核要素模板，需要根据企业岗位定制
# 每个企业6个考核要素：
s2_titles = [
    '网络运营绩效',  # → 根据岗位改
    '产品管理',      # → 根据岗位改
    '后台设置',      # → 根据岗位改
    '库存管理',      # → 根据岗位改
    '发货管理',      # → 根据岗位改
    '客户服务',      # → 根据岗位改
]
```

### 8.2 score3（学程2阶段）- 工作绩效部分

```python
s3_titles = [
    '销售完成率',  # → 根据岗位改
    '销售增长率',  # → 根据岗位改
    '客户服务',    # → 根据岗位改
    '特殊成果',    # → 根据岗位改
]
```

### 8.3 各岗位考核要素参考

| 岗位方向 | score2要素 | score3要素 |
|----------|------------|------------|
| 电商开发/运营 | 代码开发进度/代码质量/功能测试/文档规范/协作沟通/项目交付 | 系统架构设计/项目开发效率/技术攻关/创新与成果 |
| 电商内容制作/短视频 | 短视频产出量/视频剪辑质量/拍摄执行/素材管理/创意策划/团队协作 | 内容运营效果/直播执行能力/数据分析能力/创新成果 |
| 仓储设备运维 | 设备巡检完成率/故障诊断能力/维修操作规范/备件管理/设备保养执行/服务响应 | 维修效率/电气控制技能/客户现场服务/技术改进 |
| AI短视频自动化 | AI视频产出量/AI工具运用/自动化流程搭建/素材资源库/账号运营/直播协助 | 内容传播效果/数据驱动优化/跨平台运营/创新成果 |

## 9. 目录生成策略

**重要：不要使用TOC域代码**（Word版本兼容性问题会导致更新失败）。  
改为**手写带页码目录**，格式为：

```
第一部分  通用部分........................................1
  一、企业概况..........................................1
  二、安全规范与应急处理..................................2
  ...
```

关键点：
- 标题文字 + 点号填充 + 页码
- 点号数量 = max(2, 50 - 标题长度)
- 每个标题加`<w:outlineLvl>`和`<w:bookmarkStart>`为目录识别
