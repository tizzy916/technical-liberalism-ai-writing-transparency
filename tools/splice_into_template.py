#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
splice_into_template.py
=======================
把 chapters/*.md 的内容嫁接进 v3.docx 骨架，保留所有节属性、页眉页脚、页码配置。

策略（字符串级操作，避免 ElementTree 把命名空间前缀重命名为 ns0/ns1/...）：
1. 解包 v3.docx → 拿到 word/document.xml 全文 + 所有 header/footer XML
2. 字符串切分：保留 <?xml ...> 头 + <w:document ...> 开标签 / </w:body> 结束 / </w:document> 结束
3. body 内部按 <w:sectPr> 切分为 N 个 section（N=18，本论文）
4. 按 SEQUENCE 列表逐项输出新节：
   - kind="keep"    保留 v3 该节内容（封面、TOC 等）
   - kind="render"  用 pandoc 渲染 markdown，并嫁接 sectpr_from 指定节的 sectPr
5. 自动检测：若新章节 H1 与所克隆 sectPr 的 header 文字不同，
   则复制 header.xml、改写章标题、注册新 _rels 与 Content_Types
6. 全局重新编号 bookmark id 避免冲突
7. 拼接 → 写回 document.xml + _rels + Content_Types → 重新打包

★ 增 / 删 / 重排章节：编辑下方的 SEQUENCE 列表即可，无需手工改 reference.docx。

用法：
    python3 splice_into_template.py
    python3 splice_into_template.py -o ../_deliverables/word_exports/current/提交版.docx
    python3 splice_into_template.py --keep-tmp        # 保留临时目录用于排查
"""

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
PAPER_DIR = HERE.parent
CHAPTERS_DIR = PAPER_DIR / "paper" / "chapters"
REFERENCE_DOCX = HERE / "reference.docx"
DEFAULT_OUTPUT = PAPER_DIR / "_deliverables" / "word_exports" / "current" / "技术自由主义_latest_提交版.docx"

# =====================================================================
# 首行缩进后处理：给所有正文段（继承 Normal 或 Body Text 等样式的）
# 强制 inline 注入 <w:ind w:firstLine="480"/>，跳过 heading / 目录 / 标题
# 等不应缩进的段落。
#
# 为什么需要这个？
# - 中文论文要求每段首行缩进 2 字符（中文五号字下 = 480 twips）
# - reference.docx 的 Normal 样式 pPr 已带 firstLine=480
# - 但 pandoc 给段落贴 pStyle="FirstParagraph"（悬空引用）或 pStyle="19"
#   (Body Text)，Word 渲染时这些样式可能不会正确继承 Normal 的 firstLine
# - v3.docx 是靠每段 inline 写 <w:ind w:firstLine="480"/> 来保证视觉缩进的
#   这是 Word 在编辑过程中自动产生的 inline 覆盖
# - 本函数复刻这个 inline 注入行为，确保导出的 docx 视觉一致
# =====================================================================

# 这些样式的段落不注入 firstLine
# - heading 1~9（styleId 3-11）+ 中文 heading 别名（168, 169）
# - 目录各级（toc 1-9，主要是 26, 29，其余按基线扩展）
# - Title (34) / Subtitle (27) / TOC Heading (167)
# - caption (15) / footer (24) / header (25)
# - No Spacing (141)
SKIP_STYLES_FOR_INDENT = {
    "3", "4", "5", "6", "7", "8", "9", "10", "11",
    "168", "169",
    "26", "29",
    "27", "34", "167",
    "15", "24", "25", "141",
    "Title", "Subtitle",
}


def inject_first_line_indent(doc_xml: str) -> str:
    """
    给 document.xml 中所有正文段注入 inline <w:ind w:firstLine="480"/>。
    跳过：heading / 目录 / 标题 / 表格内段 / 已含 firstLine 的段 / 空段。
    返回修改后的 XML 字符串。
    """
    # 找出所有表格区域
    table_ranges = [
        (m.start(), m.end())
        for m in re.finditer(r"<w:tbl\b.*?</w:tbl>", doc_xml, re.DOTALL)
    ]

    def in_table(pos):
        return any(s <= pos < e for s, e in table_ranges)

    counters = {"mod": 0, "skip": 0}

    def process_p(m):
        full = m.group(0)
        pos = m.start()
        # 表格内不动
        if in_table(pos):
            counters["skip"] += 1
            return full
        # 已含 firstLine 不动（v3 keep 段已有 inline ind）
        if "w:firstLine" in full:
            counters["skip"] += 1
            return full
        # 跳过名单内的样式
        style_m = re.search(r'<w:pStyle w:val="([^"]+)"', full)
        style = style_m.group(1) if style_m else None
        if style and style in SKIP_STYLES_FOR_INDENT:
            counters["skip"] += 1
            return full
        # 空段（无文本 run）不动
        if not re.search(r"<w:t[^>]*>[^<]+</w:t>", full):
            counters["skip"] += 1
            return full
        # 注入 ind
        if "<w:pPr>" in full:
            new_full = re.sub(
                r"(<w:pPr>(?:(?!</w:pPr>).)*?)(</w:pPr>)",
                r'\1<w:ind w:firstLine="480"/>\2',
                full,
                count=1,
                flags=re.DOTALL,
            )
        else:
            new_full = re.sub(
                r"(<w:p\b[^>]*?>)",
                r'\1<w:pPr><w:ind w:firstLine="480"/></w:pPr>',
                full,
                count=1,
            )
        if new_full != full:
            counters["mod"] += 1
        else:
            counters["skip"] += 1
        return new_full

    new_xml = re.sub(
        r"<w:p\b[^>]*?>.*?</w:p>", process_p, doc_xml, flags=re.DOTALL
    )
    print(
        f"  [首行缩进] 注入 {counters['mod']} 段 / 跳过 {counters['skip']} 段"
        " (heading / 目录 / 表格 / 空段 / 已有 ind)"
    )
    return new_xml

# =====================================================================
# SEQUENCE：输出 docx 的节序列。改这一份就能增 / 删 / 重排章节。
# =====================================================================
#
# 列表的每一项 = 输出 docx 的一个 section。按出现顺序排列。
#
# 每项是 dict，字段：
#   kind: "keep"   保留 v3 该节的全部内容（封面、目录、评阅人等不动的页）
#         "render" 用 markdown 渲染该节内容
#   file: 章节 markdown 文件名，仅 render 时需要
#   h1:   该 markdown 内要取的一级标题文本（None = 整个文件）
#   sectpr_from: 节属性来自 v3 的哪一节（即"复制谁的页眉/页脚/页码"）
#                - keep 模式：必须等于 v3 节序号自身
#                - render 模式：通常等于 v3 中对应位置的节序号
#                - 新增章节：用一个已有"正文章节"节序号即可（如本论文用 7）
#
# ★ 增加新章节：
#     1. 在 chapters/ 下创建新的 .md 文件
#     2. 在 SEQUENCE 里复制一行 render 项，改 file 名；sectpr_from 用一个
#        正文章节的序号（比如 7 = v3 第1章），新章节就会自动获得"运行页眉 +
#        阿拉伯页码"等正文章节的节属性。
#
# ★ 删除章节：直接删 / 注释掉对应的 SEQUENCE 行即可，剩下的章节自动续编页码。
#
# ★ 重排章节：调整 SEQUENCE 顺序即可。
#
# 注意：splice 模式下，封面（v3 §0/§1）和评阅人（v3 §2）使用 kind="keep" 直接复用
# v3 中的内容（这些页含复杂的多列表格，markdown 难以无损表达）。如果需要改封面信息，
# 直接在 reference.docx 里改，或者用 build.py（它会读取 chapters/00_前置.md 渲染）。
SEQUENCE = [
    {"kind": "keep",   "sectpr_from": 0},                                                    # 中文封面（直接用 v3，不读 00_前置.md）
    {"kind": "keep",   "sectpr_from": 1},                                                    # 英文封面（同上）
    {"kind": "keep",   "sectpr_from": 2},                                                    # 公开评阅人 / 授权说明（同上）
    {"kind": "render", "file": "00_摘要.md", "h1": "摘　要",                "sectpr_from": 3},  # 中文摘要
    {"kind": "render", "file": "00_摘要.md", "h1": "Abstract",              "sectpr_from": 4},  # 英文摘要
    {"kind": "keep",   "sectpr_from": 5},                                                    # 目录（提交前在 Word 里"更新域"）
    {"kind": "render", "file": "01_技术自由主义的当代危机.md",                "sectpr_from": 6},  # 第1章
    {"kind": "render", "file": "02_信息何以自由.md",                          "sectpr_from": 7},  # 第2章
    {"kind": "render", "file": "03_审美革命及其技术问题.md",                  "sectpr_from": 8},  # 第3章
    {"kind": "render", "file": "04_技艺的分裂与重逢.md",                       "sectpr_from": 9},  # 第4章
    {"kind": "render", "file": "05_技艺自由主义.md",                          "sectpr_from": 10}, # 第5章
    {"kind": "render", "file": "06_结语_选择走出去.md",                       "sectpr_from": 11}, # 第6章
    {"kind": "render", "file": "07_参考文献.md",                              "sectpr_from": 12}, # 参考文献
    {"kind": "render", "file": "08_后记与声明.md", "h1": "致　谢",          "sectpr_from": 13},
    {"kind": "render", "file": "08_后记与声明.md", "h1": "声　明",          "sectpr_from": 14},
    {"kind": "render", "file": "08_后记与声明.md", "h1": "个人简历、在学期间完成的相关学术成果", "sectpr_from": 15},
    {"kind": "render", "file": "08_后记与声明.md", "h1": "指导小组评语",    "sectpr_from": 16},
    {"kind": "render", "file": "08_后记与声明.md", "h1": "答辩委员会决议书", "sectpr_from": 17},
]


# =====================================================================
# Markdown 提取
# =====================================================================

def strip_yaml_frontmatter(text: str) -> str:
    if text.startswith("---\n"):
        end = text.find("\n---\n", 4)
        if end != -1:
            return text[end + 5:]
    return text


def extract_section_md(full_md: str, h1_marker):
    if h1_marker is None:
        return full_md
    lines = full_md.split("\n")
    start = None
    end = len(lines)
    for i, line in enumerate(lines):
        if re.match(rf"^# {re.escape(h1_marker)}\s*$", line):
            start = i
            for j in range(i + 1, len(lines)):
                if re.match(r"^# ", lines[j]):
                    end = j
                    break
            break
    if start is None:
        raise ValueError(f"H1 marker {h1_marker!r} 未在 markdown 中找到")
    return "\n".join(lines[start:end]).rstrip() + "\n"


# =====================================================================
# 字符串级 XML 操作工具
# =====================================================================

PARA_RE = re.compile(r'<w:p\b[^>]*>.*?</w:p>|<w:p\b[^>]*/>', re.DOTALL)
TBL_RE = re.compile(r'<w:tbl\b[^>]*>.*?</w:tbl>', re.DOTALL)
SECTPR_IN_PPR_RE = re.compile(r'<w:sectPr\b[^>]*>.*?</w:sectPr>|<w:sectPr\b[^>]*/>', re.DOTALL)


def split_body_into_sections(body_xml: str):
    """把 <w:body> 内的字符串切成 N 个 section（list of strings）。
    分割点：包含 <w:sectPr ...> 的段落（其后归到下一节）；
    body 末尾的独立 <w:sectPr> 视作最后一节末尾。
    """
    # 我们用顺序扫描：找到所有顶层 <w:p>…</w:p> / <w:tbl>…</w:tbl> / <w:sectPr>…</w:sectPr>
    pos = 0
    elems = []  # list of (start, end, kind, text)
    pat = re.compile(
        r'(<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>'
        r'|<w:tbl\b[^>]*?>.*?</w:tbl>'
        r'|<w:sectPr\b[^>]*?>.*?</w:sectPr>|<w:sectPr\b[^>]*?/>)',
        re.DOTALL)
    for m in pat.finditer(body_xml):
        s, e = m.span()
        text = m.group()
        if text.startswith("<w:p"):
            kind = "p"
        elif text.startswith("<w:tbl"):
            kind = "tbl"
        else:
            kind = "sectPr"
        elems.append((s, e, kind, text))

    # 按节切：遇到 p 中含 sectPr 或 顶层 sectPr，就闭合当前节
    sections = []
    cur = []
    for (s, e, kind, text) in elems:
        cur.append(text)
        if kind == "p" and "<w:sectPr" in text:
            sections.append(cur)
            cur = []
        elif kind == "sectPr":
            sections.append(cur)
            cur = []
    if cur:
        sections.append(cur)
    return sections


def extract_sectpr(section_elems):
    """从 section 末尾摘出 sectPr 字符串（可能在 pPr 内或独立顶层）。"""
    if not section_elems:
        return None
    last = section_elems[-1]
    if last.startswith("<w:sectPr"):
        return last
    if last.startswith("<w:p"):
        m = SECTPR_IN_PPR_RE.search(last)
        if m:
            return m.group()
    return None


def attach_sectpr_to_last_paragraph(new_elems, sectpr_xml):
    """把 sectPr 注入到 new_elems 最后一个段落的 <w:pPr> 末尾。
    若该段落无 pPr，则插入。返回修改后的 new_elems。"""
    # 找到最后一个段落
    idx = None
    for i in range(len(new_elems) - 1, -1, -1):
        if new_elems[i].startswith("<w:p"):
            idx = i
            break
    if idx is None:
        # 没有段落，造一个空段落
        new_elems.append(f"<w:p><w:pPr>{sectpr_xml}</w:pPr></w:p>")
        return new_elems

    last_p = new_elems[idx]
    # 移除已存在的 sectPr
    last_p = SECTPR_IN_PPR_RE.sub("", last_p)
    # 注入到 pPr 中
    if "<w:pPr>" in last_p or "<w:pPr/>" in last_p:
        # 在 </w:pPr> 之前插入；或把 <w:pPr/> 展开
        last_p = last_p.replace("<w:pPr/>", f"<w:pPr>{sectpr_xml}</w:pPr>")
        last_p = last_p.replace("</w:pPr>", f"{sectpr_xml}</w:pPr>", 1)
    else:
        # 在 <w:p ...> 之后插入 <w:pPr>...</w:pPr>
        last_p = re.sub(r'(<w:p\b[^>]*>)', rf'\1<w:pPr>{sectpr_xml}</w:pPr>', last_p, count=1)
    new_elems[idx] = last_p
    return new_elems


# =====================================================================
# Pandoc 渲染单个 markdown 片段
# =====================================================================

def pandoc_render(md_text: str, work_dir: Path) -> list:
    md_path = work_dir / "frag.md"
    docx_path = work_dir / "frag.docx"
    md_path.write_text(md_text, encoding="utf-8")
    cmd = [
        "pandoc",
        # 用 pandoc 自身的 markdown 而非 gfm —— 后者在 CJK 字符两侧无法正确识别
        # **加粗**/*斜体* 边界（CommonMark 严格要求 ** 后跟空白或标点才能闭合，
        # 而中文标点不算 CommonMark 标点），换 markdown 即可。
        "-f", "markdown+east_asian_line_breaks-smart",
        "-t", "docx",
        "--reference-doc", str(REFERENCE_DOCX),
        str(md_path),
        "-o", str(docx_path),
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"pandoc 失败:\n{res.stderr}")
    with zipfile.ZipFile(docx_path) as z:
        doc_xml = z.read("word/document.xml").decode("utf-8")

    # 抓出 body 内部
    m = re.search(r'<w:body\b[^>]*>(.*)</w:body>', doc_xml, re.DOTALL)
    if not m:
        raise RuntimeError("pandoc 输出的 document.xml 无 <w:body>")
    body_xml = m.group(1)

    # 切出顶层段落/表，去掉末尾的 sectPr
    pat = re.compile(
        r'(<w:p\b[^>]*?/>|<w:p\b[^>]*?>.*?</w:p>'
        r'|<w:tbl\b[^>]*?>.*?</w:tbl>)',
        re.DOTALL)
    elems = [m.group() for m in pat.finditer(body_xml)]
    # 去掉每个段落里 pandoc 加的 sectPr（如果有）
    elems = [SECTPR_IN_PPR_RE.sub("", e) for e in elems]
    return elems


# =====================================================================
# Header / Relationship / Content_Types 工具
# =====================================================================
#
# 当我们克隆一个 v3 节的 sectPr 给新章节时，sectPr 里 headerReference
# 指向的 header XML 文件中硬编码着旧章节的标题文字（如"第5章 …"）。
# 为了让新章节页眉显示新标题，需要：
#   1. 复制原 header.xml → 新 headerN.xml
#   2. 把 headerN.xml 里的旧标题文本替换为新章节 H1 文本
#   3. 在 _rels/document.xml.rels 里注册新 header（分配新 rId）
#   4. 在 [Content_Types].xml 里加 Override 条目
#   5. 把克隆的 sectPr 里 headerReference 的 rId 改为新 rId
# 这一切对用户透明：用户只要在 SEQUENCE 里加一条新 chapter 即可。

def get_header_text(header_xml: str) -> str:
    """提取 header XML 中所有 <w:t> 的拼接文本（去全角空格）。"""
    texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', header_xml)
    return "".join(texts).replace("　", "").strip()


def clone_header_with_new_title(orig_header_xml: str, new_title: str) -> str:
    """复制 header XML，把其中的 <w:t> 文本替换为新标题。
    保留所有格式（rPr 字体、边框、对齐等）。"""
    # 找到第一个有非空文本的 <w:r>...<w:t>OldText</w:t>...</w:r>
    # 直接把所有 <w:t>...</w:t> 替换成 <w:t>new_title</w:t>，
    # 把第二个起的 <w:r>…</w:r> 删掉（多个 run 通常是分散的格式片段）
    # 简化做法：替换第一个 <w:t> 内的内容，删除其余 <w:t> 的内容
    new_xml = orig_header_xml
    matches = list(re.finditer(r'<w:t[^>]*>([^<]*)</w:t>', new_xml))
    if not matches:
        return new_xml
    # 用占位字符做两步替换避免冲突
    PLACEHOLDER_FIRST = "\x00FIRST\x00"
    PLACEHOLDER_REST = "\x00EMPTY\x00"
    # 先把第一个保留并替换，其余清空
    parts = []
    last_end = 0
    for i, m in enumerate(matches):
        parts.append(new_xml[last_end:m.start()])
        # 保留原标签属性
        tag_open = re.search(r'<w:t[^>]*>', m.group()).group()
        if i == 0:
            parts.append(f"{tag_open}{PLACEHOLDER_FIRST}</w:t>")
        else:
            parts.append(f"{tag_open}</w:t>")
        last_end = m.end()
    parts.append(new_xml[last_end:])
    new_xml = "".join(parts)
    # 转义 XML 特殊字符
    safe_title = (new_title.replace("&", "&amp;")
                           .replace("<", "&lt;")
                           .replace(">", "&gt;"))
    new_xml = new_xml.replace(PLACEHOLDER_FIRST, safe_title)
    return new_xml


class HeaderManager:
    """管理 v3 解包目录下的 header 文件 / 关系 / 类型注册。"""
    def __init__(self, v3_unpack: Path):
        self.v3_unpack = v3_unpack
        self.rels_path = v3_unpack / "word" / "_rels" / "document.xml.rels"
        self.ct_path = v3_unpack / "[Content_Types].xml"
        self.rels = self.rels_path.read_text(encoding="utf-8")
        self.ct = self.ct_path.read_text(encoding="utf-8")
        # 找出现有 rId 最大值
        self._max_rid = max(
            int(m.group(1)) for m in re.finditer(r'Id="rId(\d+)"', self.rels)
        )
        # 找出现有 headerN.xml 最大编号
        self._max_header = max(
            (int(m.group(1)) for m in re.finditer(r'header(\d+)\.xml', self.rels)),
            default=0
        )

    def get_header_filename_by_rid(self, rid: str):
        """rId8 → header2.xml （若不存在 / 不是 header，返回 None）"""
        # 注意：用 [^>]*? 而非 [^/]*，因为 Type URL 里含斜杠
        m = re.search(rf'<Relationship\s+Id="{rid}"[^>]*?Target="(header[^"]+)"', self.rels)
        return m.group(1) if m else None

    def add_header(self, content_xml: str) -> tuple[str, str]:
        """新增一个 header XML，返回 (新 rId, 新 filename)。"""
        self._max_rid += 1
        self._max_header += 1
        new_rid = f"rId{self._max_rid}"
        new_fname = f"header{self._max_header}.xml"
        # 写文件
        (self.v3_unpack / "word" / new_fname).write_text(content_xml, encoding="utf-8")
        # 加 _rels 关系
        self.rels = self.rels.replace(
            "</Relationships>",
            f'<Relationship Id="{new_rid}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
            f'Target="{new_fname}"/></Relationships>'
        )
        # 加 [Content_Types] override
        ct_entry = (f'<Override PartName="/word/{new_fname}" '
                    f'ContentType="application/vnd.openxmlformats-officedocument.'
                    f'wordprocessingml.header+xml"/>')
        if ct_entry not in self.ct:
            self.ct = self.ct.replace("</Types>", ct_entry + "</Types>")
        return new_rid, new_fname

    def flush(self):
        """把更新后的 _rels 和 Content_Types 写回。"""
        self.rels_path.write_text(self.rels, encoding="utf-8")
        self.ct_path.write_text(self.ct, encoding="utf-8")


def update_sectpr_header_ref(sectpr_xml: str, new_rid: str, header_type: str = "default") -> str:
    """把 sectPr 里 type=default 的 headerReference 的 r:id 改为 new_rid。
    若没有该 type 的 headerReference 就追加。"""
    pattern = rf'(<w:headerReference[^/]*?w:type="{header_type}"[^/]*?r:id=")[^"]+(")'
    if re.search(pattern, sectpr_xml):
        return re.sub(pattern, rf'\g<1>{new_rid}\g<2>', sectpr_xml)
    # type 在 r:id 之前的情况
    pattern2 = rf'(<w:headerReference[^/]*?r:id=")[^"]+("[^/]*?w:type="{header_type}")'
    if re.search(pattern2, sectpr_xml):
        return re.sub(pattern2, rf'\g<1>{new_rid}\g<2>', sectpr_xml)
    # 没有就插一个（在 sectPr 开始处）
    new_ref = f'<w:headerReference r:id="{new_rid}" w:type="{header_type}"/>'
    return sectpr_xml.replace("<w:sectPr>", f"<w:sectPr>{new_ref}", 1)


def get_header_rid_from_sectpr(sectpr_xml: str, header_type: str = "default"):
    """从 sectPr 里取出指定 type 的 headerReference 的 rId。"""
    m = re.search(
        rf'<w:headerReference[^/]*?w:type="{header_type}"[^/]*?r:id="([^"]+)"',
        sectpr_xml
    )
    if m:
        return m.group(1)
    m = re.search(
        rf'<w:headerReference[^/]*?r:id="([^"]+)"[^/]*?w:type="{header_type}"',
        sectpr_xml
    )
    return m.group(1) if m else None


def extract_first_h1(md_text: str):
    """从 markdown 文本里提取第一个 H1 标题文本（去掉 # 和首尾空白）。"""
    for line in md_text.split("\n"):
        m = re.match(r'^# (.+)$', line)
        if m:
            return m.group(1).strip()
    return None


# =====================================================================
# 全局 ID 重编号（避免 bookmark / comment / footnote 引用冲突）
# =====================================================================

def renumber_ids(xml_str: str, id_offset: int) -> str:
    """对 bookmarkStart/bookmarkEnd 等含 w:id 的元素加偏移量。
    覆盖 pandoc 自带的 id（通常 < 100），避免与 v3 中已有的 id 冲突。
    """
    # 只处理 bookmarkStart / bookmarkEnd（重要：影响目录链接）；其它一般不冲突
    def repl(m):
        return f'{m.group(1)}{int(m.group(2)) + id_offset}{m.group(3)}'
    pattern = re.compile(r'(<w:(?:bookmarkStart|bookmarkEnd)\b[^>]*?w:id=")(\d+)(")')
    return pattern.sub(repl, xml_str)


# =====================================================================
# 主流程
# =====================================================================

def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-o", "--output", type=Path, default=DEFAULT_OUTPUT)
    ap.add_argument("--keep-tmp", action="store_true")
    args = ap.parse_args()

    if shutil.which("pandoc") is None:
        sys.exit("错误：未找到 pandoc。请先 brew install pandoc")
    if not REFERENCE_DOCX.exists():
        sys.exit(f"错误：缺少 {REFERENCE_DOCX}")

    tmp_dir = Path(tempfile.mkdtemp(prefix="splice_"))
    print(f"[临时目录] {tmp_dir}")

    try:
        # 1. 解包 v3
        v3_unpack = tmp_dir / "v3"
        with zipfile.ZipFile(REFERENCE_DOCX) as z:
            z.extractall(v3_unpack)

        # 2. 读 document.xml
        doc_xml_path = v3_unpack / "word" / "document.xml"
        doc_xml = doc_xml_path.read_text(encoding="utf-8")

        # 3. 切出 body 内容（保留 prefix / suffix 不动）
        m = re.search(r'(.*?<w:body\b[^>]*>)(.*)(</w:body>.*)', doc_xml, re.DOTALL)
        if not m:
            sys.exit("v3 document.xml 没有 <w:body>")
        prefix, body_xml, suffix = m.group(1), m.group(2), m.group(3)

        v3_sections = split_body_into_sections(body_xml)
        print(f"[骨架] v3 共 {len(v3_sections)} 个 section；本次输出 {len(SEQUENCE)} 个 section")

        # HeaderManager 用于按需克隆运行页眉文件
        header_mgr = HeaderManager(v3_unpack)

        md_cache = {}
        def get_md(name):
            if name not in md_cache:
                content = (CHAPTERS_DIR / name).read_text(encoding="utf-8")
                md_cache[name] = strip_yaml_frontmatter(content)
            return md_cache[name]

        # 4. 按 SEQUENCE 顺序生成新节
        new_sections = []
        ID_BASE = 10000  # bookmark id 偏移；每节再加 1000 步长

        for out_idx, entry in enumerate(SEQUENCE):
            kind = entry.get("kind")
            sectpr_from = entry.get("sectpr_from")

            # —— 校验 sectpr_from 合法 ——
            if sectpr_from is None or not (0 <= sectpr_from < len(v3_sections)):
                sys.exit(f"\n错误：SEQUENCE[{out_idx}] 的 sectpr_from={sectpr_from!r} 越界。"
                         f"v3 只有 {len(v3_sections)} 个 section（合法范围 0..{len(v3_sections)-1}）。")

            # —— 取出 sectPr 模板（每节都需要这个）——
            sectpr_template = extract_sectpr(v3_sections[sectpr_from])
            if sectpr_template is None:
                # v3 最后一节的 sectPr 在 body 末尾而非 pPr 内；做兜底
                print(f"  [out§{out_idx:2d}] [警告] v3 §{sectpr_from} 没有 sectPr，"
                      f"输出节将缺失节属性")

            if kind == "keep":
                # 直接使用 v3 该节内容
                # （sectpr_from 应该等于该节本身；若不同则相当于复制别的节属性，
                # 但内容仍来自 sectpr_from 那一节）
                section_content = v3_sections[sectpr_from]
                n_p = sum(1 for e in section_content if e.startswith("<w:p"))
                print(f"  [out§{out_idx:2d}] keep v3 §{sectpr_from} ({n_p} paras)")
                new_sections.append(section_content)
                continue

            if kind != "render":
                sys.exit(f"\n错误：SEQUENCE[{out_idx}] 的 kind={kind!r} 不识别。"
                         f"合法值：'keep' 或 'render'。")

            # —— render 模式 ——
            chapter_file = entry.get("file")
            h1_marker = entry.get("h1")  # 可为 None
            if not chapter_file:
                sys.exit(f"\n错误：SEQUENCE[{out_idx}] kind=render 但缺 file 字段。")

            md_path = CHAPTERS_DIR / chapter_file
            if not md_path.exists():
                sys.exit(f"\n错误：SEQUENCE[{out_idx}] 指向 {chapter_file}，"
                         f"但该文件不存在（{md_path}）。\n"
                         f"修正：在 chapters/ 下创建该文件，"
                         f"或调整 SEQUENCE 该项的 file 字段。")

            md_full = get_md(chapter_file)
            try:
                md_chunk = extract_section_md(md_full, h1_marker)
            except ValueError:
                actual_h1s = [m.group(1) for m in re.finditer(r'^# (.+)$', md_full, re.MULTILINE)]
                sys.exit(f"\n错误：SEQUENCE[{out_idx}]：在 {chapter_file} 中找不到 H1 标题 {h1_marker!r}。\n"
                         f"  该文件实际存在的 H1 标题：\n"
                         + "\n".join(f"    - {h!r}" for h in actual_h1s)
                         + "\n  修正：检查全角空格 / 标点是否一致，更新 SEQUENCE 该项的 h1 字段。")

            chunk_dir = tmp_dir / f"chunk_{out_idx:02d}"
            chunk_dir.mkdir()
            new_elems = pandoc_render(md_chunk, chunk_dir)

            # bookmark id 重编号，避免段间冲突
            offset = ID_BASE + out_idx * 1000
            new_elems = [renumber_ids(e, offset) for e in new_elems]

            # —— 处理运行页眉文字 ——
            # 取出 sectpr_template 的 default header 文件，看里面文字是否与本节 H1 一致；
            # 不一致就克隆一份新 header.xml 并改写文本
            sectpr_to_use = sectpr_template
            header_note = ""
            if sectpr_template is not None:
                rid = get_header_rid_from_sectpr(sectpr_template, "default")
                if rid:
                    hdr_fname = header_mgr.get_header_filename_by_rid(rid)
                    if hdr_fname:
                        hdr_path = v3_unpack / "word" / hdr_fname
                        orig_hdr = hdr_path.read_text(encoding="utf-8")
                        orig_title = get_header_text(orig_hdr)
                        new_h1 = extract_first_h1(md_chunk)
                        # 把"第N章"前后的空白 / 全角空格归一后比较
                        norm = lambda s: re.sub(r'[\s　]+', '', s or '')
                        if new_h1 and norm(new_h1) != norm(orig_title):
                            # 克隆 header，改写标题
                            new_hdr_xml = clone_header_with_new_title(orig_hdr, new_h1)
                            new_rid, new_fname = header_mgr.add_header(new_hdr_xml)
                            sectpr_to_use = update_sectpr_header_ref(
                                sectpr_template, new_rid, "default")
                            header_note = f" [新 header→{new_fname}: {new_h1[:25]}…]"
                new_elems = attach_sectpr_to_last_paragraph(new_elems, sectpr_to_use)

            label = f"H1={h1_marker!r}" if h1_marker else ""
            print(f"  [out§{out_idx:2d}] {chapter_file:30s} {label:30s} sectpr←v3§{sectpr_from}  ({len(new_elems)} elems){header_note}")
            new_sections.append(new_elems)

        # 5. 拼回 body
        new_body = "".join("".join(sec) for sec in new_sections)
        new_doc_xml = prefix + new_body + suffix

        # 5.0 中文论文首行缩进后处理：给所有正文段强制注入 inline ind
        new_doc_xml = inject_first_line_indent(new_doc_xml)

        doc_xml_path.write_text(new_doc_xml, encoding="utf-8")

        # 5.1 把 HeaderManager 累计的 _rels / Content_Types 写回
        header_mgr.flush()

        # 6. 重新打包（先写到临时文件，再原子替换；否则在某些只读 FUSE mount 上 unlink 会失败）
        args.output.parent.mkdir(parents=True, exist_ok=True)
        tmp_output = tmp_dir / "output.docx"
        with zipfile.ZipFile(tmp_output, "w", zipfile.ZIP_DEFLATED) as z:
            for path in v3_unpack.rglob("*"):
                if path.is_file():
                    arcname = path.relative_to(v3_unpack).as_posix()
                    z.write(path, arcname)

        # 把临时文件移动到最终位置（覆盖）
        shutil.copyfile(tmp_output, args.output)

        print(f"\n[完成] 输出 {args.output}")
        print(f"      大小 {args.output.stat().st_size / 1024:.1f} KB")

    finally:
        if not args.keep_tmp:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        else:
            print(f"[保留临时目录] {tmp_dir}")


if __name__ == "__main__":
    main()
