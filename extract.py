#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 .doc 试卷提取题目为结构化 JSON（含图片、答案、解析）。

一次性运行：
    python3 extract.py [过滤关键词]

方案：
    以 Microsoft Word 转换的 docx 为权威来源（段落=题目文本，表格=共用材料），
    少数 docx 丢失的题（EMBED 公式题）用 textutil 提取的 .doc 文本补齐。
依赖：
    - 本机安装 Microsoft Word（用于 .doc -> .docx 转换）
"""
import json
import os
import re
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC_DIR = os.path.dirname(BASE_DIR)          # 行测专项练习/
DATA_DIR = os.path.join(BASE_DIR, "data")
IMG_DIR = os.path.join(DATA_DIR, "images")
TMP_DIR = os.path.join(BASE_DIR, "tmp_docx")

W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

PART_RE = re.compile(r"^第\s*[一二三四五六]\s*部分\s*\S+")
SECTION_RE = re.compile(r"^[一二三四五六七八九十]+、\S+")
ANSWER_HEAD_RE = re.compile(r"^参考答案(及解析)?")
QNUM_RE = re.compile(r"^(\d{1,3})[.．、]\s*(.*)")
ANS_QNUM_RE = re.compile(r"^(\d{1,3})[.．、]\s*([ABCD])(?![A-Z])")
DISCARD_LINES = re.compile(
    r"^(请|请看|请继续|请开始|请作答|请认真|解答|例题|共\s*\d+\s*题"
    r"|根据题目要求|注意：假设|本部分共|下列各题每题所给出的四个选项中)")


def esc(s):
    return s.replace("\\", "\\\\").replace('"', '\\"')


def list_docs():
    return sorted(f for f in os.listdir(SRC_DIR)
                  if f.lower().endswith(".doc") and not f.startswith("~"))


def doc_to_docx():
    os.makedirs(TMP_DIR, exist_ok=True)
    docs = list_docs()
    pairs = []
    script = []
    for f in docs:
        docx = os.path.join(TMP_DIR, f + ".docx")
        if os.path.exists(docx):
            pairs.append((os.path.join(SRC_DIR, f), docx))
            continue
        script.append(
            f'open POSIX file "{esc(os.path.join(SRC_DIR, f))}"\n'
            f'set d to active document\n'
            f'save as d file name POSIX file "{esc(docx)}" file format format document\n'
            f'close d saving no\n')
    if script:
        apple = ('tell application "Microsoft Word"\n' + "\n".join(script) + 'end tell\n')
        subprocess.run(["osascript", "-e", apple], check=True)
        for f in docs:
            docx = os.path.join(TMP_DIR, f + ".docx")
            if os.path.exists(docx):
                pairs.append((os.path.join(SRC_DIR, f), docx))
    return pairs


def read_docx_flow(docx_path):
    """解析 docx，返回顺序流：
    {'kind':'text','text':...} 段落文本
    {'kind':'table','text':..., 'images':[...]} 表格
    {'kind':'image','file':...} 段落内浮动图片
    """
    z = zipfile.ZipFile(docx_path)
    root = ET.fromstring(z.read("word/document.xml"))
    rels = ET.fromstring(z.read("word/_rels/document.xml.rels"))
    rid2file = {}
    for rel in rels:
        if rel.tag.endswith("Relationship"):
            rid = rel.get("Id")
            tgt = rel.get("Target")
            if "media" in tgt:
                rid2file[rid] = os.path.basename(tgt)
    body = root.find("w:body", W_NS)
    flow = []
    for el in body:
        if el.tag == end("w:p"):
            flow.extend(walk_paragraph(el, rid2file))
        elif el.tag == end("w:tbl"):
            flow.append(walk_table(el, rid2file))
    return flow


def end(tag):
    return "{%s}%s" % (W_NS["w"], tag.split(":")[1])


def _iter_pieces(el, rid2file):
    pieces = []
    pos = 0
    W_NS_VAL = "{%s}val" % W_NS["w"]
    RPR = end("w:rPr")
    VALIGN = end("w:vertAlign")
    cur_va = None
    for child in el.iter():
        tag = child.tag
        if tag == end("w:r"):
            va = None
            rpr = child.find(RPR)
            if rpr is not None:
                vaEl = rpr.find(VALIGN)
                if vaEl is not None:
                    va = vaEl.get(W_NS_VAL)
            cur_va = va
        elif tag == end("w:t"):
            if child.text:
                txt = child.text
                if cur_va == "superscript":
                    txt = "\u27e6sup\u27e7" + txt + "\u27e6/sup\u27e7"
                elif cur_va == "subscript":
                    txt = "\u27e6sub\u27e7" + txt + "\u27e6/sub\u27e7"
                pieces.append((pos, "text", txt))
            pos += 1
        elif tag in ("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline",
                     "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor"):
            for b in child.iter():
                if b.tag.endswith("}blip"):
                    rid = b.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    if rid and rid in rid2file:
                        pieces.append((pos, "image", rid2file[rid]))
            pos += 1
        elif tag == end("w:pict"):
            for v in child.iter():
                if v.tag.endswith("}imagedata"):
                    rid = v.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    if rid and rid in rid2file:
                        pieces.append((pos, "image", rid2file[rid]))
            pos += 1
        elif tag == end("w:object"):
            for v in child.iter():
                if v.tag.endswith("}imagedata"):
                    rid = v.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                    if rid and rid in rid2file:
                        pieces.append((pos, "image", rid2file[rid]))
            pos += 1
    pieces.sort(key=lambda x: x[0])
    return pieces


def walk_paragraph(p, rid2file):
    out = []
    buf = ""
    for _, kind, val in _iter_pieces(p, rid2file):
        if kind == "text":
            buf += val
        else:
            if buf:
                out.append({"kind": "text", "text": buf})
                buf = ""
            out.append({"kind": "image", "file": val})
    if buf:
        out.append({"kind": "text", "text": buf})
    return out


def walk_table(tbl, rid2file):
    rows = []
    imgs = []
    for row in tbl.iter(end("w:tr")):
        r = []
        for cell in row.iter(end("w:tc")):
            txt = "".join(t.text or "" for t in cell.iter(end("w:t")) if t.text)
            txt = normalize(txt)
            for b in cell.iter():
                if b.tag.endswith("}blip"):
                    rid = b.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    if rid and rid in rid2file:
                        imgs.append(rid2file[rid])
            if txt:
                r.append(txt)
        if r:
            rows.append(r)
    if not rows:
        return None
    return {"kind": "table", "rows": rows, "images": imgs}


def normalize(s):
    s = s.replace("\u3000", " ").replace("\xa0", " ").replace("\u2002", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = s.replace("\u2028", " ").replace("\u2029", " ")
    return s.strip()


def is_table_number(line):
    """过滤非题号行：纯数字/纯小数（页码、表格数值），或「数字.数字 + 汉字单位」的材料数据。
    数列题（如 26.125，16，3，1）和年份题号（如 42.2006 年）不会被误伤。
    """
    return bool(re.match(r"^\d+$", line) or
                re.match(r"^\d{1,3}\.\d+$", line) or
                re.match(r"^\d{1,3}\.\d{1,3}\s+[\u4e00-\u9fff]", line))


ANSWER_LINE_RE = re.compile(r"^\d{1,3}[.．、]\s*[ABCDEFGH]\s*[［【(（]?\s*解析")


def norm_part(s):
    """部分标题归一化（去空格），保证题目区/答案区一致。"""
    return re.sub(r"\s+", "", s)


def split_segments(flow):
    """顺序流 -> [ {name, lines:[...], tables:[...]} ]
    name: '题目:第X部分...' / '答案:第X部分...'
    lines 元素：{'t': text} 或 {'img': 文件名}
    """
    segments = []
    cur = None
    in_answer = False
    seen_titles = set()   # 题目区出现过的部分标题，重复出现视为答案区
    ans_streak = 0        # 连续答案解析行计数
    for item in flow:
        if item["kind"] == "table":
            if cur is not None:
                for im in item.get("images", []):
                    cur["lines"].append({"img": im})
                cur["tables"].append(item["rows"])
            continue
        if item["kind"] == "image":
            if cur is not None:
                cur["lines"].append({"img": item["file"]})
            continue
        t = normalize(item["text"])
        if not t:
            continue
        if not in_answer and ANSWER_HEAD_RE.match(t):
            in_answer = True
            cur = None
            continue
        m = PART_RE.match(t)
        if m:
            if not in_answer and t in seen_titles:
                in_answer = True   # 重复的部分标题 = 答案区开始
            if not in_answer:
                seen_titles.add(t)
            name = ("答案" if in_answer else "题目") + ":" + t
            cur = {"name": name, "lines": [], "tables": []}
            segments.append(cur)
            continue
        # 连续答案解析行 = 答案区开始（如 (十) 无参考答案标题）
        if ANSWER_LINE_RE.match(t):
            ans_streak += 1
            if ans_streak >= 3 and not in_answer:
                in_answer = True
                cur = {"name": "答案:", "lines": [], "tables": []}
                segments.append(cur)
            if cur is not None:
                cur["lines"].append({"t": t})
            continue
        ans_streak = 0
        if cur is None:
            cur = {"name": ("答案" if in_answer else "题目") + ":", "lines": [], "tables": []}
            segments.append(cur)
        cur["lines"].append({"t": t})
    return segments


def split_options(raw):
    """题干+选项混合文本 -> (stem, [option...])。选项标记可跨行/行内。"""
    raw = re.sub(r"\s+", " ", raw).strip()
    marks = []
    expected = 'A'
    for m in re.finditer(r"[ABCDEFGH][.．、]", raw):
        L = m.group(0)[0]
        if L == expected:
            marks.append((L, m.start()))
            expected = chr(ord(L) + 1)
    if len(marks) < 2:
        return raw, []
    stem = raw[:marks[0][1]].strip()
    opts = []
    for i, (L, p) in enumerate(marks):
        end_p = marks[i + 1][1] if i + 1 < len(marks) else len(raw)
        seg = raw[p + 2:end_p].strip()
        opts.append(L + ". " + seg)
    return stem, opts


_last_no = 0


def parse_questions(segments, source, images_map, formula_map=None):
    global _last_no
    questions = []
    formula_map = formula_map or {}
    for seg in segments:
        if not seg["name"].startswith("题目"):
            continue
        part = norm_part(seg["name"].split(":", 1)[1])
        section = ""
        materials = list(seg["tables"])
        cur_q = None
        raw_lines = []
        pending_images = []   # 首题前的图表，共享给本部分内后续所有题

        def finish():
            nonlocal cur_q, raw_lines
            if cur_q is not None:
                stem, opts = split_options("\n".join(raw_lines))
                cur_q["stem"] = stem
                if opts:
                    cur_q["options"] = opts
                questions.append(cur_q)
            cur_q = None
            raw_lines = []

        def settle_pending():
            """无题号的图片（如文档中图形推理的无主图）在题型结束时落账为图片题。"""
            global _last_no
            nonlocal pending_images, cur_q
            if not pending_images or cur_q is not None:
                return
            for img in pending_images:
                _last_no += 1
                questions.append({
                    "id": None, "part": part, "section": section, "stem": "",
                    "options": [], "images": [img], "answer": "",
                    "explanation": "", "source": source, "no": _last_no})
            pending_images = []

        for item in seg["lines"]:
            if "img" in item:
                if item["img"] in formula_map:
                    # 公式图 → 嵌入当前题干/选项文本
                    mark = "\u27e6img:%s\u27e7" % formula_map[item["img"]]
                    if cur_q is not None:
                        if raw_lines:
                            raw_lines[-1] = raw_lines[-1] + mark
                        else:
                            raw_lines.append(mark)
                    continue
                img_path = images_map.get(item["img"])
                if not img_path:
                    continue   # wmf 公式/符号，忽略
                if cur_q is not None:
                    cur_q["images"].append(img_path)
                else:
                    pending_images.append(img_path)   # 资料图表 / 无主题图
                continue
            line = item["t"]
            if SECTION_RE.match(line) and not line[:1].isdigit():
                finish()
                # 资料类 SECTION（根据…回答…题）的图表是材料，保留给后续题；
                # 其他（如图形推理）的无主图则结算为图片题
                if "根据" not in line:
                    settle_pending()
                section = line.split("。")[0]
                continue
            if is_table_number(line):
                continue   # 纯数字/小数行（PDF 页码、图表数据），非题号
            if DISCARD_LINES.match(line) or "【例题" in line:
                if "例题" in line or "解答" in line:
                    pending_images = []   # 例题图不属于任何题
                continue
            if "【答案】" in line:
                # 内联答案（单文档中题后直接跟答案）
                if cur_q is not None:
                    am = re.search(r"【答案】\s*([ABCDEFGH]+)", line)
                    if am:
                        cur_q["answer"] = am.group(1)
                        em = re.search(r"(?:【解析】|。解析|解析[:：])\s*(.*)", line)
                        if em:
                            cur_q["explanation"] = em.group(1).lstrip("：:。 ") 
                continue
            m = QNUM_RE.match(line)
            if m:
                n = int(m.group(1))
                if n <= _last_no:
                    continue
                finish()
                _last_no = n
                cur_q = {"id": None, "part": part, "section": section, "stem": "",
                         "options": [], "images": [], "answer": "", "explanation": "",
                         "source": source, "no": n}
                if materials:
                    cur_q["material"] = list(materials)
                if pending_images:
                    cur_q["images"] = list(pending_images)   # 图表共享给后续题
                if m.group(2):
                    raw_lines.append(m.group(2))
                continue
            if cur_q is not None:
                raw_lines.append(line)
        finish()
        if "资料分析" not in part:
            settle_pending()   # 部分结束时结算无主图（资料分析部分除外）
    return questions


def parse_answer_map(segments):
    """返回 (answer_map, seq_lines)。
    answer_map: {(part, no): {'ans':.., 'expl':..}}  带题号的答案
    seq_lines:  无题号答案区的原始行（按顺序对应全部题目）
    """
    answer_map = {}
    seq_lines = []
    for seg in segments:
        if not seg["name"].startswith("答案"):
            continue
        part = norm_part(seg["name"].split(":", 1)[1])
        for item in seg["lines"]:
            if "img" in item:
                continue
            line = item["t"]
            m = re.search(r"(\d{1,3})[.．、]\s*(?:【答案】\s*)?([ABCDEFGH]+)", line)
            if part and m:
                # 带题号答案：行内可能有多题（如 "1.B 2.B 3.C" 或 "1.【答案】B。解析：…"）
                for mm in re.finditer(r"(\d{1,3})[.．、]\s*(?:【答案】\s*)?([ABCDEFGH]+)", line):
                    no = int(mm.group(1))
                    ans = mm.group(2)
                    # 该题解析 = 从本题答案后到下一个题号之间
                    rest = line[mm.end():]
                    nm = re.search(r"\d{1,3}[.．、]", rest)
                    if nm:
                        rest = rest[:nm.start()]
                    expl = ""
                    em = re.search(r"【解析】\s*(.*)", rest)
                    if em:
                        expl = em.group(1)
                    elif rest.strip():
                        expl = rest.strip()
                    entry = answer_map.setdefault((part, no), {})
                    if "ans" not in entry:
                        entry.update({"ans": ans, "expl": expl})
            elif not part:
                # 顺序答案：字母开头，"数字.【答案】B"，"数字.字母"（可多题同行），或范围 "1-10：DBCDC"
                m2 = re.match(r"^(\d+)[.．、]\s*(?:【答案】\s*)?([ABCDEFGH])", line)
                if m2:
                    all_m = list(re.finditer(
                        r"(\d+)[.．、]\s*(?:【答案】\s*)?([ABCDEFGH])", line))
                    if len(all_m) == 1:
                        ans = all_m[0].group(2)
                        rest = line[all_m[0].end():].strip()
                        em = re.search(r"【解析】\s*(.*)", rest)
                        expl = em.group(1) if em else re.sub(r"^[。．：:\s]+", "", rest)
                        expl = re.sub(r"^解析[：:]?\s*", "", expl)
                        seq_lines.append(ans + ((" " + expl) if expl else ""))
                    else:
                        for mm in all_m:
                            seq_lines.append(mm.group(2))
                elif re.match(r"^[ABCDEFGH]", line):
                    seq_lines.append(line)
                else:
                    rm = re.match(r"^(\d+)-(\d+)[:：]\s*([A-Za-z\s]+)", line)
                    if rm:
                        letters = re.sub(r"\s+", "", rm.group(3))
                        for ch in letters:
                            if ch.upper() in "ABCDEFGH":
                                seq_lines.append(ch.upper())
    return answer_map, seq_lines


def extract_png_from_doc(doc_bin):
    """从 .doc 二进制按顺序提取完整 PNG 数据列表。"""
    pngs = []
    pos = 0
    while True:
        i = doc_bin.find(b"\x89PNG", pos)
        if i < 0:
            break
        e = doc_bin.find(b"\x00\x00\x00\x00IEND", i)
        if e < 0:
            break
        pngs.append(doc_bin[i:e + 12])
        pos = e + 12
    return pngs


def wmf_to_png(data, dest, width=150):
    """用 wmf2svg + rsvg-convert 把 WMF 公式图转成 PNG。
    wmf2svg 会把 MathType 特殊符号映射为 Symbol 字体字符，但 macOS 渲染缺失导致乱码，
    转换时删除这些 Symbol 字体的 text 元素（分数/符号本身是矢量绘制，不受影响）。
    """
    tmp = os.path.join(TMP_DIR, "formula_%d.wmf" % os.getpid())
    with open(tmp, "wb") as f:
        f.write(data)
    svg = tmp + ".svg"
    try:
        r1 = subprocess.run(["wmf2svg", "-o", svg, tmp], capture_output=True)
        if r1.returncode != 0:
            return False
        raw = open(svg, "rb").read().decode("latin-1")
        # 删除 Symbol 字体中含非 ASCII 的乱码 text（MathType 特殊符号映射错误）；
        # 保留 ASCII 的 Symbol text（如 + - 等运算符号）
        raw = re.sub(r"<text[^>]*font-family:Symbol[^>]*>[^<]*[^\x00-\x7f][^<]*</text>",
                     "", raw)
        # 转成 UTF-8 供 rsvg 解析
        with open(svg, "w", encoding="utf-8") as f:
            f.write(raw)
        r2 = subprocess.run(["rsvg-convert", "-w", str(width), "-o", dest, svg],
                            capture_output=True)
        return r2.returncode == 0 and os.path.exists(dest)
    finally:
        for p in (tmp, svg):
            if os.path.exists(p):
                os.remove(p)


def extract_images(docx_path, prefix, flow=None, doc_bin=None):
    """docx 图片复制到 IMG_DIR，返回 (found, formula_map)。
    found: media名->相对路径（题图：PNG/JPEG 或大 WMF 用 .doc 原始 PNG 补）
    formula_map: media名->相对路径（小 WMF 公式图，已转 PNG）
    """
    z = zipfile.ZipFile(docx_path)
    found = {}
    formula_map = {}
    doc_pngs = extract_png_from_doc(doc_bin) if doc_bin else []
    png_idx = 0
    WMF_BIG = 8 * 1024   # 大于此的 wmf 视为题图（公式符号仅 ~0.4KB）

    def save(data, name, key=None, fmap=None):
        dest = os.path.join(IMG_DIR, "%s_%s" % (prefix, name))
        if not os.path.exists(dest):
            with open(dest, "wb") as fh:
                fh.write(data)
        rel = os.path.join("images", "%s_%s" % (prefix, name))
        (fmap or found)[key or os.path.basename(name)] = rel

    # 优先按文档顺序（flow），否则按 zip 顺序
    if flow:
        for it in flow:
            if it["kind"] != "image":
                continue
            base = os.path.basename(it["file"])
            try:
                data = z.read("word/media/" + base)
            except KeyError:
                continue
            if data[:8] == b"\x89PNG\r\n\x1a\n" or data[:3] == b"\xff\xd8\xff":
                save(data, base)
                png_idx += 1   # 该题图在 .doc 原始图中占一个位置
            elif data[:4] == b"\x01\x00\x09\x00":
                if len(data) > WMF_BIG:
                    # WMF 题图 → 用 .doc 原始 PNG 补
                    if png_idx < len(doc_pngs):
                        save(doc_pngs[png_idx], "wmf_%s.png" % base, key=base)
                        png_idx += 1
                else:
                    # 小 WMF 公式图 → 转 PNG 嵌入题干/选项
                    dest = os.path.join(IMG_DIR, "%s_formula_%s.png" % (prefix, base))
                    if not os.path.exists(dest):
                        wmf_to_png(data, dest)
                    if os.path.exists(dest):
                        formula_map[base] = os.path.join("images", "%s_formula_%s.png" % (prefix, base))
            else:
                continue
    else:
        for n in z.namelist():
            if "media/" in n:
                base = os.path.basename(n)
                try:
                    data = z.read(n)
                except KeyError:
                    continue
                if data[:8] == b"\x89PNG\r\n\x1a\n" or data[:3] == b"\xff\xd8\xff":
                    save(data, base)
    return found, formula_map


def textutil_text(doc):
    return subprocess.run(["textutil", "-convert", "txt", "-stdout", doc],
                          capture_output=True).stdout.decode("utf-8", "replace")


def build_fallback(tlines, missing_nums, source):
    """从 textutil 文本中按题号精确提取缺失的题（docx 丢失的公式题等）。
    返回题目列表（无图片关联）。
    """
    result = []
    part = ""
    section = ""
    i = 0
    n = len(tlines)
    while i < n:
        line = tlines[i]
        if ANSWER_HEAD_RE.match(line):
            break   # 答案区不用于提取题目
        pm = PART_RE.match(line)
        if pm:
            part = line
            section = ""
            i += 1
            continue
        sm = SECTION_RE.match(line)
        if sm and not line[:1].isdigit():
            section = line.split("。")[0]
            i += 1
            continue
        if is_table_number(line):
            i += 1
            continue
        m = QNUM_RE.match(line)
        if not m or int(m.group(1)) not in missing_nums:
            i += 1
            continue
        no = int(m.group(1))
        body = [m.group(2)] if m.group(2) else []
        # 子题型说明行（如"1．每道题包含两套图形…"）不提取
        joined_body = "".join(body)
        if re.search(r"解析", joined_body) or \
           re.match(r"^(每道|每道题|包含|请|在右面|在左边|下列各题|从四个|题目要求)", joined_body) or not body:
            i += 1
            continue
        i += 1
        while i < n:
            nx = QNUM_RE.match(tlines[i])
            if nx:
                break
            if is_table_number(tlines[i]):
                i += 1
                continue
            body.append(tlines[i])
            i += 1
        stem, opts = split_options("\n".join(body))
        # 有效性验证：必须含选项或题干空白括号，否则是表格数据/说明
        if len(opts) < 2 and not re.search(r"[（(]\s*[）)]", stem):
            i += 1
            continue
        result.append({"id": None, "part": part, "section": section, "stem": stem,
                       "options": opts, "images": [], "answer": "", "explanation": "",
                       "source": source, "no": no})
    return result


def merge_fallback(questions, fallback):
    """把 fallback 中缺失 (part, no) 的题插入 questions（保持题号升序）。"""
    have = {(q["part"], q["no"]) for q in questions}
    missing = sorted((q for q in fallback if (q["part"], q["no"]) not in have),
                     key=lambda q: q["no"])
    if not missing:
        return questions
    merged = list(questions)
    for fb in missing:
        inserted = False
        for i, q in enumerate(merged):
            if q["no"] > fb["no"]:
                merged.insert(i, fb)
                inserted = True
                break
        if not inserted:
            merged.append(fb)
    return merged


def text_to_flow(text):
    lines = [normalize(l) for l in text.splitlines() if normalize(l)]
    return [{"kind": "text", "text": l} for l in lines]


def pdf_flow(path, prefix):
    """用 PyMuPDF 解析 PDF，返回 (flow, images_map)。
    图片按页面 y 坐标插入文本流，自动关联到上方最近的题号。
    页眉装饰图（y<70）跳过。
    """
    import fitz
    from collections import defaultdict
    doc = fitz.open(path)
    flow = []
    images_map = {}
    seq = 0
    for page in doc:
        items = []
        words = page.get_text("words")
        lines = defaultdict(list)
        for w in words:
            lines[(w[5], w[6])].append(w)
        for key, ws in sorted(lines.items()):
            ws.sort(key=lambda w: w[0])
            y0 = min(w[1] for w in ws)
            x0 = min(w[0] for w in ws)
            text = " ".join(w[4] for w in ws)
            items.append((y0, x0, "text", text))
        for img in page.get_images(full=True):
            xref = img[0]
            for r in page.get_image_rects(xref):
                if r.y0 < 70:   # 页眉装饰图
                    continue
                seq += 1
                fname = "%s_pdf%03d.png" % (prefix, seq)
                pm = fitz.Pixmap(doc, xref)
                if pm.n - pm.alpha > 3 or pm.alpha:
                    pm = fitz.Pixmap(fitz.csRGB, pm)
                pm.save(os.path.join(IMG_DIR, fname))
                images_map[fname] = os.path.join("images", fname)
                items.append((r.y0, r.x0, "image", fname))
        items.sort(key=lambda it: (it[0], it[1]))
        for _, _, kind, data in items:
            if kind == "text":
                flow.append({"kind": "text", "text": data})
            else:
                flow.append({"kind": "image", "file": data})
    doc.close()
    return flow, images_map


def file_to_flow(path, prefix="imp"):
    """把上传的题库/答案文件解析为顺序流 (flow, images_map)。
    支持 doc / docx / pdf / txt。
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        flow = read_docx_flow(path)
        images_map, formula_map = extract_images(path, prefix, flow)
        return flow, images_map, formula_map
    if ext == ".doc":
        txt = textutil_text(path)
    elif ext == ".pdf":
        try:
            import fitz
            flow, images_map = pdf_flow(path, prefix)
            return flow, images_map, {}
        except ImportError:
            r = subprocess.run(["pdftotext", "-layout", path, "-"], capture_output=True)
            txt = r.stdout.decode("utf-8", "replace")
    elif ext in (".txt", ".md"):
        with open(path, encoding="utf-8", errors="replace") as f:
            txt = f.read()
    else:
        raise ValueError("不支持的文件格式：%s" % ext)
    return text_to_flow(txt), {}, {}


def convert_to_docx(path):
    """单个 .doc 转 docx（Word），返回 docx 路径。"""
    name = os.path.splitext(os.path.basename(path))[0]
    docx = os.path.join(TMP_DIR, "imp_" + re.sub(r"[^\w\u4e00-\u9fff]", "_", name) + ".docx")
    if os.path.exists(docx):
        return docx
    apple = ('tell application "Microsoft Word"\n'
             f'open POSIX file "{esc(path)}"\n'
             'set d to active document\n'
             f'save as d file name POSIX file "{esc(docx)}" file format format document\n'
             'close d saving no\nend tell\n')
    subprocess.run(["osascript", "-e", apple], check=True)
    return docx


def fill_missing_from_answers(qs, answer_map, source):
    """答案区存在的题号但题目区缺失的题，用答案/解析补全为占位题。
    例如图形推理题只有答案没有题目时，保证题号连续完整。"""
    have = {(q["part"], q["no"]) for q in qs}
    added = []
    for (part, no), info in answer_map.items():
        if (part, no) in have or not info.get("ans"):
            continue
        added.append({
            "id": None, "part": part, "section": "", "stem": "（本题在原文档中缺图/缺题干，仅存答案）",
            "options": [], "images": [], "answer": info["ans"],
            "explanation": info.get("expl", ""), "source": source, "no": no})
    qs.extend(added)
    qs.sort(key=lambda q: q["no"])


def import_documents(question_path, answer_path=None, source=None):
    """从上传文档解析题目（可附加独立答案文档）。返回题目列表（无 id）。
    场景：
      1) question_path + answer_path：题目文档 + 独立答案文档
      2) question_path 单独：单文档内同时含题与答案
    doc/docx 走完整 docx 解析（含图片、textutil 补缺），pdf/txt 走文本解析。
    """
    global _last_no
    ext = os.path.splitext(question_path)[1].lower()
    if ext in (".doc", ".docx"):
        docx = convert_to_docx(question_path) if ext == ".doc" else question_path
        flow = read_docx_flow(docx)
        segments = split_segments(flow)
        prefix = re.sub(r"\s+", "", os.path.basename(question_path))[:20]
        doc_bin = open(question_path, "rb").read() if ext == ".doc" else None
        images_map, formula_map = extract_images(docx, prefix, flow, doc_bin)
        _last_no = 0
        qs = parse_questions(segments, source or os.path.basename(question_path), images_map, formula_map)
        # textutil 补齐 docx 丢失的公式题
        txt = textutil_text(question_path)
        tlines = [normalize(l) for l in txt.splitlines() if normalize(l)]
        max_no = max((q["no"] for q in qs), default=0)
        missing_nums = set(range(1, max_no + 1)) - {q["no"] for q in qs}
        fb = build_fallback(tlines, missing_nums, source or os.path.basename(question_path))
        qs = merge_fallback(qs, fb)
    else:
        prefix = re.sub(r"\s+", "", os.path.basename(question_path))[:20]
        flow, images_map, formula_map = file_to_flow(question_path, prefix)
        segments = split_segments(flow)
        _last_no = 0
        qs = parse_questions(segments, source or os.path.basename(question_path), images_map, formula_map)
    if answer_path:
        aflow, _, _ = file_to_flow(answer_path, "imp")
        asegs = split_segments(aflow)
        if not any(s["name"].startswith("答案") for s in asegs):
            # 独立答案文档：无答案区标题，整体当作答案区解析
            asegs = [{"name": "答案:", "lines": [l for s in asegs for l in s["lines"]],
                      "tables": []}]
        answer_map, seq_lines = parse_answer_map(asegs)
    else:
        answer_map, seq_lines = parse_answer_map(segments)
    for q in qs:
        info = answer_map.get((q["part"], q["no"])) or answer_map.get(("", q["no"]))
        if info:
            q["answer"] = info["ans"]
            if info.get("expl"):
                q["explanation"] = info["expl"]
    if seq_lines and qs:
        seq = [(l[0], l[1:].strip()) for l in seq_lines]
        for q, (ans, expl) in zip(sorted(qs, key=lambda x: x["no"]), seq):
            q["answer"] = ans
            if expl:
                q["explanation"] = expl
    fill_missing_from_answers(qs, answer_map, source or os.path.basename(question_path))
    return qs


def main():
    global _last_no
    filter_kw = sys.argv[1] if len(sys.argv) > 1 else ""
    os.makedirs(IMG_DIR, exist_ok=True)
    pairs = doc_to_docx()
    all_questions = []
    for doc, docx in pairs:
        if filter_kw and filter_kw not in doc:
            continue
        _last_no = 0
        name = os.path.basename(doc)
        print("解析:", name)
        flow = read_docx_flow(docx)
        segments = split_segments(flow)
        prefix = re.sub(r"\s+", "", name).replace(".doc", "")[:20]
        doc_bin = open(doc, "rb").read()
        images_map, formula_map = extract_images(docx, prefix, flow, doc_bin)
        qs = parse_questions(segments, name, images_map, formula_map)
        # 补齐 docx 丢失的公式题
        txt = textutil_text(doc)
        tlines = [normalize(l) for l in txt.splitlines() if normalize(l)]
        max_no = max((q["no"] for q in qs), default=0)
        missing_nums = set(range(1, max_no + 1)) - {q["no"] for q in qs}
        fb = build_fallback(tlines, missing_nums, name)
        qs = merge_fallback(qs, fb)
        answer_map, seq_lines = parse_answer_map(segments)
        for q in qs:
            info = answer_map.get((q["part"], q["no"]))
            if info:
                q["answer"] = info["ans"]
                q["explanation"] = info["expl"]
        # 无题号答案区：按顺序分配给全部题目（如 (十一)(十二)）
        if seq_lines and qs:
            seq = [(l[0], l[1:].strip()) for l in seq_lines]
            ordered = sorted(qs, key=lambda q: (q["no"]))
            for q, (ans, expl) in zip(ordered, seq):
                q["answer"] = ans
                if expl:
                    q["explanation"] = expl
        # 用答案区补全缺失的题（如无图的图形推理题），保证题号连续
        fill_missing_from_answers(qs, answer_map, name)
        all_questions.extend(qs)
        print("  提取 %d 题, 图片 %d 张, 带答案 %d 题" % (
            len(qs), len(images_map),
            sum(1 for q in qs if q["answer"])))
    out = os.path.join(DATA_DIR, "questions.json")
    for i, q in enumerate(all_questions):
        q["id"] = "Q%05d" % (i + 1)   # 全局唯一 id
    with open(out, "w", encoding="utf-8") as f:
        json.dump(all_questions, f, ensure_ascii=False, indent=1)
    print("完成，共 %d 题 -> %s" % (len(all_questions), out))


if __name__ == "__main__":
    main()
