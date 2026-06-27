#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os, re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_BREAK

BASE = "/home/user/6-main/01-6个主号内容创作/引流小号_发布文案"

# 缺汇总Word的批次：(起, 止)
BATCHES = [(85, 91), (92, 98), (99, 105), (106, 114), (115, 123)]

def find_txt(idx):
    for fn in os.listdir(BASE):
        if fn.startswith(f"{idx}_") and fn.endswith(".txt"):
            return os.path.join(BASE, fn)
    return None

def add_txt_to_doc(doc, idx, content, is_first):
    if not is_first:
        # 分页
        p = doc.add_paragraph()
        p.add_run().add_break(WD_BREAK.PAGE)
    # 篇序标题
    h = doc.add_paragraph()
    hr = h.add_run(f"第{idx}篇")
    hr.bold = True
    hr.font.size = Pt(14)
    for line in content.strip().split("\n"):
        p = doc.add_paragraph()
        run = p.add_run(line)
        run.font.size = Pt(12)
        if line.startswith("标题：") or line.startswith("账号："):
            run.bold = True

def gen_batch(start, end):
    doc = Document()
    first = True
    count = 0
    for idx in range(start, end + 1):
        path = find_txt(idx)
        if not path:
            print(f"  ⚠ 第{idx}篇 txt 缺失，跳过")
            continue
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        add_txt_to_doc(doc, idx, content, first)
        first = False
        count += 1
    fname = f"引流小号_第{start}-{end}篇.docx"
    out = os.path.join(BASE, fname)
    doc.save(out)
    print(f"  ✅ {fname}（{count}篇）")
    return out

if __name__ == "__main__":
    print("=== 补生成引流小号汇总Word ===")
    for s, e in BATCHES:
        gen_batch(s, e)
