#!/usr/bin/env python3
"""
统一Word文档生成脚本
支持所有文件夹的格式规范
用法: python3 generate_word.py
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

# ─────────────────────────────────────────
# 颜色常量
# ─────────────────────────────────────────
class C:
    RED      = RGBColor(0xC0, 0x39, 0x2B)   # 标题红
    DARK_RED = RGBColor(0xC8, 0x52, 0x2A)   # 03号标题红
    DARK     = RGBColor(0x1A, 0x18, 0x16)   # 深色正文
    GRAY     = RGBColor(0x6B, 0x65, 0x60)   # 灰色辅助
    BLUE     = RGBColor(0x1A, 0x5F, 0xA8)   # 话题蓝
    NAVY     = RGBColor(0x2C, 0x3E, 0x6B)   # 05号页眉蓝
    BERRY    = RGBColor(0xD9, 0x5F, 0x7A)   # 05号知识点粉
    SECTION  = RGBColor(0x2C, 0x3E, 0x50)   # 章节标题深蓝

def set_page_margins(doc, margin_cm=2.54):
    """设置页边距"""
    for section in doc.sections:
        section.top_margin    = Cm(margin_cm)
        section.bottom_margin = Cm(margin_cm)
        section.left_margin   = Cm(margin_cm)
        section.right_margin  = Cm(margin_cm)

def add_hr(doc):
    """添加分割线"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'CCCCCC')
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_colored_para(doc, text, color, bold=False, size=12, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.size = Pt(size)
    return p

def add_body_lines(doc, text, size=11):
    """将多行文本按段落添加，自动识别小标题和正文"""
    for line in text.split('\n'):
        stripped = line.strip()
        if not stripped:
            doc.add_paragraph()
            continue
        # 判断是否是小标题行（以特殊符号开头或全大写标识）
        is_header = (
            stripped.startswith(('📋', '🎯', '🔑', '⚠️', '✨', '📝', '✅', '❌', '🔄', '📅', '💪', '🌱')) or
            (stripped.startswith('**') and stripped.endswith('**'))
        )
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        # 清除markdown加粗
        clean = stripped.replace('**', '')
        run = p.add_run(clean)
        run.font.size = Pt(size)
        if is_header:
            run.font.color.rgb = C.DARK_RED
            run.font.bold = True
        else:
            run.font.color.rgb = C.DARK

def write_txt(content, out_path):
    """写入txt文件"""
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(content)
    print(f'  ✅ {os.path.basename(out_path)}')

def make_txt_06(title, body, topics, xhs_account, gzh_account, folder, filename_base):
    """06文件夹：同时生成小红书txt和公众号txt"""
    tmpl = '标题：{title}\n\n正文：\n{body}\n\n话题：{topics}\n\n时间：\n\n账号：{account}\n'
    base = f'/home/user/6-main/{folder}'
    write_txt(tmpl.format(title=title, body=body, topics=topics, account=xhs_account),
              f'{base}/{filename_base}_小红书.txt')
    write_txt(tmpl.format(title=title, body=body, topics=topics, account=gzh_account),
              f'{base}/{filename_base}_公众号.txt')

def add_footer_info(doc, topics, account, gray=C.GRAY):
    add_hr(doc)
    add_colored_para(doc, f'话题：{topics}', gray, size=9)
    add_colored_para(doc, '时间：', gray, size=9)
    add_colored_para(doc, f'账号：{account}', gray, size=9)


# ─────────────────────────────────────────
# 通用单账号Word生成器
# ─────────────────────────────────────────
def make_word_standard(title, body, topics, account, out_path,
                        title_color=C.RED, title_size=20):
    """通用格式：标题大红居中 + 正文 + 底部信息"""
    doc = Document()
    set_page_margins(doc)

    # 账号信息（顶部灰色小字）
    add_colored_para(doc, account, C.GRAY, size=9)
    add_hr(doc)

    # 标题
    add_colored_para(doc, title, title_color, bold=True, size=title_size,
                     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    add_hr(doc)

    # 正文
    add_body_lines(doc, body)

    # 底部
    add_footer_info(doc, topics, account)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    print(f'  ✅ {os.path.basename(out_path)}')


# ─────────────────────────────────────────
# 01-6个主号：6账号合并Word
# ─────────────────────────────────────────
def make_word_6accounts(accounts_data, topic_name, out_path):
    """6账号合并到一个Word，每账号独立分页"""
    doc = Document()
    set_page_margins(doc)

    # 封面页
    add_colored_para(doc, topic_name, C.RED, bold=True, size=24,
                     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    add_colored_para(doc, '6账号完整图文脚本', C.GRAY, size=12,
                     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_hr(doc)

    style_map = {
        '账号①': '主号·甄嬛传版',
        '账号②': '副号1·蜡笔小新版',
        '账号③': '副号2·甄嬛速记版',
        '账号④': '副号3·避坑吐槽版',
        '账号⑤': '副号4·纯考点总结版',
        '账号⑥': '副号5·考前冲刺版',
    }
    for i, (label, data) in enumerate(accounts_data.items()):
        if i == 0:
            doc.add_page_break()
        # 账号标识
        add_colored_para(doc, f'{label}  {style_map.get(label, "")}', C.GRAY, size=9)
        add_hr(doc)
        # 标题
        add_colored_para(doc, data['title'], C.RED, bold=True, size=18,
                         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
        add_hr(doc)
        # 正文
        add_body_lines(doc, data['body'])
        # 底部
        add_footer_info(doc, data['topics'], data['account'])
        # 分页（除最后一页）
        if i < len(accounts_data) - 1:
            doc.add_page_break()

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    print(f'  ✅ {os.path.basename(out_path)}')


# ─────────────────────────────────────────
# 05-CPA税法：5 PART结构
# ─────────────────────────────────────────
def make_word_cpa_tax(title, body, topics, account, out_path, is_small=False):
    """CPA税法格式：含PART A-E结构"""
    doc = Document()
    set_page_margins(doc)

    # 页眉
    add_colored_para(doc, 'CPA税法 · 增值税', C.NAVY, bold=True, size=10)
    add_hr(doc)

    # 标题
    add_colored_para(doc, title, C.RED, bold=True, size=20,
                     align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    add_hr(doc)

    if is_small:
        # 小号引流版直接输出正文
        add_body_lines(doc, body)
    else:
        # 大号：PART B 正文内容
        add_colored_para(doc, 'PART B  图文内容稿', C.BERRY, bold=True, size=11, space_after=6)
        add_body_lines(doc, body)
        add_hr(doc)

        # PART E 小红书正文
        add_colored_para(doc, 'PART E  小红书发布文案', C.BERRY, bold=True, size=11, space_after=6)
        add_colored_para(doc, f'标题：{title}', C.DARK, size=11)

    add_footer_info(doc, topics, account)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    print(f'  ✅ {os.path.basename(out_path)}')


# ─────────────────────────────────────────
# 主程序：生成本批次所有Word
# ─────────────────────────────────────────
BASE = '/home/user/6-main'

def main():
    print('\n开始生成Word文档...\n')

    # ── 06 公务员（每篇生成：小红书txt + 公众号txt + Word）──
    print('📁 06-公务员')
    _06_title  = '申论到底考什么？很多人连题型都没搞清楚'
    _06_body   = '''很多人备考申论，一上来就买教材刷题。
但连申论考什么都没搞清楚，刷再多题也白费。📚

申论说难，其实不难。
难的是没有方向，不知道从哪里下手。

📋 申论就4种题型，先搞清楚

**第一种：归纳概括题**
最基础，也最多见。
给你一堆材料，让你提炼要点。
核心：从材料里找答案，不要自己发挥。
不是考你会不会写，是考你会不会找。

**第二种：综合分析题**
比归纳难一级。
常见问法："谈谈你对XXX的理解"
答题逻辑：是什么→为什么→怎么办，三段走。
材料给了什么，你说什么，不要空谈。

**第三种：应用文写作**
考场最怕的其实是这种，不是大作文。
格式不对直接扣分。
常考：倡议书、简报、发言稿、调研报告。
核心：先记格式，内容从材料里找，不要自己编。

**第四种：大作文（议论文）**
很多人一上来就先练大作文，顺序搞反了。
大作文要最后练。
先把前三种小题练熟，再攻大作文。
核心：论点要鲜明，举例不要太单一。

🎯 正确备考顺序

归纳概括 → 综合分析 → 应用文 → 大作文

前期把小题练熟，时间占70%。
后期才重点练大作文，占30%。

很多人申论学不好，不是题做少了。
是方向错了，方法没对。
先把4种题型搞清楚，再按顺序刷，才是正确节奏。🌱

我整理了一份适合考公小白的入门资料包，需要可以私信：考公资料。'''
    _06_topics = '#考公 #公务员考试 #国考 #省考 #考公小白 #公考备考 #申论 #申论备考 #考公资料'
    _06_xhs    = '06红书店铺 | 考公学姐'
    _06_gzh    = '13公众号|05粉色手机'

    make_txt_06(_06_title, _06_body, _06_topics, _06_xhs, _06_gzh,
                '06-公务员', '考公文案_06_申论题型科普')
    make_word_standard(
        title      = _06_title,
        body       = _06_body,
        topics     = _06_topics,
        account    = _06_xhs,
        out_path   = f'{BASE}/06-公务员/考公文案_06_申论题型科普.docx',
        title_color= C.RED,
    )

    # ── 05 CPA税法 大号 ──
    print('📁 05-CPA税法')
    make_word_cpa_tax(
        title   = '进项税额抵扣的坑，我帮你全踩完了',
        body    = '''增值税计算，销项税大家都会算。
但进项税额这块，坑特别多。📚

你以为买了东西就能抵扣？不一定。
你以为出差餐费能抵？不行。
你以为住宿费不行？其实可以。

📋 第一关：哪些凭证可以抵扣？

🔑 口诀：专票海关，收购农，旅客运，通行费，完税凭证

逐句拆解：
· 专票：增值税专用发票（最常见）
· 海关：海关进口增值税专用缴款书
· 收购农：农产品收购发票
· 旅客运：旅客运输服务购票凭证
· 通行费：ETC通行费电子发票
· 完税凭证：进口货物完税凭证

📝 诗雨叮嘱：旅客运输可抵扣，但必须是合规扣税凭证——增值税电子普通发票或注明旅客身份信息的铁路/航空票据，普通收据不行！

❌ 第二关：哪些进项不能抵扣？

🔑 口诀：简免集福娱，贷款餐饮居民乐；非正常损失连累进项

⚠️ 易错高频对比：出差报销——机票✅、火车票✅、住宿✅（取专票）、餐饮费❌。
住宿和餐饮别搞混，这是每年必考点！

🔄 第三关：什么时候要转出已抵扣进项？

🔑 口诀：已抵后改用途转，非正常损失转，无法确定用还原法

📝 诗雨叮嘱：地震、台风等自然灾害造成的损失，不是"非正常损失"，不需要做进项转出！

✨ 本篇核心考点汇总

1. 可抵扣凭证口诀：专票海关收购农旅客运通行费完税
2. 不得抵扣口诀：简免集福娱 + 贷款餐饮居民乐
3. 进项转出：改用途 + 非正常损失 + 还原法
4. 住宿✅ vs 餐费❌——这对比每年必考！

诗雨叮嘱：进项税额几乎每年必考，"可抵扣凭证"和"不得抵扣情形"背熟，选择题稳拿分！''',
        topics  = '#注会cpa #cpa备考 #税法 #增值税 #进项税额 #注会备考 #会计考试干货 #诗雨会计',
        account = '06红书店铺｜姗姗 店铺大号',
        out_path= f'{BASE}/05-CPA税法/税法第④篇_进项税额抵扣_大号.docx',
        is_small= False,
    )

    make_word_cpa_tax(
        title   = '出差住宿可以抵税，但餐费不行！很多人搞不清',
        body    = '''出差报销，到底哪些费用能抵扣增值税进项？🤔

✅ 机票/火车票：可以抵（凭证要注明旅客信息）
✅ 住宿费：可以抵（要取得增值税专用发票）
❌ 餐饮费：不可以，明确列入不得抵扣清单

🔑 记这句口诀：贷款餐饮居民乐，统统不能抵扣掉

很多人以为出差费用全能抵扣。
结果餐费报销这块被查了还不知道原因。⚠️

住宿和餐饮，一个能抵一个不能抵。
记清楚这个对比，选择题拿稳这个考点。💪''',
        topics  = '#注会cpa #cpa备考 #税法 #增值税 #进项税额 #财税知识 #会计考试干货 #诗雨会计',
        account = '06红书店铺｜姗姗 店铺大号',
        out_path= f'{BASE}/05-CPA税法/税法第④篇_进项税额抵扣_小号引流.docx',
        is_small= True,
    )

    # ── 04 注会专业号 ──
    print('📁 04-注会专业号')
    make_word_standard(
        title   = '跟我一起备注会的同学，都在用这个方法',
        body    = '''备注会这件事，真的很孤独。📚

刷到凌晨不知道自己学到哪了。
做错的题不知道怎么归类。
知识点背了忘，忘了再背，没有成就感。

我身边坚持下来的同学，有一个共同点——
他们不是一个人在备考。

我自己备考那年，前期也是自己刷书。
刷到7月，发现还是不系统，心里很慌。

后来跟着一个带学带背的老师走，
每天告诉我哪些是真正的高频考点，哪些先跳过，
帮我把6科的框架串起来。

那种感觉完全不一样——
不是"我在学"，而是"我跟着走就行"。🌱

**带学带背班做的事情很简单：**

每天给你一个考点，带你记住它。
告诉你这个考点怎么考、怎么踩坑、怎么拿分。
你只需要每天跟着打卡30分钟就好。

不需要自己规划进度。
不需要自己找重点。
只需要你每天能花30分钟跟着走。

现在是5月，距离考试还有不到3个月。💪
这3个月，方法对了，真的能变很多。

有在备注会的同学，欢迎了解一下。
感兴趣的在评论区说"注会"，我私信你详情。''',
        topics  = '#注会cpa #注会备考 #cpa备考 #注册会计师 #会计考试干货 #备考攻略 #注会经验 #诗雨会计',
        account = '06红书店铺｜姗姗 店铺大号',
        out_path= f'{BASE}/04-注会专业号/Day7_种草.docx',
        title_color=C.RED,
    )

    # ── 03 初级专业号 ──
    print('📁 03-初级专业号')
    make_word_standard(
        title   = '现在开始备考2027年初级会计，比别人早半年',
        body    = '''很多人备考初级会计，都是等报名了才开始。📚
然后3个月赶完两本书，还要做题。
临考前手忙脚乱，焦虑到睡不着。

其实完全没必要。

2027年初级会计，考试时间大约是明年5-6月。
现在是2026年5月，距离考试整整12个月。

现在开始，你比大多数人早了半年。

🎯 早开始的3个真实优势

**第一：时间够，不用熬夜赶进度**
每天1小时，稳稳推进。
不高压、不焦虑，真正学扎实。

**第二：先建框架，后面学起来有方向**
初级不难，但知识点多。
先花1-2个月把两科目录过一遍。
知道每章考什么，后面精读不会懵。

**第三：做题量自然多，发挥更稳**
备考时间长，刷题次数多。
选择题一看就知道考哪个考点。
这就是早备考最直接的优势。

📋 建议备考节奏

🗓️ 6-8月：过一遍教材，建框架，不求背，求熟悉
🗓️ 9-12月：重点章节逐章精读 + 做配套练习
🗓️ 1-3月：整理错题，高频考点反复巩固
🗓️ 4-5月：刷真题，模拟考，查漏补缺

现在开始不是卷，是聪明备考。🌱

会计实务先学，经济法可以稍后开始。
从《会计实务》第一章，1天20页，完全来得及。

你是今年刚决定备考的吗？评论区告诉我👇''',
        topics  = '#初级会计备考 #初级会计 #会计 #备考经验 #会计小白 #大学生考证 #考证规划 #财会 #会计入门 #诗雨会计',
        account = '06红书店铺|孙咏美1530521',
        out_path= f'{BASE}/03-初级专业号/初级_现在开始备考2027年.docx',
        title_color=C.DARK_RED,
    )

    # ── 02 会计科普 ──
    print('📁 02-会计科普')
    make_word_standard(
        title   = '考了初级会计，找财务实习有没有用？',
        body    = '''这个问题，我被问过很多次。📚

"学姐，我考了初级，简历上写了，HR会看吗？"
"初级会计含金量够不够？"
"没有实习经验，考初级能加分吗？"

说实话，这个问题没有标准答案。
但我可以告诉你几个真实的情况。

🎯 初级对找实习：有用，但不是万能的

**有用的地方：**

**第一，筛简历时能过关。**
很多财务岗实习，简历筛选时会看有没有会计相关证书。
初级会计能证明你对这个行业有基本了解，不是完全的门外汉。
大家条件差不多的情况下，有证的确实会多一点优势。

**第二，面试时有话可说。**
如果你备考过初级，你就学过：
借贷记账、科目分类、固定资产折旧、增值税计算……
面试官随便问一个，你能答上来，印象就不一样了。

**第三，实习上手更快。**
财务实习最常见的工作：录凭证、整理发票、核对账单。
学过初级的同学，至少知道凭证是什么，科目怎么对应。
没学过的同学，前两周都在问"借贷是什么意思"。

**但有几件事别误解：**

初级会计不是找实习的硬性门槛。
拿了证不代表HR就必须要你。
实习看的是：你学过什么 + 有没有相关经历 + 你这个人靠不靠谱。
证书只是其中一个维度。

另外，初级不能替代实际操作能力。
如果你只是考过了，从来没用软件录过账，
实习第一天还是会懵。

🌱 我的建议

如果你大学在读，想找财务实习：
先把初级考了，打一个知识基础。
同时学一下Excel和金蝶/用友的基础操作。
两个加起来，找实习的底气就够了。

证书不能替代一切，但有了它，走进财务部门的第一步，会更顺一点。✨''',
        topics  = '#初级会计 #初级会计备考 #财务实习 #会计小白 #大学生考证 #会计 #财会 #零基础学会计 #考证规划 #诗雨会计',
        account = '06红书店铺｜胡志兰手机',
        out_path= f'{BASE}/02-会计科普/小红书文案_考初级对找财务实习有没有用_06红书店铺｜胡志兰手机.docx',
        title_color=C.RED,
    )

    # ── 01 6个主号（6账号合并Word）──
    print('📁 01-6个主号内容创作')
    accounts_data = {
        '账号①': {
            'title'  : '收入确认五步法，用甄嬛传一次讲通透',
            'body'   : '''后宫里，甄嬛每次接旨，都要走五道程序，才能算"事成了"。📚
账面上，会计确认收入，也是一样——
不是钱到了就能记，是"控制权转移了"才能记。

🎯 第一步：识别合同
皇上下旨，这就是"合同"。
双方都认可 + 有商业实质 + 收款可能性高。
合同不成立，后面四步全作废。

第二步：识别履约义务
皇上的旨意可能包含两件事——管理后宫和侍奉皇上。
这就是两个履约义务，分开计量，分开确认。

第三步：确定交易价格
谈清楚"皇上给多少彩礼"。
还要考虑可变对价、融资成分、非现金对价。

第四步：分摊交易价格
总彩礼按两件事各自的"单独售价"比例拆开。
不能随意拆，要按可观察的单独售价比例来。

第五步：履约时确认收入
甄嬛管完后宫才记管理收入，侍寝完成才记侍寝收入。
时间点：控制权一次转移（卖货）
时段：持续提供服务（装修、订阅服务）

✨ 诗雨叮嘱
收入确认的核心逻辑：控制权转移了，才能确认收入。
不是款项收到了，不是货物发出去了，是"控制权"转移了。
这个逻辑没搞清楚，五步法套了也白套。''',
            'topics' : '#中级会计备考 #中级会计 #会计 #收入确认 #甄嬛传 #会计大白话 #备考冲刺 #财会干货 #9月上岸 #诗雨会计',
            'account': '06红书店铺|粉色手机',
        },
        '账号②': {
            'title'  : '靠小新家，我终于搞懂收入确认了',
            'body'   : '''说实话，我之前一直搞不懂收入确认。📚

不是说不会背定义，是真的不理解：
为什么钱收到了，却不能马上记收入？
为什么货发出去了，有时候也不算收入？

后来我用小新家的生意来理解，突然就通了。

广志爸爸开了家定制蛋糕店。
美冴妈妈下了一个订单：要订生日蛋糕+配送到家，共500元。

第一步：这个订单算不算合同？
广志点头，美冴付了定金，合同成立✅

第二步：这个订单包含几件事？
蛋糕是一件事，配送是另一件事——两个履约义务。

第三步：500元怎么定的？
蛋糕市场价400，配送市场价100，总价500，正常。

第四步：500元怎么分给两件事？
蛋糕占80%→分400元，配送占20%→分100元。

第五步：什么时候记收入？
蛋糕做好交给美冴时，记400元（时间点）。
配送完成签收时，记100元（时间点）。

结论：不是钱到就算，是"交付了"才算。🌱''',
            'topics' : '#中级会计备考 #中级会计 #会计 #收入确认 #蜡笔小新 #会计大白话 #备考经验 #财会干货 #诗雨会计',
            'account': '06红书店铺|黄色手机',
        },
        '账号③': {
            'title'  : '收入确认5步法考点｜甄嬛传5分钟背完',
            'body'   : '''收入确认高频考点，甄嬛版速记来了。📚
建议收藏，考前回看。

🔑 五步法对照

第①步 皇上下的旨是否合法有效 → 识别合同（商业实质+收款可能性）
第②步 旨意里包含几件事 → 识别履约义务（可明确区分的商品/服务）
第③步 谈清楚给多少彩礼 → 确定交易价格（含可变对价+融资成分）
第④步 彩礼按比例分给每件事 → 分摊交易价格（按单独售价比例）
第⑤步 事情做完了才能记账 → 履约时确认收入（时间点/时段）

⚠️ 三大易考陷阱

陷阱①：收到钱 ≠ 确认收入（控制权转移才算）
陷阱②：货发出去 ≠ 确认收入（看控制权，不看货在哪）
陷阱③：一笔合同可能含多个履约义务，要分开算

🔑 时间点 vs 时段

时间点：卖货、出售软件授权（使用权型）
时段：装修服务、订阅服务、物业服务、软件维护（访问权型）

判断时段的三个条件（满足一个即可）：
① 客户在企业履约同时获得并消耗利益
② 企业创造的资产对客户有价值（即使未完工）
③ 企业的履约没有替代用途，且有权收款

建议收藏，考前回看✅''',
            'topics' : '#中级会计备考 #中级会计 #会计考点速记 #收入确认 #甄嬛传 #会计 #备考冲刺 #财会干货 #9月上岸 #诗雨会计',
            'account': '06红书店铺|蓝色ipad张菊芳',
        },
        '账号④': {
            'title'  : '学收入确认，我踩过的3个让我崩溃的坑',
            'body'   : '''备考中级会计，收入确认这章，我踩了三个坑。📚
说出来，让你少走弯路。

坑①：以为"收到钱"就能确认收入

我一开始的理解：钱到账了 = 可以记收入了。
然后做题全错。

真相：确认收入的核心是"控制权转移"，不是"款项收到"。
付款了但货还没交，是合同负债（预收账款）。
货交了但还没付款，就可以确认收入了。

坑②：一个合同只看成一笔收入

真相：一个合同可能包含多个独立的履约义务。
卖电视机附带一年上门维修服务——
卖货（时间点确认）+ 维修服务（时段确认）要分开。
把总价拆开，分别确认。

坑③：时间点和时段分不清

判断时段有三个条件，满足一个就行：
① 客户在你履约的同时就获得并消耗利益
② 你在创造的资产对客户有价值（即使没完工）
③ 你的履约没有替代用途，且对已完成部分有权收款

最常见：装修、订阅服务、物业费 = 时段
卖货、一次性软件授权 = 时间点

这三个坑，我当时全踩了。
现在发现把逻辑搞清楚，收入这章真的不难。🌱''',
            'topics' : '#中级会计备考 #中级会计 #会计 #收入确认 #备考经验 #会计大白话 #避坑 #财会干货 #诗雨会计',
            'account': '06红书店铺|孙文新134 已实名',
        },
        '账号⑤': {
            'title'  : '收入确认核心考点｜收藏这一篇就够了',
            'body'   : '''收入确认完整考点整理，建议收藏复习。📚

✅ 一、五步法框架

① 识别合同：商业实质 + 各方认可 + 收款可能性高
② 识别履约义务：可明确区分的单项商品/服务
③ 确定交易价格：固定对价 + 可变对价 + 融资成分 + 非现金对价
④ 分摊交易价格：按各履约义务单独售价比例分摊
⑤ 确认收入：时间点 or 时段，取决于控制权转移方式

✅ 二、时段确认条件（满足任一即可）

① 客户在企业履约同时获得并消耗利益
② 企业创造的资产对客户有价值（即使未完工）
③ 无替代用途 + 对已完成部分有权收款

✅ 三、可变对价

期望值法（多种可能结果）vs 最可能金额法（两种可能结果）
限制：高度可能不发生重大收入转回才能纳入

✅ 四、重大融资成分

付款与控制权转移时间差异明显 → 调整交易价格
豁免：期限≤1年

✅ 五、主要责任人 vs 代理人

主要责任人（全额）：承担库存风险、可自主定价
代理人（净额）：仅确认佣金

✅ 六、常见错误

❌ 款项收到 ≠ 确认收入
❌ 一个合同只算一笔
❌ 退货直接冲收入（应设退货负债+退货资产）

建议收藏，考前回看✅''',
            'topics' : '#中级会计备考 #中级会计 #会计考点速记 #收入确认 #备考冲刺 #财会干货 #考点整理 #9月上岸 #诗雨会计',
            'account': '06红书店铺|孙文新 小号店铺号',
        },
        '账号⑥': {
            'title'  : '距离中级考试105天｜收入确认必考考点今天搞定',
            'body'   : '''距离2026年中级会计考试，还有105天。⏰

收入确认是每年必考章节，分值重，考法多变。
今天把必考考点列出来，对照检查你掌握了多少。

📋 为什么这章是必考？

中级会计实务第十四章，历年考频极高。
客观题必出、主观题常出。
难点不在背，在于判断：时间点还是时段？是不是代理人？

🎯 5个必须会的核心考点

考点1：五步法框架
识别合同→识别履约义务→确定交易价格→分摊→确认

考点2：时段 vs 时间点确认
满足三个条件之一：同时获得并消耗 / 资产有价值 / 无替代用途+有权收款

考点3：可变对价的限制
高度可能不发生重大收入转回 → 才能纳入
期望值法 vs 最可能金额法

考点4：主要责任人 vs 代理人
主要责任人=全额；代理人=净额（佣金）
判断关键：谁承担库存风险？谁有自主定价权？

考点5：附退回条款的销售
预期不退回部分 → 确认收入
预期退回部分 → 确认退货负债

📅 备考建议

现在开始刷真题选择题，以考点4和考点5为重点。
主观题至少做3道历年真题，搞清楚答题格式。
9月5日之前，这章要能闭眼写出五步法框架。

9月5日稳稳上岸💪''',
            'topics' : '#中级会计备考 #中级会计 #收入确认 #备考冲刺 #会计考试干货 #9月上岸 #考点整理 #中级会计考试 #财会干货 #诗雨会计',
            'account': '06红书店铺|姗姗已实名',
        },
    }
    make_word_6accounts(
        accounts_data,
        topic_name='收入确认五步法',
        out_path=f'{BASE}/01-6个主号内容创作/收入确认_6账号发布文案/收入确认_6账号完整脚本.docx',
    )

    print('\n✅ 全部Word文件生成完毕！\n')

if __name__ == '__main__':
    main()
