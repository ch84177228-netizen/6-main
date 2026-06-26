#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os, zipfile

BASE = "/home/user/6-main/01-6个主号内容创作/引流小号_发布文案"
P = "#会计 #中级会计 #中级会计备考 #注会cpa #财会 #备考日记 #上班族备考 #备考焦虑 #诗雨会计"

def save(path, content):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

def write_txt(idx, title, body, account_name):
    tc = len(title)
    bc = len(body.replace("\n","").replace(" ",""))
    ok = "✅" if tc<=20 and bc<=80 else f"❌(标题{tc},正文{bc}字)"
    print(f"  {ok} 标题{tc}字 正文{bc}字 | {account_name}")
    kw = title[:8].replace("？","").replace("！","").replace("，","").replace("。","")
    fname = f"{idx}_{kw}_{account_name}.txt"
    content = f"标题：{title}\n\n正文：\n{body}\n\n话题：{P}\n\n时间：\n\n账号：14芳芳财会|{account_name}\n"
    save(os.path.join(BASE, fname), content)

def make_zip():
    zip_path = os.path.join(BASE, "引流小号_第106-114篇.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for i in range(106, 115):
            for fname in os.listdir(BASE):
                if fname.startswith(f"{i}_"):
                    full = os.path.join(BASE, fname)
                    zf.write(full, fname)
    return zip_path

articles = [
    (106, "考场找不到座位号慌了好一会儿",
     "进考场发现座位号怎么都找不到，在考场来回走了好几圈，浪费了十来分钟。😅后来看错了考场编号，差点交白卷。最近用诗雨漫画笔记复习，至少不会紧张到脑子空白了。",
     "沈小辉"),
    (107, "题目都看懂就是选不出正确答案",
     "做题有个毛病，每个选项都觉得有道理，就是选不出来。🤔知识点能背，判断就是卡着。换了漫画版笔记来看，感觉逻辑清楚了一点，选题时不那么犹豫了。",
     "孙青5263"),
    (108, "备考期间戒了剧只留一个综艺",
     "为备考把追的剧全删了，只保留一个综艺当奖励，做完一套题才准看一集。😭换了诗雨漫画笔记来配合刷题，记忆效率高了不少，奖励综艺也能少看几集了。",
     "橙子学会计"),
    (109, "半夜刷题脑子完全不知道在做啥",
     "昨晚11点还在刷题，刷到后面脑子完全不转，做完一道不知道自己选了什么。😴试着换成看诗雨漫画笔记，图多字少，困了也能看进去一点，比硬撑有用。",
     "新老婆"),
    (110, "模拟题全对真题一错就是一大片",
     "刷模拟题以为差不多了，一套真题下来错了20多分。😰才发现模拟题的坑和真题完全不一样。回头配合漫画笔记重看考点，感觉理解层次确实不一样了。",
     "孙玲小红书"),
    (111, "翻教材发现折角全是不会的地方",
     "翻开教材，折过角的页面几乎每页都有，全是做错题的地方。📚不会的比会的多太多。后来用诗雨漫画笔记做针对复习，图解那几块，容易记住一点。",
     "爹爹"),
    (112, "对着错题本发现错的全是同一类",
     "整理错题本才发现，错来错去就是那几类题，说明根本没搞懂那个考点。😤与其刷题不如补知识点。翻了诗雨漫画笔记对应章节，感觉豁然开朗了。",
     "张菊香"),
    (113, "题目没审完就填答案这次吃了亏",
     "做题速度慢，一紧张就不看完题目就写，结果好几道审题错了。😖练习时没这问题，进了考场就乱节奏。现在配合漫画笔记复习考点，减少会但没答对的情况。",
     "孙文礼"),
    (114, "走出考场感觉没一道题答对了",
     "上次考完出来感觉全错了，结果成绩还不错，虚惊一场。😂越是考完觉得完了反而越能过。这次用诗雨漫画笔记理清思路，希望考完不要再有那种全错感了。",
     "晴天"),
]

if __name__ == "__main__":
    print("=== 生成引流小号第106-114篇 ===")
    for idx, title, body, account in articles:
        write_txt(idx, title, body, account)
    zp = make_zip()
    print(f"\n  ✅ ZIP打包完成：{zp}")
