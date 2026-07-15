#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os, zipfile

BASE = "/home/user/6-main/01-6个主号内容创作/引流小号_发布文案"
P = "#会计 #中级会计 #中级会计备考 #注会cpa #财会 #备考日记 #上班族备考 #考前冲刺 #诗雨会计"

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
    zip_path = os.path.join(BASE, "引流小号_第115-123篇.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for i in range(115, 124):
            for fname in os.listdir(BASE):
                if fname.startswith(f"{i}_"):
                    full = os.path.join(BASE, fname)
                    zf.write(full, fname)
    return zip_path

articles = [
    (115, "距考试七十天我才刷完第一遍",
     "倒计时70天了，我才把教材刷完第一遍，看着别人都二刷三刷，慌得不行。😣后来用诗雨漫画笔记快速过考点，二刷速度明显快了，焦虑也少一点。",
     "沈小辉"),
    (116, "公式记完转头就忘真的崩溃",
     "所得税那些公式，背了忘忘了背，每次做题还是卡壳。😭翻了诗雨漫画笔记里的图解，把递延那块画成图，居然记住了，原来图比字好记。",
     "孙青5263"),
    (117, "上班族备考真的只能挤时间",
     "白天上班晚上带娃，能学习的只有娃睡后那一个小时。⏰时间太碎，只能看点好消化的。最近用诗雨漫画笔记，碎片时间翻几页也不费劲。",
     "橙子学会计"),
    (118, "做大题没思路看答案都看不懂",
     "综合题一上来就懵，连答案都看不明白步骤。😵后来发现是基础没打牢。用诗雨漫画笔记把分录一步步理清，再看答案就顺了很多。",
     "新老婆"),
    (119, "越临近考试越静不下心看书",
     "还剩两个多月，反而越来越焦虑，书翻开五分钟就走神。😮‍💨试着换成漫画版笔记，图多有意思，至少能看进去，慢慢找回状态。",
     "孙玲小红书"),
    (120, "长投权益法绕来绕去总搞反",
     "权益法的分红到底冲长投还是记收益，我每次都搞反。🤯后来看诗雨漫画笔记里的对比图，成本法权益法一目了然，再也不混了。",
     "爹爹"),
    (121, "报名缴费后才发现没怎么学",
     "报名费交了三个月，书还是崭新的。😅看着考试一天天近，终于逼自己每天学一点。用诗雨漫画笔记降低开始的门槛，没那么抗拒翻书了。",
     "张菊香"),
    (122, "二刷还是错一堆是不是没救了",
     "都二刷了，做题正确率还是上不去，怀疑自己是不是不适合考会计。😔后来发现是没抓重点。诗雨漫画笔记把高频考点标得很清楚，刷题更有方向。",
     "孙文礼"),
    (123, "考前冲刺阶段每天都在和困意斗争",
     "冲刺期想多学会儿，可一坐下就犯困，效率特别低。😴换成诗雨漫画笔记，图多字少看着不累，困的时候也能多撑十几分钟。",
     "晴天"),
]

if __name__ == "__main__":
    print("=== 生成引流小号第115-123篇 ===")
    for idx, title, body, account in articles:
        write_txt(idx, title, body, account)
    zp = make_zip()
    print(f"\n  ✅ ZIP打包完成：{zp}")
