#!/usr/bin/env python3
"""
坐标采集工具
移动鼠标到目标位置，按 Enter 记录；按 s+Enter 跳过（保留原值）
完成后自动输出可替换的坐标代码块
"""

import subprocess
import sys

# 当前脚本中的原始坐标（跳过时保留）
CURRENT = {
    'POS_FULLSCREEN':   (233,  68),
    'POS_MULTIBOARD':   (72,   194),
    'POS_GROUP_13':     (223,  266),
    'POS_BLANK':        (473,  833),
    'POS_PASTE_IMG':    (1010, 839),
    'POS_TITLE':        (946,  495),
    'POS_CONTENT':      (1088, 545),
    'POS_ADD_IMG':      (991,  255),
    'POS_SCROLL_AREA':  (833,  759),
    'POS_LINK':         (1357, 817),
    'POS_SHOP_GOODS':   (1426, 936),
    'POS_SEARCH_BOX':   (847,  331),
    'POS_FIRST_RESULT': (612,  479),
    'POS_INSERT':       (978,  1018),
    'POS_PUBLISH':      (1525, 1086),
    'POS_CONFIRM_PUB':  (982,  755),
    'POS_CLOSE_WX1':    (646,  54),
    'POS_CLOSE_WX2':    (405,  56),
    'POS_GROUP_13_END': (236,  264),
}

CURRENT_ACCOUNTS = {
    '01蓝色密码': (268, 304),
    '02VIPSun':   (271, 376),
    '03黄色手机': (263, 441),
    '04吴学安':   (260, 514),
    '05粉色手机': (268, 581),
}

# 采集顺序与说明
POSITIONS = [
    ('POS_FULLSCREEN',   '最大化窗口按钮'),
    ('POS_MULTIBOARD',   '多开面板按钮'),
    ('POS_GROUP_13',     '13公众号分组'),
    ('POS_BLANK',        '空白处（填完正文后的空白区域）'),
    ('POS_PASTE_IMG',    '贴图按钮'),
    ('POS_TITLE',        '标题输入框'),
    ('POS_CONTENT',      '正文输入框'),
    ('POS_ADD_IMG',      '+ 加号（添加图片按钮）'),
    ('POS_SCROLL_AREA',  '底部滚动区域'),
    ('POS_LINK',         '链接按钮'),
    ('POS_SHOP_GOODS',   '小店商品选项'),
    ('POS_SEARCH_BOX',   '搜索框'),
    ('POS_FIRST_RESULT', '第一个搜索结果'),
    ('POS_INSERT',       '插入按钮'),
    ('POS_PUBLISH',      '发表按钮'),
    ('POS_CONFIRM_PUB',  '确认发表按钮'),
    ('POS_CLOSE_WX1',    '第一次关闭公众号 ×'),
    ('POS_CLOSE_WX2',    '第二次关闭公众号 ×'),
    ('POS_GROUP_13_END', '收尾点13公众号'),
]

ACCOUNTS = [
    ('01蓝色密码', '账号 01蓝色密码'),
    ('02VIPSun',   '账号 02VIPSun'),
    ('03黄色手机', '账号 03黄色手机'),
    ('04吴学安',   '账号 04吴学安'),
    ('05粉色手机', '账号 05粉色手机'),
]


def get_mouse_pos():
    swift_code = '''
import CoreGraphics
if let event = CGEvent(source: nil) {
    let pos = event.location
    print("\\(Int(pos.x)),\\(Int(pos.y))")
}
'''
    result = subprocess.run(['swift', '-e', swift_code], capture_output=True, text=True)
    if result.returncode == 0 and ',' in result.stdout:
        x, y = result.stdout.strip().split(',')
        return int(x), int(y)
    return None


def capture(label, current):
    while True:
        inp = input(f"  → 移鼠标到【{label}】，Enter 记录 / s 跳过: ").strip().lower()
        if inp == 's':
            print(f"     跳过，保留原值 {current}")
            return current
        pos = get_mouse_pos()
        if pos:
            print(f"     ✅ 记录 {pos}")
            return pos
        print("     ❌ 读取失败，请重试")


def main():
    print("=" * 55)
    print("  坐标采集工具  —  wechat_publish.py")
    print("  Enter=记录当前鼠标位置  s=跳过保留原值")
    print("=" * 55)

    new_pos = {}
    print("\n【界面坐标】")
    for var, desc in POSITIONS:
        new_pos[var] = capture(desc, CURRENT[var])

    new_accounts = {}
    print("\n【账号坐标】")
    for acc, desc in ACCOUNTS:
        new_accounts[acc] = capture(desc, CURRENT_ACCOUNTS[acc])

    # 输出结果
    print("\n" + "=" * 55)
    print("采集完成！以下是新坐标，已写入 coords_result.txt")
    print("=" * 55)

    lines = []
    lines.append("# 账号坐标配置")
    lines.append("ACCOUNT_POS = {")
    for k, v in new_accounts.items():
        lines.append(f"    '{k}': {v},")
    lines.append("}")
    lines.append("")
    lines.append("# 界面坐标")
    for var, pos in new_pos.items():
        comment = next((desc for v, desc in POSITIONS if v == var), '')
        lines.append(f"{'POS_' + var[4:]:<20} = {str(pos):<14}  # {comment}")

    output = '\n'.join(lines)
    print(output)

    with open('coords_result.txt', 'w', encoding='utf-8') as f:
        f.write(output + '\n')
    print("\n✅ 已保存到 coords_result.txt")


if __name__ == '__main__':
    main()
