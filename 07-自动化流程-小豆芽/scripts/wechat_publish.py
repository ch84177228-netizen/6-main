#!/usr/bin/env python3
"""
公众号图文发布脚本
txt格式：
标题：
正文：
话题：
时间：
账号：13公众号|01蓝色密码
"""

import os, sys, subprocess, time
from pathlib import Path

# ─────────────────────────────────────────
# 账号坐标配置
ACCOUNT_POS = {
    '01蓝色密码': (268, 304),
    '02VIPSun':   (271, 376),
    '03黄色手机': (263, 441),
    '04吴学安':   (260, 514),
    '05粉色手机': (268, 581),
}

# 界面坐标
POS_FULLSCREEN   = (233,  68)    # 最大化窗口
POS_MULTIBOARD   = (72,   194)   # 多开面板
POS_GROUP_13     = (223,  266)   # 13公众号分组
POS_BLANK        = (473,  833)   # 空白处
POS_PASTE_IMG    = (1010, 839)   # 贴图按钮
POS_TITLE        = (946,  495)   # 标题输入框
POS_CONTENT      = (1088, 545)   # 正文输入框
POS_ADD_IMG      = (991,  255)    # +加号添加图片
POS_SCROLL_AREA  = (833,  759)   # 底部滚动区域
POS_LINK         = (1357, 817)   # 链接按钮
POS_SHOP_GOODS   = (1426, 936)   # 小店商品选项
POS_SEARCH_BOX   = (847,  331)   # 搜索框
POS_FIRST_RESULT = (612,  479)   # 第一个搜索结果
POS_INSERT       = (978,  1018)  # 插入按钮
POS_PUBLISH      = (1525, 1086)  # 发表按钮
POS_CONFIRM_PUB  = (982,  755)   # 确认发表按钮
POS_CLOSE_WX1    = (646,  54)    # 第一次关闭公众号
POS_CLOSE_WX2    = (405,  56)    # 第二次关闭公众号
POS_GROUP_13_END = (236,  264)   # 收尾点13公众号
# ─────────────────────────────────────────

def click(x, y, delay=0.8):
    swift_code = f'''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let p = CGPoint(x: {int(x)}, y: {int(y)})
let d = CGEvent(mouseEventSource: src, mouseType: .leftMouseDown, mouseCursorPosition: p, mouseButton: .left)!
let u = CGEvent(mouseEventSource: src, mouseType: .leftMouseUp, mouseCursorPosition: p, mouseButton: .left)!
d.post(tap: .cghidEventTap)
u.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(delay)

def move_mouse(x, y):
    swift_code = f'''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let p = CGPoint(x: {int(x)}, y: {int(y)})
let e = CGEvent(mouseEventSource: src, mouseType: .mouseMoved, mouseCursorPosition: p, mouseButton: .left)!
e.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(0.5)

def scroll_down(x, y, amount=300):
    swift_code = f'''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let s = CGEvent(scrollWheelEvent2Source: src, units: .pixel, wheelCount: 1, wheel1: -{amount}, wheel2: 0, wheel3: 0)!
s.location = CGPoint(x: {int(x)}, y: {int(y)})
s.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(0.3)

def scroll_up(x, y, amount=300):
    swift_code = f'''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let s = CGEvent(scrollWheelEvent2Source: src, units: .pixel, wheelCount: 1, wheel1: {amount}, wheel2: 0, wheel3: 0)!
s.location = CGPoint(x: {int(x)}, y: {int(y)})
s.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(0.3)

def type_text(text):
    process = subprocess.Popen(['pbcopy'], stdin=subprocess.PIPE)
    process.communicate(text.encode('utf-8'))
    time.sleep(0.3)
    swift_code = '''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let vDown = CGEvent(keyboardEventSource: src, virtualKey: 0x09, keyDown: true)!
vDown.flags = .maskCommand
vDown.post(tap: .cghidEventTap)
let vUp = CGEvent(keyboardEventSource: src, virtualKey: 0x09, keyDown: false)!
vUp.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(0.5)

def parse_txt(txt_path):
    data = {'标题': '', '正文': '', '话题': '', '时间': '', '账号': ''}
    with open(txt_path, 'r', encoding='utf-8') as f:
        content = f.read()
    lines = content.splitlines()
    current_key = None
    current_value = []
    for line in lines:
        matched = False
        for key in data:
            if line.startswith(f'{key}：') or line.startswith(f'{key}:'):
                if current_key:
                    data[current_key] = '\n'.join(current_value).strip()
                current_key = key
                value = line[len(key):].lstrip('：:').strip()
                current_value = [value] if value else []
                matched = True
                break
        if not matched and current_key:
            current_value.append(line)
    if current_key:
        data[current_key] = '\n'.join(current_value).strip()
    return data

def publish_one(folder_path, first=False):
    folder = Path(folder_path)
    txt_files = sorted(folder.glob('*.txt'))
    if not txt_files:
        print(f"  ❌ 找不到txt文件：{folder}")
        return False

    for idx, txt_file in enumerate(txt_files):
        # 只处理公众号的txt
        with open(txt_file, 'r', encoding='utf-8') as f:
            content = f.read()
        if '13公众号' not in content:
            continue
        print(f"  📄 处理：{txt_file.name}")
        data = parse_txt(txt_file)
        print(f"  📄 标题：{data['标题']}")
        print(f"  👤 账号：{data['账号']}")

        if '|' in data['账号']:
            _, account_name = data['账号'].split('|', 1)
            account_name = account_name.strip()
        else:
            account_name = data['账号'].strip()

        images = sorted([str(f) for f in folder.iterdir() if f.suffix.lower() in {'.jpg', '.jpeg', '.png', '.webp'}])
        if not images:
            print(f"  ❌ 没有找到图片")
            return False
        print(f"  🖼  图片：{len(images)} 张")

        _do_publish(folder, data, images, account_name, first=(first and idx == 0))

        if idx < len(txt_files) - 1:
            print(f"  ⏳ 等待30秒后发布下一个账号...")
            time.sleep(30)

    return True

def _do_publish(folder, data, images, account_name, first=False):
    # 1. 多开面板（每次都执行）
    print("  📋 多开面板...")
    click(*POS_MULTIBOARD)
    time.sleep(1)

    # 2. 13公众号分组（每次都执行）
    print("  📁 13公众号...")
    click(*POS_GROUP_13)
    time.sleep(1)

    # 3. 选择账号
    print(f"  👤 选择账号：{account_name}...")
    if account_name in ACCOUNT_POS:
        click(*ACCOUNT_POS[account_name])
    else:
        print(f"  ⚠️  找不到账号坐标：{account_name}")
        return
    time.sleep(30)  # 等待30秒加载

    # 5. 点贴图
    print("  🖼  点贴图...")
    click(*POS_PASTE_IMG)
    time.sleep(10)  # 等待10秒

    # 6. 填写标题
    print("  ✏️  填写标题...")
    click(*POS_TITLE)
    time.sleep(0.5)
    type_text(data['标题'])

    # 7. 填写正文+话题
    print("  ✏️  填写正文...")
    click(*POS_CONTENT)
    time.sleep(0.5)
    full_content = data['正文']
    if data['话题']:
        full_content += '\n' + data['话题']
    type_text(full_content)

    # 8. 上传图片（先移到空白处、滑到顶端，再点加号）
    print("  🖼  上传图片...")
    move_mouse(*POS_BLANK)
    time.sleep(0.5)
    for _ in range(10):
        scroll_up(POS_BLANK[0], POS_BLANK[1], 300)
    time.sleep(0.5)
    click(*POS_ADD_IMG, delay=0.5)
    time.sleep(0.5)
    click(*POS_ADD_IMG)
    time.sleep(2)
    swift_code = '''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let gDown = CGEvent(keyboardEventSource: src, virtualKey: 0x22, keyDown: true)!
gDown.flags = CGEventFlags(rawValue: CGEventFlags.maskCommand.rawValue | CGEventFlags.maskShift.rawValue)
gDown.post(tap: .cghidEventTap)
let gUp = CGEvent(keyboardEventSource: src, virtualKey: 0x22, keyDown: false)!
gUp.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(1)
    type_text(str(folder))
    swift_enter = '''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let rDown = CGEvent(keyboardEventSource: src, virtualKey: 0x24, keyDown: true)!
rDown.post(tap: .cghidEventTap)
let rUp = CGEvent(keyboardEventSource: src, virtualKey: 0x24, keyDown: false)!
rUp.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_enter])
    time.sleep(1)
    swift_cmda = '''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let aDown = CGEvent(keyboardEventSource: src, virtualKey: 0x00, keyDown: true)!
aDown.flags = .maskCommand
aDown.post(tap: .cghidEventTap)
let aUp = CGEvent(keyboardEventSource: src, virtualKey: 0x00, keyDown: false)!
aUp.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_cmda])
    time.sleep(0.5)
    subprocess.run(['swift', '-e', swift_enter])
    time.sleep(30)  # 等待30秒图片上传完成

    # 9. 滚动到底部
    print("  ⬇️  滚动到底部...")
    move_mouse(*POS_SCROLL_AREA)
    time.sleep(0.5)
    for _ in range(5):
        scroll_down(POS_SCROLL_AREA[0], POS_SCROLL_AREA[1], 300)
    time.sleep(1)

    # 10. 链接→小店商品
    print("  🔗 添加商品链接...")
    click(*POS_LINK)
    time.sleep(2)
    click(*POS_SHOP_GOODS)
    time.sleep(2)

    # 11. 搜索商品
    print("  🔍 搜索商品...")
    click(*POS_SEARCH_BOX)
    time.sleep(2)
    type_text('会计漫画笔记')
    time.sleep(5)
    click(*POS_FIRST_RESULT)
    time.sleep(2)

    # 12. 插入
    print("  ✅ 插入商品...")
    click(*POS_INSERT)
    time.sleep(2)

    # 13. 发表
    print("  🚀 发表...")
    click(*POS_PUBLISH)
    time.sleep(5)

    # 14. 确认发表
    print("  ✅ 确认发表...")
    click(*POS_CONFIRM_PUB)
    time.sleep(10)  # 等待10秒

    # 14.5 继续发表
    print("  ✅ 继续发表...")
    click(980, 810)
    time.sleep(35)  # 等待35秒

    # 15. 第一次关闭公众号
    print("  ❌ 关闭公众号（1）...")
    click(*POS_CLOSE_WX1)
    time.sleep(4)

    # 16. 第二次关闭公众号
    print("  ❌ 关闭公众号（2）...")
    click(*POS_CLOSE_WX2)
    time.sleep(4)

    # 17. 收尾：点13公众号
    print("  🔄 收尾点13公众号...")
    click(*POS_GROUP_13_END)
    time.sleep(4)

    print(f"  ✅ 发布完成！\n")

def pick_folder():
    script = '''
    tell application "Finder"
        set theFolder to choose folder with prompt "请选择今日发布的主文件夹"
        return POSIX path of theFolder
    end tell
    '''
    result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
    if result.returncode != 0:
        sys.exit(0)
    return result.stdout.strip()

if __name__ == "__main__":
    print("📰 公众号发布脚本启动\n")

    if len(sys.argv) > 1:
        main_folder = Path(sys.argv[1])
    else:
        main_folder = Path(pick_folder())

    txt_files = list(main_folder.glob('*.txt'))
    if txt_files:
        subfolders = [main_folder]
    else:
        subfolders = sorted([f for f in main_folder.iterdir() if f.is_dir()])

    if not subfolders:
        print("❌ 找不到内容")
        sys.exit(1)

    print(f"📂 找到 {len(subfolders)} 条内容待发布\n")

    for i, folder in enumerate(subfolders, 1):
        print(f"[{i}/{len(subfolders)}] 开始处理：{folder.name}")
        publish_one(folder, first=(i == 1))
        if i < len(subfolders):
            print(f"  ⏳ 等待10秒后开始下一篇...")
            time.sleep(10)

    print("🎉 全部发布完成！")
