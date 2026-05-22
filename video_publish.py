#!/usr/bin/env python3
"""
视频号图文发布脚本
txt格式：
标题：
正文：
话题：
时间：
账号：04视频号|01主账号黄色手机
"""

import os, sys, subprocess, time
from pathlib import Path

# ─────────────────────────────────────────
# 需要挂商品链接的账号
ACCOUNTS_WITH_GOODS = {'01主账号黄色手机', '02诗雨vip'}

# 账号坐标配置
ACCOUNT_POS = {
    '01主账号黄色手机': (282, 327),
    '02诗雨vip':       (271, 404),
    '03蓝色手机':      (268, 476),
    '04孙文礼':        (257, 538),
    '05胡芝兰':        (259, 613),
    '06吴学安':        (277, 680),
}

# 界面坐标
POS_FULLSCREEN   = (234,  70)    # 全屏按钮
POS_MULTIBOARD   = (71,   197)   # 多开面板
POS_GROUP_04     = (225,  293)   # 04视频号分组
POS_BLANK        = (473,  833)   # 空白处（移开鼠标用）
POS_CONTENT_MGR  = (466,  317)   # 内容管理
POS_GRAPHIC      = (439,  430)   # 图文
POS_RIGHT_AREA   = (1167, 259)   # 右侧区域（向右滚动用）
POS_POST_BTN     = (1624, 345)   # 发表图文按钮
POS_MUSIC        = (1231, 860)   # 音乐选项
POS_MUSIC_MOVE   = (1289, 1075)  # 音乐点击后移动鼠标位置
POS_MUSIC_FIRST  = (1489, 1078)  # 第一个音乐
POS_LINK         = (1247, 769)   # 链接选项
POS_GOODS        = (1238, 895)   # 商品选项
POS_SELECT_GOODS = (1243, 852)   # 选择需要添加的商品
POS_GOODS_FIRST  = (489,  525)   # 第一个商品（2026会计漫画笔记）
POS_ADD_GOODS    = (1555, 1076)  # 添加按钮
POS_TITLE        = (1214, 318)   # 图文标题输入框
POS_DESC         = (1213, 415)   # 图文描述输入框
POS_UPLOAD       = (895,  463)   # 上传图片+号
POS_SCROLL_AREA  = (941,  906)   # 底部空白处（向下滚动用）
POS_PUBLISH      = (1479, 925)   # 发表按钮
POS_CLOSE_TAB    = (404,  55)     # 关闭网页×按钮
POS_GROUP_04_END = (225,  294)    # 收尾点04视频号
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

def scroll(x, y, vertical=0, horizontal=0):
    swift_code = f'''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let s = CGEvent(scrollWheelEvent2Source: src, units: .pixel, wheelCount: 2, wheel1: {vertical}, wheel2: {horizontal}, wheel3: 0)!
s.location = CGPoint(x: {int(x)}, y: {int(y)})
s.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(0.5)

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
    data = {'标题': '', '正文': '', '话题': '', '时间': '', '账号': '', '商品': ''}
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

def publish_video(folder_path, first=False):
    folder = Path(folder_path)
    txt_files = sorted(folder.glob('*.txt'))
    if not txt_files:
        print(f"  ❌ 找不到txt文件：{folder}")
        return False

    for idx, txt_file in enumerate(txt_files):
        # 只处理视频号的txt
        with open(txt_file, 'r', encoding='utf-8') as f:
            content = f.read()
        if '04视频号' not in content:
            continue
        print(f"  📄 处理：{txt_file.name}")
        data = parse_txt(txt_file)
        print(f"  📄 标题：{data['标题']}")
        print(f"  👤 账号：{data['账号']}")

        # 解析账号名
        if '|' in data['账号']:
            _, account_name = data['账号'].split('|', 1)
            account_name = account_name.strip()
        else:
            account_name = data['账号'].strip()

        # 获取图片
        exts = {'.jpg', '.jpeg', '.png', '.webp'}
        images = sorted([str(f) for f in folder.iterdir() if f.suffix.lower() in exts])
        if not images:
            print(f"  ❌ 没有找到图片")
            return False
        print(f"  🖼  图片：{len(images)} 张")

        _do_video_publish(folder, data, images, account_name, first=(first and idx == 0))

        if idx < len(txt_files) - 1:
            print(f"  ⏳ 等待30秒后发布下一个账号...")
            time.sleep(30)

    return True

def _do_video_publish(folder, data, images, account_name, first=False):
    # 1. 多开面板（每次都执行）
    print("  📋 多开面板...")
    click(*POS_MULTIBOARD)
    time.sleep(1)

    # 2. 04视频号分组（每次都执行）
    print("  📁 04视频号...")
    click(*POS_GROUP_04)
    time.sleep(1)

    # 3. 选择账号
    print(f"  👤 选择账号：{account_name}...")
    if account_name in ACCOUNT_POS:
        click(*ACCOUNT_POS[account_name])
    else:
        print(f"  ⚠️  找不到账号坐标：{account_name}")
        return
    time.sleep(50)  # 等待50秒加载

    # 移鼠标到空白处
    move_mouse(*POS_BLANK)
    time.sleep(0.5)

    # 5. 内容管理
    print("  📂 内容管理...")
    click(*POS_CONTENT_MGR)
    time.sleep(1)

    # 6. 图文
    print("  🖼  图文...")
    click(*POS_GRAPHIC)
    time.sleep(3)  # 等待3秒加载

    # 7. 移到右侧并滚动
    print("  ➡️  滚动到右侧...")
    move_mouse(*POS_RIGHT_AREA)
    time.sleep(0.5)
    scroll(POS_RIGHT_AREA[0], POS_RIGHT_AREA[1], horizontal=-500)
    time.sleep(1)

    # 8. 发表图文
    print("  📝 点发表图文...")
    click(*POS_POST_BTN)
    time.sleep(2)

    # 9. 音乐
    print("  🎵 选择音乐...")
    click(*POS_MUSIC)
    time.sleep(3)  # 等待3秒
    move_mouse(*POS_MUSIC_MOVE)  # 移动鼠标到下方
    time.sleep(0.5)
    click(*POS_MUSIC_FIRST)
    time.sleep(1)

    # 10. 链接→商品（只有指定账号才挂）
    if account_name in ACCOUNTS_WITH_GOODS:
        print("  🔗 添加商品链接...")
        click(*POS_LINK)
        time.sleep(1)
        click(*POS_GOODS)
        time.sleep(1)
        click(*POS_SELECT_GOODS)
        time.sleep(3)
        click(*POS_GOODS_FIRST)
        time.sleep(1)
        click(*POS_ADD_GOODS)
        time.sleep(1)
    else:
        print("  🔗 跳过商品链接...")

    # 11. 填写标题
    print("  ✏️  填写标题...")
    click(*POS_TITLE)
    time.sleep(0.5)
    type_text(data['标题'])

    # 12. 填写描述（正文+话题）
    print("  ✏️  填写描述...")
    click(*POS_DESC)
    time.sleep(0.5)
    full_content = data['正文']
    if data['话题']:
        full_content += '\n' + data['话题']
    type_text(full_content)

    # 13. 上传图片
    print("  🖼  上传图片...")
    click(*POS_UPLOAD)
    time.sleep(2)
    # Command+Shift+G 跳转文件夹
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
    # 回车进入文件夹
    swift_code2 = '''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let rDown = CGEvent(keyboardEventSource: src, virtualKey: 0x24, keyDown: true)!
rDown.post(tap: .cghidEventTap)
let rUp = CGEvent(keyboardEventSource: src, virtualKey: 0x24, keyDown: false)!
rUp.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code2])
    time.sleep(1)
    # 全选
    swift_code3 = '''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let aDown = CGEvent(keyboardEventSource: src, virtualKey: 0x00, keyDown: true)!
aDown.flags = .maskCommand
aDown.post(tap: .cghidEventTap)
let aUp = CGEvent(keyboardEventSource: src, virtualKey: 0x00, keyDown: false)!
aUp.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code3])
    time.sleep(0.5)
    subprocess.run(['swift', '-e', swift_code2])
    time.sleep(30)  # 等待30秒图片上传完成

    # 14. 滚动到底部
    print("  ⬇️  滚动到底部...")
    move_mouse(*POS_SCROLL_AREA)
    time.sleep(0.5)
    for _ in range(5):
        scroll(POS_SCROLL_AREA[0], POS_SCROLL_AREA[1], vertical=-200)
    time.sleep(1)

    # 15. 发表
    print("  🚀 发表...")
    click(*POS_PUBLISH)
    time.sleep(30)  # 等待30秒

    # 关闭网页
    print("  ❌ 关闭网页...")
    click(*POS_CLOSE_TAB)
    time.sleep(1)

    # 收尾：点04视频号
    print("  🔄 收尾点04视频号...")
    click(*POS_GROUP_04_END)
    time.sleep(20)  # 等待20秒再开始下一篇

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
    print("🎬 视频号图文发布脚本启动\n")

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
        publish_video(folder, first=(i == 1))
        if i < len(subfolders):
            print(f"  ⏳ 等待30秒后开始下一篇...")
            time.sleep(30)

    print("🎉 全部发布完成！")
