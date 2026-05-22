#!/usr/bin/env python3
"""
新榜小豆芽自动发布脚本
文件夹结构：
📁 今日发布
   📁 第1条
      🖼 图片1.jpg
      🖼 图片2.jpg
      📄 文案.txt
   📁 第2条
      ...

文案.txt 格式：
标题：考公第一步不是刷题
正文：很多考公小白一开始就走偏了……
话题：#考公 #公务员考试 #国考
时间：09:00
账号：01CPA诗雨
"""

import os
import sys
import time
import subprocess
from pathlib import Path
from datetime import datetime

# ─────────────────────────────────────────
# ✏️  屏幕坐标配置（根据你的屏幕校准）
# Mac 坐标系原点在左下角，Y轴向上
SCREEN_HEIGHT = 1117  # 你的屏幕高度

def y(screen_y):
    """直接返回原始坐标，不翻转"""
    return screen_y

# 主界面坐标
POS_FULLSCREEN   = (234,  y(69))     # 全屏按钮
POS_ONE_KEY      = (67,   y(302))    # 左侧一键发布
POS_IMG_PUBLISH  = (64,   y(403))    # 左侧图片发布
POS_CLEAR        = (1020, y(1089))   # 一键清空
POS_CLEAR_CONFIRM= (1092, y(1034))   # 清空确认「确定」
POS_SELECT_ACCT  = (1291, y(374))    # 选择发布账号
POS_TITLE        = (397,  y(461))    # 标题输入框
POS_CONTENT      = (395,  y(550))    # 正文输入框
POS_UPLOAD       = (525,  y(246))    # 本地上传
POS_PUBLISH      = (1141, y(1087))   # 一键发布
POS_CONFIRM_PUB  = (1076, y(466))    # 发布至平台确认按钮
POS_CLOSE_RECORD = (618,  y(55))     # 关闭发布记录tab
POS_CLOSE_IMGPUB = (376,  y(57))     # 关闭图片发布tab

# 账号选择弹窗坐标
POS_TAB_LIST     = (761,  y(159))    # 按列表选择tab
POS_SEARCH_CLEAR = (1133, y(155))    # 搜索框×清除按钮
POS_GROUP_06     = (555,  y(491))    # 06红书店铺分组
POS_GROUP_14     = (547,  y(443))    # 14芳芳财会分组
POS_GROUP_07     = (552,  y(534))    # 07抖音发作品分组
POS_SEARCH_BOX   = (990,  y(153))    # 搜索框
POS_FIRST_RESULT = (929,  y(236))    # 第一个搜索结果
POS_CONFIRM      = (1175, y(782))    # 确定按钮

# 每个分组之间的垂直间距（像素）
GROUP_HEIGHT = 36
# ─────────────────────────────────────────

# 账号分组列表（顺序要和小豆芽左侧列表一致）
ACCOUNT_GROUPS = [
    "01CPA诗雨",
    "02诗雨樱桃",
    "03诗雨会计课堂",
    "13 悦初级科普",
    "04考公学姐",
    "14芳芳财会",
    "15会计",
    "05会计引流",
    "06会计引流",
    "07会计引流",
    "08会计引流",
    "09会计引流",
]

def click(x, y, delay=0.5):
    """用Swift点击指定坐标"""
    swift_code = f'''
import CoreGraphics
import Foundation

let point = CGPoint(x: {int(x)}, y: {int(y)})
let src = CGEventSource(stateID: .hidSystemState)
let down = CGEvent(mouseEventSource: src, mouseType: .leftMouseDown, mouseCursorPosition: point, mouseButton: .left)!
let up = CGEvent(mouseEventSource: src, mouseType: .leftMouseUp, mouseCursorPosition: point, mouseButton: .left)!
down.post(tap: .cghidEventTap)
Thread.sleep(forTimeInterval: 0.05)
up.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(delay)

def type_text(text):
    """输入文字（通过剪贴板粘贴，支持中文）"""
    process = subprocess.Popen(['pbcopy'], stdin=subprocess.PIPE)
    process.communicate(text.encode('utf-8'))
    time.sleep(0.3)
    # 用Swift模拟Command+A和Command+V
    swift_code = '''
import CoreGraphics
import Foundation

let src = CGEventSource(stateID: .hidSystemState)

// Command+A 全选
let aDown = CGEvent(keyboardEventSource: src, virtualKey: 0x00, keyDown: true)!
aDown.flags = .maskCommand
aDown.post(tap: .cghidEventTap)
let aUp = CGEvent(keyboardEventSource: src, virtualKey: 0x00, keyDown: false)!
aUp.post(tap: .cghidEventTap)
Thread.sleep(forTimeInterval: 0.1)

// Command+V 粘贴
let vDown = CGEvent(keyboardEventSource: src, virtualKey: 0x09, keyDown: true)!
vDown.flags = .maskCommand
vDown.post(tap: .cghidEventTap)
let vUp = CGEvent(keyboardEventSource: src, virtualKey: 0x09, keyDown: false)!
vUp.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(0.5)

def scroll(x, y, amount):
    """在指定坐标滚动（amount正数向上，负数向下）"""
    swift_code = f'''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let scroll = CGEvent(scrollWheelEvent2Source: src, units: .pixel, wheelCount: 1, wheel1: {amount}, wheel2: 0, wheel3: 0)!
scroll.location = CGPoint(x: {int(x)}, y: {int(y)})
scroll.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(0.3)

def swift_key(keycode, cmd=False, shift=False):
    """用Swift按键"""
    flags = []
    if cmd: flags.append('.maskCommand')
    if shift: flags.append('.maskShift')
    flag_line = f'down.flags = {flags[0]}' if len(flags)==1 else (f'down.flags = {flags[0]}.union({flags[1]})' if len(flags)==2 else '')
    swift_code = f'''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let down = CGEvent(keyboardEventSource: src, virtualKey: {keycode}, keyDown: true)!
{flag_line}
down.post(tap: .cghidEventTap)
let up = CGEvent(keyboardEventSource: src, virtualKey: {keycode}, keyDown: false)!
up.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(0.3)

def find_and_click_sync():
    """用图像识别找到「同步至右侧」按钮并点击"""
    import subprocess, tempfile, os
    import numpy as np

    template_path = os.path.expanduser('~/Desktop/sync_button_template.png')
    screen_path = os.path.join(tempfile.gettempdir(), 'xdouya_screen.png')

    # 截全屏
    subprocess.run(['screencapture', '-x', screen_path])
    time.sleep(0.5)

    try:
        import cv2
        from PIL import Image

        screen = cv2.imread(screen_path)
        template = cv2.imread(template_path)

        if screen is None or template is None:
            print("  ⚠️  图像识别失败，使用默认坐标")
            click(*POS_SYNC)
            return

        result = cv2.matchTemplate(screen, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)

        if max_val > 0.7:
            # Retina屏坐标除以2
            x = (max_loc[0] + template.shape[1] // 2) / 2
            y = (max_loc[1] + template.shape[0] // 2) / 2
            print(f"  🎯 找到同步至右侧按钮：({int(x)}, {int(y)})，相似度：{max_val:.2f}")
            click(int(x), int(y))
        else:
            print(f"  ⚠️  未找到按钮（相似度{max_val:.2f}），使用默认坐标")
            click(*POS_SYNC)

    except Exception as e:
        print(f"  ⚠️  图像识别出错：{e}，使用默认坐标")
        click(*POS_SYNC)
    finally:
        if os.path.exists(screen_path):
            os.remove(screen_path)


def swift_key(keycode, cmd=False, shift=False):
    """用Swift按键"""
    flags = []
    if cmd: flags.append('.maskCommand')
    if shift: flags.append('.maskShift')
    flag_line = f'down.flags = {flags[0]}' if len(flags)==1 else (f'down.flags = {flags[0]}.union({flags[1]})' if len(flags)==2 else '')
    swift_code = f'''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let down = CGEvent(keyboardEventSource: src, virtualKey: {keycode}, keyDown: true)!
{flag_line}
down.post(tap: .cghidEventTap)
let up = CGEvent(keyboardEventSource: src, virtualKey: {keycode}, keyDown: false)!
up.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(0.3)

def swift_cmdv():
    swift_key(0x09, cmd=True)

def swift_cmda():
    swift_key(0x00, cmd=True)

def swift_enter():
    swift_key(0x24)


    """用Swift按键"""
    flags = []
    if cmd: flags.append('.maskCommand')
    if shift: flags.append('.maskShift')
    flags_str = flags[0] if len(flags)==1 else (flags[0]+'.union('+flags[1]+')') if flags else 'CGEventFlags(rawValue: 0)'
    swift_code = f'''
import CoreGraphics
let src = CGEventSource(stateID: .hidSystemState)
let down = CGEvent(keyboardEventSource: src, virtualKey: {keycode}, keyDown: true)!
{"down.flags = " + flags_str if flags else ""}
down.post(tap: .cghidEventTap)
let up = CGEvent(keyboardEventSource: src, virtualKey: {keycode}, keyDown: false)!
up.post(tap: .cghidEventTap)
'''
    subprocess.run(['swift', '-e', swift_code])
    time.sleep(0.3)

def swift_cmdv():
    swift_key(0x09, cmd=True)

def swift_cmda():
    swift_key(0x00, cmd=True)

def swift_enter():
    swift_key(0x24)

def open_finder_select_files(image_paths):
    """通过 Finder 选择文件"""
    # 用 osascript 打开文件选择对话框并选中指定文件
    paths_str = ', '.join([f'POSIX file "{p}"' for p in image_paths])
    script = f'''
    tell application "System Events"
        tell process "新榜小豆芽"
            set frontmost to true
        end tell
    end tell
    '''
    subprocess.run(['osascript', '-e', script])
    time.sleep(0.5)

def get_group_index(group_name):
    """获取分组在列表中的索引"""
    try:
        return ACCOUNT_GROUPS.index(group_name)
    except ValueError:
        return -1

def parse_txt(txt_path):
    """解析文案txt文件，支持多行正文"""
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
                # 保存上一个字段
                if current_key:
                    data[current_key] = '\n'.join(current_value).strip()
                current_key = key
                value = line[len(key):].lstrip('：:').strip()
                current_value = [value] if value else []
                matched = True
                break
        if not matched and current_key:
            current_value.append(line)

    # 保存最后一个字段
    if current_key:
        data[current_key] = '\n'.join(current_value).strip()

    return data

def wait_until(time_str):
    """等待到指定时间再执行发布"""
    if not time_str:
        return
    now = datetime.now()
    target = datetime.strptime(f"{now.strftime('%Y-%m-%d')} {time_str}", '%Y-%m-%d %H:%M')
    if target < now:
        return
    wait_secs = (target - now).total_seconds()
    print(f"  ⏰ 等待到 {time_str} 发布，还有 {int(wait_secs//60)} 分钟...")
    time.sleep(wait_secs)

# 全局变量，追踪是否已经点过一键发布
_app_opened = False

def publish_one(folder_path, first=False):
    """发布一条内容，一个文件夹可以有多个txt分别发布"""
    folder = Path(folder_path)
    txt_files = sorted(folder.glob('*.txt'))
    if not txt_files:
        print(f"  ❌ 找不到txt文件：{folder}")
        return False

    for idx, txt_file in enumerate(txt_files):
        # 只处理小红书和抖音的txt
        with open(txt_file, 'r', encoding='utf-8') as f:
            content = f.read()
        if '04视频号' in content or '13公众号' in content:
            continue
        print(f"  📄 处理：{txt_file.name}")
        data = parse_txt(txt_file)
        print(f"  📄 标题：{data['标题']}")
        print(f"  👤 账号：{data['账号']}")
        print(f"  ⏰ 时间：{data['时间']}")

        images = sorted([str(f) for f in folder.iterdir() if f.suffix.lower() in {'.jpg', '.jpeg', '.png', '.webp'}])
        if not images:
            print(f"  ❌ 没有找到图片")
            return False
        print(f"  🖼  图片：{len(images)} 张")

        wait_until(data['时间'])
        _do_publish(folder, data, images, first=(first and idx == 0))

        if idx < len(txt_files) - 1:
            print(f"  ⏳ 等待30秒后发布下一个账号...")
            time.sleep(30)

    return True

def _do_publish(folder, data, images, first=False):

    # 激活小豆芽
    subprocess.run(['osascript', '-e', 'tell application "新榜小豆芽" to activate'])
    time.sleep(1)

    # 0. 打开小豆芽（只在第一条时执行）
    global _app_opened
    if not _app_opened:
        print("  🌱 打开小豆芽...")
        subprocess.run(['open', '-a', '新榜小豆芽'])
        time.sleep(3)
        print("  🖥  全屏...")
        click(*POS_FULLSCREEN)
        time.sleep(1)
        _app_opened = True

    # 每次都点一键发布
    click(*POS_ONE_KEY)
    time.sleep(0.5)

    # 每条都点图片发布
    click(*POS_IMG_PUBLISH)
    time.sleep(1.5)

    # 如果弹出恢复草稿弹窗，点「放弃」（暂时跳过，出现时再采集坐标）
    # click(*POS_DISCARD)
    # time.sleep(0.8)

    # 1. 一键清空
    print("  🧹 清空内容...")
    click(*POS_CLEAR)
    time.sleep(0.8)
    # 点确认弹窗
    click(*POS_CLEAR_CONFIRM)
    time.sleep(0.8)

    # 2. 填写标题
    print("  ✏️  填写标题...")
    click(*POS_TITLE)
    time.sleep(0.3)
    type_text(data['标题'])

    # 3. 填写正文+话题
    print("  ✏️  填写正文...")
    click(*POS_CONTENT)
    time.sleep(0.3)
    full_content = data['正文']
    # 抖音不在正文加话题
    if data['话题'] and '07抖音发作品' not in data['账号']:
        full_content += '\n' + data['话题']
    type_text(full_content)

    # 4. 上传图片
    print("  🖼  上传图片...")
    click(*POS_UPLOAD)
    time.sleep(2)
    swift_key(0x22, cmd=True, shift=True)  # Command+Shift+G
    time.sleep(1)
    process = subprocess.Popen(['pbcopy'], stdin=subprocess.PIPE)
    process.communicate(str(folder).encode('utf-8'))
    time.sleep(0.3)
    swift_cmdv()
    time.sleep(0.5)
    swift_enter()
    time.sleep(1)
    swift_cmda()
    time.sleep(0.3)
    swift_enter()
    time.sleep(2)

    # 5. 选择账号
    print("  👤 选择账号...")
    click(*POS_SELECT_ACCT)
    time.sleep(1)

    # 点按列表选择tab
    click(*POS_TAB_LIST)
    time.sleep(0.8)

    # 解析分组和账号名
    if '|' in data['账号']:
        group_name, account_name = data['账号'].split('|', 1)
        group_name = group_name.strip()
        account_name = account_name.strip()
    else:
        group_name = ''
        account_name = data['账号'].strip()

    # 先清空搜索框
    click(*POS_SEARCH_CLEAR)
    time.sleep(0.5)

    # 点击对应分组
    if '06红书店铺' in group_name:
        click(*POS_GROUP_06)
    elif '07抖音发作品' in group_name:
        click(*POS_GROUP_07)
    else:
        click(*POS_GROUP_14)
    time.sleep(0.5)

    # 点搜索框输入账号名
    click(*POS_SEARCH_BOX)
    time.sleep(0.5)
    type_text(account_name)
    time.sleep(1)

    # 点第一个搜索结果
    click(*POS_FIRST_RESULT)
    time.sleep(0.5)

    # 点确定
    click(*POS_CONFIRM)
    time.sleep(1)

    # 6. 同步至右侧（图像识别自动定位）
    print("  🔄 同步至右侧...")
    find_and_click_sync()
    time.sleep(1)

    # 7. 一键发布
    print("  🚀 发布...")
    click(*POS_PUBLISH)
    time.sleep(2)

    # 点「发布至平台」确认
    time.sleep(60)
    click(*POS_CONFIRM_PUB)
    time.sleep(60)

    # 关闭发布记录和图片发布tab
    print("  ❌ 关闭发布记录...")
    click(*POS_CLOSE_RECORD)
    time.sleep(0.5)
    print("  ❌ 关闭图片发布...")
    click(*POS_CLOSE_IMGPUB)
    time.sleep(0.5)

    # 收尾：点一键发布
    print("  🔄 收尾点一键发布...")
    click(*POS_ONE_KEY)
    time.sleep(0.5)

    print(f"  ✅ 发布完成：{folder.name}\n")
    return True

def pick_folder():
    """弹出文件夹选择"""
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
    print("🌱 新榜小豆芽自动发布脚本启动\n")

    if len(sys.argv) > 1:
        main_folder = Path(sys.argv[1])
    else:
        main_folder = Path(pick_folder())

    # 判断传入的是直接发布文件夹还是主文件夹
    txt_files = list(main_folder.glob('*.txt'))
    if txt_files:
        # 直接是发布文件夹（从总控脚本调用）
        subfolders = [main_folder]
    else:
        # 是主文件夹，找子文件夹
        subfolders = sorted([f for f in main_folder.iterdir() if f.is_dir()])

    if not subfolders:
        print("❌ 找不到内容")
        sys.exit(1)

    print(f"📂 找到 {len(subfolders)} 条内容待发布\n")

    for i, folder in enumerate(subfolders, 1):
        print(f"[{i}/{len(subfolders)}] 开始处理：{folder.name}")
        publish_one(folder, first=(i == 1))
        if i < len(subfolders):
            print(f"  ⏳ 等待30秒后开始下一篇...")
            time.sleep(30)

    print("🎉 全部发布完成！")
