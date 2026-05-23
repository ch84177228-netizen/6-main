#!/usr/bin/env python3
"""
坐标采集工具 — 只采集空白处 和 + 加号
移动鼠标到目标位置，按 Enter 记录
"""

import subprocess


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


def capture(label):
    while True:
        input(f"  → 移鼠标到【{label}】，按 Enter 记录: ")
        pos = get_mouse_pos()
        if pos:
            print(f"     ✅ {pos}\n")
            return pos
        print("     ❌ 读取失败，请重试")


def main():
    print("=" * 45)
    print("  坐标采集  —  空白处 & 加号")
    print("=" * 45 + "\n")

    blank   = capture("空白处（填完正文后的空白区域）")
    add_img = capture("+ 加号（添加图片按钮）")

    print("=" * 45)
    print("采集完成，请将以下两行替换到 wechat_publish.py：")
    print("=" * 45)
    print(f"POS_BLANK        = {blank}")
    print(f"POS_ADD_IMG      = {add_img}")


if __name__ == '__main__':
    main()
