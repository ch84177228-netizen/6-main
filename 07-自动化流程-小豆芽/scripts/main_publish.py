#!/usr/bin/env python3
"""
总控发布脚本
自动识别txt里的账号分组，选择对应的发布流程：
- 06红书店铺 / 14芳芳财会 / 07抖音发作品 → 小红书/抖音流程
- 04视频号 → 视频号流程
- 13公众号 → 公众号流程
"""

import os, sys, subprocess, time
from pathlib import Path

# 导入各平台发布模块
sys.path.insert(0, str(Path.home() / 'Desktop'))

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

def parse_account(txt_path):
    """只读取账号字段的分组名"""
    try:
        with open(txt_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if '账号' in line and ('：' in line or ':' in line):
                    # 取冒号后面的内容
                    if '：' in line:
                        value = line.split('：', 1)[1].strip()
                    else:
                        value = line.split(':', 1)[1].strip()
                    # 取|前面的分组名
                    group = value.split('|')[0].strip()
                    return group
    except:
        pass
    return ''

def get_platform(group_name):
    """根据分组名判断平台"""
    if '04视频号' in group_name:
        return 'video'
    elif '13公众号' in group_name:
        return 'wechat'
    elif '07抖音发作品' in group_name:
        return 'douyin'
    else:
        return 'xiaohongshu'

def run_script(script_name, folder_path):
    """运行对应的发布脚本"""
    script_path = Path.home() / 'Desktop' / script_name
    result = subprocess.run([
        '/Library/Developer/CommandLineTools/usr/bin/python3',
        str(script_path),
        str(folder_path)
    ])
    return result.returncode == 0

if __name__ == "__main__":
    print("🚀 总控发布脚本启动\n")

    if len(sys.argv) > 1:
        main_folder = sys.argv[1]
    else:
        main_folder = pick_folder()

    main_folder = Path(main_folder)
    subfolders = sorted([f for f in main_folder.iterdir() if f.is_dir()])

    if not subfolders:
        print("❌ 文件夹里没有子文件夹")
        sys.exit(1)

    print(f"📂 找到 {len(subfolders)} 条内容待发布\n")

    for i, folder in enumerate(subfolders, 1):
        print(f"[{i}/{len(subfolders)}] 开始处理：{folder.name}")

        # 读取该文件夹里所有txt，判断平台
        txt_files = sorted(folder.glob('*.txt'))
        if not txt_files:
            print(f"  ❌ 找不到txt文件，跳过")
            continue

        # 按平台分组txt
        douyin_txts = []
        xhs_txts = []
        video_txts = []
        wechat_txts = []

        for txt in txt_files:
            group = parse_account(txt)
            platform = get_platform(group)
            if platform == 'video':
                video_txts.append(txt)
            elif platform == 'wechat':
                wechat_txts.append(txt)
            elif '07抖音发作品' in group:
                douyin_txts.append(txt)
            else:
                xhs_txts.append(txt)

        print(f"  📊 抖音：{len(douyin_txts)}个，小红书：{len(xhs_txts)}个，视频号：{len(video_txts)}个，公众号：{len(wechat_txts)}个")

        # 按顺序发布：抖音→小红书→视频号→公众号
        if douyin_txts:
            print(f"  🎵 运行抖音发布...")
            run_script('auto_publish.py', folder)
            if xhs_txts or video_txts or wechat_txts:
                print(f"  ⏳ 等待30秒...")
                time.sleep(30)

        if xhs_txts:
            print(f"  📱 运行小红书发布...")
            run_script('auto_publish.py', folder)
            if video_txts or wechat_txts:
                print(f"  ⏳ 等待30秒...")
                time.sleep(30)

        if video_txts:
            print(f"  🎬 运行视频号发布...")
            run_script('video_publish.py', folder)
            if wechat_txts:
                print(f"  ⏳ 等待30秒...")
                time.sleep(30)

        if wechat_txts:
            print(f"  📰 运行公众号发布...")
            run_script('wechat_publish.py', folder)

        if i < len(subfolders):
            print(f"  ⏳ 等待30秒后开始下一篇...")
            time.sleep(30)

    print("\n🎉 全部发布完成！")
