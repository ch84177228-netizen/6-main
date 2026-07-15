"""
把引流主号31个txt按组整理到子文件夹，重新打包ZIP
"""
import os
import shutil
import zipfile
import glob

BASE = "/home/user/6-main/01-6个主号内容创作/引流主号_发布文案"

# 定义文件编号与组的映射
group_map = {
    "第1组_主号":    [1, 2, 3, 4],
    "第2组_副号1":   [5, 6, 7, 8],
    "第3组_副号2":   [9, 10, 11, 12],
    "第4组_副号3":   [13, 14, 15, 16],
    "第5组_副号4":   [17, 18, 19],
    "第6组_副号5":   [20, 21, 22],
    "第7组_西游记":  [23, 24],
    "第8组_小新故事": [25, 26],
    "第9组_治愈猫咪": [27, 28],
    "第10组_西游记2": [29, 30],
    "第11组_小新2":  [31],
}

# 建立 no → 组名 的反查表
no_to_group = {}
for group, nos in group_map.items():
    for n in nos:
        no_to_group[n] = group

# 创建子文件夹
for group in group_map:
    os.makedirs(os.path.join(BASE, group), exist_ok=True)

# 找到所有已存在的txt（文件名格式：01_平台_标题.txt）
txt_files = sorted(glob.glob(os.path.join(BASE, "*.txt")))
moved = []
for fpath in txt_files:
    fname = os.path.basename(fpath)
    # 取编号
    try:
        no = int(fname.split("_")[0])
    except ValueError:
        continue
    group = no_to_group.get(no)
    if group:
        dest = os.path.join(BASE, group, fname)
        shutil.copy2(fpath, dest)
        moved.append((fname, group))
        print(f"✅ {fname} → {group}/")

print(f"\n共整理 {len(moved)} 个txt文件")

# 重新打包ZIP（含文件夹结构）
zip_path = os.path.join(BASE, "主号引流文案_31篇_分组文件夹.zip")
with zipfile.ZipFile(zip_path, "w") as zf:
    for group in group_map:
        folder = os.path.join(BASE, group)
        for f in sorted(os.listdir(folder)):
            if f.endswith(".txt"):
                zf.write(os.path.join(folder, f), os.path.join(group, f))
    # 也打包两个Word
    for f in os.listdir(BASE):
        if f.endswith(".docx"):
            zf.write(os.path.join(BASE, f), f)

print(f"✅ ZIP: {zip_path}")
print("完成！")
