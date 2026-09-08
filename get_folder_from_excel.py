import re
import os
import shutil
from datetime import datetime
from openpyxl import load_workbook
import bisect
from collections import defaultdict

def read_excel_dates(excel_path):
    """从Excel的B、C列读取姓名和日期，返回列表[(姓名, 日期字符串YYYY-MM-DD), ...]"""
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active
    records = []
    for row in ws.iter_rows(min_row=2, max_col=3, values_only=True):
        name = row[1]   # B列
        date_val = row[2]  # C列
        if not name or not date_val:
            continue
        if isinstance(date_val, datetime):
            date_str = date_val.strftime("%Y-%m-%d")
        else:
            for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d"):
                try:
                    dt = datetime.strptime(str(date_val), fmt)
                    date_str = dt.strftime("%Y-%m-%d")
                    break
                except ValueError:
                    continue
            else:
                print(f"警告：无法解析日期 {date_val}，跳过")
                continue
        records.append((str(name).strip(), date_str))
    return records

def build_source_index(source_root):
    """
    遍历源根目录，建立患者到所有就诊日期路径的映射。
    递归查找所有含'Aurum'的文件夹，将其下包含日期子目录的文件夹视为患者目录。
    返回: {患者名: [(日期字符串, 完整路径), ...]}
    """
    index = {}
    # 遍历第一层日期文件夹
    for date_folder in os.listdir(source_root):
        date_path = os.path.join(source_root, date_folder)
        if not os.path.isdir(date_path):
            continue
        # 查找包含'Aurum'的子文件夹
        for sub in os.listdir(date_path):
            sub_path = os.path.join(date_path, sub)
            if not os.path.isdir(sub_path):
                continue
            if 'Aurum' not in sub:
                continue
            # 递归遍历 sub_path 下所有目录
            for root, dirs, files in os.walk(sub_path):
                # 检查当前目录下是否有符合 YYYY-MM-DD 格式的子目录
                for d in dirs:
                    if re.match(r'\d{4}-\d{2}-\d{2}', d):
                        patient_dir = root      # 患者文件夹路径
                        patient_name = os.path.basename(patient_dir)
                        date_dir_path = os.path.join(patient_dir, d)
                        if patient_name not in index:
                            index[patient_name] = []
                        index[patient_name].append((d, date_dir_path))
                # 注意：os.walk 会继续深入，但进入日期子目录后，它的子目录没有日期格式，所以不会重复添加

    # 对每个患者的日期列表排序
    for patient in index:
        index[patient].sort(key=lambda x: x[0])
    return index

def copy_adjacent_records(records, source_index, target_root):
    """
    按患者分组，每个患者基于第一次出现的日期（基准），
    在所有就诊日期中找出离基准日期最近的前一次和后一次（时间轴上相邻），复制这两个文件夹（若存在）。
    不再要求基准日期必须存在。
    """
    # 分组
    patient_groups = defaultdict(list)
    for name, date in records:
        patient_groups[name].append(date)

    total_patients = len(patient_groups)
    stats = {
        'matched_had_neighbor': 0,   # 找到前或后至少一个
        'no_neighbor': 0,           # 无前无后（只有一次就诊或完全没有）
        'no_index': 0,             # 患者不在源索引中
        'copied_at_least_one': 0,
    }

    copied_count = 0
    skipped_count = 0
    copied_targets = set()

    for name, dates in patient_groups.items():
        target_date = dates[0]  # 取第一次出现的日期作为基准

        if name not in source_index:
            print(f"未找到患者 '{name}' 的索引，跳过")
            stats['no_index'] += 1
            continue

        all_dates = source_index[name]  # list of (date_str, full_path)
        # 所有日期字符串（已排序）
        date_strs = [d[0] for d in all_dates]

        # 使用二分查找基准日期的插入位置
        idx = bisect.bisect_left(date_strs, target_date)
        neighbors = []
        # 前一次：插入位置-1（如果存在）
        if idx > 0:
            neighbors.append(all_dates[idx - 1])
        # 后一次：插入位置（如果存在且不等于基准日期本身，但若基准日期在列表中，则该位置就是基准日期本身，需取下一个）
        # 但我们要找的是大于基准日期的最小值，所以如果插入位置等于基准日期，则需后移一位
        if idx < len(date_strs) and date_strs[idx] == target_date:
            # 基准日期在列表中，取下一个
            if idx + 1 < len(date_strs):
                neighbors.append(all_dates[idx + 1])
        elif idx < len(date_strs):
            # 基准日期不在列表中，插入位置就是第一个大于基准日期的
            neighbors.append(all_dates[idx])

        # 如果基准日期本身在列表中，但 idx 指向它，我们之前已经排除了它本身，所以不会重复
        # 现在去重（可能前一次和后一次是同一个？不可能，除非只有一个日期，但那样不会有两个邻居）
        # 但若基准日期在列表最前面，idx=0，只有后一次；若在最后，只有前一次。

        if not neighbors:
            print(f"患者 '{name}' 无相邻日期，跳过")
            stats['no_neighbor'] += 1
            continue

        stats['matched_had_neighbor'] += 1
        print(f"处理患者 '{name}'，基准日期 {target_date}，复制相邻日期: {[d[0] for d in neighbors]}")

        patient_copied = False
        for date_str, full_path in neighbors:
            dest_patient_dir = os.path.join(target_root, name)
            dest_dir = os.path.join(dest_patient_dir, date_str)

            if dest_dir in copied_targets:
                continue
            if os.path.exists(dest_dir):
                print(f"  目标已存在，跳过: {dest_dir}")
                skipped_count += 1
                copied_targets.add(dest_dir)
                continue

            os.makedirs(dest_patient_dir, exist_ok=True)
            try:
                shutil.copytree(full_path, dest_dir, ignore_dangling_symlinks=True)
                print(f"  已复制: {full_path} -> {dest_dir}")
                copied_count += 1
                copied_targets.add(dest_dir)
                patient_copied = True
            except Exception as e:
                print(f"  复制失败 {full_path} : {e}")

        if patient_copied:
            stats['copied_at_least_one'] += 1

    # 打印统计报告
    print("\n" + "="*50)
    print("处理统计报告")
    print(f"总患者数（Excel去重后）   : {total_patients}")
    print(f"找到相邻日期的患者数      : {stats['matched_had_neighbor']}")
    print(f"  └─ 至少复制了1个文件夹  : {stats['copied_at_least_one']}")
    print(f"无相邻日期的患者数        : {stats['no_neighbor']}")
    print(f"未找到索引的患者数        : {stats['no_index']}")
    print(f"成功复制的文件夹总数      : {copied_count}")
    print(f"因目标已存在而跳过的文件夹: {skipped_count}")
    print("="*50)

    return copied_count, skipped_count

def main():
    source_root = r"补充地址"
    excel_path = r"补充地址"
    target_root = r"补充地址"

    print("正在读取Excel...")
    records = read_excel_dates(excel_path)
    print(f"共读取 {len(records)} 条记录。")

    print("正在构建源索引（可能需要一些时间）...")
    source_index = build_source_index(source_root)
    print(f"索引构建完成，共 {len(source_index)} 位患者。")

    print("开始复制相邻日期文件夹...")
    copied, skipped = copy_adjacent_records(records, source_index, target_root)
    print(f"处理完成。成功复制 {copied} 个文件夹，跳过 {skipped} 个（已存在或错误）。")

if __name__ == "__main__":
    main()
