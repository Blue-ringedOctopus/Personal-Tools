import os
import re
from openpyxl import Workbook

def parse_date_from_folder_name(folder_name):
    """
    从类似 '2025年11月13日 xx门诊 成员姓名' 的字符串中提取日期部分，
    支持 '日' 或 '号' 结尾，并转换为 '2025/11/13' 格式。
    """
    # 匹配 2025年11月13日 或 2025年11月13号
    pattern = r'(\d{4})年(\d{1,2})月(\d{1,2})[日号]'
    match = re.search(pattern, folder_name)
    if match:
        year, month, day = match.groups()
        return f"{year}/{int(month)}/{int(day)}"
    return None

def collect_patients(root_dir):
    """
    遍历根目录下深度为2的文件夹，返回列表 [(姓名, 日期), ...]
    """
    results = []
    for date_folder in os.listdir(root_dir):
        date_path = os.path.join(root_dir, date_folder)
        if not os.path.isdir(date_path):
            continue
        date_str = parse_date_from_folder_name(date_folder)
        if not date_str:
            print(f"警告：无法从文件夹名解析日期: {date_folder}")
            continue

        for patient_folder in os.listdir(date_path):
            patient_path = os.path.join(date_path, patient_folder)
            if not os.path.isdir(patient_path):
                continue
            patient_name = patient_folder.strip()
            if patient_name:
                results.append((patient_name, date_str))
    return results

def create_excel_with_data(excel_path, data):
    """
    新建一个 Excel 工作簿，将 data 写入 B 列（姓名）和 C 列（日期），
    并添加表头。如果文件已存在则覆盖（可改为追加时间戳）。
    """
    wb = Workbook()
    ws = wb.active

    # 写入表头
    ws['B1'] = '姓名'
    ws['C1'] = '日期'

    # 写入数据（从第2行开始）
    for idx, (name, date) in enumerate(data, start=2):
        ws[f'B{idx}'] = name
        ws[f'C{idx}'] = date

    # 保存（若文件已存在会直接覆盖）
    wb.save(excel_path)
    print(f"✅ 成功生成新表格，共写入 {len(data)} 条记录，保存至：{excel_path}")

def main():
    # ========= 请根据实际情况修改以下两个路径 =========
    root_dir = r"更换地址"
    # 指定新 Excel 文件的完整路径（扩展名可用 .xlsx 或 .xlsm）
    new_excel_path = r"更换地址"
    # =================================================

    # 可选：如果不想覆盖已有文件，可自动添加时间戳
    # from datetime import datetime
    # base, ext = os.path.splitext(new_excel_path)
    # new_excel_path = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"

    print("开始遍历文件夹...")
    patient_data = collect_patients(root_dir)
    if not patient_data:
        print("⚠️ 未找到任何患者数据，请检查根目录路径是否正确。")
        return

    print(f"共找到 {len(patient_data)} 条记录。")
    create_excel_with_data(new_excel_path, patient_data)
    print("🎯 处理完成！")

if __name__ == "__main__":
    main()
