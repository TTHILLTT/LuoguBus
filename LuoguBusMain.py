import requests
import time
import sys
import json
import csv
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side


def fetch_luogu_submissions(luogu_uid, client_id, count=50):
    """获取洛谷提交记录（按时间升序排列）"""
    url = "https://www.luogu.com.cn/record/list"

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
        "Referer": f"https://www.luogu.com.cn/user/{luogu_uid}",
        "Cookie": f"__client_id={client_id}; _uid={luogu_uid}",
        "X-Luogu-Type": "content-only",
        "X-Requested-With": "XMLHttpRequest"
    }

    params = {
        "user": luogu_uid,
        "page": 1,
        "_contentOnly": 1
    }

    try:
        response = requests.get(url, headers=headers, params=params, timeout=15)
        response.raise_for_status()
        data = response.json()

        if data.get('code', 0) != 200:
            print(f"API错误: {data.get('currentData', {}).get('errorMessage', '未知错误')}")
            return []

        records = data['currentData']['records']['result']
        records.sort(key=lambda x: x['submitTime'])  # 按时间升序排序
        return records[:min(count, len(records))]
    except Exception as e:
        print(f"获取数据失败: {str(e)}")
        return []


def create_excel(records, filename):
    """创建美观的Excel文件"""
    wb = Workbook()
    ws = wb.active
    ws.title = "刷题记录"

    # 设置列宽
    column_widths = {
        'A': 25, 'B': 25, 'C': 25,
        'D': 25, 'E': 25, 'F': 25
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # 表头样式
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = "FF4F81BD"  # 蓝色背景
    alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # 写入表头
    headers = [
        "提交日期", "题号", "题目名称", "状态","运行时间", "内存占用"
    ]
    ws.append(headers)

    # 应用表头样式
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = openpyxl.styles.PatternFill(start_color=header_fill, end_color=header_fill, fill_type="solid")
        cell.alignment = alignment
        cell.border = thin_border

    # 数据样式
    data_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    status_colors = {
        "AC": "FF00B050",  # 绿色
        "WA": "FFFF0000",  # 红色
        "TLE": "FFFFC000",  # 橙色
        "MLE": "FF7030A0",  # 紫色
        "RE": "FFFF0000",  # 红色
        "CE": "FF000000"  # 黑色
    }

    # 写入数据
    for record in records:
        submit_time = datetime.fromtimestamp(record['submitTime'])
        problem = record.get('problem', {})

        # 状态和难度处理
        status = {
            12: "AC", 7: "WA", 4: "TLE",
            5: "MLE", 6: "RE", 2: "CE", 14: "Half AC"
        }.get(record['status'], f"未知({record['status']})")

        row = [
            submit_time.strftime("%Y-%m-%d %H:%M"),
            problem.get('pid', '未知'),
            problem.get('title', '未知'),
            status,
            f"{record.get('time', 0)}ms",
            f"{record.get('memory', 0)}KB",
        ]
        ws.append(row)

        # 应用数据样式
        last_row = ws.max_row
        for col in range(1, 11):
            cell = ws.cell(row=last_row, column=col)
            cell.alignment = data_alignment
            cell.border = thin_border

            # 状态列着色
            if col == 5 and status in status_colors:
                cell.fill = openpyxl.styles.PatternFill(
                    start_color=status_colors[status],
                    end_color=status_colors[status],
                    fill_type="solid"
                )
                cell.font = Font(color="FFFFFF" if status in ["WA", "RE", "CE"] else "000000")

    # 冻结首行
    ws.freeze_panes = "A2"

    # 保存文件
    wb.save(filename)
    print(f"✓ 已生成Excel文件: {filename}")


def create_csv(records, filename):
    """创建CSV文件"""
    if not records:
        return

    # 准备表头
    fieldnames = [
        "submit_time", "problem_id", "problem_name","status", "run_time", "memory_usage"
    ]

    # 准备数据
    data = []
    for record in records:
        submit_time = datetime.fromtimestamp(record['submitTime'])
        problem = record.get('problem', {})

        status = {
            12: "AC", 7: "WA", 4: "TLE",
            5: "MLE", 6: "RE", 2: "CE", 14: "Half AC"
        }.get(record['status'], f"Unknown({record['status']})")

        data.append({
            "submit_time": submit_time.strftime("%Y-%m-%d %H:%M:%S"),
            "problem_id": problem.get('pid', 'Unknown'),
            "problem_name": problem.get('title', 'Unknown'),
            "status": status,
            "run_time": record.get('time', 0),
            "memory_usage": record.get('memory', 0)
        })

    # 写入CSV
    with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)

    print(f"✓ 已生成CSV文件: {filename}")


def main():
    print("洛谷做题日记生成器 v3.1")
    print("=" * 60)
    print("将同时生成.xlsx和.csv文件，按提交时间升序排列")
    print("=" * 60)

    # 获取必要信息
    client_id = input("请输入__client_id的值: ").strip()
    luogu_uid = input("请输入_uid的值: ").strip()

    if not client_id or not luogu_uid:
        print("错误: 必须提供Cookie信息")
        sys.exit(1)

    # 获取提交记录
    print(f"\n正在获取用户 {luogu_uid} 的提交记录...")
    records = fetch_luogu_submissions(luogu_uid, client_id, 100)

    if not records:
        print("获取提交记录失败，请检查：")
        print("1. Cookie信息是否正确（需包含__client_id和_uid）")
        print("2. 账号是否有公开提交记录")
        sys.exit(1)

    # 生成文件名
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    base_filename = f"Luogu_Diary_{luogu_uid}_{timestamp}"

    # 生成两种格式文件
    create_excel(records, f"{base_filename}.xlsx")
    create_csv(records, f"{base_filename}.csv")

    # 使用提示
    print("\n使用说明:")
    print(f"1. Excel文件 ({base_filename}.xlsx):")
    print("   - 美观的格式化表格")
    print("   - 状态自动着色")
    print("   - 适合直接查看和编辑")
    print(f"2. CSV文件 ({base_filename}.csv):")
    print("   - 纯文本格式")
    print("   - 适合程序处理或导入数据库")
    print("3. 两个文件内容相同，按提交时间升序排列")


if __name__ == "__main__":
    try:
        import openpyxl
    except ImportError:
        print("错误: 需要安装openpyxl库，请执行: pip install openpyxl")
        sys.exit(1)

    main()