import requests
import time
import sys
import json
import csv
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


def fetch_luogu_submissions(luogu_uid, client_id, count=50):
    """获取洛谷提交记录，确保获取最新记录并按时间升序排列"""
    url = "https://www.luogu.com.cn/record/list"

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
        "Referer": f"https://www.luogu.com.cn/user/{luogu_uid}",
        "Cookie": f"__client_id={client_id}; _uid={luogu_uid}",
        "X-Luogu-Type": "content-only",
        "X-Requested-With": "XMLHttpRequest"
    }

    all_records = []
    page = 1
    per_page = 20  # 洛谷每页固定返回20条记录
    total_needed = count

    try:
        # 计算需要抓取的页数（向上取整）
        pages_needed = (total_needed + per_page - 1) // per_page

        # 从第一页开始获取最新记录
        while len(all_records) < total_needed and page <= pages_needed:
            params = {
                "user": luogu_uid,
                "page": page,
                "_contentOnly": 1
            }

            response = requests.get(url, headers=headers, params=params, timeout=15)
            response.raise_for_status()
            data = response.json()

            if data.get('code', 0) != 200:
                error_msg = data.get('currentData', {}).get('errorMessage', '未知错误')
                print(f"第 {page} 页API错误: {error_msg}")
                break

            records = data['currentData']['records']['result']
            if not records:
                break  # 没有更多记录了

            # 将本页记录按时间升序排列后添加到总列表
            records.sort(key=lambda x: x['submitTime'])
            all_records.extend(records)

            # 显示进度
            print(f"已获取第 {page} 页，共 {len(records)} 条记录")

            # 如果已经获取到足够数量的记录
            if len(all_records) >= total_needed:
                break

            page += 1  # 获取下一页
            time.sleep(0.5)  # 避免请求过于频繁

        # 按时间升序排序（越早的记录越靠前）
        all_records.sort(key=lambda x: x['submitTime'])

        # 确保不超过请求的数量（取最新的count条）
        return all_records[-min(total_needed, len(all_records)):]
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
        'A': 20, 'B': 15, 'C': 40,
        'D': 10, 'E': 15, 'F': 15
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    # 表头样式
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # 写入表头
    headers = [
        "提交日期", "题号", "题目名称", "状态", "运行时间", "内存占用"
    ]
    ws.append(headers)

    # 应用表头样式
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = alignment
        cell.border = thin_border

    # 数据样式
    data_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    status_colors = {
        "AC": "00B050",  # 绿色
        "WA": "FF0000",  # 红色
        "TLE": "FFC000",  # 橙色
        "MLE": "7030A0",  # 紫色
        "RE": "FF0000",  # 红色
        "CE": "000000",  # 黑色
        "Half AC": "00CED1",  # 青色
        "UKE": "000000"  # 黑色
    }

    # 写入数据
    for record in records:
        submit_time = datetime.fromtimestamp(record['submitTime'])
        problem = record.get('problem', {})

        # 状态处理（包含UKE状态）
        status = {
            12: "AC", 7: "WA", 4: "TLE",
            5: "MLE", 6: "RE", 2: "CE",
            14: "Half AC", 11: "UKE"
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
        for col in range(1, 7):  # 共6列
            cell = ws.cell(row=last_row, column=col)
            cell.alignment = data_alignment
            cell.border = thin_border

            # 状态列着色（第4列）
            if col == 4 and status in status_colors:
                cell.fill = PatternFill(
                    start_color=status_colors[status],
                    end_color=status_colors[status],
                    fill_type="solid"
                )
                # 深色背景使用白色字体
                cell.font = Font(color="FFFFFF" if status in ["WA", "RE", "CE", "UKE"] else "000000")

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
        "submit_time", "problem_id", "problem_name", "status", "run_time", "memory_usage"
    ]

    # 准备数据
    data = []
    for record in records:
        submit_time = datetime.fromtimestamp(record['submitTime'])
        problem = record.get('problem', {})

        status = {
            12: "AC", 7: "WA", 4: "TLE",
            5: "MLE", 6: "RE", 2: "CE",
            14: "Half AC", 11: "UKE"
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
    banner = f"""
                ██╗     ██╗   ██╗ ██████╗  ██████╗ ██╗   ██╗
                ██║     ██║   ██║██╔═══██╗██╔════╝ ██║   ██║
                ██║     ██║   ██║██║   ██║██║  ███╗██║   ██║
                ██║     ██║   ██║██║   ██║██║   ██║██║   ██║  
                ███████╗╚██████╔╝╚██████╔╝╚██████╔╝╚██████╔╝
                ╚══════╝ ╚═════╝  ╚═════╝  ╚═════╝  ╚═════╝
                ██████╗ ██╗   ██╗███████╗
                ██╔══██╗██║   ██║██╔════╝
                ██████╔╝██║   ██║███████╗
                ██╔══██╗██║   ██║╚════██║
                ██████╔╝╚██████╔╝███████║
                ╚═════╝  ╚═════╝ ╚══════╝
            """
    print(banner)
    print("洛谷做题日记生成器 v3.8")
    print("=" * 60)
    print("优化记录获取逻辑：确保获取最新提交并按时间升序排列")
    print("支持多页抓取，最多可获取1000条提交记录（避免频繁请求被Ban）")
    print("=" * 60)

    # 获取必要信息
    client_id = input("请输入__client_id的值: ").strip()
    luogu_uid = input("请输入_uid的值: ").strip()

    # 询问记录数量（最大1000）
    while True:
        try:
            count_input = input("请输入要获取的记录数量 (1-1000, 默认50): ").strip()
            if not count_input:
                count = 50
                break

            count = int(count_input)
            if 1 <= count <= 1000:
                break
            else:
                print("请输入1到1000之间的整数！")
        except ValueError:
            print("请输入有效的整数！")

    if not client_id or not luogu_uid:
        print("错误: 必须提供Cookie信息")
        sys.exit(1)

    # 获取提交记录
    print(f"\n正在获取用户 {luogu_uid} 的最新 {count} 条提交记录...")
    records = fetch_luogu_submissions(luogu_uid, client_id, count)

    if not records:
        print("获取提交记录失败，请检查：")
        print("1. Cookie信息是否正确（需包含__client_id和_uid）")
        print("2. 账号是否有公开提交记录")
        sys.exit(1)

    actual_count = len(records)
    if actual_count < count:
        print(f"⚠️ 注意: 只获取到 {actual_count} 条记录（请求数量: {count}）")
    else:
        print(f"✅ 成功获取 {actual_count} 条提交记录")
        # 显示时间范围
        first_submit = datetime.fromtimestamp(records[0]['submitTime']).strftime("%Y-%m-%d %H:%M")
        last_submit = datetime.fromtimestamp(records[-1]['submitTime']).strftime("%Y-%m-%d %H:%M")
        print(f"📅 时间范围: {first_submit} 至 {last_submit}")

    # 生成文件名
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    base_filename = f"Luogu_Diary_{luogu_uid}_{timestamp}"

    # 生成两种格式文件
    create_excel(records, f"{base_filename}.xlsx")
    create_csv(records, f"{base_filename}.csv")

    # 使用提示
    print("\n使用说明:")
    print(f"1. Excel文件 ({base_filename}.xlsx):")
    print("   - 表格按提交时间升序排列（最早的在最上面）")
    print("   - 美观的格式化表格，状态自动着色")
    print(f"2. CSV文件 ({base_filename}.csv):")
    print("   - 纯文本格式，适合程序处理")
    print(f"3. 包含 {actual_count} 条记录，时间从 {first_submit} 到 {last_submit}")
    print("\n提示：避免频繁请求大量数据，以防被洛谷封禁！")


if __name__ == "__main__":
    try:
        import openpyxl
    except ImportError:
        print("错误: 需要安装openpyxl库，请执行: pip install openpyxl")
        sys.exit(1)
    try:
        import requests
    except ImportError:
        print("错误: 需要安装requests库，请执行: pip install requests")
        sys.exit(1)

    main()
