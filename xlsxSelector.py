import pandas as pd
import numpy as np
import os
import sys
from pathlib import Path


# ========================
# 工具函数
# ========================

def get_user_choice(prompt, valid_choices, default=None):
    """通用选择函数"""
    while True:
        choice = input(prompt).strip()
        if not choice and default is not None:
            return default
        if choice in valid_choices:
            return choice
        print(f"输入无效，请输入 {'/'.join(valid_choices)}")


def exit_or_continue():
    """询问是否继续使用工具"""
    print("\n" + "="*50)
    choice = get_user_choice(
        "请选择: 1) 返回主菜单  2) 退出程序（默认 2）: ",
        ['1', '2'], '2'
    )
    if choice == '1':
        return True
    else:
        print("👋 感谢使用，再见！")
        sys.exit(0)


# ========================
# 合并功能（修复版）
# ========================

def merge_files():
    print("\n" + "="*40)
    print("=== CSV/XLSX 文件合并工具 ===")
    print("="*40 + "\n")

    # 1. 输入文件或文件夹
    input_type = get_user_choice(
        "请选择输入方式: 1) 多个文件路径 (空格分隔)  2) 整个文件夹（默认 2）: ",
        ['1', '2'], '2'
    )

    file_paths = []
    if input_type == "1":
        paths_input = input("请输入文件路径（空格分隔）: ").strip()
        if not paths_input:
            print("❌ 未输入文件路径！")
            return
        file_paths = paths_input.split()
        for fp in file_paths:
            if not os.path.exists(fp):
                print(f"❌ 文件不存在: {fp}")
                return
    else:
        folder_path = input("请输入文件夹路径: ").strip()
        if not folder_path:
            folder_path = "."
        if not os.path.exists(folder_path) or not os.path.isdir(folder_path):
            print(f"❌ 文件夹不存在或无效: {folder_path}")
            return

        folder = Path(folder_path)
        file_paths = list(folder.glob("*.csv")) + list(folder.glob("*.xlsx")) + list(folder.glob("*.xls"))
        file_paths = [str(p) for p in file_paths]

        if not file_paths:
            print(f"❌ 在 {folder_path} 中未找到 .csv 或 Excel 文件！")
            return

        print(f"✅ 找到 {len(file_paths)} 个文件:")
        for i, fp in enumerate(file_paths):
            print(f"  {i + 1}. {fp}")

    # 2. 排序
    sort_choice = get_user_choice("是否按文件名排序？(y/n, 默认 y): ", ['y', 'n'], 'y')
    if sort_choice == 'y':
        file_paths.sort()

    # 3. 读取文件
    dataframes = []
    all_columns = set()
    common_columns = None

    print("\n正在读取文件...")
    for file in file_paths:
        ext = os.path.splitext(file)[1].lower()
        try:
            if ext == '.csv':
                total_lines, encoding = count_csv_lines(file)
                if total_lines is None:
                    print(f"❌ 无法读取文件（编码不支持）: {file}")
                    continue
                print(f"🔍 使用编码 {encoding} 读取 {os.path.basename(file)}")
                df = pd.read_csv(file, encoding=encoding)
                data_rows = len(df)
                print(f"✓ {os.path.basename(file)}: "
                      f"总行数（含表头）= {total_lines} 行, "
                      f"实际数据行 = {data_rows} 行, "
                      f"列数 = {len(df.columns)}")
            elif ext in ['.xlsx', '.xls']:
                df = pd.read_excel(file)
                data_rows = len(df)
                total_lines = data_rows + 1
                print(f"✓ {os.path.basename(file)}: "
                      f"总行数（含表头）≈ {total_lines} 行 (估算), "
                      f"实际数据行 = {data_rows} 行, "
                      f"列数 = {len(df.columns)}")
            else:
                print(f"跳过不支持的格式: {file}")
                continue

            dataframes.append((file, df))
            all_columns.update(df.columns)
            if common_columns is None:
                common_columns = set(df.columns)
            else:
                common_columns &= set(df.columns)
        except Exception as e:
            print(f"❌ 读取失败 {file}: {type(e).__name__}: {e}")

    if not dataframes:
        print("❌ 没有成功读取任何文件！")
        return

    print(f"\n📋 所有文件中出现过的列: {sorted(all_columns)}")
    print(f"🔹 所有文件共有的列: {sorted(common_columns)}")

    # 4. 选择列
    print(f"\n当前所有数据共包含 {len(all_columns)} 个不同的列。")
    col_choice = get_user_choice("是否只保留所有文件共有的列？(y/n, 默认 n): ", ['y', 'n'], 'n')

    if col_choice == 'y' and common_columns:
        selected_columns = sorted(common_columns)
        print(f"✅ 已选择共有的列: {selected_columns}")
    else:
        selected_columns_input = input("请输入要保留的列名（英文逗号分隔，留空表示全部列）: ").strip()
        if selected_columns_input:
            selected_columns = [col.strip() for col in selected_columns_input.split(',')]
            existing_cols = [col for col in selected_columns if col in all_columns]
            if not existing_cols:
                print("❌ 警告：你指定的列在任何文件中都不存在！")
                return
            selected_columns = existing_cols
        else:
            selected_columns = sorted(all_columns)

    # 5. 重命名列
    print(f"\n📤 当前输出列: {selected_columns}")
    rename_choice = get_user_choice("是否要重命名输出列？(y/n, 默认 n): ", ['y', 'n'], 'n')
    column_mapping = {}

    if rename_choice == 'y':
        print("请为每一列输入新的列名（留空则保持原名）:")
        for col in selected_columns:
            new_name = input(f"将 '{col}' 重命名为（留空不变）: ").strip()
            column_mapping[col] = new_name if new_name else col
    else:
        column_mapping = {col: col for col in selected_columns}

    final_columns = [column_mapping[col] for col in selected_columns]

    # 6. 空值处理
    print(f"\n🧹 是否删除全为空（或全空白）的行？")
    print("   （空字符串、空格、制表符等将被视为缺失值）")
    clean_empty = get_user_choice("是否删除？(y/n, 默认 y): ", ['y', 'n'], 'y') == 'y'

    # 7. 合并
    merged_rows = 0
    combined_df = pd.DataFrame(columns=final_columns)

    print("\n🔄 正在合并数据...")

    for i, (file, df) in enumerate(dataframes):
        temp_df = df.reindex(columns=selected_columns)
        temp_df.columns = final_columns

        if clean_empty:
            temp_df.replace(r'^\s*$', np.nan, regex=True, inplace=True)
            temp_df.dropna(how='all', inplace=True)

        combined_df = pd.concat([combined_df, temp_df], ignore_index=True)
        merged_rows += len(temp_df)
        print(f"  ✔️ 已合并: {os.path.basename(file)} -> {len(temp_df)} 行")

    print(f"✅ 合并完成！共合并 {merged_rows} 行数据。")

    expected_data_rows = sum([len(df) for _, df in dataframes])
    if merged_rows != expected_data_rows:
        print(f"⚠️  注意：实际合并 {merged_rows} 行，预期 {expected_data_rows} 行。")
        print(f"    可能原因：检测到 {expected_data_rows - merged_rows} 行全空（或全空白），已被删除。")
    else:
        print(f"✅ 数据行数匹配，合并完整。")

    # 8. 输出
    output_format = get_user_choice("请选择输出格式: 1) CSV  2) XLSX（默认 1）: ", ['1', '2'], '1')
    output_ext = ".xlsx" if output_format == "2" else ".csv"

    output_file = input("请输入输出文件路径（含文件名）: ").strip()
    if not output_file:
        base_name = "merged_output"
        output_file = f"{base_name}{output_ext}"
    elif not output_file.lower().endswith(('.csv', '.xlsx')):
        output_file += output_ext

    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 9. 保存
    try:
        if output_ext == ".csv":
            combined_df.to_csv(output_file, index=False, encoding='utf-8-sig')
            print(f"🎉 成功保存为 CSV: {output_file}")
        else:
            combined_df.to_excel(output_file, index=False, sheet_name="MergedData")
            print(f"🎉 成功保存为 Excel: {output_file}")
        print(f"📊 输出文件总行数（含表头）: {len(combined_df) + 1} 行（数据行数: {len(combined_df)}）")
    except Exception as e:
        print(f"❌ 保存失败: {type(e).__name__}: {e}")


def count_csv_lines(file_path):
    encodings = ['utf-8', 'gbk', 'utf-8-sig', 'cp1252', 'latin1']
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                return sum(1 for _ in f), encoding
        except:
            continue
    return None, None


# ========================
# 分割功能
# ========================

def split_excel_or_csv():
    print("\n" + "="*40)
    print("=== CSV/XLSX 文件分割工具 ===")
    print("="*40 + "\n")

    file_path = input("请输入文件路径（支持 .csv 或 .xlsx）: ").strip()
    if not os.path.exists(file_path):
        print(f"❌ 文件 '{file_path}' 不存在！")
        return

    _, ext = os.path.splitext(file_path)
    ext = ext.lower()

    if ext == '.csv':
        sheet_name = None
    elif ext in ['.xls', '.xlsx']:
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            print(f"可用的工作表: {sheet_names}")
            sheet_name = input(f"请输入要读取的工作表名称（默认 '{sheet_names[0]}'）: ").strip()
            if not sheet_name:
                sheet_name = sheet_names[0]
            if sheet_name not in sheet_names:
                print(f"❌ 工作表 '{sheet_name}' 不存在！")
                return
        except Exception as e:
            print(f"无法读取 Excel 文件: {e}")
            return
    else:
        print(f"❌ 不支持的文件格式: {ext}，仅支持 .csv、.xls、.xlsx")
        return

    # 读取数据
    try:
        if ext == '.csv':
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"❌ 读取文件失败: {e}")
        return

    total_rows = len(df)
    columns = df.columns.tolist()
    print(f"数据总行数（不含表头）: {total_rows}")
    print(f"原始列名: {columns}")

    # 选择列
    selected_columns_input = input("请输入要提取的列名（英文逗号分隔，留空表示全部列）: ").strip()
    if selected_columns_input:
        selected_columns = [col.strip() for col in selected_columns_input.split(',')]
        missing_cols = [col for col in selected_columns if col not in columns]
        if missing_cols:
            print(f"❌ 错误：以下列不存在: {missing_cols}")
            return
    else:
        selected_columns = columns

    # 重命名
    print(f"\n当前选中的列: {selected_columns}")
    rename_choice = get_user_choice("是否要重命名输出列？(y/n，留空为 n): ", ['y', 'n'], 'n')
    column_mapping = {}
    if rename_choice == 'y':
        print("\n请为每一列输入新的列名（留空则保持原列名）:")
        for col in selected_columns:
            new_name = input(f"将 '{col}' 重命名为（留空保持不变）: ").strip()
            column_mapping[col] = new_name if new_name else col
    else:
        column_mapping = {col: col for col in selected_columns}

    final_columns = [column_mapping[col] for col in selected_columns]

    # 起始/结束行
    try:
        start_row_input = input("请输入起始行号（从 1 开始，留空为 1）: ").strip()
        start_row = int(start_row_input) - 1 if start_row_input else 0
        if start_row < 0 or start_row >= total_rows:
            print(f"❌ 起始行必须在 1 到 {total_rows} 之间")
            return
    except ValueError:
        print("❌ 请输入有效数字！")
        return

    end_row_input = input("请输入结束行号（从 1 开始，留空表示到最后）: ").strip()
    if end_row_input:
        try:
            end_row = int(end_row_input)
            if end_row <= 0 or end_row > total_rows:
                print(f"❌ 结束行必须在 1 到 {total_rows} 之间")
                return
            if end_row <= start_row + 1:
                print("❌ 结束行必须大于起始行！")
                return
        except ValueError:
            print("❌ 请输入有效数字或留空！")
            return
    else:
        end_row = total_rows

    # 每文件行数
    try:
        chunk_size_input = input("请输入每个输出文件的最大行数（如 500）: ").strip()
        chunk_size = int(chunk_size_input)
        if chunk_size <= 0:
            print("❌ 行数必须大于 0！")
            return
    except ValueError:
        print("❌ 请输入有效数字！")
        return

    # 输出格式
    output_format = get_user_choice("请选择输出格式 (1: CSV, 2: XLSX)（默认 1）: ", ['1', '2'], '1')
    output_ext = ".xlsx" if output_format == "2" else ".csv"
    print(f"✅ 输出格式: {output_ext}")

    # 输出目录
    output_dir = input("请输入输出目录（留空为当前目录）: ").strip()
    if not output_dir:
        output_dir = "."
    os.makedirs(output_dir, exist_ok=True)

    # 提取并分割
    subset_df = df.iloc[start_row:end_row][selected_columns].copy()
    subset_df.columns = final_columns
    total_subset_rows = len(subset_df)

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    num_files = 0

    print(f"\n🔄 正在分割 {total_subset_rows} 行数据，每文件最多 {chunk_size} 行...\n")

    for i in range(0, total_subset_rows, chunk_size):
        chunk_df = subset_df.iloc[i:i + chunk_size]
        output_file = os.path.join(output_dir, f"{base_name}_part_{num_files + 1}{output_ext}")

        try:
            if output_ext == ".csv":
                chunk_df.to_csv(output_file, index=False, encoding='utf-8-sig')
            else:
                chunk_df.to_excel(output_file, index=False, sheet_name="Sheet1")
            print(f"✓ 已保存: {output_file} ({len(chunk_df)} 行)")
            num_files += 1
        except Exception as e:
            print(f"❌ 保存失败 {output_file}: {e}")

    print(f"\n✅ 分割完成！共生成 {num_files} 个文件，保存在 '{output_dir}' 目录下。")


# ========================
# 主程序入口
# ========================

def main():
    print("🚀 欢迎使用 xlsxSelector")
    print("支持功能：")
    print("  1) 合并多个 CSV/Excel 文件")
    print("  2) 分割单个 CSV/Excel 文件")

    while True:
        print("\n" + "="*50)
        choice = get_user_choice(
            "请选择功能: 1) 合并  2) 分割  3) 退出（默认 1）: ",
            ['1', '2', '3'], '1'
        )

        if choice == '1':
            merge_files()
        elif choice == '2':
            split_excel_or_csv()
        elif choice == '3':
            print("👋 感谢使用，再见！")
            sys.exit(0)

        # 询问是否继续
        exit_or_continue()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 程序被用户中断，再见！")
        sys.exit(0)