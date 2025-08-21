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
    print("\n" + "=" * 50)
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
# 合并功能
# ========================

def merge_files():
    print("\n" + "=" * 40)
    print("=== CSV/XLSX 文件合并工具 ===")
    print("=" * 40 + "\n")

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

def get_file_path(prompt):
    """获取有效的文件路径"""
    while True:
        file_path = input(prompt).strip().replace("'", "").replace('"', '')
        if not os.path.exists(file_path):
            print("错误：文件路径不存在，请重新输入。")
        elif not (file_path.endswith('.xlsx') or file_path.endswith('.csv')):
            print("错误：文件格式不支持，请确保是 .xlsx 或 .csv 文件。")
        else:
            return file_path


def get_output_dir(prompt):
    """获取有效的输出目录"""
    while True:
        output_dir = input(prompt).strip().replace("'", "").replace('"', '')
        if not output_dir:
            output_dir = "."
        if not os.path.isdir(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
                print(f"目录不存在，已为您创建：{output_dir}")
                return output_dir
            except Exception as e:
                print(f"无法创建目录：{e}，请重新输入。")
        else:
            return output_dir


def read_and_process_file(file_path):
    """读取文件并处理列名"""
    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            df = pd.read_csv(file_path, dtype=str, encoding='utf-8', on_bad_lines='skip')

        # 将所有列强制转换为字符串，以保留大数字的精度
        df = df.astype(str)

        columns = df.columns.tolist()
        print("\n当前文件的列名如下：")
        print("['" + "', '".join(columns) + "']")

        while True:
            selection_input = input("\n请输入您想保留的列名（以逗号分隔，留空默认选择所有）：").strip()
            if not selection_input:
                selected_columns = columns
            else:
                selected_columns = [col.strip() for col in selection_input.split(',')]
                invalid_cols = [col for col in selected_columns if col not in columns]
                if invalid_cols:
                    print(f"错误：以下列名不存在：{invalid_cols}，请重新输入。")
                    continue
            break

        rename_choice = input("是否需要重命名这些列？(y/n)，留空默认不重命名: ").strip().lower()
        if rename_choice == 'y':
            new_names = []
            for col in selected_columns:
                new_name = input(f"请输入 '{col}' 的新名称：").strip()
                new_names.append(new_name if new_name else col)
            rename_map = dict(zip(selected_columns, new_names))
            df = df.rename(columns=rename_map)
            print("\n列名已更新为：['" + "', '".join(new_names) + "']")
            selected_columns = new_names
        else:
            print("\n已选择不重命名列。")

        return df[selected_columns]

    except Exception as e:
        print(f"读取文件时发生错误：{e}")
        return None


def slice_by_count(df):
    """按行数截取"""
    while True:
        try:
            start_row = int(input("\n请输入开始截取的行数（从1开始计数）："))
            if start_row <= 0:
                print("输入无效，请输入一个大于0的整数。")
                continue
            start_row -= 1  # 转换为0-based索引
            if start_row >= len(df):
                print(f"开始行超出文件总行数（{len(df)}），请重新输入。")
                continue
            break
        except ValueError:
            print("输入无效，请输入一个非负整数。")

    while True:
        try:
            row_count = int(input("请输入每次截取的行数："))
            if row_count <= 0:
                print("输入无效，请输入一个大于0的整数。")
                continue
            break
        except ValueError:
            print("输入无效，请输入一个大于0的整数。")

    while True:
        try:
            slice_times = int(input("请输入需要这样截取几次："))
            if slice_times <= 0:
                print("输入无效，请输入一个大于0的整数。")
                continue
            break
        except ValueError:
            print("输入无效，请输入一个大于0的整数。")

    sliced_dfs = []
    for i in range(slice_times):
        start = start_row + i * row_count
        end = start + row_count
        if start >= len(df):
            break
        sliced_dfs.append(df.iloc[start:end])
    return sliced_dfs


def slice_by_end_row(df):
    """按行范围截取"""
    while True:
        try:
            start_row = int(input("\n请输入开始截取的行数（从1开始计数）："))
            if start_row <= 0:
                print("输入无效，请输入一个大于0的整数。")
                continue
            if start_row > len(df):
                print(f"开始行超出文件总行数（{len(df)}），请重新输入。")
                continue
            break
        except ValueError:
            print("输入无效，请输入一个非负整数。")

    while True:
        try:
            end_row = int(input("请输入截取到第几行（不包含此行，从1开始计数）："))
            if end_row <= start_row:
                print("结束行必须大于开始行。")
                continue
            if end_row > len(df) + 1:
                print(f"结束行超出文件总行数（{len(df)}），请重新输入。")
                continue
            break
        except ValueError:
            print("输入无效，请选择 'csv' 或 'xlsx'。")

    start_row -= 1  # 转换为0-based索引
    end_row -= 1  # 转换为0-based索引

    while True:
        try:
            num_slices = int(input("请输入需要截取几段："))
            if num_slices <= 0:
                print("输入无效，请输入一个大于0的整数。")
                continue
            break
        except ValueError:
            print("输入无效，请输入一个大于0的整数。")

    total_rows = end_row - start_row
    slice_length = total_rows // num_slices

    sliced_dfs = []
    for i in range(num_slices):
        start = start_row + i * slice_length
        end = start + slice_length
        sliced_dfs.append(df.iloc[start:end])
    return sliced_dfs


def split_excel_or_csv():
    print("\n" + "=" * 40)
    print("=== CSV/XLSX 文件分割工具 ===")
    print("=" * 40 + "\n")

    file_path = get_file_path("请输入您要截取的 Excel 或 CSV 文件路径：")

    processed_df = read_and_process_file(file_path)
    if processed_df is None:
        return

    while True:
        slice_method = input(
            "\n请选择截取方式：\n1. 指定截取多少行，并重复截取相同行数几次\n2. 指定截取到第几行，并将截取到的部分划为几段\n请选择 (1/2): ").strip()
        if slice_method == '1':
            sliced_dataframes = slice_by_count(processed_df)
            break
        elif slice_method == '2':
            sliced_dataframes = slice_by_end_row(processed_df)
            break
        else:
            print("无效的选择，请重新输入。")

    if not sliced_dataframes:
        print("未生成任何截取数据，程序结束。")
        return

    output_dir = get_output_dir("\n请输入保存截取文件的目录地址（留空则为当前目录）：")
    if not output_dir:
        output_dir = "."

    output_filename_base = input("\n请输入输出文件的名称前缀（例如：my_data，留空则为 'output'）：").strip()
    if not output_filename_base:
        output_filename_base = "output"

    while True:
        output_format = input("请选择输出文件格式 (csv/xlsx): ").strip().lower()
        if output_format in ['csv', 'xlsx']:
            break
        else:
            print("无效的格式，请选择 'csv' 或 'xlsx'。")

    for i, df_slice in enumerate(sliced_dataframes):
        df = df_slice.copy()

        output_filename = f"{output_filename_base}_part_{i + 1}.{output_format}"
        output_path = os.path.join(output_dir, output_filename)

        try:
            if output_format == 'xlsx':
                df.to_excel(output_path, index=False)
            else:
                # 使用 utf-8-sig 编码
                df.to_csv(output_path, index=False, encoding='utf-8-sig')
            print(f"文件已保存至：{output_path}")
        except Exception as e:
            print(f"❌ 保存失败 {output_path}: {e}")

    print("\n所有截取操作已完成！")


# ========================
# 查重功能
# ========================

def read_file(file_path, sheet=None):
    ext = file_path.suffix.lower()
    try:
        if ext in ['.xlsx', '.xls']:
            excel_file = pd.ExcelFile(file_path)
            if sheet is None:
                sheet = excel_file.sheet_names[0]
            df = pd.read_excel(file_path, sheet_name=sheet, dtype=str)
            return df, excel_file.sheet_names
        elif ext == '.csv':
            df = pd.read_csv(file_path, dtype=str, low_memory=False)
            return df, ["CSV"]
        else:
            raise ValueError(f"不支持的文件格式: {ext}")
    except Exception as e:
        raise RuntimeError(f"读取文件失败 {file_path}: {e}")


def get_column_data(df, column):
    """
    支持三种输入方式：
    1. 列名（如 'email'）
    2. 列字母（如 'B'）
    3. 列序号（如 '2' 或 2）
    返回该列所有非空值的集合（字符串）
    """
    col_key = str(column).strip()

    # 情况1：先尝试当作列名查找
    if col_key in df.columns:
        return set(df[col_key].dropna().astype(str))

    # 情况2：如果不是列名，再尝试当作列序号（纯数字）
    if col_key.isdigit():
        idx = int(col_key) - 1  # 转为从0开始
        if 0 <= idx < len(df.columns):
            return set(df.iloc[:, idx].dropna().astype(str))
        else:
            raise ValueError(f"列序号 {int(col_key)} 超出范围 [1, {len(df.columns)}]")

    # 情况3：再尝试当作列字母（如 'B'）
    if len(col_key) == 1 and col_key.isalpha():
        idx = ord(col_key.upper()) - ord('A')
        if 0 <= idx < len(df.columns):
            return set(df.iloc[:, idx].dropna().astype(str))
        else:
            raise ValueError(f"列字母 '{col_key}' 超出范围 [A-{chr(ord('A') + len(df.columns) - 1)}]")

    # 都不匹配，报错
    available_cols = list(df.columns)
    raise ValueError(f"无法找到列 '{col_key}'。可用列名：{available_cols}")


def select_sheet(sheet_names):
    # 每个 Sheet 名用 ' ' 包裹，逗号分隔，不加 [ ]
    sheets_quoted = ", ".join(f"'{name}'" for name in sheet_names)
    print(f"可用的 Sheet 列表: {sheets_quoted}")

    default_sheet = sheet_names[0]
    print(f"提示：直接回车使用默认 [{default_sheet}]")

    choice = input("请选择 Sheet（输入序号或名称）: ").strip()
    if not choice:
        return default_sheet

    if choice.isdigit():
        idx = int(choice) - 1
        if 0 <= idx < len(sheet_names):
            return sheet_names[idx]
        else:
            print(f"序号超出范围，使用默认 [{default_sheet}]")
            return default_sheet
    else:
        if choice in sheet_names:
            return choice
        else:
            print(f"未找到 Sheet '{choice}'，使用默认 [{default_sheet}]")
            return default_sheet


def deduplicate_files():
    print("\n" + "=" * 40)
    print("=== CSV/XLSX 文件查重删除工具 ===")
    print("=" * 40 + "\n")

    # 1. 输入主文件路径
    main_path = input("请输入主文件路径（被查重的文件）: ").strip().strip('"\'')
    main_file = Path(main_path)
    if not main_file.exists():
        print(f"文件不存在: {main_file}")
        return

    # 读取主文件
    try:
        main_df, main_sheets = read_file(main_file)
        print(f"成功读取主文件，共 {len(main_sheets)} 个 Sheet。")
        main_sheet = select_sheet(main_sheets)

        # 重新读取用户选择的 sheet（保持 dtype=str）
        if main_file.suffix.lower() in ['.xlsx', '.xls']:
            main_df = pd.read_excel(main_file, sheet_name=main_sheet, dtype=str)
        # CSV 已读取，无需再处理

    except Exception as e:
        print(f"读取主文件失败: {e}")
        return

    # 显示列名（每个列名加 ' '，逗号分隔，不加 [ ]）
    columns_quoted = ", ".join(f"'{col}'" for col in main_df.columns)
    print(f"\n主文件 '{main_sheet}' 的原始列名: {columns_quoted}")
    main_column = input("请输入主文件用于比较的列名: ").strip()
    if not main_column:
        print("列名不能为空！")
        return

    # 2. 输入对比文件
    print("\n请输入对比文件路径（多个用分号 ; 分隔，或一行一个，空行结束）:")
    ref_input = input().strip()
    ref_files = []
    if ';' in ref_input:
        ref_files = [Path(p.strip().strip('"\'')) for p in ref_input.split(';') if p.strip()]
    else:
        if ref_input:
            ref_files.append(Path(ref_input.strip('"\'')))
        while True:
            line = input().strip()
            if not line:
                break
            ref_files.append(Path(line.strip('"\'')))

    valid_ref_files = [f for f in ref_files if f.exists()]
    if not valid_ref_files:
        print("没有有效的对比文件！")
        return

    for f in ref_files:
        if not f.exists():
            print(f"跳过不存在的文件: {f}")

    # 3. 配置对比文件
    ref_configs = []
    print("\n配置每个对比文件的 Sheet 和比较列:")
    for file in valid_ref_files:
        print(f"\n--- {file.name} ---")
        try:
            _, sheets = read_file(file)
            sheet = select_sheet(sheets)

            # 读取指定 sheet
            if file.suffix.lower() in ['.xlsx', '.xls']:
                df_temp = pd.read_excel(file, sheet_name=sheet, dtype=str)
            else:
                df_temp = pd.read_csv(file, dtype=str, low_memory=False)

            # 显示列名（加引号，去括号）
            columns_quoted = ", ".join(f"'{col}'" for col in df_temp.columns)
            print(f"列名: {columns_quoted}")

            col = input("比较列名: ").strip()
            if not col:
                print("列不能为空，跳过此文件。")
                continue

            ref_configs.append({
                'file': file,
                'sheet': sheet,
                'column': col,
                'df': df_temp
            })
        except Exception as e:
            print(f"读取失败: {e}")

    if not ref_configs:
        print("没有配置任何有效的对比文件！")
        return

    # 4. 查重处理
    print("\n开始查重处理...")
    try:
        main_values_set = get_column_data(main_df, main_column)
        print(f"主文件 '{main_column}' 列共 {len(main_values_set)} 个唯一值（仅用于检查）。")

        all_ref_values = set()
        for config in ref_configs:
            df = config['df']
            col = config['column']
            print(f"处理: {config['file'].name} [{config['sheet']}] 列 '{col}'")
            values = get_column_data(df, col)
            all_ref_values.update(values)
            print(f"添加 {len(values)} 个值，累计 {len(all_ref_values)} 个。")

        print(f"总共 {len(all_ref_values)} 个用于查重的值。")

        def is_duplicate(row):
            key = row.get(main_column)
            if pd.isna(key) or key is None:
                return False
            return str(key).strip() in all_ref_values

        mask = main_df.apply(is_duplicate, axis=1)
        removed_count = mask.sum()
        filtered_df = main_df[~mask]

        print(f"查重完成！删除 {removed_count} 行，剩余 {len(filtered_df)} 行。")

        # 5. 保存结果
        output_path = input("\n请输入保存路径（如 result.xlsx）: ").strip().strip('"\'')
        if not output_path:
            print("未指定保存路径！")
            return

        output_file = Path(output_path)
        try:
            if output_file.suffix.lower() == '.csv':
                filtered_df.to_csv(output_file, index=False, encoding='utf-8-sig')
            else:
                filtered_df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"成功保存至:\n   {output_file.resolve()}")
        except Exception as e:
            print(f"保存失败: {e}")

    except Exception as e:
        print(f"处理失败: {e}")
        return


# ========================
# 清理空行功能
# ========================

def clean_spreadsheet_main():
    print("\n" + "=" * 40)
    print("=== 表格文件行过滤工具（删除指定列为空的行） ===")
    print("=" * 40 + "\n")

    # 1. 输入文件路径
    input_path = input("📌 请输入或拖入 CSV/XLSX 文件路径: ").strip().strip('"\'')
    if not input_path or not os.path.exists(input_path):
        print("❌ 文件路径无效或不存在！")
        return

    # 2. 自动读取列名
    ext = os.path.splitext(input_path)[1].lower()
    try:
        if ext == '.csv':
            df = pd.read_csv(input_path, nrows=0)  # 只读标题
        elif ext in ['.xlsx', '.xls']:
            df = pd.read_excel(input_path, nrows=0)
        else:
            print("❌ 不支持的文件格式！仅支持 .csv、.xlsx、.xls")
            return
        columns = df.columns.tolist()
    except Exception as e:
        print(f"❌ 无法读取文件列名: {e}")
        return

    if not columns:
        print("⚠️  文件中没有找到任何列。")
        return

    # 显示列名
    print("\n📋 文件中的列名如下：")
    for i, col in enumerate(columns, 1):
        print(f"   {i:2d}. {col}")

    # 3. 让用户选择列（支持输入列名或序号）
    print("\n💡 请输入要检查空白的列：")
    print("   • 可输入列名，多个用英文逗号分隔，如：Email,Name")
    print("   • 或输入列序号，如：1,3,5")
    choice = input("\n👉 请输入: ").strip()

    if not choice:
        print("❌ 未输入任何列信息！")
        return

    selected_columns = []
    choices = [c.strip() for c in choice.split(',')]

    for c in choices:
        if c.isdigit():
            idx = int(c) - 1
            if 0 <= idx < len(columns):
                selected_columns.append(columns[idx])
            else:
                print(f"❌ 序号 {c} 超出范围！")
                return
        else:
            if c in columns:
                selected_columns.append(c)
            else:
                print(f"❌ 列名 '{c}' 不存在！")
                return

    if not selected_columns:
        print("❌ 未选择任何有效列！")
        return

    print(f"\n✅ 已选择检查以下列的空白: {selected_columns}")

    # 4. 输入输出目录
    output_dir = input("📁 请输入保存结果的目录路径（留空则为当前目录）: ").strip().strip('"\'')
    if not output_dir:
        output_dir = "."
    if not os.path.exists(output_dir):
        confirm = input(f"📁 目录不存在，是否创建? (y/n): ").strip().lower()
        if confirm != 'y':
            print("❌ 用户取消创建目录。")
            return
        try:
            os.makedirs(output_dir, exist_ok=True)
            print(f"✅ 已创建目录: {output_dir}")
        except Exception as e:
            print(f"❌ 创建目录失败: {e}")
            return

    # 5. 输入输出文件名
    output_filename = input("📄 请输入输出文件名（如 result.csv 或 result.xlsx）: ").strip().strip('"\'')
    if not output_filename:
        print("❌ 未输入文件名！")
        return

    output_path = os.path.join(output_dir, output_filename)

    # 6. 执行处理
    try:
        clean_spreadsheet(input_path, output_path, selected_columns)
    except Exception as e:
        print(f"\n❌ 程序执行出错: {e}")


def clean_spreadsheet(input_path, output_path, check_columns):
    """
    读取 CSV/XLSX 文件，删除指定列中为空的行，并保存结果。
    """
    ext = os.path.splitext(input_path)[1].lower()
    try:
        if ext == '.csv':
            df = pd.read_csv(input_path, encoding='utf-8')
            print(f"✅ 已读取 CSV 文件: {input_path}")
        elif ext in ['.xlsx', '.xls']:
            df = pd.read_excel(input_path)
            print(f"✅ 已读取 Excel 文件: {input_path}")
        else:
            raise ValueError(f"不支持的文件格式: {ext}")
    except Exception as e:
        raise Exception(f"读取文件失败: {e}")

    # 检查列是否存在
    missing_cols = [col for col in check_columns if col not in df.columns]
    if missing_cols:
        raise ValueError(f"以下列在文件中未找到: {missing_cols}")

    # 检查空白
    mask = pd.Series([True] * len(df), index=df.index)
    for col in check_columns:
        # 同时检查 NaN 和空字符串/纯空格
        col_not_empty = df[col].notna() & (df[col].astype(str).str.strip() != '')
        mask &= col_not_empty

    cleaned_df = df[mask].reset_index(drop=True)

    # 确保输出目录存在
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 保存
    out_ext = os.path.splitext(output_path)[1].lower()
    try:
        if out_ext == '.csv':
            cleaned_df.to_csv(output_path, index=False, encoding='utf-8-sig')
        elif out_ext in ['.xlsx', '.xls']:
            cleaned_df.to_excel(output_path, index=False)
        else:
            raise ValueError(f"不支持的输出格式: {out_ext}")
        print(f"\n✅ 处理完成！")
        print(f"📊 原始行数: {len(df)}")
        print(f"🧹 清理后行数: {len(cleaned_df)}")
        print(f"💾 已保存到: {output_path}")
    except Exception as e:
        raise Exception(f"保存文件失败: {e}")


# ========================
# 主程序入口
# ========================

def main():
    print("🚀 欢迎使用 xlsxSelector")
    print("支持功能：")
    print("  1) 合并多个 CSV/Excel 文件")
    print("  2) 分割单个 CSV/Excel 文件")
    print("  3) 对主文件进行查重和删除")
    print("  4) 清理空行")

    while True:
        print("\n" + "=" * 50)
        choice = get_user_choice(
            "请选择功能: 1) 合并  2) 分割  3) 查重  4) 清理空行  5) 退出（默认 1）: ",
            ['1', '2', '3', '4', '5'], '1'
        )

        if choice == '1':
            merge_files()
        elif choice == '2':
            split_excel_or_csv()
        elif choice == '3':
            deduplicate_files()
        elif choice == '4':
            clean_spreadsheet_main()
        elif choice == '5':
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