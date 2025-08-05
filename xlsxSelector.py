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


# 查重删除工具相关函数
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
    col_key = str(column).strip()

    if col_key in df.columns:
        return set(df[col_key].dropna().astype(str))

    if col_key.isdigit():
        idx = int(col_key) - 1
        if 0 <= idx < len(df.columns):
            return set(df.iloc[:, idx].dropna().astype(str))
        else:
            raise ValueError(f"列序号 {int(col_key)} 超出范围 [1, {len(df.columns)}]")

    if len(col_key) == 1 and col_key.isalpha():
        idx = ord(col_key.upper()) - ord('A')
        if 0 <= idx < len(df.columns):
            return set(df.iloc[:, idx].dropna().astype(str))
        else:
            raise ValueError(f"列字母 '{col_key}' 超出范围 [A-{chr(ord('A') + len(df.columns) - 1)}]")

    available_cols = list(df.columns)
    raise ValueError(f"无法找到列 '{col_key}'。可用列名：{available_cols}")


def select_sheet(sheet_names):
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
    print("Excel 多文件查重删除")
    print("=" * 60)

    # 主文件路径
    main_path = input("请输入主文件路径（被查重的文件）: ").strip().strip('"\'')
    main_file = Path(main_path)
    if not main_file.exists():
        print(f"文件不存在: {main_file}")
        sys.exit(1)

    try:
        main_df, main_sheets = read_file(main_file)
        print(f"成功读取主文件，共 {len(main_sheets)} 个 Sheet。")
        main_sheet = select_sheet(main_sheets)

        if main_file.suffix.lower() in ['.xlsx', '.xls']:
            main_df = pd.read_excel(main_file, sheet_name=main_sheet, dtype=str)

        columns_quoted = ", ".join(f"'{col}'" for col in main_df.columns)
        print(f"\n主文件 '{main_sheet}' 的原始列名: {columns_quoted}")
        main_column = input("请输入主文件用于比较的列名: ").strip()
        if not main_column:
            print("列名不能为空！")
            sys.exit(1)

        ref_input = input("\n请输入对比文件路径（多个用分号 ; 分隔），直接回车结束: ").strip()
        ref_files = []
        if ref_input:
            ref_files = [Path(p.strip().strip('"\'')) for p in ref_input.split(';') if p.strip()]

        valid_ref_files = [f for f in ref_files if f.exists()]
        if not valid_ref_files:
            print("没有有效的对比文件！")
            sys.exit(1)

        ref_configs = []
        for file in valid_ref_files:
            print(f"\n--- {file.name} ---")
            try:
                _, sheets = read_file(file)
                sheet = select_sheet(sheets)

                if file.suffix.lower() in ['.xlsx', '.xls']:
                    df_temp = pd.read_excel(file, sheet_name=sheet, dtype=str)
                else:
                    df_temp = pd.read_csv(file, dtype=str, low_memory=False)

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
            sys.exit(1)

        print("\n开始查重处理...")
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
            key = row[main_column]
            if pd.isna(key):
                return False
            return str(key).strip() in all_ref_values

        mask = main_df.apply(is_duplicate, axis=1)
        removed_count = mask.sum()
        filtered_df = main_df[~mask]

        print(f"查重完成！删除 {removed_count} 行，剩余 {len(filtered_df)} 行。")

        output_path = input("\n请输入保存路径（如 result.xlsx）: ").strip().strip('"\'')
        if not output_path:
            print("未指定保存路径！")
            sys.exit(1)

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
        sys.exit(1)


# 合并功能（修复版）
def merge_files():
    print("\n" + "=" * 40)
    print("=== CSV/XLSX 文件合并工具 ===")
    print("=" * 40 + "\n")

    # 输入文件路径
    file_input = input("请输入要合并的文件路径（多个用分号;分隔）: ").strip()
    if not file_input:
        print("❌ 未输入任何文件路径！")
        return

    file_paths = [Path(p.strip().strip('"\'')) for p in file_input.split(';') if p.strip()]

    # 过滤存在的文件
    valid_files = []
    for path in file_paths:
        if path.exists():
            valid_files.append(path)
        else:
            print(f"⚠️ 跳过不存在的文件: {path}")

    if not valid_files:
        print("❌ 没有有效的文件可以合并！")
        return

    print(f"✅ 找到 {len(valid_files)} 个有效文件")

    # 读取所有文件
    data_frames = []
    for file in valid_files:
        try:
            if file.suffix.lower() in ['.xlsx', '.xls']:
                df = pd.read_excel(file, dtype=str)
            elif file.suffix.lower() == '.csv':
                df = pd.read_csv(file, dtype=str, low_memory=False)
            else:
                print(f"跳过不支持的格式: {file}")
                continue
            print(f"✔️ 读取成功: {file.name} ({len(df)} 行)")
            data_frames.append(df)
        except Exception as e:
            print(f"❌ 读取失败 {file}: {e}")

    if not data_frames:
        print("❌ 没有成功读取任何文件！")
        return

    # 合并
    try:
        merged_df = pd.concat(data_frames, ignore_index=True)
        print(f"✅ 合并完成！总行数: {len(merged_df)}")

        # 保存
        output_path = input("\n请输入保存路径（如 merged.xlsx）: ").strip().strip('"\'')
        if not output_path:
            print("❌ 未指定保存路径！")
            return

        output_file = Path(output_path)
        try:
            if output_file.suffix.lower() == '.csv':
                merged_df.to_csv(output_file, index=False, encoding='utf-8-sig')
            else:
                merged_df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"🎉 成功保存至:\n   {output_file.resolve()}")
        except Exception as e:
            print(f"❌ 保存失败: {e}")

    except Exception as e:
        print(f"❌ 合并失败: {e}")


# 分割功能
def split_excel_or_csv():
    print("\n" + "=" * 40)
    print("=== CSV/XLSX 文件分割工具 ===")
    print("=" * 40 + "\n")

    file_path = input("请输入要分割的文件路径: ").strip().strip('"\'')
    file = Path(file_path)

    if not file.exists():
        print(f"❌ 文件不存在: {file}")
        return

    try:
        if file.suffix.lower() in ['.xlsx', '.xls']:
            df = pd.read_excel(file, dtype=str)
        elif file.suffix.lower() == '.csv':
            df = pd.read_csv(file, dtype=str, low_memory=False)
        else:
            print("❌ 不支持的文件格式！")
            return

        print(f"✅ 成功读取文件，共 {len(df)} 行")

        # 输入每份大小
        while True:
            size_input = input("请输入每份文件的行数: ").strip()
            if size_input.isdigit() and int(size_input) > 0:
                chunk_size = int(size_input)
                break
            print("❌ 请输入有效的正整数！")

        # 分割
        total_rows = len(df)
        num_files = (total_rows + chunk_size - 1) // chunk_size  # 向上取整

        print(f"开始分割为 {num_files} 个文件...")

        base_name = file.stem
        suffix = file.suffix

        for i in range(num_files):
            start_idx = i * chunk_size
            end_idx = min((i + 1) * chunk_size, total_rows)
            chunk_df = df.iloc[start_idx:end_idx]

            output_file = file.parent / f"{base_name}_part{i + 1:03d}{suffix}"
            try:
                if suffix.lower() == '.csv':
                    chunk_df.to_csv(output_file, index=False, encoding='utf-8-sig')
                else:
                    chunk_df.to_excel(output_file, index=False, engine='openpyxl')
                print(f"✅ 保存: {output_file.name} ({len(chunk_df)} 行)")
            except Exception as e:
                print(f"❌ 保存失败 {output_file}: {e}")

        print("🎉 分割完成！")

    except Exception as e:
        print(f"❌ 处理失败: {e}")


# ========================
# 主程序入口
# ========================
def main():
    print("🚀 欢迎使用 xlsxSelector")
    print("支持功能：")
    print("  1) 合并多个 CSV/Excel 文件")
    print("  2) 分割单个 CSV/Excel 文件")
    print("  3) 查重删除")

    while True:
        print("\n" + "=" * 50)
        choice = get_user_choice(
            "请选择功能: 1) 合并  2) 分割  3) 查重删除  4) 退出（默认 1）: ",
            ['1', '2', '3', '4'], '1'
        )

        if choice == '1':
            merge_files()
        elif choice == '2':
            split_excel_or_csv()
        elif choice == '3':
            deduplicate_files()
        elif choice == '4':
            print("👋 感谢使用，再见！")
            sys.exit(0)

            # 询问是否继续
        if not exit_or_continue():
            print("👋 感谢使用，再见！")
            sys.exit(0)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 程序被用户中断，再见！")
        sys.exit(0)