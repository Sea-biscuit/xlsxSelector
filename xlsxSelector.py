import pandas as pd
import os
import sys

def split_excel_or_csv():
    print("=== CSV/XLSX 文件分割工具 ===\n")

    # 1. 输入文件路径
    file_path = input("请输入文件路径（支持 .csv 或 .xlsx）: ").strip()
    if not os.path.exists(file_path):
        print(f"错误：文件 '{file_path}' 不存在！")
        sys.exit(1)

    # 判断文件类型
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
                print(f"错误：工作表 '{sheet_name}' 不存在！")
                sys.exit(1)
        except Exception as e:
            print(f"无法读取 Excel 文件: {e}")
            sys.exit(1)
    else:
        print(f"不支持的文件格式: {ext}，仅支持 .csv、.xls、.xlsx")
        sys.exit(1)

    # 2. 读取数据
    try:
        if ext == '.csv':
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"读取文件失败: {e}")
        sys.exit(1)

    total_rows = len(df)
    columns = df.columns.tolist()
    print(f"数据总行数（不含表头）: {total_rows}")
    print(f"原始列名: {columns}")

    # 3. 选择要提取的列
    selected_columns_input = input("请输入要提取的列名（英文逗号分隔，留空表示全部列）: ").strip()
    if selected_columns_input:
        selected_columns = [col.strip() for col in selected_columns_input.split(',')]
        missing_cols = [col for col in selected_columns if col not in columns]
        if missing_cols:
            print(f"错误：以下列不存在: {missing_cols}")
            sys.exit(1)
    else:
        selected_columns = columns

    # 4. 列名编辑功能
    print(f"\n当前选中的列: {selected_columns}")
    rename_choice = input("是否要重命名输出列？(y/n，留空为 n): ").strip().lower()
    column_mapping = {}

    if rename_choice in ('y', 'yes', 'Y'):
        print("\n请为每一列输入新的列名（留空则保持原列名）:")
        for col in selected_columns:
            new_name = input(f"将 '{col}' 重命名为（留空保持不变）: ").strip()
            column_mapping[col] = new_name if new_name else col
    else:
        # 不重命名，保持原名
        column_mapping = {col: col for col in selected_columns}

    # 构建最终列名列表
    final_columns = [column_mapping[col] for col in selected_columns]

    # 5. 指定起始行和结束行
    try:
        start_row_input = input("请输入起始行号（从 0 开始）（如果您需要从第5行开始截取，您需要输入4）: ").strip()
        start_row = int(start_row_input) if start_row_input else 0
        if start_row < 0 or start_row >= total_rows:
            print(f"起始行必须在 0 到 {total_rows - 1} 之间")
            sys.exit(1)
    except ValueError:
        print("请输入有效数字！")
        sys.exit(1)

    end_row_input = input("请输入结束行号（留空表示到最后）: ").strip()
    if end_row_input:
        try:
            end_row = int(end_row_input)
            if end_row <= start_row or end_row > total_rows:
                print(f"结束行必须大于起始行且不超过 {total_rows}")
                sys.exit(1)
        except ValueError:
            print("请输入有效数字或留空！")
            sys.exit(1)
    else:
        end_row = total_rows

    # 6. 指定每个文件最大行数
    try:
        chunk_size_input = input("请输入每个输出文件的最大行数（如 500）: ").strip()
        chunk_size = int(chunk_size_input)
        if chunk_size <= 0:
            print("行数必须大于 0！")
            sys.exit(1)
    except ValueError:
        print("请输入有效数字！")
        sys.exit(1)

    # 7. 输出格式选择
    output_format = input("请选择输出格式 (1: CSV, 2: XLSX)（默认 1）: ").strip()
    if output_format == "2":
        output_ext = ".xlsx"
    else:
        output_ext = ".csv"
    print(f"输出格式: {output_ext}")

    # 8. 输出目录
    output_dir = input("请输入输出目录（留空为当前目录）: ").strip()
    if not output_dir:
        output_dir = "."
    os.makedirs(output_dir, exist_ok=True)

    # 9. 提取数据并重命名列
    subset_df = df.iloc[start_row:end_row][selected_columns].copy()
    subset_df.columns = final_columns  # 应用新列名
    total_subset_rows = len(subset_df)

    # 10. 分块并输出
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    num_files = 0

    print(f"\n正在分割 {total_subset_rows} 行数据，每文件最多 {chunk_size} 行...\n")

    for i in range(0, total_subset_rows, chunk_size):
        chunk_df = subset_df.iloc[i:i + chunk_size]
        output_file = os.path.join(output_dir, f"{base_name}_part_{num_files + 1}{output_ext}")

        try:
            if output_ext == ".csv":
                chunk_df.to_csv(output_file, index=False, encoding='utf-8')
            else:
                chunk_df.to_excel(output_file, index=False, sheet_name="Sheet1")
            print(f"✓ 已保存: {output_file} ({len(chunk_df)} 行)")
            num_files += 1
        except Exception as e:
            print(f"✗ 保存失败 {output_file}: {e}")

    print(f"\n✅ 分割完成！共生成 {num_files} 个文件，保存在 '{output_dir}' 目录下。")

if __name__ == "__main__":
    split_excel_or_csv()