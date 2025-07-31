import pandas as pd
import os
import sys

def split_excel():
    print("=== Excel 文件分割工具 ===\n")

    # 1. 输入文件路径
    file_path = input("请输入 Excel 文件路径: ").strip()
    if not os.path.exists(file_path):
        print(f"错误：文件 '{file_path}' 不存在！")
        sys.exit(1)

    # 2. 读取所有 sheet 名称
    try:
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        print(f"可用的工作表: {sheet_names}")
    except Exception as e:
        print(f"无法读取 Excel 文件: {e}")
        sys.exit(1)

    sheet_name = input(f"请输入要读取的工作表名称（默认第一个 '{sheet_names[0]}'）: ").strip()
    if not sheet_name:
        sheet_name = sheet_names[0]
    if sheet_name not in sheet_names:
        print(f"错误：工作表 '{sheet_name}' 不存在！")
        sys.exit(1)

    # 3. 读取数据（用于获取列名和行数）
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"读取工作表 '{sheet_name}' 失败: {e}")
        sys.exit(1)

    total_rows = len(df)
    columns = df.columns.tolist()
    print(f"数据总行数: {total_rows}")
    print(f"可用列名: {columns}")

    # 4. 指定列
    selected_columns_input = input("请输入要提取的列名（多个列用英文逗号分隔，留空表示全部列）: ").strip()
    if selected_columns_input:
        selected_columns = [col.strip() for col in selected_columns_input.split(',')]
        # 检查列是否存在
        missing_cols = [col for col in selected_columns if col not in columns]
        if missing_cols:
            print(f"错误：以下列不存在: {missing_cols}")
            sys.exit(1)
    else:
        selected_columns = columns

    # 5. 指定起始行和结束行
    try:
        start_row_input = input("请输入起始行号（从 1 开始，0 表示第一行数据）(比如你要读取第3行数据，你需要输入2）: ").strip()
        start_row = int(start_row_input) if start_row_input else 0
        if start_row < 0:
            print("起始行不能小于 0！")
            sys.exit(1)
        if start_row >= total_rows:
            print(f"起始行 {start_row} 超出数据范围（共 {total_rows} 行）")
            sys.exit(1)
    except ValueError:
        print("请输入有效的数字！")
        sys.exit(1)

    end_row_input = input("请输入结束行号（留空表示到最后）: ").strip()
    if end_row_input:
        try:
            end_row = int(end_row_input)
            if end_row < start_row:
                print("结束行不能小于起始行！")
                sys.exit(1)
            if end_row > total_rows:
                print(f"结束行超过最大行数，自动设为 {total_rows}")
                end_row = total_rows
        except ValueError:
            print("请输入有效的数字或留空！")
            sys.exit(1)
    else:
        end_row = total_rows

    # 6. 指定每个文件的最大行数
    try:
        chunk_size_input = input("请输入每个输出文件的最大行数（例如 500）: ").strip()
        chunk_size = int(chunk_size_input)
        if chunk_size <= 0:
            print("行数必须大于 0！")
            sys.exit(1)
    except ValueError:
        print("请输入有效的数字！")
        sys.exit(1)

    # 7. 输出目录
    output_dir = input("请输入输出目录（留空为当前目录）: ").strip()
    if not output_dir:
        output_dir = "."
    os.makedirs(output_dir, exist_ok=True)

    # 8. 提取指定范围的数据
    subset_df = df.iloc[start_row:end_row][selected_columns].reset_index(drop=True)

    # 9. 分割并保存
    total_subset_rows = len(subset_df)
    num_chunks = (total_subset_rows + chunk_size - 1) // 2  # 向上取整
    print(f"\n正在分割 {total_subset_rows} 行数据，每文件最多 {chunk_size} 行...")

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    for i in range(0, total_subset_rows, chunk_size):
        chunk_df = subset_df.iloc[i:i + chunk_size]
        output_file = os.path.join(output_dir, f"{base_name}_part_{i//chunk_size + 1}.xlsx")
        try:
            chunk_df.to_excel(output_file, index=False, sheet_name="Sheet1")
            print(f"已保存: {output_file} ({len(chunk_df)} 行)")
        except Exception as e:
            print(f"保存文件 {output_file} 失败: {e}")

    print(f"\n✅ 分割完成！共生成 { (total_subset_rows - 1) // chunk_size + 1 } 个文件，保存在 '{output_dir}' 目录下。")

if __name__ == "__main__":
    split_excel()