import os
import io
import zipfile
import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # 启用 CORS，允许前端跨域访问

# 工具函数
def read_uploaded_file(file_storage, dtype=None):
    filename = file_storage.filename
    ext = os.path.splitext(filename)[1].lower()
    df = None
    if ext == '.csv':
        try:
            df = pd.read_csv(file_storage, dtype=dtype)
        except UnicodeDecodeError:
            df = pd.read_csv(file_storage, dtype=dtype, encoding='gbk')
    elif ext in ['.xlsx', '.xls']:
        df = pd.read_excel(file_storage, dtype=dtype)
    return df

@app.route('/api/get_file_info', methods=['POST'])
def get_file_info():
    file_storage = request.files.get('file')
    if not file_storage:
        return jsonify({'error': '未找到文件'}), 400
    try:
        df = read_uploaded_file(file_storage)
        if df is None:
            return jsonify({'error': '无法读取文件，请确保文件格式正确'}), 400
        columns = df.columns.tolist()
        row_count = len(df)
        return jsonify({'columns': columns, 'row_count': row_count})
    except Exception as e:
        return jsonify({'error': f'获取文件信息失败: {e}'}), 500

@app.route('/api/merge', methods=['POST'])
def merge_files_api():
    try:
        files = request.files.getlist('files')
        if not files:
            return "未选择任何文件！", 400
        
        merge_columns_str = request.form.get('merge_cols')
        rename_columns_str = request.form.get('rename_cols')
        
        if merge_columns_str:
            merge_columns = [col.strip() for col in merge_columns_str.split(',')]
            rename_columns = [col.strip() for col in rename_columns_str.split(',')]
            if len(merge_columns) != len(rename_columns):
                return "原始列名和新列名数量不一致", 400
            rename_map = dict(zip(merge_columns, rename_columns))
        else:
            merge_columns = None
            rename_map = None

        combined_df = pd.DataFrame()
        
        for file_storage in files:
            df = read_uploaded_file(file_storage)
            if df is not None:
                if merge_columns:
                    df_to_concat = pd.DataFrame(columns=rename_columns)
                    for original_col, new_col in rename_map.items():
                        if original_col in df.columns:
                            df_to_concat[new_col] = df[original_col]
                        else:
                            df_to_concat[new_col] = np.nan
                    combined_df = pd.concat([combined_df, df_to_concat], ignore_index=True, sort=False)
                else:
                    combined_df = pd.concat([combined_df, df], ignore_index=True, sort=False)

        output_format = request.form.get('format')
        output_name = request.form.get('output_name', 'merged_output')
        
        output_buffer = io.BytesIO()
        if output_format == 'csv':
            combined_df.to_csv(output_buffer, index=False, encoding='utf-8-sig')
            output_buffer.seek(0)
            return send_file(output_buffer, as_attachment=True, download_name=f'{output_name}.csv', mimetype='text/csv')
        else:
            combined_df.to_excel(output_buffer, index=False)
            output_buffer.seek(0)
            return send_file(output_buffer, as_attachment=True, download_name=f'{output_name}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return f"合并失败: {e}", 500

@app.route('/api/split', methods=['POST'])
def split_files_api():
    try:
        file_storage = request.files['file']
        output_name = request.form.get('output_name', 'split_output')
        output_format = request.form.get('format')
        
        df = read_uploaded_file(file_storage, dtype=str)
        if df is None:
            return "文件读取失败", 400

        output_columns_str = request.form.get('output_cols')
        rename_columns_str = request.form.get('rename_cols')
        
        if output_columns_str:
            output_columns = [col.strip() for col in output_columns_str.split(',')]
            rename_columns = [col.strip() for col in rename_columns_str.split(',')]
            if len(output_columns) != len(rename_columns):
                return "原始列名和新列名数量不一致", 400
            rename_map = dict(zip(output_columns, rename_columns))
            
            missing_cols = [col for col in output_columns if col not in df.columns]
            if missing_cols:
                return f"指定的输出列 '{', '.join(missing_cols)}' 不存在", 400
            
            df = df.rename(columns=rename_map)
            df = df[[rename_map[col] for col in output_columns]]
        
        slice_method = request.form.get('method')
        sliced_dfs = []
        
        if slice_method == 'count':
            start_row = int(request.form.get('start_row')) - 1
            row_count = int(request.form.get('row_count'))
            slice_times = int(request.form.get('times'))
            
            for i in range(slice_times):
                start = start_row + i * row_count
                end = start + row_count
                if start >= len(df):
                    break
                sliced_dfs.append(df.iloc[start:end])
        elif slice_method == 'range':
            start_row = int(request.form.get('range_start')) - 1
            end_row = int(request.form.get('range_end'))
            num_slices = int(request.form.get('sections'))
            
            total_rows = end_row - start_row
            slice_length = total_rows // num_slices
            
            for i in range(num_slices):
                start = start_row + i * slice_length
                end = start + slice_length
                sliced_dfs.append(df.iloc[start:end])
        else:
            sliced_dfs.append(df)

        if not sliced_dfs:
            return "没有可分割的数据", 400

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for i, df_slice in enumerate(sliced_dfs):
                temp_buffer = io.BytesIO()
                part_name = f"{output_name}_part_{i+1}"
                
                if output_format == 'csv':
                    df_slice.to_csv(temp_buffer, index=False, encoding='utf-8-sig')
                    zip_file.writestr(f'{part_name}.csv', temp_buffer.getvalue())
                else:
                    df_slice.to_excel(temp_buffer, index=False)
                    zip_file.writestr(f'{part_name}.xlsx', temp_buffer.getvalue())
        
        zip_buffer.seek(0)
        return send_file(zip_buffer, as_attachment=True, download_name=f'{output_name}.zip', mimetype='application/zip')
    except Exception as e:
        return f"分割失败: {e}", 500

@app.route('/api/deduplicate', methods=['POST'])
def deduplicate_api():
    try:
        main_file = request.files['main_file']
        ref_files = request.files.getlist('ref_files')
        main_column = request.form.get('main_col')
        output_name = request.form.get('output_name', 'deduplicated')
        dedupe_mode = request.form.get('mode')
        fuzzy_threshold = int(request.form.get('fuzzy_threshold', 80))
        output_format = request.form.get('format')
        
        main_df = read_uploaded_file(main_file, dtype=str)
        if main_df is None or main_column not in main_df.columns:
            return "主文件或列名无效", 400
        
        all_ref_values = set()
        for i, file_storage in enumerate(ref_files):
            ref_col_name = request.form.get(f'ref_col_{i}')
            ref_df = read_uploaded_file(file_storage, dtype=str)
            if ref_df is not None:
                if ref_col_name in ref_df.columns:
                    values = set(ref_df[ref_col_name].astype(str).str.strip().dropna())
                    all_ref_values.update(values)
        
        if dedupe_mode == 'exact':
            def is_duplicate(row):
                key = str(row.get(main_column)).strip()
                return key in all_ref_values if pd.notna(key) else False
        elif dedupe_mode == 'email':
            lower_ref_values = {v.lower() for v in all_ref_values}
            def is_duplicate(row):
                key = str(row.get(main_column)).strip().lower()
                return key in lower_ref_values if pd.notna(key) else False
        elif dedupe_mode == 'fuzzy':
            def is_duplicate(row):
                key = str(row.get(main_column)).strip()
                if not key:
                    return False
                for ref_value in all_ref_values:
                    if fuzz.ratio(key, ref_value) >= fuzzy_threshold:
                        return True
                return False
        
        mask = main_df.apply(is_duplicate, axis=1)
        filtered_df = main_df[~mask]
        
        output_buffer = io.BytesIO()
        if output_format == 'csv':
            filtered_df.to_csv(output_buffer, index=False, encoding='utf-8-sig')
            output_buffer.seek(0)
            return send_file(output_buffer, as_attachment=True, download_name=f'{output_name}.csv', mimetype='text/csv')
        else:
            filtered_df.to_excel(output_buffer, index=False)
            output_buffer.seek(0)
            return send_file(output_buffer, as_attachment=True, download_name=f'{output_name}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return f"查重失败: {e}", 500

@app.route('/api/clean', methods=['POST'])
def clean_api():
    try:
        file_storage = request.files['file']
        check_columns_str = request.form.get('cols')
        check_columns = [col.strip() for col in check_columns_str.split(',')]
        output_name = request.form.get('output_name', 'cleaned_output')

        df = read_uploaded_file(file_storage)
        if df is None:
            return "文件读取失败", 400
        
        missing_cols = [col for col in check_columns if col not in df.columns]
        if missing_cols:
            return f"以下列在文件中未找到: {', '.join(missing_cols)}", 400

        mask = pd.Series([True] * len(df), index=df.index)
        for col in check_columns:
            col_not_empty = df[col].notna() & (df[col].astype(str).str.strip() != '')
            mask &= col_not_empty
        
        cleaned_df = df[mask].reset_index(drop=True)

        output_buffer = io.BytesIO()
        cleaned_df.to_excel(output_buffer, index=False)
        output_buffer.seek(0)
        return send_file(output_buffer, as_attachment=True, download_name=f'{output_name}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        return f"清理失败: {e}", 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)