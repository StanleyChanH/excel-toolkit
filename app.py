"""
Excel批量操作工具箱 - 后端Flask应用
为非技术用户提供简单的Excel文件批量处理功能
Author: StanleyChanH
License: MIT
"""

import os
import io
import zipfile
import tempfile
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import pandas as pd
import openpyxl

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}

def allowed_file(filename):
    """检查文件扩展名是否被允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def create_temp_excel(dataframe, filename=None):
    """创建临时的Excel文件并返回字节流"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, sheet_name='Sheet1', index=False)
    output.seek(0)
    return output

def create_temp_csv(dataframe, filename=None):
    """创建临时的CSV文件并返回字节流"""
    output = io.BytesIO()
    dataframe.to_csv(output, index=False, encoding='utf-8-sig')
    output.seek(0)
    return output

def create_zip_file(files_dict):
    """创建包含多个文件的ZIP压缩包"""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, file_content in files_dict.items():
            zip_file.writestr(filename, file_content.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

@app.route('/')
def index():
    """主页"""
    return render_template('index.html')

@app.route('/api/merge-files', methods=['POST'])
def merge_files():
    """
    合并多个Excel文件
    """
    try:
        if 'files' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        files = request.files.getlist('files')
        if not files or files[0].filename == '':
            return jsonify({'error': '请选择文件'}), 400

        keep_headers = request.form.get('keep_headers', 'true').lower() == 'true'
        add_source_column = request.form.get('add_source_column', 'false').lower() == 'true'

        all_data = []
        header_saved = False

        for file in files:
            if not allowed_file(file.filename):
                continue

            # 读取Excel文件
            if file.filename.endswith('.csv'):
                df = pd.read_csv(file, encoding='utf-8-sig')
            else:
                df = pd.read_excel(file, engine='openpyxl')

            # 添加来源文件列
            if add_source_column:
                df['来源文件'] = file.filename

            # 处理标题行
            if not header_saved and keep_headers:
                all_data.append(df)
                header_saved = True
            elif header_saved and keep_headers:
                all_data.append(df.iloc[1:])  # 跳过标题行
            else:
                all_data.append(df)

        if not all_data:
            return jsonify({'error': '没有有效的Excel文件'}), 400

        # 合并所有数据
        merged_df = pd.concat(all_data, ignore_index=True)

        # 创建输出文件
        output = create_temp_excel(merged_df)
        filename = f"合并结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

@app.route('/api/merge-sheets', methods=['POST'])
def merge_sheets():
    """
    合并单个文件的多个Sheet
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '请选择文件'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': '只支持Excel文件'}), 400

        add_sheet_column = request.form.get('add_sheet_column', 'false').lower() == 'true'

        # 读取所有sheet
        with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
            file.save(tmp_file.name)
            tmp_file_path = tmp_file.name

        try:
            excel_file = pd.ExcelFile(tmp_file_path)
            all_data = []

            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(tmp_file_path, sheet_name=sheet_name, engine='openpyxl')

                # 添加来源Sheet列
                if add_sheet_column:
                    df['来源Sheet'] = sheet_name

                # 跳过第一个sheet的标题行，保留其他sheet的标题行
                if excel_file.sheet_names.index(sheet_name) > 0:
                    df = df.iloc[1:] if len(df) > 0 else df

                all_data.append(df)

            # 合并所有数据
            merged_df = pd.concat(all_data, ignore_index=True)

            # 创建输出文件
            output = create_temp_excel(merged_df)
            filename = f"合并Sheet结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

            return send_file(
                output,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        finally:
            # 删除临时文件
            os.unlink(tmp_file_path)

    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

@app.route('/api/split-by-column', methods=['POST'])
def split_by_column():
    """
    按列拆分Sheet
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        file = request.files['file']
        column_name = request.form.get('column_name', '').strip()

        if file.filename == '':
            return jsonify({'error': '请选择文件'}), 400

        if not column_name:
            return jsonify({'error': '请输入列名'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': '只支持Excel文件'}), 400

        # 读取文件
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file, encoding='utf-8-sig')
        else:
            df = pd.read_excel(file, engine='openpyxl')

        # 检查列是否存在
        if column_name not in df.columns:
            return jsonify({'error': f'列名 "{column_name}" 不存在'}), 400

        # 按列的唯一值拆分
        unique_values = df[column_name].unique()
        files_dict = {}

        for value in unique_values:
            # 过滤数据
            filtered_df = df[df[column_name] == value]

            # 创建文件
            output = create_temp_excel(filtered_df)
            # 安全文件名
            safe_value = str(value).replace('/', '_').replace('\\', '_')
            filename = f"{safe_value}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            files_dict[filename] = output

        # 创建ZIP文件
        zip_buffer = create_zip_file(files_dict)
        zip_filename = f"按列拆分结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )

    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

@app.route('/api/split-by-rows', methods=['POST'])
def split_by_rows():
    """
    按行数拆分Sheet
    """
    try:
        if 'file' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        file = request.files['file']
        rows_per_file = request.form.get('rows_per_file', '').strip()

        if file.filename == '':
            return jsonify({'error': '请选择文件'}), 400

        if not rows_per_file:
            return jsonify({'error': '请输入每文件的行数'}), 400

        try:
            rows_per_file = int(rows_per_file)
            if rows_per_file <= 0:
                raise ValueError()
        except ValueError:
            return jsonify({'error': '请输入有效的正整数'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': '只支持Excel文件'}), 400

        # 读取文件
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file, encoding='utf-8-sig')
        else:
            df = pd.read_excel(file, engine='openpyxl')

        # 按行数拆分
        total_rows = len(df)
        num_files = (total_rows + rows_per_file - 1) // rows_per_file
        files_dict = {}

        for i in range(num_files):
            start_idx = i * rows_per_file
            end_idx = min((i + 1) * rows_per_file, total_rows)

            # 切片数据
            split_df = df.iloc[start_idx:end_idx]

            # 创建文件
            output = create_temp_excel(split_df)
            filename = f"第{i+1}部分_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            files_dict[filename] = output

        # 创建ZIP文件
        zip_buffer = create_zip_file(files_dict)
        zip_filename = f"按行拆分结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )

    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

@app.route('/api/find-replace', methods=['POST'])
def find_replace():
    """
    批量查找与替换
    """
    try:
        if 'files' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        files = request.files.getlist('files')
        find_text = request.form.get('find_text', '').strip()
        replace_text = request.form.get('replace_text', '').strip()

        if not files or files[0].filename == '':
            return jsonify({'error': '请选择文件'}), 400

        if not find_text:
            return jsonify({'error': '请输入查找内容'}), 400

        files_dict = {}

        for file in files:
            if not allowed_file(file.filename):
                continue

            # 读取Excel文件的所有sheet
            if file.filename.endswith('.csv'):
                df = pd.read_csv(file, encoding='utf-8-sig')
                # 执行查找替换
                df = df.replace(find_text, replace_text, regex=False)

                # 创建输出文件
                output = create_temp_csv(df)
                filename = f"替换_{secure_filename(file.filename)}"
                files_dict[filename] = output
            else:
                # 处理Excel文件
                with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
                    file.save(tmp_file.name)
                    tmp_file_path = tmp_file.name

                try:
                    excel_file = pd.ExcelFile(tmp_file_path)
                    sheet_data = {}

                    for sheet_name in excel_file.sheet_names:
                        df = pd.read_excel(tmp_file_path, sheet_name=sheet_name, engine='openpyxl')
                        # 执行查找替换
                        df = df.replace(find_text, replace_text, regex=False)
                        sheet_data[sheet_name] = df

                    # 创建新的Excel文件
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for sheet_name, df in sheet_data.items():
                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                    output.seek(0)

                    filename = f"替换_{secure_filename(file.filename)}"
                    files_dict[filename] = output

                finally:
                    # 删除临时文件
                    os.unlink(tmp_file_path)

        if not files_dict:
            return jsonify({'error': '没有有效的文件'}), 400

        # 创建ZIP文件
        zip_buffer = create_zip_file(files_dict)
        zip_filename = f"查找替换结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )

    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

@app.route('/api/delete-columns', methods=['POST'])
def delete_columns():
    """
    批量删除指定列
    """
    try:
        if 'files' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        files = request.files.getlist('files')
        columns_text = request.form.get('columns', '').strip()

        if not files or files[0].filename == '':
            return jsonify({'error': '请选择文件'}), 400

        if not columns_text:
            return jsonify({'error': '请输入要删除的列名'}), 400

        # 解析列名
        columns_to_delete = [col.strip() for col in columns_text.split(',')]
        columns_to_delete = [col for col in columns_to_delete if col]  # 移除空字符串

        if not columns_to_delete:
            return jsonify({'error': '请输入有效的列名'}), 400

        files_dict = {}

        for file in files:
            if not allowed_file(file.filename):
                continue

            # 读取文件
            if file.filename.endswith('.csv'):
                df = pd.read_csv(file, encoding='utf-8-sig')
            else:
                df = pd.read_excel(file)

            # 删除列（只删除存在的列）
            existing_columns = [col for col in columns_to_delete if col in df.columns]
            if existing_columns:
                df = df.drop(columns=existing_columns)

            # 创建输出文件
            if file.filename.endswith('.csv'):
                output = create_temp_csv(df)
            else:
                output = create_temp_excel(df)

            filename = f"删除列_{secure_filename(file.filename)}"
            if file.filename.endswith('.csv'):
                filename = filename.replace('.xlsx', '.csv')
            files_dict[filename] = output

        if not files_dict:
            return jsonify({'error': '没有有效的文件'}), 400

        # 创建ZIP文件
        zip_buffer = create_zip_file(files_dict)
        zip_filename = f"删除列结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )

    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

@app.route('/api/filter-data', methods=['POST'])
def filter_data():
    """
    批量数据筛选
    """
    try:
        if 'files' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        files = request.files.getlist('files')
        column_name = request.form.get('column_name', '').strip()
        condition = request.form.get('condition', '').strip()
        value = request.form.get('value', '').strip()

        if not files or files[0].filename == '':
            return jsonify({'error': '请选择文件'}), 400

        if not all([column_name, condition, value]):
            return jsonify({'error': '请填写所有筛选条件'}), 400

        # 筛选条件映射
        conditions = {
            '等于': lambda df, col, val: df[col] == val,
            '不等于': lambda df, col, val: df[col] != val,
            '包含': lambda df, col, val: df[col].astype(str).str.contains(val, na=False),
            '不包含': lambda df, col, val: ~df[col].astype(str).str.contains(val, na=False),
            '大于': lambda df, col, val: pd.to_numeric(df[col], errors='coerce') > pd.to_numeric(val, errors='coerce'),
            '小于': lambda df, col, val: pd.to_numeric(df[col], errors='coerce') < pd.to_numeric(val, errors='coerce'),
            '大于等于': lambda df, col, val: pd.to_numeric(df[col], errors='coerce') >= pd.to_numeric(val, errors='coerce'),
            '小于等于': lambda df, col, val: pd.to_numeric(df[col], errors='coerce') <= pd.to_numeric(val, errors='coerce'),
        }

        if condition not in conditions:
            return jsonify({'error': '无效的筛选条件'}), 400

        all_filtered_data = []

        for file in files:
            if not allowed_file(file.filename):
                continue

            # 读取文件
            if file.filename.endswith('.csv'):
                df = pd.read_csv(file, encoding='utf-8-sig')
            else:
                df = pd.read_excel(file)

            # 检查列是否存在
            if column_name not in df.columns:
                continue

            # 添加来源文件列
            df['来源文件'] = file.filename

            # 应用筛选条件
            try:
                if condition in ['大于', '小于', '大于等于', '小于等于']:
                    # 数值比较
                    mask = conditions[condition](df, column_name, value)
                else:
                    # 文本比较
                    mask = conditions[condition](df, column_name, value)

                filtered_df = df[mask]
                if not filtered_df.empty:
                    all_filtered_data.append(filtered_df)
            except Exception as e:
                # 如果筛选失败，跳过这个文件
                continue

        if not all_filtered_data:
            return jsonify({'error': '没有找到符合条件的数据'}), 400

        # 合并所有筛选结果
        result_df = pd.concat(all_filtered_data, ignore_index=True)

        # 创建输出文件
        output = create_temp_excel(result_df)
        filename = f"筛选结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

@app.route('/api/convert-format', methods=['POST'])
def convert_format():
    """
    格式转换（XLSX <-> CSV）
    """
    try:
        if 'files' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        files = request.files.getlist('files')
        convert_type = request.form.get('convert_type', '').strip()

        if not files or files[0].filename == '':
            return jsonify({'error': '请选择文件'}), 400

        if not convert_type:
            return jsonify({'error': '请选择转换类型'}), 400

        files_dict = {}

        for file in files:
            if not allowed_file(file.filename):
                continue

            filename = secure_filename(file.filename)
            base_name = os.path.splitext(filename)[0]

            if convert_type == 'xlsx_to_csv':
                # XLSX转CSV
                if file.filename.endswith('.xlsx') or file.filename.endswith('.xls'):
                    # 读取Excel文件
                    excel_file = pd.ExcelFile(file)

                    # 每个sheet转换为一个CSV文件
                    for sheet_name in excel_file.sheet_names:
                        df = pd.read_excel(file, sheet_name=sheet_name, engine='openpyxl')

                        # 创建CSV文件
                        output = create_temp_csv(df)
                        csv_filename = f"{base_name}_{sheet_name}.csv"
                        files_dict[csv_filename] = output

            elif convert_type == 'csv_to_xlsx':
                # CSV转XLSX
                if file.filename.endswith('.csv'):
                    # 读取CSV文件
                    df = pd.read_csv(file, encoding='utf-8-sig')

                    # 创建Excel文件
                    output = create_temp_excel(df)
                    xlsx_filename = f"{base_name}.xlsx"
                    files_dict[xlsx_filename] = output

        if not files_dict:
            return jsonify({'error': '没有可以转换的文件'}), 400

        # 创建ZIP文件
        zip_buffer = create_zip_file(files_dict)
        zip_filename = f"格式转换结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )

    except Exception as e:
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

@app.errorhandler(413)
def too_large(e):
    """文件过大错误处理"""
    return jsonify({'error': '文件过大，请上传小于100MB的文件'}), 413

@app.errorhandler(404)
def not_found(e):
    """404错误处理"""
    return jsonify({'error': '页面未找到'}), 404

@app.errorhandler(500)
def internal_error(e):
    """500错误处理"""
    return jsonify({'error': '服务器内部错误'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)