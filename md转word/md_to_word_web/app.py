# -*- coding: utf-8 -*-
"""
Markdown 转 Word Web 应用
基于 Flask 框架的简单 Web 界面
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
from werkzeug.utils import secure_filename
from flask import Flask, render_template, request, send_file, jsonify

# 导入转换模块
from convert import convert_single_file

app = Flask(__name__)

# 配置
UPLOAD_DIR = Path(__file__).parent / 'uploads'
UPLOAD_DIR.mkdir(exist_ok=True)

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'.md', '.markdown'}
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB

# 临时转换文件存储
TEMP_DIR = UPLOAD_DIR / 'temp'
TEMP_DIR.mkdir(exist_ok=True)


def allowed_file(filename):
    """检查文件是否被允许"""
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """显示上传页面"""
    return render_template('index.html')


@app.route('/api/convert', methods=['POST'])
def convert():
    """处理文件上传和转换"""

    try:
        # 检查是否有文件在请求中
        if 'file' not in request.files:
            return jsonify({'success': False, 'message': '没有选择文件'}), 400

        file = request.files['file']

        # 检查文件名
        if file.filename == '':
            return jsonify({'success': False, 'message': '文件名为空'}), 400

        # 检查文件类型
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'message': '只支持 .md 和 .markdown 文件'}), 400

        # 保存上传的文件
        filename = secure_filename(file.filename)
        input_path = TEMP_DIR / filename
        file.save(str(input_path))

        # 检查文件大小
        if input_path.stat().st_size > MAX_FILE_SIZE:
            input_path.unlink()
            return jsonify({'success': False, 'message': '文件超过 50MB 限制'}), 400

        # 生成输出文件路径
        output_filename = filename.replace('.md', '.docx').replace('.markdown', '.docx')
        output_path = TEMP_DIR / output_filename

        # 解析转换选项
        generate_toc = request.form.get('generate_toc', 'true').lower() == 'true'
        toc_depth = request.form.get('toc_depth', '3')
        highlight_style = request.form.get('highlight_style', 'tango')

        # 暂时不支持这些选项的修改（保留未来扩展）
        # 直接调用 convert_single_file 使用默认设置

        print(f"开始转换: {filename}")
        print(f"选项: TOC={generate_toc}, 深度={toc_depth}, 高亮={highlight_style}")

        # 执行转换
        success = convert_single_file(str(input_path), str(output_path))

        # 删除上传的文件
        if input_path.exists():
            input_path.unlink()

        if success and output_path.exists():
            # 返回下载链接
            download_url = f'/api/download/{output_filename}'
            return jsonify({
                'success': True,
                'message': '转换成功！',
                'download_url': download_url,
                'filename': output_filename
            })
        else:
            if output_path.exists():
                output_path.unlink()
            return jsonify({
                'success': False,
                'message': '转换失败，请检查文件内容是否正确'
            }), 500

    except Exception as e:
        print(f"错误: {str(e)}")
        return jsonify({
            'success': False,
            'message': f'服务器错误: {str(e)}'
        }), 500


@app.route('/api/download/<filename>')
def download(filename):
    """下载转换后的文件"""
    try:
        # 安全性检查：防止目录遍历
        filename = secure_filename(filename)
        file_path = TEMP_DIR / filename

        # 检查文件是否存在
        if not file_path.exists() or not file_path.is_file():
            return jsonify({'error': '文件不存在'}), 404

        # 发送文件
        response = send_file(
            str(file_path),
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

        # 在后台删除文件（下载后）
        # 这里我们先返回文件，让客户端下载后再请求删除
        return response

    except Exception as e:
        print(f"下载错误: {str(e)}")
        return jsonify({'error': '下载失败'}), 500


@app.route('/api/cleanup/<filename>', methods=['POST'])
def cleanup(filename):
    """清理临时文件"""
    try:
        filename = secure_filename(filename)
        file_path = TEMP_DIR / filename

        if file_path.exists():
            file_path.unlink()
            return jsonify({'success': True, 'message': '文件已删除'})
        else:
            return jsonify({'success': False, 'message': '文件不存在'})

    except Exception as e:
        print(f"清理错误: {str(e)}")
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/health')
def health():
    """健康检查接口"""
    return jsonify({
        'status': 'healthy',
        'app': 'Markdown to Word Converter',
        'version': '1.0'
    })


def cleanup_old_files():
    """定期清理 24 小时前的临时文件"""
    import time
    now = time.time()
    max_age = 24 * 3600  # 24 小时

    if TEMP_DIR.exists():
        for file_path in TEMP_DIR.glob('*'):
            if file_path.is_file():
                if now - file_path.stat().st_mtime > max_age:
                    try:
                        file_path.unlink()
                    except Exception as e:
                        print(f"删除文件 {file_path} 失败: {str(e)}")


if __name__ == '__main__':
    # 在启动前清理旧文件
    cleanup_old_files()

    # 运行 Flask 应用
    print("启动 Markdown 转 Word Web 应用...")
    print("打开浏览器访问: http://localhost:5000")

    app.run(debug=False, host='localhost', port=5000, threaded=True)
