# -*- coding: utf-8 -*-
"""
转换服务类
封装 Markdown 到 Word 的转换逻辑
"""

import os
import sys
from typing import Tuple, Dict, Any, Optional

# 添加父目录到路径，以便导入 convert 模块
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 导入转换模块
import convert as convert_module


class ConverterService:
    """Markdown 转 Word 转换服务"""

    def __init__(self):
        """初始化转换服务"""
        self.pandoc_path = convert_module.PANDOC_PATH

    def check_pandoc(self) -> Tuple[bool, str]:
        """
        检查 Pandoc 是否可用

        Returns:
            (是否可用, 消息)
        """
        if os.path.exists(self.pandoc_path):
            return True, f"Pandoc 路径: {self.pandoc_path}"
        return False, f"找不到 Pandoc: {self.pandoc_path}"

    def convert(
        self,
        input_file: str,
        output_file: str,
        options: Optional[Dict[str, Any]] = None
    ) -> Tuple[bool, str]:
        """
        转换 Markdown 文件到 Word 文档

        Args:
            input_file: 输入的 Markdown 文件路径
            output_file: 输出的 Word 文件路径
            options: 转换选项
                - generate_toc: 是否生成目录 (默认 True)
                - toc_depth: 目录深度 (默认 3)
                - highlight_style: 代码高亮样式 (默认 'tango')

        Returns:
            (是否成功, 消息)
        """
        # 默认选项
        if options is None:
            options = {}

        generate_toc = options.get('generate_toc', True)
        toc_depth = options.get('toc_depth', 3)
        highlight_style = options.get('highlight_style', 'tango')

        # 检查输入文件
        if not os.path.exists(input_file):
            return False, f"输入文件不存在: {input_file}"

        # 检查 Pandoc
        pandoc_ok, pandoc_msg = self.check_pandoc()
        if not pandoc_ok:
            return False, pandoc_msg

        try:
            # 调用转换函数（带选项）
            success = convert_with_options(
                input_file,
                output_file,
                generate_toc=generate_toc,
                toc_depth=toc_depth,
                highlight_style=highlight_style
            )

            if success:
                file_size = os.path.getsize(output_file)
                size_kb = file_size / 1024
                return True, f"转换成功! 文件大小: {size_kb:.1f} KB"
            else:
                return False, "转换失败，请检查 Markdown 文件格式"

        except Exception as e:
            return False, f"转换过程中发生错误: {str(e)}"


def convert_with_options(
    input_md: str,
    output_docx: str,
    generate_toc: bool = True,
    toc_depth: int = 3,
    highlight_style: str = 'tango'
) -> bool:
    """
    带选项的转换函数

    Args:
        input_md: 输入 Markdown 文件路径
        output_docx: 输出 Word 文件路径
        generate_toc: 是否生成目录
        toc_depth: 目录深度
        highlight_style: 代码高亮样式

    Returns:
        是否成功
    """
    import subprocess
    import tempfile

    # 参考模板放在输出文件同目录
    output_dir = os.path.dirname(output_docx) or '.'
    reference_docx = os.path.join(output_dir, 'reference_template.docx')

    # 步骤 1: 创建参考模板
    template_created = convert_module.create_reference_docx(reference_docx)

    # 步骤 2: 预处理 Markdown 内容
    with open(input_md, 'r', encoding='utf-8') as f:
        content = f.read()

    processed_content = convert_module.preprocess_markdown(content)

    # 创建临时文件
    input_dir = os.path.dirname(os.path.abspath(input_md))
    temp_md = os.path.join(input_dir, '_temp_processed.md')

    with open(temp_md, 'w', encoding='utf-8') as f:
        f.write(processed_content)

    # 步骤 3: 构建 Pandoc 命令
    cmd = [
        convert_module.PANDOC_PATH,
        temp_md,
        '-o', output_docx,
        '--from', 'markdown+tex_math_dollars+raw_tex',
        '--to', 'docx',
        '--standalone',
        '--wrap=auto',
        '--resource-path', input_dir,
        '--highlight-style', highlight_style,
    ]

    # 添加目录选项
    if generate_toc:
        cmd.append('--toc')
        cmd.append(f'--toc-depth={toc_depth}')

    # 添加参考模板
    if template_created and os.path.exists(reference_docx):
        cmd.extend(['--reference-doc', reference_docx])

    try:
        # 执行 Pandoc 命令
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            cwd=input_dir
        )

        success = result.returncode == 0

        if success:
            # 步骤 4: 后处理字体设置
            convert_module.apply_chinese_fonts_to_docx(output_docx)

        return success

    except Exception as e:
        print(f"转换错误: {e}")
        return False

    finally:
        # 清理临时文件
        if os.path.exists(temp_md):
            try:
                os.remove(temp_md)
            except Exception:
                pass

        if os.path.exists(reference_docx):
            try:
                os.remove(reference_docx)
            except Exception:
                pass
