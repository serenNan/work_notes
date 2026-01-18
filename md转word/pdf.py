"""
docx转pdf并按序号合并（通用版）

用法:
    python convert_and_merge.py [文件夹路径] [--keep|--delete]

    不指定路径则使用当前目录
    输出PDF以文件夹名称命名
    --keep   保留单独的PDF文件（默认）
    --delete 删除单独的PDF文件

需要安装: pip install pymupdf pywin32
"""

import os
import sys
import glob
import argparse
import win32com.client
import fitz  # PyMuPDF - 更好地保留PDF资源


def convert_docx_to_pdf(folder):
    """使用 Word 将所有 docx 转换为 pdf"""
    print("正在启动 Word...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    docx_files = sorted(glob.glob(os.path.join(folder, "*.docx")))
    pdf_files = []

    try:
        for docx_path in docx_files:
            pdf_path = docx_path.replace(".docx", ".pdf")
            print(f"  转换: {os.path.basename(docx_path)}")

            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)  # 17 = PDF格式
            doc.Close()
            pdf_files.append(pdf_path)
    finally:
        word.Quit()

    print(f"转换完成，共 {len(pdf_files)} 个文件\n")
    return pdf_files


def merge_pdfs(pdf_files, output_path):
    """使用 PyMuPDF 合并 PDF（完整保留所有资源）"""
    print("正在合并 PDF...")
    result = fitz.open()

    for pdf_path in pdf_files:
        print(f"  添加: {os.path.basename(pdf_path)}")
        with fitz.open(pdf_path) as pdf:
            result.insert_pdf(pdf)

    result.save(output_path)
    result.close()
    print(f"\n合并完成: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="DOCX 转 PDF 并合并工具")
    parser.add_argument("folder", nargs="?", default=os.getcwd(), help="目标文件夹路径")
    parser.add_argument("--delete", action="store_true", help="删除单独的PDF文件")
    args = parser.parse_args()

    folder = os.path.abspath(args.folder)

    if not os.path.isdir(folder):
        print(f"错误: 文件夹不存在 - {folder}")
        sys.exit(1)

    # 输出文件名 = 文件夹名称.pdf
    folder_name = os.path.basename(folder.rstrip(os.sep))
    output_file = os.path.join(folder, f"{folder_name}.pdf")

    print("=" * 50)
    print("DOCX 转 PDF 并合并工具")
    print("=" * 50)
    print(f"目标文件夹: {folder}")
    print(f"输出文件: {folder_name}.pdf\n")

    # 转换
    pdf_files = convert_docx_to_pdf(folder)

    if not pdf_files:
        print("未找到 docx 文件")
        return

    # 合并
    merge_pdfs(pdf_files, output_file)

    # 删除中间文件
    if args.delete:
        print("\n正在删除单独的 PDF 文件...")
        for pdf in pdf_files:
            os.remove(pdf)
        print("已删除中间文件")
    else:
        print("\n单独的 PDF 文件已保留（使用 --delete 可删除）")


if __name__ == "__main__":
    main()
