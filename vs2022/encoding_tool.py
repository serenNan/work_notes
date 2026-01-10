"""
通用编码转换工具
"""
import sys
import os
import argparse

sys.stdout.reconfigure(encoding='utf-8')

# 支持的编码
ENCODINGS = ['utf-8', 'gbk', 'gb2312', 'gb18030', 'utf-16', 'ascii', 'latin-1']
# 支持的文件扩展名
EXTENSIONS = ['.cpp', '.h', '.hpp', '.c', '.cc', '.cxx', '.txt', '.py', '.java', '.cs']

def convert_file(file_path, from_enc, to_enc):
    """转换单个文件的编码"""
    try:
        with open(file_path, 'r', encoding=from_enc) as f:
            content = f.read()
        with open(file_path, 'w', encoding=to_enc) as f:
            f.write(content)
        return True
    except Exception as e:
        print(f"  错误: {e}")
        return False

def copy_and_convert(src_path, dest_path, from_enc, to_enc):
    """复制文件并转换编码"""
    try:
        with open(src_path, 'r', encoding=from_enc) as f:
            content = f.read()
        with open(dest_path, 'w', encoding=to_enc) as f:
            f.write(content)
        return True
    except Exception as e:
        print(f"  错误: {e}")
        return False

def deploy(src_dir, dest_dir, from_enc='utf-8', to_enc='gbk', recursive=False):
    """部署: 复制目录并转换编码"""
    print(f"源目录: {src_dir}")
    print(f"目标目录: {dest_dir}")
    print(f"编码转换: {from_enc} -> {to_enc}")
    print()

    if not os.path.exists(src_dir):
        print("错误: 源目录不存在!")
        return 0

    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)

    count = 0

    if recursive:
        for root, dirs, files in os.walk(src_dir):
            rel_path = os.path.relpath(root, src_dir)
            dest_root = os.path.join(dest_dir, rel_path) if rel_path != '.' else dest_dir

            if not os.path.exists(dest_root):
                os.makedirs(dest_root)

            for filename in files:
                ext = os.path.splitext(filename)[1].lower()
                if ext in EXTENSIONS:
                    src_path = os.path.join(root, filename)
                    dest_path = os.path.join(dest_root, filename)
                    rel_file = os.path.relpath(src_path, src_dir)
                    print(f"  {rel_file}")
                    if copy_and_convert(src_path, dest_path, from_enc, to_enc):
                        count += 1
    else:
        for filename in os.listdir(src_dir):
            src_path = os.path.join(src_dir, filename)
            if os.path.isfile(src_path):
                ext = os.path.splitext(filename)[1].lower()
                if ext in EXTENSIONS:
                    dest_path = os.path.join(dest_dir, filename)
                    print(f"  {filename}")
                    if copy_and_convert(src_path, dest_path, from_enc, to_enc):
                        count += 1

    print()
    print(f"完成! 共处理 {count} 个文件")
    return count

def convert(path, from_enc, to_enc, recursive=False):
    """转换文件或目录的编码 (原地转换)"""
    if os.path.isfile(path):
        print(f"文件: {path}")
        print(f"转换: {from_enc} -> {to_enc}")
        if convert_file(path, from_enc, to_enc):
            print("完成!")
        return

    if os.path.isdir(path):
        print(f"目录: {path}")
        print(f"转换: {from_enc} -> {to_enc}")
        print()

        count = 0
        if recursive:
            for root, dirs, files in os.walk(path):
                for filename in files:
                    ext = os.path.splitext(filename)[1].lower()
                    if ext in EXTENSIONS:
                        file_path = os.path.join(root, filename)
                        rel_file = os.path.relpath(file_path, path)
                        print(f"  {rel_file}")
                        if convert_file(file_path, from_enc, to_enc):
                            count += 1
        else:
            for filename in os.listdir(path):
                file_path = os.path.join(path, filename)
                if os.path.isfile(file_path):
                    ext = os.path.splitext(filename)[1].lower()
                    if ext in EXTENSIONS:
                        print(f"  {filename}")
                        if convert_file(file_path, from_enc, to_enc):
                            count += 1

        print()
        print(f"完成! 共处理 {count} 个文件")
        return

    print(f"错误: 路径不存在: {path}")

def main():
    parser = argparse.ArgumentParser(
        description='通用编码转换工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例:
  %(prog)s convert file.cpp -f utf-8 -t gbk           # 单文件转换
  %(prog)s convert src/ -f utf-8 -t gbk               # 目录转换
  %(prog)s convert src/ -f utf-8 -t gbk -r            # 递归转换
  %(prog)s deploy src/ dest/ -f utf-8 -t gbk          # 复制并转换
  %(prog)s deploy src/ dest/ -f utf-8 -t gbk -r       # 递归复制并转换

支持的编码: utf-8, gbk, gb2312, gb18030, utf-16, ascii, latin-1
'''
    )

    subparsers = parser.add_subparsers(dest='command', help='命令')

    # convert 命令
    p_convert = subparsers.add_parser('convert', help='原地转换编码')
    p_convert.add_argument('path', help='文件或目录路径')
    p_convert.add_argument('-f', '--from', dest='from_enc', default='utf-8', help='源编码 (默认: utf-8)')
    p_convert.add_argument('-t', '--to', dest='to_enc', default='gbk', help='目标编码 (默认: gbk)')
    p_convert.add_argument('-r', '--recursive', action='store_true', help='递归处理子目录')

    # deploy 命令
    p_deploy = subparsers.add_parser('deploy', help='复制并转换编码')
    p_deploy.add_argument('src', help='源目录')
    p_deploy.add_argument('dest', help='目标目录')
    p_deploy.add_argument('-f', '--from', dest='from_enc', default='utf-8', help='源编码 (默认: utf-8)')
    p_deploy.add_argument('-t', '--to', dest='to_enc', default='gbk', help='目标编码 (默认: gbk)')
    p_deploy.add_argument('-r', '--recursive', action='store_true', help='递归处理子目录')

    args = parser.parse_args()

    if args.command == 'convert':
        convert(args.path, args.from_enc, args.to_enc, args.recursive)
    elif args.command == 'deploy':
        deploy(args.src, args.dest, args.from_enc, args.to_enc, args.recursive)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
