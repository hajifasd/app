import argparse
import os
import sys
from yaml import safe_load


def main():
    parser = argparse.ArgumentParser(description="课程文件解析统计工具（命令行）")
    parser.add_argument("--input", "-i", dest="input_folder", help="要解析的课程文件目录（必填）")
    parser.add_argument("--config", "-c", dest="config", default=None, help="可选：config.yaml 路径，默认使用脚本同目录下的 config.yaml")
    args = parser.parse_args()

    # 计算脚本目录并加入 sys.path，确保能导入本包内模块（从任意工作目录运行）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)

    # 延迟导入包内模块（依赖 sys.path）
    from src.file_parser import parse_files
    from src.data_cleaner import clean_courses
    from src.stat_export import stat_and_export

    # 加载配置
    config_path = args.config if args.config else os.path.join(script_dir, "config.yaml")
    if not os.path.exists(config_path):
        print(f"警告：配置文件未找到：{config_path}（将继续，但 stat_and_export 可能使用默认输出路径）")
        config = {}
    else:
        with open(config_path, "r", encoding="utf-8") as f:
            config = safe_load(f) or {}

    input_folder = args.input_folder
    if not input_folder:
        print("错误：必须提供 --input 参数指定要解析的目录")
        parser.print_help()
        sys.exit(2)

    if not os.path.isdir(input_folder):
        print(f"错误：输入目录不存在或不是文件夹：{input_folder}")
        sys.exit(1)

    print(f"开始扫描目录：{input_folder}")
    all_courses = parse_files(input_folder)
    print(f"解析完成，共找到 {len(all_courses)} 条原始课程记录")

    cleaned_courses = clean_courses(all_courses)
    print(f"清洗完成，有效课程数：{len(cleaned_courses)}")

    # 如果 config 中有 output.path，可写回临时配置给 stat_and_export
    if "output" in config and "path" in config["output"]:
        # stat_and_export 会读取 config.yaml，本处不强制覆盖文件，但可以 setenv 或直接 call with modified behavior
        pass

    stat_and_export(cleaned_courses)
    print("任务完成！")


if __name__ == "__main__":
    main()