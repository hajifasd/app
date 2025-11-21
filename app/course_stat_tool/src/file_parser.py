import os
import pandas as pd
import pdfplumber
import re
from yaml import safe_load
from typing import List, Dict

def _load_config():
    pkg_root = os.path.dirname(os.path.dirname(__file__))
    cfg_path = os.path.join(pkg_root, "config.yaml")
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                return safe_load(f) or {}
        except Exception:
            return {}
    return {}

CONFIG = _load_config()

def get_file_list(folder_path):
    """扫描文件夹，获取所有Excel和PDF文件路径"""
    file_list = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(('.xlsx', '.xls', '.pdf')):
                file_path = os.path.join(root, file)
                file_list.append(file_path)
    return file_list

def parse_excel(file_path):
    """解析Excel文件，提取课程信息"""
    courses = []
    try:
        # 读取Excel所有sheet
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            # 清理列名（去除空格）
            df.columns = [col.strip() for col in df.columns]
            
            # 匹配课程相关字段（按config.yaml配置）
            matched_cols = {}
            for target_col, possible_names in CONFIG["field_mapping"].items():
                for col in df.columns:
                    if any(name in col for name in possible_names):
                        matched_cols[target_col] = col
                        break
            
            # 提取每行数据
            for _, row in df.iterrows():
                course = {
                    "文件来源": file_path,
                    "sheet名称": sheet_name
                }
                for target_col, actual_col in matched_cols.items():
                    course[target_col] = row[actual_col] if pd.notna(row[actual_col]) else None
                courses.append(course)
    except Exception as e:
        print(f"解析Excel失败：{file_path} -> {str(e)}")
    return courses

def parse_pdf(file_path):
    """解析文本型PDF文件，提取课程信息"""
    # 使用表格结构解析（适配 时间段-节次-星期 的课表表格）
    def parse_course_cell(cell_text: str, time_period: str, section: str, weekday: str) -> Dict:
        """解析单个课程单元格内容，拆分核心字段"""
        # 1. 提取课程分类（根据★/☆/◆/◇）
        category_map = {
            "★": "理论",
            "☆": "实验",
            "◆": "上机",
            "◇": "实践"
        }
        category = ""
        ct = cell_text
        for symbol, cat in category_map.items():
            if symbol in ct:
                category = cat
                ct = ct.replace(symbol, "").strip()
                break

        # 2. 提取周次（格式：1-18周/1-19周(单)）
        week_pattern = r'(\d+-\d+周(?:\(单\)|\(双\))?)'
        week_match = re.search(week_pattern, ct)
        week = week_match.group(1) if week_match else ""
        if week:
            ct = ct.replace(week, "").strip()

        # 3. 提取地点（格式：日新楼/力行楼A104等）
        location_pattern = r'([^/]+?/[^/]+?)(?=/|$)'
        location_match = re.search(location_pattern, ct)
        location = location_match.group(1).strip() if location_match else ""
        if location:
            ct = ct.replace(location, "").strip()

        # 4. 提取讲师
        teacher_pattern = r'([^/]+?)(?=/[^/]+?本|$)'
        teacher_match = re.search(teacher_pattern, ct)
        teacher = teacher_match.group(1).strip() if teacher_match else "未知讲师"
        exclude_teacher = ["23计算机", "23数学", "23大数据", "课程设计", "未安排"]
        if any(keyword in teacher for keyword in exclude_teacher) or len(teacher) > 30:
            teacher = "未知讲师"

        # 5. 提取课程名称（剩余文本，去除冗余分隔符）
        course_name = re.sub(r'[()（）/:；,，]+', '', ct).strip()
        if not course_name:
            return {}

        # 6. 计算课时（节次→课时：1-2节=1课时，3-4节=1课时，以此类推）
        class_hour = 1 if section else 0
        if "-" in section:
            class_hour = 1
        elif section in ["1", "3", "5", "7"]:
            class_hour = 1

        return {
            "课程名称": course_name,
            "讲师": teacher,
            "课时": class_hour,
            "分类": category,
            "周次": week,
            "地点": location,
            "节次": section,
            "时间段": f"{weekday}-{time_period}",
            "来源原文_课程名": cell_text
        }

    courses = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                # 优先尝试提取表格
                table = page.extract_table()
                if not table:
                    # 也尝试 extract_tables
                    tables = page.extract_tables()
                    if not tables:
                        continue
                    # 合并多个表格为一个列表的行序列
                    rows = []
                    for t in tables:
                        rows.extend(t)
                    table = rows

                # 遍历表格行（跳过表头）
                for row_idx, row in enumerate(table[1:], 2):
                    if not row or all((cell is None or (isinstance(cell, str) and cell.strip() == "")) for cell in row):
                        continue

                    time_period = row[0].strip() if row[0] else ""
                    section = row[1].strip() if len(row) > 1 and row[1] else ""
                    weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
                    for col_idx, weekday in enumerate(weekdays, 2):
                        if col_idx >= len(row):
                            continue
                        cell_content = row[col_idx].strip() if row[col_idx] else ""
                        if not cell_content or "未安排" in cell_content:
                            continue
                        course_info = parse_course_cell(cell_content, time_period, section, weekday)
                        if course_info:
                            course_info.update({
                                "文件来源": file_path,
                                "sheet/页码": f"第{page_num}页-{weekday}",
                                "来源标识": f"{file_path}|page{page_num}|{weekday}|sec{section}|col{col_idx}"
                            })
                            courses.append(course_info)
    except Exception as e:
        print(f"解析PDF失败：{file_path} -> {str(e)}")
    return courses

def parse_files(folder_path):
    """统一解析所有Excel和PDF文件"""
    file_list = get_file_list(folder_path)
    all_courses = []
    for file in file_list:
        print(f"正在解析：{file}")
        if file.endswith(('.xlsx', '.xls')):
            all_courses.extend(parse_excel(file))
        else:
            all_courses.extend(parse_pdf(file))
    return all_courses