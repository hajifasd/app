import os
import pandas as pd
import pdfplumber
import re
from yaml import safe_load

# 加载配置文件
with open("config.yaml", "r", encoding="utf-8") as f:
    CONFIG = safe_load(f)

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
    courses = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                # 提取页面文本
                text = page.extract_text() or ""
                text = text.replace("\n", " ").strip()
                
                # 用正则提取课程信息（匹配「课程名称：XXX 讲师：XXX 课时：XXX」格式）
                pattern = r"课程名称[:：]\s*([^，,；;]+)\s*讲师[:：]\s*([^，,；;]+)\s*课时[:：]\s*([^，,；;]+)"
                matches = re.findall(pattern, text)
                
                for match in matches:
                    course = {
                        "文件来源": file_path,
                        "页码": page_num,
                        "课程名称": match[0].strip(),
                        "讲师": match[1].strip(),
                        "课时": match[2].strip(),
                        "分类": None  # 无分类时默认None
                    }
                    courses.append(course)
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