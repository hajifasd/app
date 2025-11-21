import pandas as pd
from yaml import safe_load
import re
from typing import List, Dict

# 加载配置文件
with open("config.yaml", "r", encoding="utf-8") as f:
    CONFIG = safe_load(f)

def stat_and_export(cleaned_courses):
    """统计课程数据并导出为Excel"""
    if not cleaned_courses:
        print("没有有效课程数据可统计！")
        return
    
    # 转为DataFrame方便处理
    df = pd.DataFrame(cleaned_courses)
    
    # 1. 总体统计
    # 计算总课程数（经过清洗和去重）
    total_courses = len(df)
    total_hours = int(df["课时_标准化"].sum()) if "课时_标准化" in df.columns else 0

    # 涉及讲师数：排除占位值 '未知讲师' 或空
    if "讲师" in df.columns:
        teacher_series = df["讲师"].replace({None: "", "": "", "未知讲师": ""})
        teacher_count = int(teacher_series[teacher_series != ""].nunique())
    else:
        teacher_count = 0

    category_count = int(df["分类"].dropna().nunique()) if "分类" in df.columns else 0

    df_total = pd.DataFrame({
        "统计项": ["总课程数", "总课时（小时）", "涉及讲师数", "涉及分类数"],
        "数值": [total_courses, total_hours, teacher_count, category_count]
    })
    
    # 2. 原始清洗后数据（只保留关键列）
    key_columns = ["课程名称", "讲师", "课时", "课时_标准化", "分类", "文件来源"]
    df_raw = df[key_columns].copy()
    
    # 3. 导出到Excel
    output_path = CONFIG["output"]["path"]
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_raw.to_excel(writer, sheet_name="清洗后课程数据", index=False)
        df_total.to_excel(writer, sheet_name="总体统计", index=False)
        
        # 5. 周次分布（根据周次字符串解析为区间，并按区间聚合）
        week_counts = {}
        if "周次" in df_raw.columns:
            for val in df_raw["周次"].fillna(""):
                intervals = parse_week_numbers(str(val))
                if intervals:
                    for itv in intervals:
                        week_counts[itv] = week_counts.get(itv, 0) + 1
                else:
                    week_counts.setdefault('未知周次', 0)
                    week_counts['未知周次'] += 1
        if week_counts:
            week_df = pd.DataFrame({
                "周次": list(week_counts.keys()),
                "课程数量": list(week_counts.values())
            })
            week_df.to_excel(writer, sheet_name="周次分布", index=False)
    
def parse_week_numbers(week_str):
    """把周次字符串解析为区间字符串列表（区间聚合）。

    示例输入："1-16周", "1-12周(单周)", "3周", "1,3,5周", "1-8周,10周"
    返回示例：['1-16', '1-12(单)', '3', '1', '3', '5', '1-8', '10']
    注意：对于包含单/双周信息，区间字符串会附带 '(单)' 或 '(双)'.
    """
    if not week_str or not isinstance(week_str, str):
        return []
    s = week_str.replace('周', '').replace('第', '')
    parts = re.split(r'[，,；;\s]+', s)
    intervals = []
    for p in parts:
        if not p:
            continue
        parity_suffix = ''
        if '(' in p and ')' in p:
            base, paren = p.split('(', 1)
            p = base
            if '单' in paren:
                parity_suffix = '(单)'
            elif '双' in paren:
                parity_suffix = '(双)'

        m = re.match(r'^(\d+)-(\d+)$', p)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a > b:
                a, b = b, a
            intervals.append(f"{a}-{b}" + parity_suffix)
            continue

        m2 = re.match(r'^(\d+)$', p)
        if m2:
            intervals.append(m2.group(1))
            continue

        # 其它格式忽略
    return intervals


def export_courses_to_csv(courses: List[Dict], output_path: str) -> str:
    """优化CSV导出，只保留核心字段"""
    df = pd.DataFrame(courses)
    core_columns = [
        "文件来源", "sheet/页码", "课程名称", "讲师", "课时",
        "分类", "周次", "地点", "节次", "时间段"
    ]
    # 仅保留存在的核心列
    cols = [c for c in core_columns if c in df.columns]
    df = df[cols]
    # 删除无课程名称的行
    if "课程名称" in df.columns:
        df = df.dropna(subset=["课程名称"])
    df.to_csv(output_path, index=False, encoding="utf-8-sig")
    return output_path