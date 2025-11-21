import re
import os
from yaml import safe_load


def _load_config():
    """尝试加载 package 根目录下的 config.yaml，用于读取可选的去重键配置"""
    pkg_root = os.path.dirname(os.path.dirname(__file__))
    cfg_path = os.path.join(pkg_root, "config.yaml")
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                return safe_load(f) or {}
        except Exception:
            return {}
    return {}


def _normalize_course_name(raw_name):
    """增强版课程名称标准化，保留多语种和特殊符号"""
    if not raw_name:
        return ""
    # 仅移除明显干扰字符（保留字母、数字、汉字、常见符号）
    return re.sub(r"[^\w\u4e00-\u9fa5·•\-()（）【】:：,，;；/\/\s]", "", str(raw_name)).strip()


def clean_courses(raw_courses, dedupe_keys=None):
    """优化后的数据清洗：去重+字段补全+冗余过滤

    - 按 (课程名称, 周次, 节次) 去重
    - 过滤无效/过短课程名称
    - 补全讲师/分类字段，并只保留核心字段
    """
    cleaned = []
    seen = set()

    for course in raw_courses:
        name = course.get("课程名称") or ""
        name = _normalize_course_name(name)
        # 过滤无效课程（无名称或名称过短）
        if not name or len(name) < 2:
            continue

        week = course.get("周次", "") or ""
        section = course.get("节次", "") or ""

        # 去重：课程名称+周次+节次
        key = (name, week, section)
        if key in seen:
            continue
        seen.add(key)

        # 字段补全（处理未识别的讲师/分类）
        teacher = course.get("讲师", "")
        if teacher == "未知讲师" or not teacher:
            teacher = "未安排"

        category = course.get("分类", "") or ""
        if not category:
            category = "未知"

        # 课时字段保留或兼容处理
        class_hour = course.get("课时", course.get("课时_标准化", 0))
        try:
            class_hour = int(class_hour)
        except Exception:
            # 尝试从字符串中提取数字
            m = re.findall(r'\d+', str(class_hour))
            class_hour = int(m[0]) if m else 0

        cleaned_course = {
            "文件来源": course.get("文件来源", ""),
            "sheet/页码": course.get("sheet/页码", course.get("sheet名称", "")),
            "课程名称": name,
            "讲师": teacher,
            "课时": class_hour,
            "分类": category,
            "周次": week,
            "地点": course.get("地点", ""),
            "节次": section,
            "时间段": course.get("时间段", ""),
            "来源原文_课程名": course.get("来源原文_课程名", "")
        }
        cleaned.append(cleaned_course)

    return cleaned