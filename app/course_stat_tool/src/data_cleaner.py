import os
try:
    import regex as re  # 支持 \p{L} 等 Unicode 属性
    _HAS_REGEX_UNICODE = True
except Exception:
    import re  # 回退到标准库 re
    _HAS_REGEX_UNICODE = False
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
    # 使用 regex 库时允许 Unicode 字母/数字（支持日文、韩文等）
    s = str(raw_name)
    if _HAS_REGEX_UNICODE:
        return re.sub(r"[^\p{L}\p{N}·•\-()（）【】:：,，;；/\\/\s]", "", s).strip()
    # fallback: 保持以前的宽松规则（包含常见汉字和 \w）
    return re.sub(r"[^\w\u4e00-\u9fa5·•\-()（）【】:：,，;；/\\/\s]", "", s).strip()


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

        # 课时字段保留或兼容处理，增强对 '36h','2-4 小时' 等格式的识别
        raw_hour = course.get("课时", course.get("课时_标准化", 0))
        class_hour = 0
        if isinstance(raw_hour, (int, float)):
            try:
                class_hour = int(raw_hour)
            except Exception:
                class_hour = 0
        else:
            s = str(raw_hour)
            # 优先匹配类似 '36 课时' / '36小时' / '36h' 等
            m = re.search(r'(\d+)\s*(?:课时|小时|h|H)?', s)
            if m:
                class_hour = int(m.group(1))
            else:
                # 兜底：提取任意数字
                m2 = re.findall(r'\d+', s)
                class_hour = int(m2[0]) if m2 else 0

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