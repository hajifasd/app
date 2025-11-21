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


def clean_courses(courses, dedupe_keys=None):
    """清洗课程数据：去重、标准化字段

    特性：
    - 支持从 `config.yaml` 中读取 `dedupe_keys`（示例：['课程名称','讲师','周次']），若无配置则使用默认 ['课程名称','讲师']。
    - 保证返回字段包含 `课程名称`、`讲师`、`课时_标准化`、`分类`。
    - 去重时对缺失字段进行容错（使用占位字符串）。
    """
    config = _load_config()
    if dedupe_keys is None:
        dedupe_keys = config.get("dedupe_keys") or ["课程名称", "讲师"]

    cleaned_courses = []
    seen = set()

    for course in courses:
        # 跳过无课程名称的记录
        raw_name = course.get("课程名称")
        if not raw_name:
            continue

        # 标准化课程名称
        standard_name = re.sub(r'[^\\u4e00-\\u9fa5a-zA-Z0-9]', '', str(raw_name)).strip()

        # 标准化讲师并保留原文
        raw_teacher = course.get("讲师", None)
        course["来源原文_讲师"] = raw_teacher
        def _is_probable_name(s):
            if not s or not isinstance(s, str):
                return False
            s = s.strip()
            # 去掉括号内备注
            s = re.sub(r'[()（）].*?$', '', s).strip()
            # 含数字、含明显非人名关键词则否
            if re.search(r'\d', s):
                return False
            if any(k in s for k in ["计算机", "教室", "楼", "班", "级", "学院", "专业", "实验", "本"]):
                return False
            # 简单中文姓名匹配：2-4个汉字或含中点
            if re.match(r'^[\u4e00-\u9fa5·•]{2,8}$', s):
                return True
            # 英文名（首字母大写）或其他短字符串也可接受
            if re.match(r'^[A-Za-z\-\s]{2,30}$', s):
                return True
            return False

        if raw_teacher and _is_probable_name(str(raw_teacher)):
            teacher = str(raw_teacher).strip()
        else:
            teacher = "未知讲师"

        # 确保课时字段
        duration = course.get("课时", course.get("课时_标准化", 0))
        duration_str = str(duration).strip()
        if duration_str.isdigit():
            hours = int(duration_str)
        else:
            m = re.findall(r'\\d+', duration_str)
            hours = int(m[0]) if m else 0
        course["课时_标准化"] = hours

        # 确保分类字段并保留原文，尝试标准化为 '星期X-上午/下午/晚上'
        raw_category = course.get("分类", None)
        course["来源原文_分类"] = raw_category
        def _normalize_category(cat, section=""):
            if not cat or not isinstance(cat, str):
                # 尝试用节次反推（如 '1-2' -> 上午）
                if section:
                    if any(s in section for s in ["1", "2", "3", "4", "1-2", "3-4"]):
                        return "未知星期-上午"
                return "未分类"
            s = cat.strip()
            # 提取星期
            wk = None
            m = re.search(r'(星期[一二三四五六日]|周[一二三四五六日])', s)
            if m:
                wk = m.group(1).replace('周', '星期')
            # 时间段关键词
            if any(k in s for k in ['上午', '上']):
                period = '上午'
            elif any(k in s for k in ['下午', '下']):
                period = '下午'
            elif any(k in s for k in ['晚', '夜']):
                period = '晚上'
            else:
                # 数字时间或节次
                m2 = re.search(r'(第?\d+[-~–—]\d+节|\d+-\d+节|\d+节|\d+小时|\d{1,2}:\d{2})', s)
                if m2:
                    # 简单用节次判断上午/下午
                    if any(n in s for n in ['1-2', '1-4', '1-3']):
                        period = '上午'
                    else:
                        period = '下午'
                else:
                    period = None

            if wk and period:
                return f"{wk}-{period}"
            if wk and not period:
                return f"{wk}-未知时段"
            # 无星期信息，返回原文标记
            return '未分类'

        normalized_cat = _normalize_category(raw_category, course.get('节次', ''))
        course["分类"] = normalized_cat

        # 生成去重 key（按配置的字段顺序），缺失字段用占位
        key_parts = []
        for k in dedupe_keys:
            if k == "课程名称":
                key_parts.append(standard_name)
            elif k == "讲师":
                key_parts.append(teacher)
            else:
                # 其它字段直接使用原始字符串或占位
                val = course.get(k)
                key_parts.append(str(val).strip() if val is not None else "__MISSING__")

        unique_key = tuple(key_parts)
        if unique_key in seen:
            continue
        seen.add(unique_key)

        # 写回标准化字段
        course["课程名称"] = standard_name
        course["讲师"] = teacher

        cleaned_courses.append(course)

    return cleaned_courses