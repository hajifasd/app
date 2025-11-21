import os
import logging
try:
    import regex as re  # 支持 \p{L} 等 Unicode 属性
    _HAS_REGEX_UNICODE = True
except Exception:
    import re  # 回退到标准库 re
    _HAS_REGEX_UNICODE = False

# logging
logging.basicConfig(filename=os.path.join(os.path.dirname(os.path.dirname(__file__)), 'parse_errors.log'),
                    level=logging.WARNING,
                    format='%(asctime)s %(levelname)s %(name)s: %(message)s')
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


def _parse_hours(duration_str):
    """解析课时字符串，支持多数字与单位，如 '36课时（12实验）','36h','2-4 小时' 等，返回一个整数课时估计"""
    if not duration_str and duration_str != 0:
        return 0
    s = str(duration_str).lower()
    # 处理带单位的常见形式
    # 优先匹配明显的小时/课时数
    m = re.findall(r'\d+', s)
    if not m:
        return 0
    nums = list(map(int, m))
    # 如果只有一个数字，直接返回
    if len(nums) == 1:
        return nums[0]
    # 若有多个数字，优先取最大（如 36课时(12实验) -> 36）
    return max(nums)


def _clean_teacher(teacher_raw: str, source_text: str, course_name: str, config: dict) -> str:
    """对讲师字段进行后处理，过滤专业/年级噪声并尝试从原文中提取真实人名。"""
    if not teacher_raw:
        teacher_raw = ""
    t = str(teacher_raw).strip()
    # 读取黑名单配置
    teacher_blacklist = [x.lower() for x in config.get('teacher_blacklist', [])]

    # 如果 t 含有明显噪声词或数字，则视为无效
    noise_patterns = ['未安排', '课', '本', '班', '计算', '大数据', '电信', '信息', '实验室']
    if any(tok in t for tok in teacher_blacklist) or any(tok in t for tok in noise_patterns) or re.search(r'\d', t) or len(t) > 30:
        t = ''

    # 如果教师与课程名高度重合（任一方向包含），也认为是误识别
    if t and course_name and (course_name.replace(' ', '') in t.replace(' ', '') or t.replace(' ', '') in course_name.replace(' ', '')):
        t = ''

    # 讲师字段应主要为中文姓名（2-6个汉字或含·）；否则尝试从原文提取
    if t and not re.match(r'^[\u4e00-\u9fff·•]{2,6}$', t):
        t = ''

    # 若仍无有效讲师，尝试从原文中提取中文姓名（2-4个汉字）
    if not t and source_text:
        def _preclean_source_text(s: str) -> str:
            """预清理来源文本：移除末尾的班级/专业/编号噪声、未安排、短码等，便于提取人名。"""
            s0 = str(s)
            # 替换常见分隔并去掉 '未安排' 等标记
            s1 = re.sub(r'未安排', '', s0)
            # 去掉类似 /23计算机本 ; /23 数学本 ; -0001 等
            s1 = re.sub(r'/?\s*23[\u4e00-\u9fa5\w\s-]{0,20}本', '', s1)
            s1 = re.sub(r'-\d{3,}', '', s1)
            s1 = re.sub(r'\(\d+-\d+节\)', '', s1)
            s1 = re.sub(r'\(.*?\)', '', s1)
            # 去掉末尾以 / 分隔的班级/专业碎片
            s1 = re.sub(r'/[\u4e00-\u9fa5\w\s-]{1,20}$', '', s1)
            # 移除多余空白和重复斜杠
            s1 = re.sub(r'[\s\u00A0]+', ' ', s1).strip()
            s1 = s1.replace('//', '/').strip()
            return s1

        src = _preclean_source_text(source_text)
        # 常见分隔符拆分后优先考虑短片段
        parts = [p.strip() for p in re.split(r'[\n/;；|]', src) if p.strip()]
        candidates = []
        for p in parts:
            # 从片段中找可能的人名
            found = re.findall(r'[\u4e00-\u9fff·•]{2,6}', p)
            for f in found:
                # 过滤包含课程名或黑名单或数字的候选
                if course_name and f in course_name:
                    continue
                if any(tok in f for tok in teacher_blacklist):
                    continue
                # 过滤明显为学院/专业/年级/编号之类的噪声片段
                candidate_noise_subs = ['学','科','本','班','院','系','专','实验','楼','室','号']
                if any(sub in f for sub in candidate_noise_subs):
                    continue
                if re.search(r'\d', f):
                    continue
                if 2 <= len(f) <= 6:
                    candidates.append(f)
        if candidates:
            def _is_valid_chinese_name(name: str, cfg: dict) -> bool:
                if not name:
                    return False
                # 必须为 2-4 个汉字（允许·）
                if not re.match(r'^[\u4e00-\u9fff·•]{2,4}$', name):
                    return False
                # 白名单优先
                white = [x for x in cfg.get('name_whitelist', []) if x]
                if name in white:
                    return True
                # 常见姓氏首字判断
                common_surnames = set(list("赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜"))
                if name[0] in common_surnames:
                    return True
                return False

            dept_blacklist = [x.lower() for x in config.get('department_blacklist', [])]
            # 先从候选中去掉明显为院系/专业的候选
            filtered = []
            for c in candidates:
                low = c.lower()
                if any(db in low for db in dept_blacklist):
                    continue
                filtered.append(c)
            if not filtered:
                filtered = candidates

            chosen = None
            # 在过滤后的候选中优先选白名单或以常见姓氏开头的短名
            for c in reversed(filtered):
                if _is_valid_chinese_name(c, config):
                    chosen = c
                    break
                if not chosen and re.match(r'^[\u4e00-\u9fff·•]{2,4}$', c):
                    chosen = c
            if not chosen:
                chosen = filtered[-1]

            if _is_valid_chinese_name(chosen, config):
                t = chosen
            else:
                t = ''
    return t or '未安排'


def clean_courses(raw_courses, dedupe_keys=None):
    """优化后的数据清洗：去重+字段补全+冗余过滤

    - 按 (课程名称, 周次, 节次) 去重
    - 过滤无效/过短课程名称
    - 补全讲师/分类字段，并只保留核心字段
    """
    config = _load_config()
    if dedupe_keys is None:
        dedupe_keys = config.get('dedupe_keys') or ["课程名称", "讲师"]

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
        # 后处理讲师字段，过滤噪声并尝试从原文中提取真实姓名
        source_text = course.get("来源原文_课程名", "") or course.get("来源原文", "")
        teacher = _clean_teacher(teacher, source_text, name, config)

        category = course.get("分类", "") or ""
        if not category:
            category = "未知"

        # 课时字段保留或兼容处理，使用 _parse_hours 解析复杂格式
        raw_hour = course.get("课时", course.get("课时_标准化", 0))
        try:
            class_hour = _parse_hours(raw_hour)
        except Exception:
            logging.exception(f"课时解析失败，输入: {raw_hour}")
            # 兜底
            m2 = re.findall(r'\d+', str(raw_hour))
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