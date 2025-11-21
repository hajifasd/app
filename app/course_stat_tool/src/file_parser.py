import os
import pandas as pd
import pdfplumber
import re
from yaml import safe_load
from typing import List, Dict
import logging
from typing import List

# 临时调试开关，调试完成后会移除或置为 False
DEBUG_PARSER = False

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

# logging
logging.basicConfig(filename=os.path.join(os.path.dirname(os.path.dirname(__file__)), 'parse_errors.log'),
                    level=logging.WARNING,
                    format='%(asctime)s %(levelname)s %(name)s: %(message)s')
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
            # 先不指定表头，方便容错查找真实表头行
            df0 = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str)
            header_row = 0
            # 在前几行寻找包含至少 N 个目标字段的行作为表头，阈值可在 config.yaml 中配置
            header_threshold = int(CONFIG.get("header_match_threshold", 2))
            header_blacklist = [x.lower() for x in CONFIG.get("header_blacklist", [])]
            max_check = min(6, len(df0))
            for i in range(max_check):
                row_text = ' '.join([str(x) for x in df0.iloc[i].dropna()]).lower()
                # 如果行文本包含在黑名单中，跳过
                if any(blk in row_text for blk in header_blacklist):
                    continue
                matched = 0
                for _, names in CONFIG.get("field_mapping", {}).items():
                    if any(n.lower() in row_text for n in names):
                        matched += 1
                if matched >= header_threshold:
                    header_row = i
                    break
            # 重新读取带表头的表格
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, dtype=str)
            except Exception:
                df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
            # 清理列名（去除空格）
            df.columns = [str(col).strip() for col in df.columns]

            # 匹配课程相关字段（按config.yaml配置）
            matched_cols = {}
            for target_col, possible_names in CONFIG.get("field_mapping", {}).items():
                for col in df.columns:
                    col_lower = str(col).lower()
                    if any(name.lower() in col_lower for name in possible_names):
                        matched_cols[target_col] = col
                        break

            # 提取每行数据
            for _, row in df.iterrows():
                course = {
                    "文件来源": file_path,
                    "sheet名称": sheet_name
                }
                for target_col, actual_col in matched_cols.items():
                    try:
                        val = row[actual_col]
                    except Exception:
                        val = None
                    course[target_col] = val if pd.notna(val) else None
                courses.append(course)
    except Exception as e:
        logging.exception(f"解析Excel失败：{file_path} -> {str(e)}")
        print(f"解析Excel失败：{file_path} -> {str(e)}")
    return courses

def parse_pdf(file_path):
    """解析文本型PDF文件，提取课程信息"""
    # 使用表格结构解析（适配 时间段-节次-星期 的课表表格）
    header_tokens = set(["星期一","星期二","星期三","星期四","星期五","星期六","星期日",
                         "节次","时间段","时间","序号","周次","节","上课时间"])

    def is_header_row(row: List[str]) -> bool:
        """检测一行是否为表头/标题行（例如包含星期、节次、时间段等关键词）。"""
        if not row:
            return False
        non_empty = [c for c in row if c and isinstance(c, str) and c.strip()]
        if not non_empty:
            return False
        # 如果大多数非空单元格是已知表头词，则认为是表头
        cnt_header = 0
        for c in non_empty:
            txt = c.strip()
            if txt in header_tokens:
                cnt_header += 1
            # 包含“星期”关键字也判断为表头
            elif txt.startswith("星期"):
                cnt_header += 1
        return cnt_header >= max(1, len(non_empty) // 2)
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
        ct = cell_text or ""
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
        # 尝试更稳健地匹配地点，优先匹配包含“楼”“室”“教”“号”之类的短语
        location = ""
        loc_match = re.search(r'([\u4e00-\u9fff\w\-]{2,20}(楼|室|教|号)[\w\-\d]*)', ct)
        if loc_match:
            location = loc_match.group(0).strip()
            ct = ct.replace(location, "").strip()

        # 4. 提取讲师
        # 讲师匹配：优先按分隔符拆分，常见分隔符为 /、\n、;、；、| 等
        teacher = "未知讲师"
        parts = [p.strip() for p in re.split(r'[\n/;；|]', ct) if p.strip()]
        # 可能的课程名优先取 parts[0]
        probable_course_name = parts[0] if parts else ""
        # 移除明显含有地点/周次/数字的片段，剩余短片段可能为讲师或课程名
        candidate_teachers = []
        for p in parts:
            if re.search(r'\d+周', p) or re.search(r'第?\d+节', p):
                continue
            if any(tok in p for tok in ['楼','室','教','号']):
                continue
            # 避免把纯时间/星期误判为讲师
            if p.startswith('星期') or re.match(r'^\d{1,2}:', p):
                continue
            # 排除明显的班级/专业/编号等噪声
            if '未安排' in p or re.search(r'23\w{0,6}本', p) or re.search(r'-?\d{3,}', p):
                continue
            if 1 <= len(p) <= 30:
                candidate_teachers.append(p)
        if candidate_teachers:
            # 优先选最后一个符合中文人名格式的候选（2-6个汉字）
            teacher_candidate = None
            for p in reversed(candidate_teachers):
                if re.match(r'^[\u4e00-\u9fff·•]{2,6}$', p):
                    teacher_candidate = p
                    break
            if not teacher_candidate:
                for p in reversed(candidate_teachers):
                    if p != probable_course_name and not re.search(r'\d', p):
                        teacher_candidate = p
                        break
            if teacher_candidate:
                teacher = teacher_candidate
                try:
                    ct = ct.replace(teacher, '').strip()
                except Exception:
                    pass

        # 5. 提取课程名称（剩余文本，去除冗余分隔符）
        # 剩余文本作为课程名的候选，去掉多余符号
        # 如果 parts 存在且首段看起来像课程名，优先使用
        if probable_course_name and not probable_course_name.startswith('(') and len(probable_course_name) > 1:
            course_name = probable_course_name
        else:
            course_name = re.sub(r'[()（）/:；,，\n]+', ' ', ct).strip()
        # 去掉末尾的编号和专业标签
        course_name = re.sub(r'[-–—]\d+|\b23\w{0,6}\b', '', course_name).strip()
        if DEBUG_PARSER:
            try:
                print(f"[DEBUG] parse_course_cell ct='{ct[:120]}' parts={parts[:3]} candidates={candidate_teachers[:3]} course_name='{course_name}'")
            except Exception:
                pass
        # 如果课程名不合理（如为星期/节次等），则认为不是课程单元
        if not course_name or course_name in header_tokens or course_name.startswith('星期'):
            return {}

        # 6. 计算课时（节次→课时：1-2节=1课时，3-4节=1课时，以此类推）
        # 简单根据节次估算课时：通常每两节为1课时；如果未给出节次则为0
        class_hour = 0
        try:
            if section:
                # 支持形式 1-2 或 3-4
                if '-' in section:
                    parts_sec = [int(s) for s in re.findall(r'\d+', section)]
                    if len(parts_sec) >= 2:
                        class_hour = max(1, (abs(parts_sec[1] - parts_sec[0]) + 1) // 2)
                else:
                    nums = re.findall(r'\d+', section)
                    if nums:
                        class_hour = 1
        except Exception:
            class_hour = 1 if section else 0

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
                if DEBUG_PARSER:
                    try:
                        print(f"[DEBUG] page {page_num} extract_table -> {('None' if table is None else len(table))} rows")
                    except Exception:
                        pass
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

                if DEBUG_PARSER:
                    # 打印前几行以便调试表头判断
                    try:
                        for i, r in enumerate(table[:6]):
                            print(f"[DEBUG] page {page_num} row[{i}]: {r}")
                    except Exception:
                        pass

                # 遍历表格行（尝试跳过表头或标题行）
                for row_idx, row in enumerate(table[1:], 2):
                    if not row or all((cell is None or (isinstance(cell, str) and cell.strip() == "")) for cell in row):
                        continue

                    # 如果这一整行看起来像表头，则跳过
                    try:
                        if is_header_row(row):
                            logging.debug(f"跳过表头行: page {page_num} row {row_idx} -> {row}")
                            continue
                    except Exception:
                        pass

                    time_period = row[0].strip() if row[0] else ""
                    section = row[1].strip() if len(row) > 1 and row[1] else ""
                    weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
                    for col_idx, weekday in enumerate(weekdays, 2):
                        if col_idx >= len(row):
                            continue
                        cell_content = row[col_idx].strip() if row[col_idx] else ""
                        # 仅在单元格完全为“未安排”时跳过，避免忽略包含课程与“未安排”标记的复合单元格
                        if not cell_content or cell_content.strip() in ("未安排", "未 安排"):
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