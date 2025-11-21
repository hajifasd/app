import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd
import pdfplumber
import re
from yaml import safe_load
try:
    # 包模式：相对导入
    from .src.stat_export import parse_week_numbers
except Exception:
    # 脚本模式：绝对导入（同级目录 src）
    from src.stat_export import parse_week_numbers
try:
    # 当作为包导入时使用相对导入
    from .src.data_cleaner import clean_courses
except Exception:
    # 当直接以脚本运行时使用绝对导入（工作目录应为本文件所在目录）
    from src.data_cleaner import clean_courses
from datetime import datetime


def extract_teacher_from_cell(text):
    """尝试从单元格文本中提取讲师，按多种常见格式优先匹配。

    返回讲师名或 '未知讲师'
    """
    if not text:
        return "未知讲师"
    # 优先匹配带关键字的格式
    patterns = [r'讲师[:：]\s*([^，,;/\\()]+)', r'教师[:：]\s*([^，,;/\\()]+)', r'授课人[:：]\s*([^，,;/\\()]+)']
    for p in patterns:
        m = re.search(p, text)
        if m:
            name = m.group(1).strip()
            if name:
                return name

    # 常见斜杠分隔形式，如 '课程/张三/地点' 或 '/张三/'，以及多候选形式 '张三/李四'
    # 先按斜杠或逗号分割，优先选择第一个看起来像人名的候选
    candidates = re.split(r'[\/，,;]', text)
    def is_name(s):
        s = s.strip()
        if not s:
            return False
        # 排除含有关键词或数字的候选
        if any(k in s for k in ["本", "计算机", "数学", "电信工", "大数据"]):
            return False
        if re.search(r'\d', s):
            return False
        # 限制长度（1到30）
        return 1 <= len(s) <= 30

    for cand in candidates:
        c = cand.strip()
        # 去掉括号内说明，如 '张三(主讲)'
        c = re.sub(r'[()（）].*?$', '', c).strip()
        if is_name(c):
            return c

    # 结尾处格式如 '...：张三' 或 '... / 张三'
    m = re.search(r'[:：/]\s*([^/，,;\n]+)$', text)
    if m:
        name = re.sub(r'[()（）].*?$', '', m.group(1).strip())
        if is_name(name):
            return name

    return "未知讲师"


def normalize_time_period(time_period, section=""):
    """将任意时间段或节次映射为 '上午'/'下午'/'晚上'。"""
    if not time_period and section:
        # 用节次反推
        if any(s in section for s in ["1", "2", "3", "4", "1-2", "3-4"]):
            return "上午"
        if any(s in section for s in ["5", "6", "7", "8", "5-6", "7-8"]):
            return "下午"
        if any(s in section for s in ["9", "10", "9-10"]):
            return "晚上"

    tp = str(time_period)
    if not tp:
        return "未知时段"
    tp = tp.strip()
    # 关键词判断
    if any(k in tp for k in ["上午", "早"]):
        return "上午"
    if any(k in tp for k in ["下午", "午"]):
        return "下午"
    if any(k in tp for k in ["晚", "夜"]):
        return "晚上"

    # 若包含时间点，如 08:00 或 8:00-9:40
    m = re.search(r'(\d{1,2})(?::|点)?', tp)
    if m:
        hour = int(m.group(1))
        if 6 <= hour <= 11:
            return "上午"
        if 12 <= hour <= 17:
            return "下午"
        if hour >= 18 or hour <= 5:
            return "晚上"

    return "未知时段"

# 加载配置文件
def load_config():
    # 获取项目根目录（与 data_cleaner.py 保持一致：course_stat_tool 的父目录）
    pkg_root = os.path.dirname(os.path.dirname(__file__))
    cfg_path = os.path.join(pkg_root, "config.yaml")
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, "r", encoding="utf-8") as f:
                return safe_load(f) or {}
        except Exception:
            return {}
    # 若未找到配置文件，返回默认配置（避免报错）
    return {
        "field_mapping": {
            "课程名称": ["课程名称", "课程"],
            "讲师": ["讲师", "教师"],
            "课时": ["课时", "时长"],
            "分类": ["分类", "类型"]
        }
    }

CONFIG = load_config()

# ---------------------- 核心解析逻辑（修复时间段+讲师识别，适配课表） ----------------------
def parse_single_file(file_path):
    """解析单个Excel/PDF文件，返回课程列表（修复未知时间段+讲师误识别）"""
    courses = []
    file_type = os.path.splitext(file_path)[1].lower()
    
    try:
        if file_type in ['.xlsx', '.xls']:
            # 解析Excel（原有逻辑不变，保持兼容性）
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                df.columns = [col.strip() for col in df.columns]
                
                # 匹配字段（按config.yaml配置）
                matched_cols = {}
                for target_col, possible_names in CONFIG["field_mapping"].items():
                    for col in df.columns:
                        if any(name in col for name in possible_names):
                            matched_cols[target_col] = col
                            break
                
                # 提取每行数据
                for row_idx, row in enumerate(df.itertuples(index=False), start=1):
                    # 当匹配到列时，从 tuple 中取值；保留原始行索引用于来源追踪
                    def _get(col):
                        if col in matched_cols:
                            val = getattr(row, matched_cols[col]) if hasattr(row, matched_cols[col]) else None
                            return val
                        return None

                    course_name = _get("课程名称")
                    teacher_val = _get("讲师")
                    hours_val = _get("课时")
                    category_val = _get("分类")

                    course = {
                        "文件来源": os.path.basename(file_path),
                        "sheet/页码": sheet_name,
                        "来源标识": f"{os.path.basename(file_path)}|{sheet_name}|row{row_idx}",
                        "课程名称": course_name if pd.notna(course_name) else None,
                        "讲师": teacher_val if pd.notna(teacher_val) else "未知",
                        "课时": hours_val if pd.notna(hours_val) else 0,
                        "分类": category_val if pd.notna(category_val) else "未分类"
                    }
                    if course["课程名称"]:
                        courses.append(course)
        
        elif file_type == '.pdf':
            # 解析PDF（修复合并单元格：记忆上一行时间段）
            with pdfplumber.open(file_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    tables = page.extract_tables()
                    if not tables:
                        continue
                    
                    # ---------------------- 新增：记忆上一行的有效时间段（解决合并单元格） ----------------------
                    last_valid_time_period = ""  # 存储上一行的非空时间段，用于填充合并行
                    
                    # 遍历所有表格，提取课程数据（跳过表头行，从内容行开始）
                    for table in tables:
                        for row_idx, row in enumerate(table):
                            if row_idx == 0:  # 跳过表头（时间段、节次、星期列）
                                continue
                            if not row or len(row) < 3:  # 过滤无效行（至少包含时间段、节次、1个星期列）
                                continue
                            
                            # ---------------------- 1. 提取节次（不变） ----------------------
                            section = str(row[1]).strip() if (row[1] and row[1] not in ["None", "", "/未安排"]) else ""
                            
                            # ---------------------- 2. 修复：优先用记忆的时间段，再提取/反推 ----------------------
                            # ① 先尝试提取当前行的时间段
                            current_time_period = str(row[0]).strip() if (row[0] and row[0] not in ["None", "", "/未安排"]) else ""
                            
                            # ② 若当前行时间段为空，但有记忆的有效时间段（合并单元格场景），直接复用
                            if not current_time_period and last_valid_time_period:
                                time_period = last_valid_time_period
                            # ③ 若当前行有有效时间段，更新记忆变量，并使用当前时间段
                            elif current_time_period:
                                time_period = current_time_period
                                last_valid_time_period = current_time_period  # 记忆当前有效时间段
                            # ④ 无记忆且当前行空，用节次反推（兜底逻辑）
                            else:
                                time_period = "未知时段"
                                if section:
                                    if any(s in section for s in ["1", "2", "3", "4", "1-2", "3-4"]):
                                        time_period = "上午"
                                        last_valid_time_period = "上午"  # 反推后也记忆，供后续行使用
                                    elif any(s in section for s in ["5", "6", "7", "8", "5-6", "7-8"]):
                                        time_period = "下午"
                                        last_valid_time_period = "下午"
                                    elif any(s in section for s in ["9", "10", "9-10"]):
                                        time_period = "晚上"
                                        last_valid_time_period = "晚上"
                            
                            # ---------------------- 3. 后续遍历星期列、提取课程信息（不变） ----------------------
                            weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
                            for col_idx, weekday in enumerate(weekdays, start=2):
                                if col_idx >= len(row):
                                    continue
                                course_cell = str(row[col_idx]).strip() if row[col_idx] else ""
                                if not course_cell or course_cell in ["/未安排", "None", "", " "]:
                                    continue
                                
                                # （以下课程名称、讲师、周次等提取逻辑不变，保持原代码）
                                name_match = re.search(r'^([^★☆◆◇(（]+)', course_cell)
                                course_name = name_match.group(1).strip() if name_match else "未知课程"
                                week_match = re.search(r'(\d+-\d+周|\d+周\(\w+\)|\d+周)', course_cell)
                                week = week_match.group(1) if week_match else "未知周次"
                                location_match = re.search(r'/(.*?)/', course_cell)
                                location = location_match.group(1).strip() if location_match else "未知地点"
                                class_hour = 2 if "-" in section else 1 if section.isdigit() else 0
                                
                                # 讲师识别（改进）
                                teacher = extract_teacher_from_cell(course_cell)
                                
                                # 规范时间段为 上午/下午/晚上
                                period = normalize_time_period(time_period, section)
                                # 构建来源标识，包含页、weekday、行位置（col_idx 可视为 weekday 列索引）
                                source_id = f"{os.path.basename(file_path)}|page{page_num}|{weekday}|sec{section}|col{col_idx}"
                                # 添加课程到列表（规范分类为 星期-时间段），并包含来源标识
                                courses.append({
                                    "文件来源": os.path.basename(file_path),
                                    "sheet/页码": f"第{page_num}页-{weekday}",
                                    "来源标识": source_id,
                                    "课程名称": course_name,
                                    "讲师": teacher,
                                    "课时": class_hour,
                                    "分类": f"{weekday}-{period}",
                                    "周次": week,
                                    "地点": location,
                                    "节次": section
                                })
        
        return courses, f"✅ 成功解析：{os.path.basename(file_path)}（{len(courses)}条记录）"
    
    except Exception as e:
        return [], f"❌ 解析失败：{os.path.basename(file_path)} -> {str(e)[:50]}..."

# 使用统一的清洗实现（在 src/data_cleaner.py 中实现）

def stat_courses(cleaned_courses):
    """统计课程数据（保留课表专属字段）"""
    if not cleaned_courses:
        return None
    df = pd.DataFrame(cleaned_courses)
    
    # 基础统计+课表专属统计（周次分布）
    # 规范分类列（缺失时标为 '未分类'）
    if "分类" not in df.columns:
        df["分类"] = "未分类"
    else:
        df["分类"] = df["分类"].fillna("未分类")

    # 课时字段兼容性
    if "课时_标准化" not in df.columns and "课时" in df.columns:
        df["课时_标准化"] = pd.to_numeric(df["课时"], errors='coerce').fillna(0).astype(int)

    # 优先只统计有效课时 >0 的记录，避免占位行影响统计
    df_valid = df[df["课时_标准化"] > 0].copy()

    # 统计讲师时排除占位值
    teacher_series = df_valid["讲师"].replace({None: "", "": "", "未知讲师": ""})
    teacher_count = int(teacher_series[teacher_series != ""].nunique())

    # 分类分布（基于有效记录）
    category_counts = df_valid["分类"].value_counts().to_dict()

    # 周次分布：使用 parse_week_numbers 解析后聚合为周编号计数（基于有效记录）
    week_counts = {}
    if "周次" in df_valid.columns:
        for val in df_valid["周次"].fillna(""):
            weeks = parse_week_numbers(str(val))
            if weeks:
                for w in weeks:
                    week_counts[w] = week_counts.get(w, 0) + 1
            else:
                week_counts.setdefault('未知周次', 0)
                week_counts['未知周次'] += 1

    total_stats = {
        # 下面的统计均基于有效课程（课时>0），以避免占位行扭曲结果
        "总课程数": int(len(df_valid)),
        "总课时": int(df_valid["课时_标准化"].sum()),
        "涉及讲师数": int(teacher_count),
        "涉及分类数": int(df_valid["分类"].nunique()),
        "涉及周次": int(len(week_counts)) if week_counts else 0,
        "讲师分布": df_valid["讲师"].value_counts().to_dict(),
        "分类分布": category_counts,
        "周次分布": week_counts
    }
    return total_stats, df

# ---------------------- GUI 界面（保持原有功能，优化提示文本） ----------------------
class CourseParserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("课程文件解析统计工具（课表适配版）")
        self.root.geometry("1000x700")
        self.root.resizable(True, True)
        
        # 全局变量
        self.selected_files = []
        self.all_courses = []
        self.cleaned_courses = []
        self.stat_result = None
        
        # ---------------------- 界面布局 ----------------------
        # 1. 顶部文件选择区
        self.file_frame = ttk.Frame(root, padding="10")
        self.file_frame.pack(fill=tk.X, expand=False)
        
        ttk.Label(self.file_frame, text="选择Excel/PDF文件（支持课表）：", font=("Arial", 11)).pack(side=tk.LEFT, padx=5)
        self.select_btn = ttk.Button(self.file_frame, text="浏览文件", command=self.select_files)
        self.select_btn.pack(side=tk.LEFT, padx=5)
        self.clear_btn = ttk.Button(self.file_frame, text="清空选择", command=self.clear_files)
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 文件列表显示
        self.file_listbox = tk.Listbox(root, height=6, font=("Arial", 10))
        self.file_listbox.pack(fill=tk.BOTH, expand=False, padx=10, pady=5)
        
        # 2. 解析日志区
        self.log_frame = ttk.Frame(root, padding="10")
        self.log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        ttk.Label(self.log_frame, text="解析日志：", font=("Arial", 11)).pack(anchor=tk.W)
        self.log_text = tk.Text(self.log_frame, height=10, font=("Arial", 9), state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 3. 操作按钮区
        self.btn_frame = ttk.Frame(root, padding="10")
        self.btn_frame.pack(fill=tk.X, expand=False)
        
        self.parse_btn = ttk.Button(self.btn_frame, text="开始解析", command=self.start_parse, state=tk.DISABLED)
        self.parse_btn.pack(side=tk.LEFT, padx=5)
        self.export_btn = ttk.Button(self.btn_frame, text="导出Excel", command=self.export_excel, state=tk.DISABLED)
        self.export_btn.pack(side=tk.LEFT, padx=5)
        self.debug_export_btn = ttk.Button(self.btn_frame, text="导出调试CSV", command=self.export_debug_csv, state=tk.DISABLED)
        self.debug_export_btn.pack(side=tk.LEFT, padx=5)

        # 3.1 去重选项（可配置去重键）
        self.dedupe_frame = ttk.Frame(root, padding="6")
        self.dedupe_frame.pack(fill=tk.X, expand=False, padx=10)
        ttk.Label(self.dedupe_frame, text="去重键（选中用于判定是否为重复课程）：", font=("Arial", 10)).pack(anchor=tk.W)
        # 默认选项
        self.dedupe_options = ["课程名称", "讲师", "周次", "地点", "节次"]
        self.dedupe_vars = {}
        var_frame = ttk.Frame(self.dedupe_frame)
        var_frame.pack(anchor=tk.W)
        for opt in self.dedupe_options:
            var = tk.BooleanVar(value=True if opt in ["课程名称", "讲师"] else False)
            cb = ttk.Checkbutton(var_frame, text=opt, variable=var)
            cb.pack(side=tk.LEFT, padx=6)
            self.dedupe_vars[opt] = var
        
        # 4. 统计结果区
        self.stat_frame = ttk.Frame(root, padding="10")
        self.stat_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=5)
        
        ttk.Label(self.stat_frame, text="统计结果：", font=("Arial", 11, "bold")).pack(anchor=tk.W, pady=2)
        self.stat_text = tk.Text(self.stat_frame, height=8, font=("Arial", 10), state=tk.DISABLED)
        self.stat_text.pack(fill=tk.BOTH, expand=False, padx=5, pady=2)
    
    # ---------------------- 功能函数 ----------------------
    def select_files(self):
        """选择文件（支持多文件）"""
        file_types = [
            ("支持的文件", "*.xlsx;*.xls;*.pdf"),
            ("Excel文件", "*.xlsx;*.xls"),
            ("PDF课表", "*.pdf"),
            ("所有文件", "*.*")
        ]
        files = filedialog.askopenfilenames(
            title="选择课程文件（Excel/PDF课表）",
            filetypes=file_types,
            initialdir=os.path.expanduser("~")
        )
        if files:
            new_files = [f for f in files if f not in self.selected_files]
            self.selected_files.extend(new_files)
            # 更新文件列表
            self.file_listbox.delete(0, tk.END)
            for file in self.selected_files:
                self.file_listbox.insert(tk.END, os.path.basename(file))
            self.parse_btn.config(state=tk.NORMAL)
            self.log(f"已选择 {len(self.selected_files)} 个文件")
    
    def clear_files(self):
        """清空选择"""
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.log("已清空所有选择文件")
        self.parse_btn.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.DISABLED)
        self.debug_export_btn.config(state=tk.DISABLED)
        # 清空统计结果
        self.stat_text.config(state=tk.NORMAL)
        self.stat_text.delete(1.0, tk.END)
        self.stat_text.config(state=tk.DISABLED)
    
    def log(self, msg):
        """添加日志"""
        self.log_text.config(state=tk.NORMAL)
        time_str = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{time_str}] {msg}\n")
        self.log_text.see(tk.END)  # 滚动到最新日志
        self.log_text.config(state=tk.DISABLED)
    
    def start_parse(self):
        """开始解析"""
        self.log("="*50 + " 开始解析 " + "="*50)
        self.all_courses.clear()
        
        # 逐个解析文件
        for file_path in self.selected_files:
            courses, msg = parse_single_file(file_path)
            self.log(msg)
            self.all_courses.extend(courses)
        
        # 清洗数据
        # 获取用户选择的去重键
        selected_keys = [k for k, v in self.dedupe_vars.items() if v.get()]
        self.log(f"使用去重键：{selected_keys}")
        self.cleaned_courses = clean_courses(self.all_courses, dedupe_keys=selected_keys)
        self.log(f"数据清洗完成：去重后剩余 {len(self.cleaned_courses)} 条有效记录")
        
        # 统计与显示
        if self.cleaned_courses:
            self.stat_result, self.df_cleaned = stat_courses(self.cleaned_courses)
            self.show_stat()
            self.export_btn.config(state=tk.NORMAL)
            self.debug_export_btn.config(state=tk.NORMAL)
            messagebox.showinfo("解析成功", f"共解析 {len(self.all_courses)} 条记录，去重后 {len(self.cleaned_courses)} 条有效记录！")
        else:
            self.stat_text.config(state=tk.NORMAL)
            self.stat_text.insert(tk.END, "没有解析到有效课程数据！\n提示：请确认文件是Excel或表格型课表PDF（非图片扫描件）")
            self.stat_text.config(state=tk.DISABLED)
            messagebox.showwarning("解析结果", "没有解析到有效课程数据，请检查文件格式和内容！")
        
        self.log("="*50 + " 解析结束 " + "="*50 + "\n")
    
    def show_stat(self):
        """显示统计结果（优化格式，避免过长）"""
        if not self.stat_result:
            return
        
        # 构建统计文本（显示前N项，避免内容溢出）
        stat_str = f"""
基础统计：
  - 总课程数：{self.stat_result['总课程数']} 门
  - 总课时：{self.stat_result['总课时']} 小时
  - 涉及讲师数：{self.stat_result['涉及讲师数']} 位
  - 涉及分类数：{self.stat_result['涉及分类数']} 个
  - 涉及周次：{self.stat_result['涉及周次']} 种

讲师分布（前5位）：
{chr(10).join([f"  - {t}: {c} 门" for t, c in list(self.stat_result['讲师分布'].items())[:5]])}

分类分布（星期-时间段，前8个）：
{chr(10).join([f"  - {cat}: {c} 门" for cat, c in list(self.stat_result['分类分布'].items())[:8]])}

周次分布（课表专属）：
{chr(10).join([f"  - {w}: {c} 门" for w, c in self.stat_result['周次分布'].items()]) if self.stat_result['周次分布'] else "  - 无周次数据"}
        """
        self.stat_text.config(state=tk.NORMAL)
        self.stat_text.delete(1.0, tk.END)
        self.stat_text.insert(tk.END, stat_str.strip())
        self.stat_text.config(state=tk.DISABLED)
    
    def export_excel(self):
        """导出Excel（包含课表专属字段）"""
        if not self.stat_result:
            messagebox.showwarning("导出失败", "没有可导出的统计数据！")
            return

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"course_stats_{ts}.xlsx"
        save_path = filedialog.asksaveasfilename(title="保存统计结果", defaultextension=".xlsx", initialfile=default_name,
                                                 filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*")])
        if not save_path:
            return

        try:
            with pd.ExcelWriter(save_path) as writer:
                df_cleaned = pd.DataFrame(self.cleaned_courses)
                # 列顺序优先级
                preferred = ["文件来源", "sheet/页码", "来源标识", "课程名称", "讲师", "来源原文_讲师", "课时", "课时_标准化", "分类", "来源原文_分类", "周次", "地点", "节次"]
                cols = [c for c in preferred if c in df_cleaned.columns] + [c for c in df_cleaned.columns if c not in preferred]
                df_cleaned[cols].to_excel(writer, sheet_name="详细课程数据", index=False)

                # 基础统计
                basic_df = pd.DataFrame({
                    "统计项": ["总课程数", "总课时(小时)", "涉及讲师数", "涉及分类数", "涉及周次种类"],
                    "数值": [
                        self.stat_result["总课程数"],
                        self.stat_result["总课时"],
                        self.stat_result["涉及讲师数"],
                        self.stat_result["涉及分类数"],
                        self.stat_result["涉及周次"]
                    ]
                })
                basic_df.to_excel(writer, sheet_name="基础统计", index=False)

                # 讲师分布
                teacher_df = pd.DataFrame({
                    "讲师": list(self.stat_result["讲师分布"].keys()),
                    "课程数量": list(self.stat_result["讲师分布"].values())
                })
                teacher_df.to_excel(writer, sheet_name="讲师分布", index=False)

                # 分类分布
                category_df = pd.DataFrame({
                    "分类(星期-时间段)": list(self.stat_result["分类分布"].keys()),
                    "课程数量": list(self.stat_result["分类分布"].values())
                })
                category_df.to_excel(writer, sheet_name="分类分布", index=False)

                # 周次分布
                if self.stat_result.get("周次分布"):
                    week_df = pd.DataFrame({
                        "周次": list(self.stat_result["周次分布"].keys()),
                        "课程数量": list(self.stat_result["周次分布"].values())
                    })
                    week_df.to_excel(writer, sheet_name="周次分布", index=False)

            self.log(f"导出成功：{os.path.basename(save_path)}")
            messagebox.showinfo("导出成功", f"结果已保存到：\n{save_path}\n包含详细数据与统计表。")
        except Exception as e:
            self.log(f"导出失败：{str(e)}")
            messagebox.showerror("导出失败", f"保存错误：{str(e)}")
    
    def export_debug_csv(self):
        """导出解析前的原始记录和清洗后的记录为 CSV，便于人工核查"""
        if not self.all_courses and not self.cleaned_courses:
            messagebox.showwarning("导出失败", "当前没有解析数据可导出！")
            return

        folder = filedialog.askdirectory(title="选择保存调试 CSV 的目录", initialdir=os.path.expanduser("~"))
        if not folder:
            return

        try:
            raw_path = os.path.join(folder, f"parsed_raw_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
            clean_path = os.path.join(folder, f"parsed_cleaned_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
            if self.all_courses:
                pd.DataFrame(self.all_courses).to_csv(raw_path, index=False, encoding='utf-8-sig')
            if self.cleaned_courses:
                pd.DataFrame(self.cleaned_courses).to_csv(clean_path, index=False, encoding='utf-8-sig')
            self.log(f"调试CSV已导出：{raw_path} ， {clean_path}")
            messagebox.showinfo("导出成功", f"已导出调试 CSV：\n{raw_path}\n{clean_path}")
        except Exception as e:
            self.log(f"调试CSV导出失败：{str(e)}")
            messagebox.showerror("导出失败", f"导出调试 CSV 失败：{str(e)}")

# ---------------------- 运行GUI ----------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = CourseParserGUI(root)
    root.mainloop()