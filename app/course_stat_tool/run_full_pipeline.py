import traceback
import json
import os
from src.file_parser import parse_pdf
from src.stat_export import export_courses_to_csv, stat_and_export, CONFIG
from src.data_cleaner import clean_courses

PDF_PATH = r"c:\Users\321\Desktop\app\------\test\查煜明(2025-2026-1)课表 (1).pdf"

try:
    print('Parsing PDF:', PDF_PATH)
    courses = parse_pdf(PDF_PATH)
    print('Parsed courses count:', len(courses))
    if courses:
        print('Sample parsed entry:')
        print(json.dumps(courses[0], ensure_ascii=False, indent=2))

    csv_out = os.path.join(os.path.dirname(PDF_PATH), 'parsed_output.csv')
    try:
        export_courses_to_csv(courses, csv_out)
        print('Exported CSV to:', csv_out)
    except Exception as e:
        print('CSV export failed:', e)
        traceback.print_exc()

    cleaned = clean_courses(courses)
    print('Cleaned courses count:', len(cleaned))
    if cleaned:
        print('Sample cleaned entry:')
        print(json.dumps(cleaned[0], ensure_ascii=False, indent=2))

    # Use configured output path if present
    out_excel = CONFIG.get('output', {}).get('path')
    if not out_excel:
        out_excel = os.path.join(os.path.dirname(PDF_PATH), 'courses_stat.xlsx')
    try:
        stat_and_export(cleaned)
        print('Stat and export attempted, Excel path (from config):', out_excel)
    except Exception as e:
        print('Stat export failed:', e)
        traceback.print_exc()

except Exception as e:
    print('Pipeline failed:', e)
    traceback.print_exc()
