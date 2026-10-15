# -*- coding: utf-8 -*-
"""
Korean Double Negation (이중 부정) Grammar Extractor
=====================================================
Quet cac file JSON tu kho ngu lieu 국립국어원 va tim cac cau tieng Han
chua cau truc phu dinh kep.

Cach dung:
    python extract_double_negation.py --input <thu_muc_goc> --output <file_excel>

Vi du:
    python extract_double_negation.py --input "E:/data" --output "ket_qua.xlsx"
    python extract_double_negation.py --input "E:/data/2021" --output "ket_qua.xlsx"
"""

import io
import sys

# Force UTF-8 output on Windows to avoid emoji/Unicode encoding errors
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import json
import os
import re
import argparse
import glob
from pathlib import Path
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ============================================================================
# CẤU TRÚC PHỦ ĐỊNH KÉP (이중 부정) - REGEX PATTERNS
# ============================================================================
#
# Thiết kế regex:
# - Tiếng Hàn khẩu ngữ hay chèn từ đệm, trợ từ giữa các thành phần
#   → dùng .{0,N} cho phép khoảng cách linh hoạt
# - Mỗi cấu trúc có nhiều dạng chia (않다→않아/않고/않은/않는/않을/않았/않겠...)
#   → dùng character class [았었겠]? và nhóm đuôi phổ biến
# - "form" trong JSON đã được chuẩn hoá chính tả, nên dùng form làm chính
#
# Quy ước:
#   NEG_AN   = 안
#   NEG_MOT  = 못
#   NEG_STEM = 없|모르|모를|아니|아닌
#   V_GAP    = .{1,15}  (khoảng cách giữa phần phủ định đầu và phần sau)
#   JI       = 지
#   AN_CONJ  = 않 + đuôi
#   MOT_CONJ = 못하 + đuôi  /  못 해 / 못 한
#   MAL_CONJ = 말 + đuôi (말아/말고/말자/말라...)
#   SU_RI    = 수|리  (을/ㄹ 수 없다, 을/ㄹ 리 없다)
#   GEOT     = 것|건|게|거  (것 contractions: 것이→게/건/거)
#   ANIDA    = 아니 + đuôi (아니다/아니야/아닌/아닐...)
#   EOPDA    = 없 + đuôi (없다/없어/없고/없는/없을/없었...)

# --- Đuôi (endings) phổ biến cho từng gốc ---
# Dùng để match các dạng chia của 않다, 못하다, 말다, 없다, 아니다

# 않다 endings
AN_ENDINGS = r'않[다아고은는을았었겠지며자도서면니습까요야]?'
# 못하다 endings  
MOT_ENDINGS = r'(?:못하[다아고는을였겠지며면니습까요야]?|못\s*(?:하[다아고는을였겠지며면니습까요야]?|해[다요]?|한))'
# Tap hop toan bo cac chu Hangul mang patchim ㄹ va ㄴ trong Unicode (399 chu moi loai)
RIEUL_CHARS = "".join([chr(c) for c in range(0xac00, 0xd7a4) if (c - 0xac00) % 28 == 8])
NIEUN_CHARS = "".join([chr(c) for c in range(0xac00, 0xd7a4) if (c - 0xac00) % 28 == 4])

# Dong tu/tinh tu co patchim ㄹ (vi du: 갈, 할, 볼) hoac theo sau boi 을 (vi du: 먹을)
V_RIEUL = rf'[가-힣]*(?:[{RIEUL_CHARS}]|[가-힣]+을)'

# Dong tu/tinh tu dang dinh ngu: patchim ㄴ (간, 한, 본) hoac 은/는 (먹은, 가는)
V_MOD = rf'[가-힣]*(?:[{NIEUN_CHARS}]|[가-힣]+(?:은|는))'

# Dong tu/tinh tu tieng Han co ban (1-10 ky tu)
V_WORD = r'[가-힣]{1,10}'

# Cac duoi chia pho bien cua vi tu phu dinh
AN_ENDINGS = r'않[다아고은는을았었겠지며자도서면니습까요야네죠거든래]?'
MOT_ENDINGS = r'(?:못하[다아고는을였겠지며면니습까요야네죠거든래]?|못했[다어고는을겠지며면니습까요야네죠거든래]?|못\s*(?:하[다아고는을였겠지며면니습까요야네죠거든래]?|했[다어고는을겠지며면니습까요야네죠거든래]?|해[다요]?|한))'
MAL_ENDINGS = r'(?:말[다아고자라려면지며니습까요야았겠네죠거든래]?|마[라아요]?)'
EOP_ENDINGS = r'없[다어고는을었겠지며면니습까요야네죠거든래]?'
ANI_ENDINGS = r'아니[다야고라면며는ㄴ란린ㄹ었겠지니습까요네죠거든래]?'

PATTERNS = [
    {
        'id': '1', 'name': '안 –지 않다', 'desc': '안 V-지 않다',
        'regex': re.compile(r'(?<![가-힣])안\s+' + V_WORD + r'지(?:는|도)?\s*' + AN_ENDINGS, re.UNICODE)
    },
    {
        'id': '2', 'name': '안 –지 못하다', 'desc': '안 V-지 못하다',
        'regex': re.compile(r'(?<![가-힣])안\s+' + V_WORD + r'지(?:는|도)?\s*' + MOT_ENDINGS, re.UNICODE)
    },
    {
        'id': '3', 'name': '못 –지 않다', 'desc': '못 V-지 않다',
        'regex': re.compile(r'(?<![가-힣])(?:못\s+' + V_WORD + r'지|못지)(?:는|도)?\s*' + AN_ENDINGS, re.UNICODE)
    },
    {
        'id': '4', 'name': '안 –지 말-', 'desc': '안 V-지 말다',
        'regex': re.compile(r'(?<![가-힣])안\s+' + V_WORD + r'지(?:는|도)?\s*' + MAL_ENDINGS, re.UNICODE)
    },
    {
        'id': '5', 'name': '없-지 않다', 'desc': '없-지 않다',
        'regex': re.compile(r'(?<![가-힣])[가-힣]*없지(?:는|도)?\s*' + AN_ENDINGS, re.UNICODE)
    },
    {
        'id': '6', 'name': '모르-지 않다', 'desc': '모르-지 않다',
        'regex': re.compile(r'(?<![가-힣])모르지(?:는|도)?\s*' + AN_ENDINGS, re.UNICODE)
    },
    {
        'id': '7', 'name': '안+수/리+없다', 'desc': '안 V-(으)ㄹ 수/리 없다',
        'regex': re.compile(r'(?<![가-힣])안\s+' + V_RIEUL + r'\s*(?:수|리)(?:가|는|도)?\s*' + EOP_ENDINGS, re.UNICODE)
    },
    {
        'id': '8', 'name': '못+수/리+없다', 'desc': '못 V-(으)ㄹ 수/리 없다',
        'regex': re.compile(r'(?<![가-힣])못\s+' + V_RIEUL + r'\s*(?:수|리)(?:가|는|도)?\s*' + EOP_ENDINGS, re.UNICODE)
    },
    {
        'id': '9', 'name': '못+것+아니다', 'desc': '못 V 것 아니다',
        'regex': re.compile(r'(?<![가-힣])못\s+' + V_MOD + r'\s*(?:것|건|게|거)(?:이|은|는|도|가)?\s*' + ANI_ENDINGS, re.UNICODE)
    },
    {
        'id': '10', 'name': '안+것+아니다', 'desc': '안 V 것 아니다',
        'regex': re.compile(r'(?<![가-힣])안\s+' + V_MOD + r'\s*(?:것|건|게|거)(?:이|은|는|도|가)?\s*' + ANI_ENDINGS, re.UNICODE)
    },
    {
        'id': '11', 'name': '없-+것+아니다', 'desc': '없- 것 아니다',
        'regex': re.compile(r'(?<![가-힣])[가-힣]*없는\s*(?:것|건|게|거)(?:이|은|는|도|가)?\s*' + ANI_ENDINGS, re.UNICODE)
    },
    {
        'id': '12', 'name': '모르-+것+아니다', 'desc': '모르- 것 아니다',
        'regex': re.compile(r'(?<![가-힣])모르는\s*(?:것|건|게|거)(?:이|은|는|도|가)?\s*' + ANI_ENDINGS, re.UNICODE)
    },
    {
        'id': '13', 'name': '아니-+것+아니다', 'desc': '아니- 것 아니다',
        'regex': re.compile(r'(?<![가-힣])(?:아닌|아니라는)\s*(?:것|건|게|거)(?:이|은|는|도|가)?\s*' + ANI_ENDINGS, re.UNICODE)
    },
    {
        'id': '14', 'name': '없-+수/리+없다', 'desc': '없- 수/리 없다',
        'regex': re.compile(r'(?<![가-힣])[가-힣]*없을\s*(?:수|리)(?:가|는|도)?\s*' + EOP_ENDINGS, re.UNICODE)
    },
    {
        'id': '15', 'name': '모르-+수/리+없다', 'desc': '모르- 수/리 없다',
        'regex': re.compile(r'(?<![가-힣])모를\s*(?:수|리)(?:가|는|도)?\s*' + EOP_ENDINGS, re.UNICODE)
    },
    {
        'id': '16', 'name': '아니-+수/리+없다', 'desc': '아니- 수/리 없다',
        'regex': re.compile(r'(?<![가-힣])아닐\s*(?:수|리)(?:가|는|도)?\s*' + EOP_ENDINGS, re.UNICODE)
    },
]


# ============================================================================
# HÀM XỬ LÝ
# ============================================================================

def extract_year_from_path(filepath):
    """
    Trích năm từ đường dẫn file hoặc tên file.
    - Thử tìm trong tên thư mục cha (ví dụ: .../2021/...)
    - Thử tìm trong tên file (ví dụ: SDRW2100000001 → 21 → 2021)
    - Nếu không tìm được, trả về 'N/A'
    """
    path_str = str(filepath)
    
    # Tìm năm 4 chữ số trong đường dẫn (ưu tiên thư mục)
    year_in_path = re.findall(r'[/\\](20[12]\d)[/\\]', path_str)
    if year_in_path:
        return year_in_path[-1]  # lấy thư mục gần nhất
    
    # Tìm trong tên file: SDRW21... → 2021, SDRW22... → 2022
    basename = os.path.basename(path_str)
    m = re.search(r'SDRW(\d{2})', basename)
    if m:
        short_year = m.group(1)
        return '20' + short_year
    
    # Fallback: tìm bất kỳ 4 chữ số giống năm
    m = re.search(r'(20[12]\d)', path_str)
    if m:
        return m.group(1)
    
    return 'N/A'


def extract_year_from_json(data):
    """Trích năm từ metadata bên trong file JSON."""
    try:
        year = data.get('metadata', {}).get('year', '')
        if year and re.match(r'20\d{2}', str(year)):
            return str(year)
    except:
        pass
    return None


def get_utterances(data):
    """
    Trích tất cả các utterance từ cấu trúc JSON của 국립국어원.
    
    Cấu trúc: data → document[] → utterance[]
    Mỗi utterance có:
      - "form": văn bản đã chuẩn hoá chính tả
      - "original_form": văn bản gốc (chưa chuẩn hoá)
      - "id": mã định danh
    """
    utterances = []
    
    documents = data.get('document', [])
    if not isinstance(documents, list):
        documents = [documents]
    
    for doc in documents:
        utts = doc.get('utterance', [])
        if not isinstance(utts, list):
            utts = [utts]
        for utt in utts:
            form = utt.get('form', '').strip()
            original_form = utt.get('original_form', '').strip()
            utt_id = utt.get('id', '')
            if form:
                utterances.append({
                    'form': form,
                    'original_form': original_form,
                    'id': utt_id
                })
    
    return utterances


def search_patterns_in_text(text, patterns=None):
    """
    Tìm tất cả các cấu trúc phủ định kép trong một câu.
    Trả về list các pattern match được.
    """
    if patterns is None:
        patterns = PATTERNS
    
    matches = []
    for pat in patterns:
        m = pat['regex'].search(text)
        if m:
            matches.append({
                'pattern_id': pat['id'],
                'pattern_name': pat['name'],
                'pattern_desc': pat['desc'],
                'matched_text': m.group()
            })
    return matches


def process_single_file(filepath, patterns=None):
    """
    Xử lý một file JSON: đọc, trích utterance, tìm pattern.
    Trả về list các kết quả tìm được.
    """
    results = []
    
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (json.JSONDecodeError, UnicodeDecodeError, IOError) as e:
        print(f"  [LỖI] Không đọc được {filepath}: {e}")
        return results
    
    # Lấy năm
    year = extract_year_from_json(data) or extract_year_from_path(filepath)
    
    # Lấy metadata
    file_id = data.get('id', os.path.basename(filepath))
    category = data.get('metadata', {}).get('category', '')
    
    # Lấy utterances
    utterances = get_utterances(data)
    
    for utt in utterances:
        # Tìm trên form (đã chuẩn hoá) — ưu tiên
        matches = search_patterns_in_text(utt['form'], patterns)
        
        # Nếu form không match, thử original_form
        if not matches and utt['original_form']:
            matches = search_patterns_in_text(utt['original_form'], patterns)
        
        for match in matches:
            results.append({
                'year': year,
                'file': os.path.basename(filepath),
                'file_id': file_id,
                'category': category,
                'utterance_id': utt['id'],
                'sentence_form': utt['form'],
                'sentence_original': utt['original_form'],
                'pattern_id': match['pattern_id'],
                'pattern_name': match['pattern_name'],
                'pattern_desc': match['pattern_desc'],
                'matched_text': match['matched_text']
            })
    
    return results


def find_json_files(input_path):
    """
    Tìm tất cả file .json trong thư mục (đệ quy).
    Nếu input_path là file đơn lẻ, trả về [input_path].
    """
    input_path = Path(input_path)
    
    if input_path.is_file() and input_path.suffix.lower() == '.json':
        return [str(input_path)]
    
    if input_path.is_dir():
        files = sorted(input_path.rglob('*.json'))
        return [str(f) for f in files]
    
    return []


def save_to_excel(results, output_path):
    """Lưu kết quả vào file Excel (.xlsx) với format đẹp."""
    
    if not HAS_OPENPYXL:
        # Fallback: lưu CSV
        csv_path = output_path.replace('.xlsx', '.csv')
        print(f"\n⚠ Không tìm thấy thư viện openpyxl. Lưu CSV thay: {csv_path}")
        import csv
        with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=[
                'year', 'file', 'utterance_id', 'category',
                'sentence_form', 'sentence_original',
                'pattern_name', 'matched_text'
            ])
            writer.writeheader()
            for r in results:
                writer.writerow({
                    'year': r['year'],
                    'file': r['file'],
                    'utterance_id': r['utterance_id'],
                    'category': r['category'],
                    'sentence_form': r['sentence_form'],
                    'sentence_original': r['sentence_original'],
                    'pattern_name': r['pattern_name'],
                    'matched_text': r['matched_text']
                })
        print(f"✅ Đã lưu {len(results)} kết quả vào {csv_path}")
        return csv_path
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "이중 부정 결과"
    
    # --- Header ---
    headers = [
        ('STT', 6),
        ('Năm', 8),
        ('File nguồn', 22),
        ('Mã phát ngôn', 28),
        ('Thể loại', 25),
        ('Câu tiếng Hàn (chuẩn hoá)', 55),
        ('Câu tiếng Hàn (gốc)', 55),
        ('Cấu trúc tìm thấy', 22),
        ('Phần khớp', 25),
    ]
    
    header_font = Font(name='맑은 고딕', bold=True, size=11, color='FFFFFF')
    header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    thin_border = Border(
        left=Side(style='thin', color='B4C6E7'),
        right=Side(style='thin', color='B4C6E7'),
        top=Side(style='thin', color='B4C6E7'),
        bottom=Side(style='thin', color='B4C6E7'),
    )
    
    for col_idx, (header_text, width) in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width
    
    # Freeze top row
    ws.freeze_panes = 'A2'
    
    # --- Data rows ---
    data_font = Font(name='맑은 고딕', size=10)
    data_font_kr = Font(name='맑은 고딕', size=11)
    highlight_font = Font(name='맑은 고딕', size=10, bold=True, color='C00000')
    even_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
    
    for row_idx, r in enumerate(results, 2):
        row_data = [
            row_idx - 1,
            r['year'],
            r['file'],
            r['utterance_id'],
            r['category'],
            r['sentence_form'],
            r['sentence_original'],
            r['pattern_name'],
            r['matched_text'],
        ]
        
        for col_idx, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.border = thin_border
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            
            if col_idx in (6, 7):  # Câu tiếng Hàn
                cell.font = data_font_kr
            elif col_idx == 9:  # Phần khớp
                cell.font = highlight_font
            else:
                cell.font = data_font
            
            # Tô màu xen kẽ
            if row_idx % 2 == 0:
                cell.fill = even_fill
    
    # --- Thêm sheet thống kê ---
    ws_stats = wb.create_sheet(title="Thống kê")
    ws_stats.cell(row=1, column=1, value="Cấu trúc").font = Font(bold=True)
    ws_stats.cell(row=1, column=2, value="Mô tả").font = Font(bold=True)
    ws_stats.cell(row=1, column=3, value="Số câu tìm thấy").font = Font(bold=True)
    ws_stats.column_dimensions['A'].width = 25
    ws_stats.column_dimensions['B'].width = 30
    ws_stats.column_dimensions['C'].width = 18
    
    # Đếm theo pattern
    from collections import Counter
    pattern_counts = Counter(r['pattern_name'] for r in results)
    
    for row_idx, pat in enumerate(PATTERNS, 2):
        ws_stats.cell(row=row_idx, column=1, value=pat['name'])
        ws_stats.cell(row=row_idx, column=2, value=pat['desc'])
        ws_stats.cell(row=row_idx, column=3, value=pattern_counts.get(pat['name'], 0))
    
    total_row = len(PATTERNS) + 2
    ws_stats.cell(row=total_row, column=1, value="TỔNG CỘNG").font = Font(bold=True)
    ws_stats.cell(row=total_row, column=3, value=len(results)).font = Font(bold=True)
    
    # Thêm thống kê theo năm
    ws_stats.cell(row=total_row + 2, column=1, value="Theo năm").font = Font(bold=True, size=12)
    year_counts = Counter(r['year'] for r in results)
    for i, (year, count) in enumerate(sorted(year_counts.items()), total_row + 3):
        ws_stats.cell(row=i, column=1, value=year)
        ws_stats.cell(row=i, column=3, value=count)
    
    # Save
    wb.save(output_path)
    print(f"\n✅ Đã lưu {len(results)} kết quả vào {output_path}")
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description='Trích câu tiếng Hàn chứa cấu trúc phủ định kép (이중 부정)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ví dụ sử dụng:
  python extract_double_negation.py --input "E:/data" --output "ket_qua.xlsx"
  python extract_double_negation.py --input "E:/data/2021" --output "2021_results.xlsx"
  python extract_double_negation.py --input "E:/data" --output "ket_qua.xlsx" --patterns 1 2 5
        """
    )
    parser.add_argument(
        '--input', '-i',
        required=True,
        help='Đường dẫn thư mục chứa file JSON (quét đệ quy) hoặc một file JSON đơn lẻ'
    )
    parser.add_argument(
        '--output', '-o',
        default=None,
        help='Đường dẫn file kết quả (.xlsx). Mặc định: ket_qua_YYYYMMDD_HHMMSS.xlsx'
    )
    parser.add_argument(
        '--patterns', '-p',
        nargs='*',
        default=None,
        help='Chỉ tìm một số cấu trúc (theo số thứ tự 1-16). Mặc định: tìm tất cả.'
    )
    
    args = parser.parse_args()
    
    # Output path
    if args.output is None:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        args.output = f'ket_qua_{timestamp}.xlsx'
    
    # Lọc patterns nếu cần
    if args.patterns:
        selected = [p for p in PATTERNS if p['id'] in args.patterns]
        if not selected:
            print("❌ Không tìm thấy pattern nào khớp với số đã chọn.")
            sys.exit(1)
        print(f"🔍 Chỉ tìm {len(selected)} cấu trúc: {', '.join(p['name'] for p in selected)}")
    else:
        selected = PATTERNS
        print(f"🔍 Tìm tất cả {len(selected)} cấu trúc phủ định kép")
    
    # Tìm file JSON
    print(f"\n📂 Đang quét thư mục: {args.input}")
    json_files = find_json_files(args.input)
    
    if not json_files:
        print("❌ Không tìm thấy file JSON nào.")
        sys.exit(1)
    
    print(f"📄 Tìm thấy {len(json_files)} file JSON")
    
    # Xử lý từng file
    all_results = []
    errors = 0
    
    for idx, filepath in enumerate(json_files, 1):
        if idx % 100 == 0 or idx == len(json_files):
            print(f"  ⏳ Đang xử lý: {idx}/{len(json_files)} files...", end='\r')
        
        results = process_single_file(filepath, selected)
        all_results.extend(results)
    
    print(f"\n\n📊 KẾT QUẢ:")
    print(f"   Tổng file đã quét: {len(json_files)}")
    print(f"   Tổng câu tìm được: {len(all_results)}")
    
    if all_results:
        # Thống kê nhanh
        from collections import Counter
        pattern_counts = Counter(r['pattern_name'] for r in all_results)
        print(f"\n   Theo cấu trúc:")
        for pat_name, count in pattern_counts.most_common():
            print(f"     • {pat_name}: {count} câu")
        
        # Lưu file
        save_to_excel(all_results, args.output)
    else:
        print("\n   ⚠ Không tìm thấy câu nào khớp.")


if __name__ == '__main__':
    main()
