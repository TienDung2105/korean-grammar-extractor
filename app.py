# -*- coding: utf-8 -*-
"""
Korean Double Negation Web Extractor
=====================================
Web app cho phep upload file JSON va tim cau tieng Han chua cau truc phu dinh kep.
"""

import os
import re
import tempfile
from flask import Flask, Request, jsonify


class DiskUploadRequest(Request):
    def _get_file_stream(self, total_content_length, content_type, filename=None, content_length=None):
        # Even thousands of small files must not stay in multipart RAM buffers.
        return tempfile.TemporaryFile(mode='w+b')

app = Flask(__name__)
app.request_class = DiskUploadRequest
# Không giới hạn dung lượng upload
app.config['MAX_CONTENT_LENGTH'] = None

@app.errorhandler(413)
def request_entity_too_large(error):
    return jsonify({'error': 'Yêu cầu tải lên vượt giới hạn máy chủ. Hãy nén JSON thành ZIP hoặc chia dữ liệu thành các ZIP nhỏ hơn.'}), 413

@app.errorhandler(500)
def internal_error(error):
    return jsonify({'error': 'Lỗi server: ' + str(error)}), 500

@app.after_request
def add_no_cache_headers(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response

# ============================================================================
# REGEX PATTERNS - 16 cau truc phu dinh kep (DA TOI UU DO CHINH XAC CAO)
# ============================================================================

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

ALL_PATTERNS = [
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
# PROCESSING FUNCTIONS
# ============================================================================

def build_custom_pattern(raw_text, custom_id):
    """Tu dong tao regex thong minh tu chuoi ngu phap nguoi dung tu nhap."""
    raw = raw_text.strip()
    if not raw:
        return None

    name = raw.strip('-~ ')
    p = raw.strip('~ ')
    if p.startswith('-'):
        p = p[1:]

    # Xu ly cac ky hieu ngu phap pho bien tieng Han
    p = p.replace('(으)ㄹ', f'(?:[{RIEUL_CHARS}]|[가-힣]+을)')
    p = p.replace('(으)ㄴ', f'(?:[{NIEUN_CHARS}]|[가-힣]+은)')
    p = p.replace('(으)', '(?:으)?')

    parts = p.split()
    escaped_parts = []
    for part in parts:
        if RIEUL_CHARS in part or NIEUN_CHARS in part or '(?:' in part:
            escaped_parts.append(part)
        else:
            escaped_parts.append(re.escape(part))
    regex_str = r'\s*'.join(escaped_parts)

    try:
        rx = re.compile(regex_str, re.UNICODE)
    except Exception:
        rx = re.compile(re.escape(name), re.UNICODE)

    return {
        'id': f'custom_{custom_id}',
        'name': name,
        'desc': f'Tự nhập: {name}',
        'regex': rx
    }


from low_memory import register

register(app, ALL_PATTERNS, build_custom_pattern)


if __name__ == '__main__':
    os.makedirs('templates', exist_ok=True)
    print("=" * 60)
    print("  Korean Double Negation Extractor")
    print("  Mo trinh duyet va vao: http://localhost:5000")
    print("=" * 60)
    app.run(debug=False, host='0.0.0.0', port=5000)

