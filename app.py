import os
import io
import re
import json
import time
import base64
import unicodedata
import streamlit as st
import streamlit.components.v1 as components
from openai import OpenAI

try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# --- CẤU HÌNH ĐÚNG THÔNG TIN TỪ VIETCOD ---
API_KEY = "sk-cef0946a13d9b842-7zfdhi-44349205"
BASE_URL = "https://router.vietcod.com/v1"

PERSISTENT_STORE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "platform_persistent_data.json")

try:
    client = OpenAI(
        api_key=API_KEY,
        base_url=BASE_URL
    )
except Exception as e:
    client = None
    print(f"Lỗi khởi tạo client: {e}")


def load_persistent_store() -> dict:
    """Nạp dữ liệu hồ sơ học sinh, báo cáo tiến bộ, bài giao và điểm số từ file lưu trữ bền vững."""
    default_data = {
        "student_profiles": [],
        "reports": [],
        "assigned_tasks": [],
        "student_progress": {},
    }
    if not os.path.exists(PERSISTENT_STORE_PATH):
        return default_data
    try:
        with open(PERSISTENT_STORE_PATH, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        if not isinstance(loaded, dict):
            return default_data
        return {
            "student_profiles": loaded.get("student_profiles", []) if isinstance(loaded.get("student_profiles"), list) else [],
            "reports": loaded.get("reports", []) if isinstance(loaded.get("reports"), list) else [],
            "assigned_tasks": loaded.get("assigned_tasks", []) if isinstance(loaded.get("assigned_tasks"), list) else [],
            "student_progress": loaded.get("student_progress", {}) if isinstance(loaded.get("student_progress"), dict) else {},
        }
    except Exception:
        return default_data


def save_persistent_store():
    """Lưu toàn bộ hồ sơ học sinh, báo cáo tiến bộ, bài giao và tiến độ học sinh xuống file JSON."""
    payload = {
        "student_profiles": st.session_state.get("student_profiles", []),
        "reports": st.session_state.get("reports", []),
        "assigned_tasks": st.session_state.get("assigned_tasks", []),
        "student_progress": st.session_state.get("student_progress", {}),
    }
    try:
        with open(PERSISTENT_STORE_PATH, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Lỗi lưu dữ liệu bền vững: {e}")


def get_illustration_svg(kind: str = "hero") -> str:
    """Trả về mã SVG minh họa trực quan, sinh động cho từng khu vực và dạng bài tập."""
    illustrations = {
        "hero": """
        <svg viewBox="0 0 320 120" width="100%" height="115" xmlns="http://www.w3.org/2000/svg">
            <defs>
                <linearGradient id="gSky" x1="0%" y1="0%" x2="100%" y2="100%">
                    <stop offset="0%" stop-color="#60A5FA"/>
                    <stop offset="55%" stop-color="#34D399"/>
                    <stop offset="100%" stop-color="#FBBF24"/>
                </linearGradient>
            </defs>
            <rect x="8" y="10" width="304" height="100" rx="22" fill="url(#gSky)" opacity="0.25"/>
            <circle cx="62" cy="60" r="34" fill="#FDE68A"/>
            <circle cx="62" cy="60" r="22" fill="#F59E0B"/>
            <rect x="115" y="28" width="82" height="64" rx="12" fill="#FFFFFF" stroke="#2563EB" stroke-width="3"/>
            <line x1="128" y1="46" x2="184" y2="46" stroke="#3B82F6" stroke-width="4" stroke-linecap="round"/>
            <line x1="128" y1="60" x2="174" y2="60" stroke="#10B981" stroke-width="4" stroke-linecap="round"/>
            <line x1="128" y1="74" x2="160" y2="74" stroke="#F59E0B" stroke-width="4" stroke-linecap="round"/>
            <circle cx="242" cy="54" r="24" fill="#A7F3D0" stroke="#059669" stroke-width="3"/>
            <path d="M231 54 L239 62 L254 46" fill="none" stroke="#047857" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
            <circle cx="282" cy="34" r="8" fill="#F472B6"/>
            <circle cx="278" cy="86" r="10" fill="#818CF8"/>
        </svg>
        """,
        "true_false": """
        <svg viewBox="0 0 180 70" width="170" height="64" xmlns="http://www.w3.org/2000/svg">
            <rect x="4" y="6" width="172" height="58" rx="14" fill="#F0FDF4" stroke="#16A34A" stroke-width="2.5"/>
            <circle cx="36" cy="35" r="14" fill="#DCFCE7" stroke="#16A34A" stroke-width="2.5"/>
            <path d="M30 35 L35 40 L43 30" fill="none" stroke="#15803D" stroke-width="3" stroke-linecap="round"/>
            <circle cx="74" cy="35" r="14" fill="#FEE2E2" stroke="#DC2626" stroke-width="2.5"/>
            <path d="M69 30 L79 40 M79 30 L69 40" fill="none" stroke="#B91C1C" stroke-width="3" stroke-linecap="round"/>
            <rect x="98" y="25" width="64" height="8" rx="4" fill="#166534"/>
            <rect x="98" y="40" width="48" height="7" rx="3.5" fill="#86EFAC"/>
        </svg>
        """,
        "multi_choice": """
        <svg viewBox="0 0 180 70" width="170" height="64" xmlns="http://www.w3.org/2000/svg">
            <rect x="4" y="6" width="172" height="58" rx="14" fill="#ECFDF5" stroke="#10B981" stroke-width="2.5"/>
            <rect x="20" y="17" width="16" height="16" rx="4" fill="#10B981"/>
            <path d="M23 25 L27 29 L33 21" fill="none" stroke="#FFFFFF" stroke-width="2.5" stroke-linecap="round"/>
            <rect x="20" y="38" width="16" height="16" rx="4" fill="#10B981"/>
            <path d="M23 46 L27 50 L33 42" fill="none" stroke="#FFFFFF" stroke-width="2.5" stroke-linecap="round"/>
            <rect x="46" y="21" width="108" height="8" rx="4" fill="#047857"/>
            <rect x="46" y="42" width="86" height="8" rx="4" fill="#6EE7B7"/>
        </svg>
        """,
        "short_answer": """
        <svg viewBox="0 0 180 70" width="170" height="64" xmlns="http://www.w3.org/2000/svg">
            <rect x="4" y="6" width="172" height="58" rx="14" fill="#FFFBEB" stroke="#F59E0B" stroke-width="2.5"/>
            <rect x="20" y="20" width="106" height="30" rx="8" fill="#FFFFFF" stroke="#D97706" stroke-width="2" stroke-dasharray="5,3"/>
            <line x1="30" y1="35" x2="96" y2="35" stroke="#B45309" stroke-width="3" stroke-linecap="round"/>
            <polygon points="136,46 156,22 163,28 143,52" fill="#F97316"/>
        </svg>
        """,
        "essay": """
        <svg viewBox="0 0 180 70" width="170" height="64" xmlns="http://www.w3.org/2000/svg">
            <rect x="4" y="6" width="172" height="58" rx="14" fill="#F5F3FF" stroke="#7C3AED" stroke-width="2.5"/>
            <rect x="20" y="16" width="112" height="38" rx="8" fill="#FFFFFF" stroke="#8B5CF6" stroke-width="2"/>
            <line x1="28" y1="26" x2="120" y2="26" stroke="#6D28D9" stroke-width="3" stroke-linecap="round"/>
            <line x1="28" y1="35" x2="112" y2="35" stroke="#A78BFA" stroke-width="3" stroke-linecap="round"/>
            <line x1="28" y1="44" x2="94" y2="44" stroke="#C4B5FD" stroke-width="3" stroke-linecap="round"/>
            <circle cx="150" cy="35" r="11" fill="#DDD6FE" stroke="#7C3AED" stroke-width="2"/>
        </svg>
        """,
        "matching": """
        <svg viewBox="0 0 180 70" width="170" height="64" xmlns="http://www.w3.org/2000/svg">
            <rect x="4" y="6" width="172" height="58" rx="14" fill="#FDF2F8" stroke="#EC4899" stroke-width="2.5"/>
            <circle cx="34" cy="24" r="7" fill="#EC4899"/>
            <circle cx="34" cy="46" r="7" fill="#8B5CF6"/>
            <circle cx="144" cy="24" r="7" fill="#8B5CF6"/>
            <circle cx="144" cy="46" r="7" fill="#EC4899"/>
            <path d="M41 24 C88 24, 90 46, 137 46" fill="none" stroke="#EC4899" stroke-width="3" stroke-linecap="round"/>
            <path d="M41 46 C88 46, 90 24, 137 24" fill="none" stroke="#8B5CF6" stroke-width="3" stroke-linecap="round"/>
        </svg>
        """
    }
    return illustrations.get(kind, illustrations["hero"])


def play_feedback_sound(kind: str = "correct", trigger_id: str = ""):
    """Phát âm thanh phản hồi đúng / sai / hoàn thành bằng Web Audio API."""
    if not st.session_state.get("sound_enabled", True):
        return

    if kind == "correct":
        js_notes = """
        const notes = [523.25, 659.25, 783.99, 1046.50];
        notes.forEach((freq, idx) => {
            const osc = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.type = 'triangle';
            osc.frequency.setValueAtTime(freq, ctx.currentTime + idx * 0.09);
            gain.gain.setValueAtTime(0.001, ctx.currentTime + idx * 0.09);
            gain.gain.exponentialRampToValueAtTime(0.18, ctx.currentTime + idx * 0.09 + 0.02);
            gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + idx * 0.09 + 0.24);
            osc.connect(gain);
            gain.connect(ctx.destination);
            osc.start(ctx.currentTime + idx * 0.09);
            osc.stop(ctx.currentTime + idx * 0.09 + 0.25);
        });
        """
    elif kind == "wrong":
        js_notes = """
        const notes = [311.13, 233.08];
        notes.forEach((freq, idx) => {
            const osc = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.type = 'sine';
            osc.frequency.setValueAtTime(freq, ctx.currentTime + idx * 0.16);
            gain.gain.setValueAtTime(0.001, ctx.currentTime + idx * 0.16);
            gain.gain.exponentialRampToValueAtTime(0.14, ctx.currentTime + idx * 0.16 + 0.02);
            gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + idx * 0.16 + 0.28);
            osc.connect(gain);
            gain.connect(ctx.destination);
            osc.start(ctx.currentTime + idx * 0.16);
            osc.stop(ctx.currentTime + idx * 0.16 + 0.30);
        });
        """
    else:
        js_notes = """
        const notes = [523.25, 659.25, 783.99, 1046.50, 1318.51];
        notes.forEach((freq, idx) => {
            const osc = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.type = 'triangle';
            osc.frequency.setValueAtTime(freq, ctx.currentTime + idx * 0.08);
            gain.gain.setValueAtTime(0.001, ctx.currentTime + idx * 0.08);
            gain.gain.exponentialRampToValueAtTime(0.20, ctx.currentTime + idx * 0.08 + 0.02);
            gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + idx * 0.08 + 0.32);
            osc.connect(gain);
            gain.connect(ctx.destination);
            osc.start(ctx.currentTime + idx * 0.08);
            osc.stop(ctx.currentTime + idx * 0.08 + 0.34);
        });
        """

    components.html(
        f"""
        <!-- sound_trigger:{trigger_id}_{time.time_ns()} -->
        <script>
        (function() {{
            try {{
                const AudioCtx = window.AudioContext || window.webkitAudioContext;
                const ctx = new AudioCtx();
                if (ctx.state === 'suspended') {{
                    ctx.resume();
                }}
                {js_notes}
            }} catch (e) {{}}
        }})();
        </script>
        """,
        height=0,
        width=0,
    )


def normalize_vietnamese_and_clean(text: str) -> str:
    """Chuẩn hóa unicode NFC tiếng Việt, loại bỏ emoji, ký tự điều khiển và ký tự không cần thiết."""
    if not text:
        return ""
    s = unicodedata.normalize("NFC", str(text))
    s = s.replace("\ufffd", "")
    s = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", s)
    s = re.sub(r"[\U00010000-\U0010FFFF]", "", s)
    s = re.sub(r"[\u2190-\u2BFF]", "", s)
    s = s.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\s+([,.;:!?])", r"\1", s)
    return s.strip()


def rule_fix_common_typos(text: str) -> str:
    """Sửa nhanh các từ viết tắt / lỗi gõ phím tiếng Việt phổ biến trước khi gửi AI rà soát."""
    if not text:
        return ""
    typo_map = {
        r"\bko\b": "không",
        r"\bk\b": "không",
        r"\bcx\b": "cũng",
        r"\bcs\b": "có",
        r"\bphk\b": "phải",
        r"\bveef\b": "về",
        r"\btrc\b": "trước",
        r"\bnhung\b": "nhưng",
        r"\bhs\b": "học sinh",
        r"\bgv\b": "giáo viên",
        r"\bthcs\b": "THCS",
        r"\biep\b": "IEP",
    }
    out = unicodedata.normalize("NFC", text)
    for pat, repl in typo_map.items():
        out = re.sub(pat, repl, out, flags=re.IGNORECASE)
    return out


@st.cache_data(show_spinner=False)
def ai_proofread_vietnamese_cached(text_to_check: str) -> str:
    """Rà soát và chuẩn hóa chính tả tiếng Việt bằng AI cho nội dung do người dùng nhập."""
    if not text_to_check.strip() or not client:
        return text_to_check
    if len(text_to_check) > 6000:
        return text_to_check

    prompt = f"""
    Hãy rà soát lỗi chính tả, dấu thanh tiếng Việt, lỗi gõ phím và viết hoa đầu câu cho đoạn văn bản dưới đây.
    Giữ nguyên toàn bộ ý nghĩa, cấu trúc dòng, số liệu và thuật ngữ chuyên môn. Không thêm lời bình luận hay giải thích.
    Chỉ trả về đúng văn bản đã chuẩn hóa chính tả:

    {text_to_check}
    """
    try:
        resp = client.chat.completions.create(
            model="gemini-3.8-flash-high",
            messages=[{"role": "user", "content": prompt}]
        )
        corrected = (resp.choices[0].message.content or "").strip()
        return corrected if corrected else text_to_check
    except Exception:
        return text_to_check


def proofread_and_normalize_user_input(raw_text: str, use_ai: bool = True) -> str:
    """Bước rà soát và chuẩn hóa chính tả tự động cho nội dung người dùng nhập trước khi xuất file."""
    step1 = rule_fix_common_typos(normalize_vietnamese_and_clean(raw_text))
    if not step1:
        return ""
    if use_ai and st.session_state.get("auto_spellcheck_export", True):
        step1 = ai_proofread_vietnamese_cached(step1)
    return normalize_vietnamese_and_clean(step1)


def clean_text_for_export(raw_text: str, proofread_user_input: bool = True) -> str:
    """Làm sạch văn bản xuất file: rà soát chính tả, bỏ thao tác thừa, bỏ dấu Markdown/emoji, chuẩn hóa dấu câu."""
    prepared = proofread_and_normalize_user_input(raw_text, use_ai=proofread_user_input)

    redundant_patterns = [
        r"(?i)\b(?:gv|giáo viên)\s+đặt\s+bút(?:\s+chì)?\s+vào\s+tay\s+(?:hs|học sinh)[^:\n]*:?",
        r"(?i)\bviết\s+sẵn\s+khung\s+giàn\s+giáo(?:\s+cho\s+bài)?[^:\n]*:?",
        r"(?i)\bđặt\s+bút(?:\s+chì)?\s+vào\s+tay\s+(?:hs|học sinh)[^:\n]*:?",
        r"(?i)\bgiáo viên\s+(?:thực hiện|làm mẫu|hướng dẫn thao tác)\s+như sau[^:\n]*:?",
    ]

    cleaned_lines = []
    for line in prepared.splitlines():
        s = normalize_vietnamese_and_clean(line)
        if not s:
            cleaned_lines.append("")
            continue

        s = re.sub(r"^[•●▪▸►\-*]+\s*", "", s)
        for pat in redundant_patterns:
            s = re.sub(pat, "", s).strip()

        s = re.sub(r"^[:\-–]+\s*", "", s).strip()
        s = re.sub(r"\*\*(.*?)\*\*", r"\1", s)
        s = re.sub(r"__(.*?)__", r"\1", s)
        s = s.replace("`", "").replace("#", "")
        s = normalize_vietnamese_and_clean(s)

        if s:
            if s[0].islower():
                s = s[0].upper() + s[1:]
            cleaned_lines.append(s)

    normalized = "\n".join(cleaned_lines)
    normalized = re.sub(r"\n{3,}", "\n\n", normalized).strip()
    return normalized


def build_docx_bytes(title_text: str, markdown_body: str, proofread_user_input: bool = True) -> bytes:
    """Tạo file Word (.docx) hoặc văn bản sạch nội dung, đã rà soát chính tả và loại bỏ ký tự không cần thiết."""
    clean_title = clean_text_for_export(title_text, proofread_user_input=proofread_user_input)
    cleaned_body = clean_text_for_export(markdown_body, proofread_user_input=proofread_user_input)
    if not DOCX_AVAILABLE:
        combined = f"{clean_title}\n\n{cleaned_body}".strip()
        return combined.encode("utf-8")

    doc = Document()
    doc.add_heading(clean_title, level=1)
    for line in cleaned_body.splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue
        doc.add_paragraph(cleaned)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def export_clean_file(
    title_text: str,
    raw_body: str,
    base_filename: str,
    label: str,
    key: str,
    proofread_user_input: bool = True,
):
    """Hàm xuất file thống nhất cho mọi khu vực: tự động rà soát chính tả và làm sạch ký tự không cần thiết."""
    file_ext = "docx" if DOCX_AVAILABLE else "txt"
    mime_type = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        if DOCX_AVAILABLE
        else "text/plain"
    )
    payload = build_docx_bytes(title_text, raw_body, proofread_user_input=proofread_user_input)
    safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", normalize_vietnamese_and_clean(base_filename)).strip("_") or "tai_lieu"
    st.download_button(
        label=f"{label} (.{file_ext})",
        data=payload,
        file_name=f"{safe_name}.{file_ext}",
        mime=mime_type,
        key=key,
        use_container_width=True
    )


def format_questions_for_export(questions: list[dict], include_answers: bool = False) -> str:
    """Định dạng danh sách câu hỏi thành văn bản chuẩn để xuất file (có thể ẩn đáp án cho học sinh)."""
    lines = []
    for idx, q in enumerate(questions, start=1):
        qtype = q.get("type", "true_false")
        q_text = clean_text_for_export(q.get("question", ""), proofread_user_input=False)
        lines.append(f"Câu {idx}: {q_text}")

        if qtype == "true_false":
            lines.append("   Chọn: Đúng / Sai")
        elif qtype in ("single_choice", "multi_choice"):
            for opt in q.get("options", []):
                opt_clean = clean_text_for_export(str(opt), proofread_user_input=False)
                lines.append(f"   - {opt_clean}")
        elif qtype == "short_answer":
            lines.append("   Trả lời ngắn: ............................")
        elif qtype == "essay":
            lines.append("   Bài làm tự luận: ......................................................")
        elif qtype == "matching":
            pairs = q.get("pairs", {})
            right_choices = q.get("right_choices", list(pairs.values()))
            lines.append("   Cột trái:")
            for k in pairs.keys():
                lines.append(f"     - {clean_text_for_export(str(k), proofread_user_input=False)}")
            lines.append("   Cột phải:")
            for r_item in right_choices:
                lines.append(f"     - {clean_text_for_export(str(r_item), proofread_user_input=False)}")

        if include_answers:
            if qtype == "true_false":
                tf_ans = "Đúng" if q.get("is_true", True) else "Sai"
                lines.append(f"   Đáp án đúng: {tf_ans}")
            elif qtype == "single_choice":
                lines.append(f"   Đáp án đúng: {clean_text_for_export(str(q.get('answer', '')), proofread_user_input=False)}")
            elif qtype == "multi_choice":
                ans_joined = ", ".join(clean_text_for_export(str(a), proofread_user_input=False) for a in q.get("answers", []))
                lines.append(f"   Các đáp án đúng: {ans_joined}")
            elif qtype == "short_answer":
                lines.append(f"   Đáp án mẫu: {clean_text_for_export(str(q.get('sample_answer', '')), proofread_user_input=False)}")
            elif qtype == "essay":
                lines.append(f"   Gợi ý đáp án tự luận: {clean_text_for_export(str(q.get('sample_answer', '')), proofread_user_input=False)}")
            elif qtype == "matching":
                pair_strs = [
                    f"{clean_text_for_export(str(k), proofread_user_input=False)} - {clean_text_for_export(str(v), proofread_user_input=False)}"
                    for k, v in q.get("pairs", {}).items()
                ]
                lines.append(f"   Ghép cặp chuẩn: {' ; '.join(pair_strs)}")
            exp_clean = clean_text_for_export(str(q.get("explanation", "")), proofread_user_input=False)
            if exp_clean:
                lines.append(f"   Giải thích: {exp_clean}")

        lines.append("")

    return "\n".join(lines).strip()


def format_error_report_for_export(rep: dict) -> str:
    """Định dạng báo cáo tổng hợp lỗi sai, phạm trù, lý do sai, cách làm đúng dạng và bài tập tương tự để xuất file."""
    lines = [
        f"Học sinh: {clean_text_for_export(rep.get('student', ''), proofread_user_input=False)}",
        f"Môn học: {clean_text_for_export(rep.get('subject', ''), proofread_user_input=False)}",
        f"Bài làm: {clean_text_for_export(rep.get('exam_title', rep.get('subject', '')), proofread_user_input=False)}",
        f"Tổng số câu: {rep.get('total_learned', 0)} | Số câu đúng: {rep.get('correct_count', 0)} | Số câu sai: {rep.get('wrong_count', 0)}",
        ""
    ]

    err_items = rep.get("error_items", [])
    if err_items:
        lines.append("CHI TIẾT TỪNG CÂU SAI:")
        for err in err_items:
            lines.append(f"Câu {err.get('q_number', '')}: {clean_text_for_export(err.get('question_text', ''), proofread_user_input=False)}")
            lines.append(f"  - Phạm trù kiến thức: {clean_text_for_export(err.get('category', ''), proofread_user_input=False)}")
            lines.append(f"  - Câu trả lời của học sinh: {clean_text_for_export(err.get('user_answer', ''), proofread_user_input=True)}")
            lines.append(f"  - Đáp án chuẩn: {clean_text_for_export(err.get('correct_answer', ''), proofread_user_input=False)}")
            lines.append(f"  - Lý do sai: {clean_text_for_export(err.get('reason', ''), proofread_user_input=False)}")
            lines.append(f"  - Cách làm đúng dạng này: {clean_text_for_export(err.get('how_to_fix', ''), proofread_user_input=False)}")
            lines.append("")

    remedial_qs = rep.get("remedial_questions", [])
    if remedial_qs:
        lines.append("BÀI TẬP TƯƠNG TỰ GỢI Ý GIAO LẠI:")
        lines.append(format_questions_for_export(remedial_qs, include_answers=True))

    return "\n".join(lines).strip()


def extract_uploaded_doc_text(uploaded_file) -> str:
    """Đọc nội dung cơ bản từ file đính kèm (.txt, .docx, .pdf)."""
    if uploaded_file is None:
        return ""
    name = uploaded_file.name.lower()
    raw_bytes = uploaded_file.getvalue()

    if name.endswith(".txt"):
        return normalize_vietnamese_and_clean(raw_bytes.decode("utf-8", errors="ignore"))

    if name.endswith(".docx") and DOCX_AVAILABLE:
        try:
            doc = Document(io.BytesIO(raw_bytes))
            paragraphs = [normalize_vietnamese_and_clean(p.text) for p in doc.paragraphs if p.text.strip()]
            return "\n".join(paragraphs) if paragraphs else f"[Đã đính kèm file: {uploaded_file.name}]"
        except Exception:
            return f"[Đã đính kèm file: {uploaded_file.name}]"

    return f"[Đã đính kèm file bài học: {uploaded_file.name}]"


# --- CẤU HÌNH GIAO DIỆN TRANG WEB ---
st.set_page_config(
    page_title="Nền tảng Hỗ trợ Giáo dục Hòa nhập THCS",
    page_icon="🌈",
    layout="wide"
)

# --- TÙY CHỈNH GIAO DIỆN RỰC RỠ, SẮC NÉT, NHIỀU HÌNH MINH HỌA ---
st.markdown("""
    <style>
    html, body, [class*="css"] {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-size: 20px !important;
        color: #1E293B;
    }
    .stApp {
        background: radial-gradient(circle at top left, #FFF1F2 0%, #EFF6FF 42%, #ECFDF5 100%);
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #E0F2FE 0%, #DCFCE7 52%, #FEF3C7 100%);
        border-right: 4px solid #0EA5E9;
    }
    [data-testid="stSidebar"] * {
        font-size: 19px !important;
    }
    h1 {
        font-size: 38px !important;
        color: #1D4ED8 !important;
        font-weight: 800 !important;
    }
    h2 {
        font-size: 30px !important;
        color: #0F766E !important;
        font-weight: 800 !important;
    }
    h3 {
        font-size: 23px !important;
        color: #312E81 !important;
        font-weight: 700 !important;
    }
    .hero-banner {
        background: linear-gradient(125deg, #1D4ED8 0%, #0284C7 42%, #059669 80%, #D97706 100%);
        color: #FFFFFF;
        padding: 22px 26px;
        border-radius: 24px;
        box-shadow: 0 14px 28px rgba(29, 78, 216, 0.22);
        margin-bottom: 18px;
        border: 2px solid rgba(255,255,255,0.45);
    }
    .hero-banner h2, .hero-banner h3, .hero-banner p {
        color: #FFFFFF !important;
        margin: 0;
    }
    .visual-card {
        background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 100%);
        border-radius: 20px;
        padding: 18px;
        box-shadow: 0 10px 22px rgba(15, 23, 42, 0.08);
        border-top: 6px solid #3B82F6;
        margin-bottom: 14px;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .visual-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 14px 26px rgba(15, 23, 42, 0.12);
    }
    .visual-card.green { border-top-color: #10B981; background: linear-gradient(180deg, #FFFFFF 0%, #F0FDF4 100%); }
    .visual-card.amber { border-top-color: #F59E0B; background: linear-gradient(180deg, #FFFFFF 0%, #FFFBEB 100%); }
    .visual-card.purple { border-top-color: #8B5CF6; background: linear-gradient(180deg, #FFFFFF 0%, #F5F3FF 100%); }
    .visual-card.rose { border-top-color: #F43F5E; background: linear-gradient(180deg, #FFFFFF 0%, #FFF1F2 100%); }

    .question-box {
        background: linear-gradient(135deg, #FFFFFF 0%, #F8FAFF 100%);
        border-radius: 20px;
        padding: 20px 24px;
        border-left: 8px solid #4F46E5;
        box-shadow: 0 8px 20px rgba(79, 70, 229, 0.11);
        margin-bottom: 14px;
    }
    .badge-type {
        display: inline-block;
        padding: 5px 14px;
        border-radius: 999px;
        font-size: 15px !important;
        font-weight: 800;
        background: linear-gradient(90deg, #E0E7FF 0%, #FCE7F3 100%);
        color: #312E81;
        margin-bottom: 8px;
    }
    .stAlert {
        border-radius: 14px;
        font-size: 19px !important;
        font-weight: 600;
    }
    .stButton>button {
        background: linear-gradient(90deg, #2563EB 0%, #0D9488 55%, #16A34A 100%);
        color: white;
        border-radius: 14px;
        padding: 0.78rem 1.45rem;
        font-size: 20px !important;
        font-weight: 700;
        border: none;
        box-shadow: 0 6px 14px rgba(13, 148, 136, 0.25);
        transition: 0.25s;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        background: linear-gradient(90deg, #1D4ED8 0%, #0F766E 100%);
    }
    .stTextInput input, .stTextArea textarea, .stSelectbox select {
        font-size: 19px !important;
        border-radius: 12px !important;
        border: 2px solid #60A5FA !important;
        padding: 10px !important;
        background-color: #FFFFFF !important;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 19px !important;
        font-weight: 700 !important;
        padding: 10px 20px;
        border-radius: 12px 12px 0 0;
    }
    </style>
""", unsafe_allow_html=True)

# --- HÀM GỌI AI CÓ TÍCH HỢP CACHING VÀ DÙNG ĐÚNG MODEL GEMINI TỪ GÓI ĐÃ MUA ---
@st.cache_data(show_spinner=False)
def call_ai_cached(prompt_text):
    if not client:
        return "⚠️ Chưa cấu hình client AI."

    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model="gemini-3.8-flash-high",
                messages=[{"role": "user", "content": prompt_text}]
            )
            return response.choices[0].message.content
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                raise e


def call_ai_multimodal(prompt_text: str, image_bytes: bytes | None = None, mime_type: str = "image/jpeg"):
    """Gọi AI hỗ trợ kèm hình ảnh bài học."""
    if not client:
        return "⚠️ Chưa cấu hình client AI."

    if image_bytes:
        b64_img = base64.b64encode(image_bytes).decode("utf-8")
        content_payload = [
            {"type": "text", "text": prompt_text},
            {
                "type": "image_url",
                "image_url": {"url": f"data:{mime_type};base64,{b64_img}"}
            }
        ]
    else:
        content_payload = prompt_text

    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model="gemini-3.8-flash-high",
                messages=[{"role": "user", "content": content_payload}]
            )
            return response.choices[0].message.content
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                raise e


def sanitize_question_for_student(q: dict) -> dict:
    """Loại bỏ đáp án, gợi ý đáp án và lời giải thích khi giao bài cho học sinh."""
    return {
        "id": q.get("id", f"q_{int(time.time() * 1000)}"),
        "type": q.get("type", "true_false"),
        "type_label": q.get("type_label", "Câu hỏi"),
        "category": q.get("category", ""),
        "question": q.get("question", ""),
        "is_true": q.get("is_true", True),
        "options": list(q.get("options", [])),
        "answers": list(q.get("answers", [])),
        "answer": q.get("answer", ""),
        "keywords": list(q.get("keywords", [])),
        "pairs": dict(q.get("pairs", {})),
        "right_choices": list(q.get("right_choices", [])),
        "hide_answers": True,
    }


def build_question_spec(
    use_mcq: bool,
    mcq_true_false: bool,
    mcq_multi: bool,
    mcq_short: bool,
    cnt_tf: int,
    cnt_multi: int,
    cnt_short: int,
    use_essay: bool,
    cnt_essay: int,
) -> list[tuple[str, int]]:
    """Tạo danh sách (loại_câu_hỏi, số_lượng) theo đúng lựa chọn Trắc nghiệm/Tự luận và số câu."""
    specs: list[tuple[str, int]] = []
    if use_mcq:
        if mcq_true_false and cnt_tf > 0:
            specs.append(("true_false", cnt_tf))
        if mcq_multi and cnt_multi > 0:
            specs.append(("multi_choice", cnt_multi))
        if mcq_short and cnt_short > 0:
            specs.append(("short_answer", cnt_short))
    if use_essay and cnt_essay > 0:
        specs.append(("essay", cnt_essay))
    return specs


def get_sample_custom_questions(subject: str, topic: str, specs: list[tuple[str, int]]) -> list[dict]:
    """Sinh bộ câu hỏi mẫu đúng theo các dạng và số lượng người dùng chọn."""
    t_label = topic.strip() if topic.strip() else f"Ôn tập {subject}"
    questions: list[dict] = []
    seq = 1

    for qtype, count in specs:
        for i in range(1, count + 1):
            if qtype == "true_false":
                is_t = (i % 2 == 1)
                stmt = (
                    f"Chia nhỏ kiến thức thành từng bước ngắn giúp học sinh {subject} ({t_label}) dễ ghi nhớ hơn."
                    if is_t
                    else f"Học vẹt thuộc lòng toàn bộ văn bản mà không cần hiểu ý nghĩa là cách tốt nhất cho {subject} ({t_label})."
                )
                questions.append({
                    "id": f"q_tf_{seq}",
                    "type": "true_false",
                    "type_label": "✅ Trắc nghiệm Đúng / Sai",
                    "category": f"{subject} — {t_label} (Nhận định Đúng/Sai)",
                    "question": f"Phát biểu ({i}/{count}): {stmt}",
                    "is_true": is_t,
                    "explanation": (
                        "Đúng vì chia nhỏ kiến thức giúp giảm tải nhận thức và nhớ lâu hơn."
                        if is_t
                        else "Sai vì hiểu bản chất và gắn ví dụ thực tế mới giúp vận dụng tốt lâu dài."
                    )
                })
            elif qtype == "multi_choice":
                questions.append({
                    "id": f"q_multi_{seq}",
                    "type": "multi_choice",
                    "type_label": "☑️ Trắc nghiệm Nhiều lựa chọn",
                    "category": f"{subject} — {t_label} (Tổng hợp nhiều ý đúng)",
                    "question": f"Câu hỏi nhiều lựa chọn ({i}/{count}) về '{t_label}' ({subject}): Chọn tất cả cách học hiệu quả:",
                    "options": [
                        "Tóm tắt ý chính bằng từ khóa nổi bật",
                        "Tự giải thích lại kiến thức bằng lời của em",
                        "Kiểm tra lại lý do làm sai ở từng câu",
                        "Học dồn thức khuya trước ngày kiểm tra"
                    ],
                    "answers": [
                        "Tóm tắt ý chính bằng từ khóa nổi bật",
                        "Tự giải thích lại kiến thức bằng lời của em",
                        "Kiểm tra lại lý do làm sai ở từng câu"
                    ],
                    "explanation": "Nắm từ khóa, tự diễn đạt lại và rút kinh nghiệm từ câu sai giúp tiến bộ rõ rệt."
                })
            elif qtype == "short_answer":
                questions.append({
                    "id": f"q_short_{seq}",
                    "type": "short_answer",
                    "type_label": "✏️ Trắc nghiệm Trả lời ngắn",
                    "category": f"{subject} — {t_label} (Từ khóa trọng tâm)",
                    "question": f"Trả lời ngắn ({i}/{count}) - Chủ đề '{t_label}' ({subject}): Khi gặp bài khó, em nên chia thành từng ______ nhỏ để dễ hiểu?",
                    "keywords": ["bước", "phần", "ý", "đoạn"],
                    "sample_answer": "bước",
                    "explanation": "Chia thành từng bước nhỏ giúp em tập trung từng ý một mà không bị ngợp."
                })
            elif qtype == "essay":
                questions.append({
                    "id": f"q_essay_{seq}",
                    "type": "essay",
                    "type_label": "📝 Tự luận",
                    "category": f"{subject} — {t_label} (Vận dụng & nêu ví dụ)",
                    "question": f"Câu hỏi tự luận ({i}/{count}) - Chủ đề '{t_label}' ({subject}): Hãy nêu một ví dụ thực tế liên quan đến '{t_label}' và giải thích ngắn gọn cách em áp dụng kiến thức đó.",
                    "sample_answer": f"Nêu được ví dụ gắn với '{t_label}' và giải thích rõ ràng từng bước thực hiện.",
                    "explanation": "Bài tự luận tốt cần có ví dụ cụ thể, trình bày mạch lạc từng ý và nêu được kết luận."
                })
            seq += 1

    return questions


def generate_ai_custom_questions(
    subject: str,
    topic_text: str,
    specs: list[tuple[str, int]],
    extra_file_text: str = "",
    image_bytes: bytes | None = None,
    mime_type: str = "image/jpeg",
) -> list[dict]:
    """Dùng AI sinh câu hỏi JSON theo đúng các loại và số lượng câu được yêu cầu."""
    spec_desc = ", ".join(f"{cnt} câu dạng '{qtype}'" for qtype, cnt in specs)
    prompt = f"""
    Bạn là giáo viên THCS chuyên thiết kế bài kiểm tra và ôn tập hòa nhập.
    Hãy tạo chính xác bộ câu hỏi theo yêu cầu sau cho môn {subject}:
    - Chủ đề/Nội dung: {topic_text}
    - Nội dung file bổ sung: {extra_file_text}
    - Cấu trúc bắt buộc: {spec_desc}.

    Các giá trị "type" hợp lệ:
    - "true_false": câu hỏi trắc nghiệm Đúng/Sai -> gồm "question", "category" (phạm trù kiến thức cụ thể), "is_true" (true hoặc false), "explanation".
    - "multi_choice": câu hỏi trắc nghiệm nhiều lựa chọn -> gồm "question", "category", "options" (4 lựa chọn), "answers" (mảng các lựa chọn đúng), "explanation".
    - "short_answer": câu hỏi trắc nghiệm trả lời ngắn -> gồm "question", "category", "keywords" (mảng từ khóa ngắn), "sample_answer", "explanation".
    - "essay": câu hỏi tự luận -> gồm "question", "category", "sample_answer" (gợi ý đáp án), "explanation".

    BẮT BUỘC trả về DUY NHẤT một chuỗi JSON hợp lệ (không bọc trong markdown ```json hay lời dẫn ngoài JSON):
    [
      {{
        "id": "q_1",
        "type": "true_false",
        "type_label": "✅ Trắc nghiệm Đúng / Sai",
        "category": "{subject} - {topic_text}",
        "question": "Nội dung phát biểu...",
        "is_true": true,
        "explanation": "Giải thích..."
      }}
    ]
    """
    raw = call_ai_multimodal(prompt, image_bytes=image_bytes, mime_type=mime_type)
    cleaned = raw.strip()
    if cleaned.startswith("```"):
        cleaned = cleaned.split("\n", 1)[-1]
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3]
        cleaned = cleaned.strip()

    parsed = json.loads(cleaned)
    if not isinstance(parsed, list) or len(parsed) == 0:
        raise ValueError("JSON câu hỏi không hợp lệ")
    return parsed


def build_similar_question_for_error(q: dict, q_number: int, category: str, subject: str, attempt_num: int = 1) -> dict:
    """Tự động thiết kế một câu hỏi tương tự đúng dạng & phạm trù của câu làm sai."""
    qtype = q.get("type", "true_false")
    base_q = q.get("question", "").strip()
    suffix = f" (Lần luyện #{attempt_num})" if attempt_num > 1 else ""

    if qtype == "true_false":
        flipped_bool = not bool(q.get("is_true", True))
        return {
            "id": f"rem_tf_{q.get('id', q_number)}_{attempt_num}_{int(time.time() * 1000)}",
            "type": "true_false",
            "type_label": "✅ Trắc nghiệm Đúng / Sai",
            "category": category,
            "question": f"[Bài tập tương tự Câu {q_number}{suffix}] Xét phạm trù '{category}': Phát biểu sau đúng hay sai — {base_q}",
            "is_true": flipped_bool,
            "explanation": q.get("explanation", "")
        }
    if qtype == "multi_choice":
        opts = list(q.get("options", [])) or ["Ý đúng 1", "Ý đúng 2", "Ý chưa chính xác", "Ý gây nhiễu"]
        ans = list(q.get("answers", [])) or [opts[0]]
        return {
            "id": f"rem_multi_{q.get('id', q_number)}_{attempt_num}_{int(time.time() * 1000)}",
            "type": "multi_choice",
            "type_label": "☑️ Trắc nghiệm Nhiều lựa chọn",
            "category": category,
            "question": f"[Bài tập tương tự Câu {q_number}{suffix}] Thuộc phạm trù '{category}': Chọn tất cả phương án chính xác liên quan đến: {base_q}",
            "options": opts,
            "answers": ans,
            "explanation": q.get("explanation", "")
        }
    if qtype == "short_answer":
        sample = q.get("sample_answer", "từ khóa")
        kws = list(q.get("keywords", [])) or [str(sample)]
        return {
            "id": f"rem_short_{q.get('id', q_number)}_{attempt_num}_{int(time.time() * 1000)}",
            "type": "short_answer",
            "type_label": "✏️ Trắc nghiệm Trả lời ngắn",
            "category": category,
            "question": f"[Bài tập tương tự Câu {q_number}{suffix}] Phạm trù '{category}': Điền từ/cụm từ trọng tâm tương tự câu hỏi '{base_q}'",
            "keywords": kws,
            "sample_answer": sample,
            "explanation": q.get("explanation", "")
        }
    return {
        "id": f"rem_essay_{q.get('id', q_number)}_{attempt_num}_{int(time.time() * 1000)}",
        "type": "essay",
        "type_label": "📝 Tự luận",
        "category": category,
        "question": f"[Bài tập tương tự Câu {q_number}{suffix}] Phạm trù '{category}': Trình bày lại cách giải và nêu một ví dụ khác cho bài toán/yêu cầu: {base_q}",
        "sample_answer": q.get("sample_answer", ""),
        "explanation": q.get("explanation", "")
    }


def build_error_analysis_and_remedial(
    q: dict,
    q_number: int,
    user_val,
    subject: str,
    exam_title: str,
) -> dict:
    """Phân tích chi tiết 1 câu sai: phạm trù kiến thức, lý do sai, cách làm đúng dạng đó và sinh câu hỏi tương tự."""
    qtype = q.get("type", "true_false")
    q_text = q.get("question", "")
    category = q.get("category") or f"{subject} — {exam_title or q.get('type_label', 'Kiến thức trọng tâm')}"

    if qtype == "true_false":
        user_str = str(user_val) if user_val else "(Chưa chọn)"
        correct_str = "Đúng" if q.get("is_true", True) else "Sai"
        reason = f"Em chọn '{user_str}' trong khi nhận định đúng là '{correct_str}'. Có thể em đọc lướt điều kiện hoặc từ khóa khẳng định, phủ định trong câu."
        how_to_fix = f"Đọc chậm từng mệnh đề, gạch chân các từ khóa quan trọng và đối chiếu với quy tắc cốt lõi của '{category}' trước khi chọn Đúng hoặc Sai."
    elif qtype == "multi_choice":
        user_str = ", ".join(user_val or []) or "(Chưa tích chọn)"
        correct_str = ", ".join(q.get("answers", []))
        reason = f"Lựa chọn [{user_str}] chưa khớp đầy đủ đáp án chuẩn [{correct_str}]. Có thể em chọn thiếu ý đúng hoặc tích nhầm ý chưa phù hợp."
        how_to_fix = f"Kiểm tra độc lập từng phương án một (đúng hay sai với '{category}') thay vì chọn vội theo cảm tính, sau đó rà soát lại tất cả các ô đã chọn."
    elif qtype == "short_answer":
        user_str = str(user_val or "").strip() or "(Để trống)"
        correct_str = str(q.get("sample_answer", "")).strip()
        reason = f"Câu trả lời '{user_str}' chưa chứa từ khóa trọng tâm '{correct_str}' hoặc diễn đạt chưa đúng khái niệm cần điền."
        how_to_fix = f"Xác định từ khóa cốt lõi gắn với '{category}', viết ngắn gọn 1–3 từ chuẩn khái niệm và đọc lại cả câu sau khi điền để xem có hợp nghĩa không."
    elif qtype == "essay":
        user_str = str(user_val or "").strip() or "(Chưa viết bài làm)"
        correct_str = str(q.get("sample_answer", "")).strip()
        reason = "Câu trả lời tự luận còn quá ngắn hoặc thiếu ví dụ minh họa, lập luận rõ ràng theo yêu cầu của đề."
        how_to_fix = f"Chia bài làm thành 2 bước: (1) Nêu khái niệm hoặc quy tắc thuộc '{category}', (2) Đưa 1 ví dụ thực tế cụ thể và giải thích ngắn gọn."
    else:
        user_str = str(user_val)
        correct_str = str(q.get("answer", ""))
        reason = f"Lựa chọn '{user_str}' chưa đúng với đáp án chuẩn '{correct_str}'."
        how_to_fix = f"Ôn lại kiến thức nền tảng của '{category}' và làm thêm các câu tương tự theo từng bước."

    similar_q = build_similar_question_for_error(q, q_number=q_number, category=category, subject=subject, attempt_num=1)

    return {
        "q_number": q_number,
        "raw_question": dict(q),
        "question_text": q_text,
        "category": category,
        "user_answer": user_str,
        "correct_answer": correct_str,
        "reason": reason,
        "how_to_fix": how_to_fix,
        "similar_question": similar_q,
    }


def submit_error_report_to_teacher(
    student_name: str,
    subject: str,
    exam_title: str,
    total_q: int,
    correct_q: int,
    error_items: list[dict],
) -> dict:
    """Lưu tổng hợp lỗi sai chi tiết từng câu + bài tập tương tự vào session_state và file bền vững."""
    wrong_q = len(error_items)
    remedial_questions = [item["similar_question"] for item in error_items if item.get("similar_question")]
    detail_blocks = []
    for err in error_items:
        detail_blocks.append(
            f"- **Câu {err['q_number']}:** *{err['question_text']}*\n"
            f"  • **Phạm trù kiến thức:** {err['category']}\n"
            f"  • **Em chọn/trả lời:** {err['user_answer']}\n"
            f"  • **Lý do sai:** {err['reason']}\n"
            f"  • **Cách làm đúng dạng này:** {err['how_to_fix']}"
        )
    summary_md = "\n".join(detail_blocks) if detail_blocks else "Em làm đúng tất cả các câu trong bài!"

    report_obj = {
        "student": student_name,
        "subject": subject or "Tổng hợp",
        "exam_title": exam_title or subject,
        "total_learned": total_q,
        "correct_count": correct_q,
        "wrong_count": wrong_q,
        "error_items": error_items,
        "remedial_questions": remedial_questions,
        "remedial_send_count": 0,
        "content": f"Bài **{exam_title or subject}**: Đúng {correct_q}/{total_q} câu, sai {wrong_q} câu.\n{summary_md}"
    }

    if student_name:
        if student_name not in st.session_state.student_progress:
            st.session_state.student_progress[student_name] = {"total": 0, "correct": 0, "wrong": 0}
        st.session_state.student_progress[student_name]["total"] += total_q
        st.session_state.student_progress[student_name]["correct"] += correct_q
        st.session_state.student_progress[student_name]["wrong"] += wrong_q
        st.session_state.reports.append(report_obj)
        save_persistent_store()

    return report_obj


def evaluate_single_question(q: dict, user_val, prefix: str, hide_answers: bool = False) -> tuple[bool, str]:
    """Chấm điểm 1 câu hỏi bất kỳ và trả về (is_correct, feedback_message). Khi hide_answers=True tuyệt đối không lộ đáp án/giải thích."""
    qtype = q.get("type")

    if qtype == "true_false":
        target_bool = bool(q.get("is_true", True))
        if user_val == "Đúng":
            chosen_bool = True
        elif user_val == "Sai":
            chosen_bool = False
        else:
            return False, "💡 **Em hãy chọn Đúng hoặc Sai trước khi kiểm tra nhé!**"

        ok = (chosen_bool == target_bool)
        target_label = "Đúng" if target_bool else "Sai"
        if hide_answers:
            msg = "🔒 **Đã ghi nhận lựa chọn của em (khóa 1 lần chọn)!**"
        else:
            msg = (
                f"🎉 **Chính xác! Tuyệt vời lắm!** {q.get('explanation', '')}"
                if ok
                else f"💡 **Chưa đúng rồi!** Nhận định này là **{target_label}**. {q.get('explanation', '')}"
            )
        return ok, msg

    if qtype == "single_choice":
        correct_ans = q.get("answer", "")
        ok = (user_val == correct_ans)
        if hide_answers:
            msg = "🔒 **Đã ghi nhận lựa chọn của em (khóa 1 lần chọn)!**"
        else:
            msg = (
                f"🎉 **Chính xác! Tuyệt vời lắm!** {q.get('explanation', '')}"
                if ok
                else f"💡 **Chưa đúng rồi!** Đáp án đúng là: **{correct_ans}**. {q.get('explanation', '')}"
            )
        return ok, msg

    if qtype == "multi_choice":
        selected_set = set(user_val or [])
        target_set = set(q.get("answers", []))
        ok = (selected_set == target_set and len(target_set) > 0)
        target_str = ", ".join(q.get("answers", []))
        if hide_answers:
            msg = "🔒 **Đã ghi nhận lựa chọn của em (khóa 1 lần chọn)!**"
        else:
            msg = (
                f"🌟 **Xuất sắc! Em đã chọn đúng tất cả đáp án!** {q.get('explanation', '')}"
                if ok
                else f"💡 **Suýt trúng rồi!** Các đáp án đúng là: **{target_str}**. {q.get('explanation', '')}"
            )
        return ok, msg

    if qtype == "short_answer":
        ans_text = (user_val or "").strip().lower()
        kws = [k.strip().lower() for k in q.get("keywords", []) if k.strip()]
        sample = q.get("sample_answer", "")
        ok = bool(ans_text) and (not kws or any(kw in ans_text for kw in kws))
        if hide_answers:
            msg = "🔒 **Đã ghi nhận câu trả lời của em (khóa 1 lần chọn)!**"
        else:
            msg = (
                f"🎯 **Hay quá! Câu trả lời của em khớp ý chuẩn!** ({sample}) — {q.get('explanation', '')}"
                if ok
                else f"💡 **Gợi ý đáp án chuẩn:** **{sample}**. {q.get('explanation', '')}"
            )
        return ok, msg

    if qtype == "essay":
        ans_text = (user_val or "").strip()
        ok = len(ans_text) >= 8
        sample = q.get("sample_answer", "")
        if hide_answers:
            msg = "🔒 **Đã ghi nhận bài tự luận của em (khóa 1 lần chọn)!**"
        else:
            msg = (
                f"✍️ **Bài tự luận đã được ghi nhận!** Gợi ý tham khảo: **{sample}** — {q.get('explanation', '')}"
                if ok
                else f"💡 **Hãy viết chi tiết hơn nhé.** Gợi ý đáp án: **{sample}**. {q.get('explanation', '')}"
            )
        return ok, msg

    if qtype == "matching":
        pairs = q.get("pairs", {})
        user_map = user_val or {}
        if not pairs:
            return False, "Không có dữ liệu ghép cặp."
        matched_count = sum(1 for k, v in pairs.items() if user_map.get(k) == v)
        ok = (matched_count == len(pairs))
        detail_lines = [f"• *{k}* → **{v}**" for k, v in pairs.items()]
        detail_str = " | ".join(detail_lines)
        if hide_answers:
            msg = "🔒 **Đã ghi nhận các cặp ghép của em (khóa 1 lần chọn)!**"
        else:
            msg = (
                f"🏆 **Tuyệt đỉnh! Em ghép đúng cả {matched_count}/{len(pairs)} cặp!** {q.get('explanation', '')}"
                if ok
                else f"💡 **Em ghép đúng {matched_count}/{len(pairs)} cặp.** Đáp án chuẩn: {detail_str}. {q.get('explanation', '')}"
            )
        return ok, msg

    return False, "Chưa hỗ trợ dạng câu hỏi này."


def render_student_post_submit_report(report_obj: dict):
    """Hiển thị báo cáo lỗi sai & tiến bộ ngay trên giao diện học sinh sau khi nộp bài (không lộ đáp án chuẩn)."""
    total_q = report_obj.get("total_learned", 0)
    correct_q = report_obj.get("correct_count", 0)
    wrong_q = report_obj.get("wrong_count", 0)
    err_items = report_obj.get("error_items", [])

    st.markdown('<div class="visual-card rose">', unsafe_allow_html=True)
    st.markdown(f"### 📋 Báo cáo Tiến bộ & Lỗi sai — **{report_obj.get('exam_title', report_obj.get('subject', ''))}**")
    acc = (correct_q / total_q * 100) if total_q > 0 else 0.0
    st.markdown(
        f"📊 **Tiến bộ bài này:** Đúng **{correct_q}/{total_q}** câu | Sai: **{wrong_q}** câu | Tỷ lệ đúng: **{acc:.1f}%**"
    )
    if total_q > 0:
        st.progress(acc / 100.0)

    if err_items:
        st.markdown("#### 🔍 Các câu em làm chưa đúng & hướng cải thiện:")
        for err in err_items:
            st.markdown(
                f"- **Câu {err['q_number']}:** {err['question_text']}\n"
                f"  - **Phạm trù kiến thức:** `{err['category']}`\n"
                f"  - **Câu trả lời của em:** `{err['user_answer']}`\n"
                f"  - **Lý do chưa đúng:** {err['reason']}\n"
                f"  - **Cách làm đúng dạng này:** {err['how_to_fix']}"
            )
    else:
        st.success("🌟 Tuyệt vời! Em đã làm đúng toàn bộ các câu trong bài này!")
    st.markdown('</div>', unsafe_allow_html=True)


def render_interactive_question_list(
    questions: list[dict],
    ns_key: str,
    student_name: str = "",
    subject: str = "",
    exam_title: str = "",
    hide_answers: bool = False,
    show_report_on_finish: bool = False,
    report_state_key: str = "",
    on_submitted_callback=None,
):
    """Hiển thị & chấm điểm trực tiếp bộ câu hỏi kèm hình minh họa và âm thanh chúc mừng / phản hồi."""
    if not questions:
        st.info("ℹ️ Chưa có câu hỏi nào để hiển thị.")
        return

    score_correct = 0
    error_items: list[dict] = []

    for idx, q in enumerate(questions):
        q_id = q.get("id", f"q_{idx}")
        qtype = q.get("type", "true_false")
        state_key = f"q_eval_{ns_key}_{q_id}"
        lock_key = f"q_locked_{ns_key}_{q_id}"
        is_locked = bool(hide_answers and st.session_state.get(lock_key, False))

        col_q_txt, col_q_svg = st.columns([3.4, 1.1])
        with col_q_txt:
            lock_badge = " • 🔒 Đã khóa lựa chọn" if is_locked else ""
            st.markdown(f"""
                <div class="question-box">
                    <span class="badge-type">{q.get('type_label', 'Câu hỏi')}{lock_badge}</span>
                    <h3>Câu {idx + 1}: {q.get('question', '')}</h3>
                </div>
            """, unsafe_allow_html=True)
        with col_q_svg:
            st.markdown(get_illustration_svg(qtype), unsafe_allow_html=True)

        user_response = None

        if qtype == "true_false":
            user_response = st.radio(
                "Chọn Đúng hoặc Sai:",
                ["Đúng", "Sai"],
                disabled=is_locked,
                key=f"inp_{ns_key}_{q_id}_tf"
            )
        elif qtype == "single_choice":
            user_response = st.radio(
                "Chọn 1 đáp án đúng:",
                q.get("options", []),
                disabled=is_locked,
                key=f"inp_{ns_key}_{q_id}_single"
            )
        elif qtype == "multi_choice":
            chosen = []
            for opt_i, opt in enumerate(q.get("options", [])):
                if st.checkbox(opt, disabled=is_locked, key=f"inp_{ns_key}_{q_id}_multi_{opt_i}"):
                    chosen.append(opt)
            user_response = chosen
        elif qtype == "short_answer":
            user_response = st.text_input(
                "Nhập câu trả lời ngắn của em:",
                placeholder="Gõ câu trả lời ngắn gọn tại đây...",
                disabled=is_locked,
                key=f"inp_{ns_key}_{q_id}_short"
            )
        elif qtype == "essay":
            user_response = st.text_area(
                "Trình bày câu trả lời tự luận của em:",
                placeholder="Viết câu trả lời tự luận chi tiết tại đây...",
                disabled=is_locked,
                key=f"inp_{ns_key}_{q_id}_essay"
            )
        elif qtype == "matching":
            pairs = q.get("pairs", {})
            right_choices = q.get("right_choices", list(pairs.values()))
            chosen_map = {}
            for p_i, left_item in enumerate(pairs.keys()):
                col_l, col_r = st.columns([1.2, 1.3])
                with col_l:
                    st.markdown(f"🧩 **Vế trái:** {left_item}")
                with col_r:
                    chosen_map[left_item] = st.selectbox(
                        f"Chọn vế ghép cho mục {p_i + 1}:",
                        ["-- Chọn nội dung tương ứng --"] + right_choices,
                        disabled=is_locked,
                        key=f"inp_{ns_key}_{q_id}_match_{p_i}"
                    )
            user_response = chosen_map

        btn_label = f"🔒 Đã khóa câu {idx + 1}" if is_locked else f"🔊 Kiểm tra câu {idx + 1}"
        if st.button(btn_label, disabled=is_locked, key=f"check_btn_{ns_key}_{q_id}"):
            is_ok, fb_msg = evaluate_single_question(q, user_response, f"{ns_key}_{q_id}", hide_answers=hide_answers)
            st.session_state[state_key] = {"is_ok": is_ok, "msg": fb_msg, "user_response": user_response}
            if hide_answers:
                st.session_state[lock_key] = True
            play_feedback_sound("correct" if is_ok else "wrong", trigger_id=f"{ns_key}_{q_id}")
            if hide_answers:
                st.rerun()

        if state_key in st.session_state:
            res = st.session_state[state_key]
            if res["is_ok"]:
                score_correct += 1
                st.success(res["msg"])
            else:
                st.warning(res["msg"])
                error_items.append(
                    build_error_analysis_and_remedial(
                        q=q,
                        q_number=idx + 1,
                        user_val=res.get("user_response"),
                        subject=subject,
                        exam_title=exam_title,
                    )
                )

        st.markdown("---")

    if st.button("🎊 Nộp bài & Gửi báo cáo lỗi sai", key=f"finish_all_{ns_key}", use_container_width=True):
        if hide_answers:
            for q_i, q_item in enumerate(questions):
                q_id = q_item.get("id", f"q_{q_i}")
                st.session_state[f"q_locked_{ns_key}_{q_id}"] = True

        total_q = len(questions)
        wrong_q = len(error_items)
        if score_correct == total_q and total_q > 0:
            st.balloons()
            play_feedback_sound("finish", trigger_id=f"finish_{ns_key}")
            st.success(f"🏆 **CHÚC MỪNG RỰC RỠ!** Em đạt **{score_correct}/{total_q}** câu đúng (100%)!")
        else:
            play_feedback_sound("correct" if score_correct >= wrong_q else "wrong", trigger_id=f"finish_{ns_key}")
            st.info(f"📊 **Kết quả hiện tại:** Đúng **{score_correct}/{total_q}** câu | Sai/Chưa đạt: **{wrong_q}** câu.")

        if student_name:
            rep_obj = submit_error_report_to_teacher(
                student_name=student_name,
                subject=subject or "Tổng hợp",
                exam_title=exam_title or subject,
                total_q=total_q,
                correct_q=score_correct,
                error_items=error_items,
            )
            if report_state_key:
                st.session_state[report_state_key] = rep_obj
            if callable(on_submitted_callback):
                on_submitted_callback()
            st.success("📨 Đã lưu vĩnh viễn báo cáo lỗi sai, tiến bộ và bài tập tương tự cho cả học sinh và giáo viên!")

    if show_report_on_finish and report_state_key and st.session_state.get(report_state_key):
        render_student_post_submit_report(st.session_state[report_state_key])


def render_question_builder_controls(ns_prefix: str):
    """Giao diện chọn dạng Trắc nghiệm (Đúng/Sai, Nhiều lựa chọn, Trả lời ngắn), Tự luận và số câu từng dạng."""
    st.markdown("#### 🎯 Chọn dạng câu hỏi & số lượng câu muốn tạo:")
    c_type1, c_type2 = st.columns(2)
    with c_type1:
        use_mcq = st.checkbox("☑️ Trắc nghiệm", value=True, key=f"{ns_prefix}_use_mcq")
    with c_type2:
        use_essay = st.checkbox("📝 Tự luận", value=True, key=f"{ns_prefix}_use_essay")

    mcq_tf = False
    mcq_multi = False
    mcq_short = False
    cnt_tf = 0
    cnt_multi = 0
    cnt_short = 0
    cnt_essay = 0

    if use_mcq:
        st.markdown("##### 🔹 Tùy chọn dạng Trắc nghiệm & số câu:")
        m1, m2, m3 = st.columns(3)
        with m1:
            mcq_tf = st.checkbox("Trắc nghiệm Đúng / Sai", value=True, key=f"{ns_prefix}_mcq_tf")
            if mcq_tf:
                cnt_tf = st.number_input("Số câu Đúng / Sai:", min_value=1, max_value=15, value=2, step=1, key=f"{ns_prefix}_cnt_tf")
        with m2:
            mcq_multi = st.checkbox("Trắc nghiệm Nhiều lựa chọn", value=True, key=f"{ns_prefix}_mcq_multi")
            if mcq_multi:
                cnt_multi = st.number_input("Số câu Nhiều lựa chọn:", min_value=1, max_value=15, value=2, step=1, key=f"{ns_prefix}_cnt_multi")
        with m3:
            mcq_short = st.checkbox("Trắc nghiệm Trả lời ngắn", value=True, key=f"{ns_prefix}_mcq_short")
            if mcq_short:
                cnt_short = st.number_input("Số câu Trả lời ngắn:", min_value=1, max_value=15, value=2, step=1, key=f"{ns_prefix}_cnt_short")

    if use_essay:
        st.markdown("##### 🔹 Tùy chọn dạng Tự luận & số câu:")
        cnt_essay = st.number_input("Số câu Tự luận:", min_value=1, max_value=10, value=2, step=1, key=f"{ns_prefix}_cnt_essay")

    specs = build_question_spec(
        use_mcq=use_mcq,
        mcq_true_false=mcq_tf,
        mcq_multi=mcq_multi,
        mcq_short=mcq_short,
        cnt_tf=int(cnt_tf),
        cnt_multi=int(cnt_multi),
        cnt_short=int(cnt_short),
        use_essay=use_essay,
        cnt_essay=int(cnt_essay),
    )
    total_selected = sum(n for _, n in specs)
    st.caption(f"📌 Tổng số câu sẽ tạo: **{total_selected}** câu.")
    return specs


# --- QUẢN LÝ TRẠNG THÁI ĐĂNG NHẬP & DỮ LIỆU BỀN VỮNG ---
_persisted = load_persistent_store()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""
if "role" not in st.session_state:
    st.session_state.role = ""
if "email" not in st.session_state:
    st.session_state.email = ""
if "province" not in st.session_state:
    st.session_state.province = ""
if "ward" not in st.session_state:
    st.session_state.ward = ""
if "school" not in st.session_state:
    st.session_state.school = ""
if "class_name" not in st.session_state:
    st.session_state.class_name = ""
if "student_profiles" not in st.session_state:
    st.session_state.student_profiles = _persisted["student_profiles"]
if "reports" not in st.session_state:
    st.session_state.reports = _persisted["reports"]
if "assigned_tasks" not in st.session_state:
    st.session_state.assigned_tasks = _persisted["assigned_tasks"]
if "student_progress" not in st.session_state:
    st.session_state.student_progress = _persisted["student_progress"]
if "iep_result" not in st.session_state:
    st.session_state.iep_result = ""
if "theory_result" not in st.session_state:
    st.session_state.theory_result = ""
if "practice_quiz_result" not in st.session_state:
    st.session_state.practice_quiz_result = ""
if "student_practice_questions" not in st.session_state:
    st.session_state.student_practice_questions = []
if "teacher_draft_questions" not in st.session_state:
    st.session_state.teacher_draft_questions = []
if "sound_enabled" not in st.session_state:
    st.session_state.sound_enabled = True
if "auto_spellcheck_export" not in st.session_state:
    st.session_state.auto_spellcheck_export = True

# =========================================================================
# MÀN HÌNH ĐĂNG NHẬP
# =========================================================================
if not st.session_state.logged_in:
    col_b1, col_b2 = st.columns([2.8, 1.2])
    with col_b1:
        st.markdown("""
            <div class="hero-banner">
                <h2>🌈 NỀN TẢNG SỐ HỖ TRỢ GIÁO DỤC HÒA NHẬP THCS 🌈</h2>
                <p style="font-size: 20px; margin-top: 8px;">
                    🎨 Giao diện tươi sáng • 🖼️ Minh họa trực quan • 🔊 Âm thanh chúc mừng • 🧩 Trắc nghiệm & Tự luận tương tác
                </p>
            </div>
        """, unsafe_allow_html=True)
    with col_b2:
        st.markdown(get_illustration_svg("hero"), unsafe_allow_html=True)

    c_v1, c_v2, c_v3, c_v4 = st.columns(4)
    with c_v1:
        st.markdown('<div class="visual-card"><h3>📋 Lập IEP AI</h3><p>Soạn kế hoạch cá nhân hóa & xuất file Word gọn sạch.</p></div>', unsafe_allow_html=True)
    with c_v2:
        st.markdown('<div class="visual-card green"><h3>✨ Học liệu Sinh động</h3><p>Tóm tắt ngắn gọn, dễ nhớ theo từng bước nhỏ.</p></div>', unsafe_allow_html=True)
    with c_v3:
        st.markdown('<div class="visual-card amber"><h3>🧩 Đa Dạng Câu Hỏi</h3><p>Đúng/Sai, nhiều lựa chọn, trả lời ngắn & tự luận.</p></div>', unsafe_allow_html=True)
    with c_v4:
        st.markdown('<div class="visual-card purple"><h3>🔊 Âm thanh Vui nhộn</h3><p>Chúc mừng khi làm đúng & gợi ý nhẹ nhàng khi sai.</p></div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown('<div class="visual-card rose">', unsafe_allow_html=True)
        selected_role = st.selectbox(
            "📌 Chọn vai trò sử dụng hệ thống:",
            ["Giáo viên", "Học sinh"],
            key="login_role_selector"
        )

        with st.form("google_login_form"):
            st.markdown("<h3 style='text-align: center;'>🔐 Đăng nhập hệ thống</h3>", unsafe_allow_html=True)
            google_email = st.text_input("📧 Nhập email Google của bạn:", placeholder="nguyenhoangtuquyen@gmail.com")
            user_name = st.text_input("👤 Họ và tên hiển thị:", placeholder="Nhập tên của bạn...")

            st.markdown("---")
            st.markdown("📍 **Thông tin đơn vị / khu vực:**")

            province_input = st.text_input("1. Tỉnh / Thành phố:", placeholder="Ví dụ: Tỉnh Quảng Nam")
            ward_input = st.text_input("2. Phường / Xã / Thị trấn:", placeholder="Ví dụ: Phường An Mỹ")

            if selected_role == "Giáo viên":
                school_input = st.text_input("Tên trường THCS:", placeholder="Ví dụ: THCS Huỳnh Thúc Kháng")
                class_input = ""
            else:
                col_s1, col_s2 = st.columns(2)
                with col_s1:
                    school_input = st.text_input("Tên trường THCS:", placeholder="Ví dụ: THCS Huỳnh Thúc Kháng")
                with col_s2:
                    class_input = st.text_input("Lớp:", placeholder="Ví dụ: Lớp 9/3")

            submit_login = st.form_submit_button("🚀 Đăng nhập ngay", use_container_width=True)

            if submit_login:
                missing_school_or_class = (
                    not school_input.strip()
                    or (selected_role == "Học sinh" and not class_input.strip())
                )
                if not google_email.strip() or "@" not in google_email:
                    st.warning("⚠️ Vui lòng nhập địa chỉ email Google hợp lệ!")
                elif not user_name.strip():
                    st.warning("⚠️ Vui lòng nhập họ và tên của bạn!")
                elif not province_input.strip() or not ward_input.strip():
                    st.warning("⚠️ Vui lòng nhập đầy đủ Tỉnh/Thành phố và Phường/Xã!")
                elif missing_school_or_class:
                    if selected_role == "Giáo viên":
                        st.warning("⚠️ Vui lòng nhập đầy đủ Tên trường!")
                    else:
                        st.warning("⚠️ Vui lòng nhập đầy đủ Tên trường và Lớp!")
                else:
                    fresh_store = load_persistent_store()
                    st.session_state.student_profiles = fresh_store["student_profiles"]
                    st.session_state.reports = fresh_store["reports"]
                    st.session_state.assigned_tasks = fresh_store["assigned_tasks"]
                    st.session_state.student_progress = fresh_store["student_progress"]

                    st.session_state.logged_in = True
                    st.session_state.username = user_name.strip()
                    st.session_state.email = google_email.strip()
                    st.session_state.role = selected_role
                    st.session_state.province = province_input.strip()
                    st.session_state.ward = ward_input.strip()
                    st.session_state.school = school_input.strip()
                    st.session_state.class_name = class_input.strip() if selected_role == "Học sinh" else ""
                    play_feedback_sound("correct", trigger_id="login_ok")
                    st.success(f"🎉 Đăng nhập thành công với tài khoản: {google_email}")
                    st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    st.stop()

# =========================================================================
# GIAO DIỆN CHÍNH SAU KHI ĐĂNG NHẬP
# =========================================================================
st.sidebar.markdown(get_illustration_svg("hero"), unsafe_allow_html=True)
st.sidebar.markdown("### 👤 Xin chào!")
school_line = (
    f"🏫 **Trường:** {st.session_state.school} - {st.session_state.class_name}"
    if st.session_state.role == "Học sinh" and st.session_state.class_name
    else f"🏫 **Trường:** {st.session_state.school}"
)
st.sidebar.info(
    f"**{st.session_state.username}**\n\n"
    f"📧 `{st.session_state.email}`\n\n"
    f"📌 Vai trò: **{st.session_state.role}**\n\n"
    f"📍 **Khu vực:** {st.session_state.ward}, {st.session_state.province}\n\n"
    f"{school_line}"
)

st.session_state.sound_enabled = st.sidebar.toggle(
    "🔊 Âm thanh phản hồi đúng/sai",
    value=st.session_state.sound_enabled
)
st.session_state.auto_spellcheck_export = st.sidebar.toggle(
    "✍️ Tự động rà soát chính tả khi xuất file",
    value=st.session_state.auto_spellcheck_export
)

if st.sidebar.button("🚪 Đăng xuất hệ thống", use_container_width=True):
    save_persistent_store()
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.email = ""
    st.session_state.role = ""
    st.session_state.province = ""
    st.session_state.ward = ""
    st.session_state.school = ""
    st.session_state.class_name = ""
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### 🧭 Hệ thống Điều hướng")

if st.session_state.role == "Giáo viên":
    menu = st.sidebar.selectbox(
        "Chọn chức năng:",
        [
            "1. Trạm AI: Lập Kế hoạch IEP & Tối ưu Lý thuyết",
            "2. Quản lý Hồ sơ, Giao bài & Theo dõi Tiến bộ"
        ]
    )
else:
    menu = st.sidebar.selectbox(
        "Chọn chức năng:",
        [
            "1. Kho học liệu tương tác (Tóm tắt Lý thuyết Tự học)",
            "2. Làm bài tập giáo viên gửi & Đánh giá tiến bộ"
        ]
    )

# =========================================================================
# SỬ DỤNG @st.fragment CHO CÁC KHU VỰC ĐỂ TĂNG TỐC ĐỘ
# =========================================================================

@st.fragment
def render_teacher_iep_section():
    b_col1, b_col2 = st.columns([3.0, 1.0])
    with b_col1:
        st.markdown("""
            <div class="hero-banner">
                <h2>📋 TRẠM AI THÔNG MINH: IEP & CHUYỂN HÓA LÝ THUYẾT</h2>
                <p style="font-size: 19px; margin-top: 6px;">✨ Thiết kế kế hoạch cá nhân hóa, rút gọn bài học & xuất file gọn sạch</p>
            </div>
        """, unsafe_allow_html=True)
    with b_col2:
        st.markdown(get_illustration_svg("hero"), unsafe_allow_html=True)

    tab_iep, tab_theory = st.tabs(["📑 Lập Kế hoạch IEP theo File Mẫu", "✨ Tối ưu hóa Lý thuyết Khô khan"])

    with tab_iep:
        st.write("Giáo viên đính kèm file mẫu quy chuẩn, chọn bộ môn, dạng khuyết tật và nhập yêu cầu cần đạt.")
        uploaded_file = st.file_uploader(
            "📁 Đính kèm file mẫu IEP (.txt, .docx, .pdf):",
            type=["txt", "docx", "pdf"]
        )

        file_content_text = ""
        if uploaded_file is not None:
            try:
                file_content_text = extract_uploaded_doc_text(uploaded_file)
                st.success(f"✅ Đã tải lên thành công: **{uploaded_file.name}**")
            except Exception as e:
                st.error(f"❌ Lỗi đọc file: {e}")

        col1, col2 = st.columns(2)
        with col1:
            mon_hoc = st.selectbox(
                "Chọn bộ môn giảng dạy:",
                ["Môn Toán", "Môn Ngữ Văn", "Môn Khoa học Tự nhiên", "Môn Lịch sử & Địa lý", "Môn Tin học"]
            )
        with col2:
            dang_khuyet_tat = st.selectbox(
                "Dạng khuyết tật / Đối tượng học sinh:",
                [
                    "Khuyết tật trí tuệ nhẹ (chậm tiếp thu, khó tập trung, ghi nhớ ngắn hạn)",
                    "Khuyết tật vận động tay (khó thao tác chuột/bàn phím chính xác)",
                    "Khó khăn về nhận thức và ngôn ngữ"
                ]
            )

        yeu_cau_goc = st.text_area(
            "Nhập Yêu cầu cần đạt hoặc Tên bài học chính khóa:",
            placeholder="Ví dụ: Nêu khái niệm thuật toán và ví dụ thực tế..."
        )

        if st.button("🤖 AI tự động lập Kế hoạch IEP chi tiết", use_container_width=True):
            if not yeu_cau_goc.strip():
                st.warning("⚠️ Vui lòng nhập Yêu cầu cần đạt!")
            else:
                with st.spinner("🤖 AI đang phân tích và lập bản IEP chuyên sâu..."):
                    try:
                        req_clean = proofread_and_normalize_user_input(yeu_cau_goc)
                        prompt_text = f"""
                        Lập Kế hoạch IEP chi tiết cho môn {mon_hoc}, học sinh bị {dang_khuyet_tat}, yêu cầu cần đạt: {req_clean}. 
                        Dựa vào file mẫu: {file_content_text}. Trình bày rõ hành động cụ thể của giáo viên bộ môn bằng tiếng Việt chuẩn chính tả, định dạng Markdown.
                        """
                        ket_qua_ai = call_ai_cached(prompt_text)
                        st.session_state.iep_result = ket_qua_ai
                        play_feedback_sound("correct", trigger_id="iep_done")
                        st.success("🎉 Hoàn thành tạo Kế hoạch IEP tự động!")
                    except Exception as e:
                        st.error(f"❌ Lỗi AI: {e}")

        if st.session_state.iep_result:
            st.markdown('<div class="visual-card">', unsafe_allow_html=True)
            st.markdown(st.session_state.iep_result)
            st.markdown('</div>', unsafe_allow_html=True)
            if not DOCX_AVAILABLE:
                st.caption("💡 Gợi ý: Cài `pip install python-docx` để xuất chuẩn định dạng `.docx`.")
            export_clean_file(
                title_text=f"Kế hoạch IEP - {mon_hoc}",
                raw_body=st.session_state.iep_result,
                base_filename="Ke_hoach_IEP",
                label="📥 Xuất Kế hoạch IEP",
                key="dl_iep_clean",
                proofread_user_input=True
            )

    with tab_theory:
        st.subheader("✨ Biến đổi lý thuyết khô khan thành sinh động, dễ hiểu")
        st.write("Giáo viên có thể **dán nội dung lý thuyết**, **tải file bài học** hoặc **tải hình ảnh bài học** để AI cô đọng thành các bước ngắn và ví dụ trực quan.")

        theory_subject = st.selectbox(
            "Chọn môn học cho lý thuyết:",
            ["Môn Toán", "Môn Ngữ Văn", "Môn Khoa học Tự nhiên", "Môn Lịch sử & Địa lý", "Môn Tin học"],
            key="th_sub"
        )
        raw_theory = st.text_area(
            "1. Dán nội dung lý thuyết gốc cần tinh gọn:",
            placeholder="Dán đoạn lý thuyết dài trong sách giáo khoa vào đây..."
        )
        theory_file = st.file_uploader(
            "2. Hoặc tải file lý thuyết / bài học (.txt, .docx, .pdf):",
            type=["txt", "docx", "pdf"],
            key="th_file"
        )
        theory_image = st.file_uploader(
            "3. Hoặc tải hình ảnh trang SGK / bảng / sơ đồ bài học (.png, .jpg, .jpeg, .webp):",
            type=["png", "jpg", "jpeg", "webp"],
            key="th_img"
        )

        if st.button("🪄 AI Rút gọn & Trực quan hóa Bài học", use_container_width=True):
            extracted_th_file = extract_uploaded_doc_text(theory_file) if theory_file else ""
            has_th_input = bool(raw_theory.strip() or extracted_th_file.strip() or theory_image is not None)

            if not has_th_input:
                st.warning("⚠️ Vui lòng dán nội dung lý thuyết, tải file hoặc tải hình ảnh bài học!")
            else:
                with st.spinner("🪄 AI đang cô đọng và sinh động hóa kiến thức..."):
                    try:
                        th_clean = proofread_and_normalize_user_input(raw_theory)
                        prompt_theory = f"""
                        Bạn là chuyên gia giáo dục hòa nhập. Hãy chuyển đổi nội dung lý thuyết môn {theory_subject} sau đây (bao gồm văn bản dán vào, nội dung file và/hoặc hình ảnh đính kèm) thành dạng cực kỳ ngắn gọn, sinh động, dễ hiểu, dùng ví dụ thực tế gần gũi và chia nhỏ thành các bước (micro-learning) phù hợp với học sinh khuyết tật trí tuệ nhẹ/chậm tiếp thu, viết tiếng Việt chuẩn chính tả:
                        Nội dung dán vào: {th_clean}
                        Nội dung từ file: {extracted_th_file}
                        """
                        if theory_image is not None:
                            img_b = theory_image.getvalue()
                            m_type = theory_image.type or "image/jpeg"
                            optimized_theory = call_ai_multimodal(prompt_theory, image_bytes=img_b, mime_type=m_type)
                        else:
                            optimized_theory = call_ai_cached(prompt_theory)

                        st.session_state.theory_result = optimized_theory
                        play_feedback_sound("correct", trigger_id="theory_done")
                        st.success("🌟 Đã tối ưu hóa lý thuyết thành công!")
                    except Exception as e:
                        st.error(f"❌ Lỗi AI: {e}")

        if st.session_state.theory_result:
            st.markdown('<div class="visual-card green">', unsafe_allow_html=True)
            st.markdown(st.session_state.theory_result)
            st.markdown('</div>', unsafe_allow_html=True)
            export_clean_file(
                title_text=f"Lý thuyết tối ưu - {theory_subject}",
                raw_body=st.session_state.theory_result,
                base_filename="Ly_thuyet_toi_uu",
                label="📥 Xuất lý thuyết tối ưu",
                key="dl_theory_clean",
                proofread_user_input=True
            )


@st.fragment
def render_teacher_management_section():
    b_col1, b_col2 = st.columns([3.0, 1.0])
    with b_col1:
        st.markdown("""
            <div class="hero-banner">
                <h2>📊 QUẢN LÝ HỒ SƠ, GIAO BÀI KIỂM TRA ĐA DẠNG & THEO DÕI TIẾN BỘ</h2>
                <p style="font-size: 19px; margin-top: 6px;">🧩 Chọn dạng Trắc nghiệm/Tự luận & số câu, lọc theo trường/phường/tỉnh của GV, chọn lớp & học sinh giao bài</p>
            </div>
        """, unsafe_allow_html=True)
    with b_col2:
        st.markdown(get_illustration_svg("matching"), unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📁 Danh sách Hồ sơ Học sinh", "📤 Soạn & Giao Bài Kiểm Tra Đa Dạng", "📈 Báo cáo Tiến bộ & Lỗi sai từ AI"])

    with tab1:
        st.subheader("Thêm mới hồ sơ học sinh hòa nhập (tự động lưu sau khi đăng xuất/đăng nhập lại)")
        with st.form("add_student_form"):
            s_name = st.text_input("Họ và tên học sinh:")
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                s_province = st.text_input("Tỉnh / Thành phố:", value=st.session_state.province)
            with col_p2:
                s_ward = st.text_input("Phường / Xã / Thị trấn:", value=st.session_state.ward)
            col_a, col_b = st.columns(2)
            with col_a:
                s_school = st.text_input("Tên trường THCS:", value=st.session_state.school)
            with col_b:
                s_class = st.text_input("Lớp học:", placeholder="Ví dụ: Lớp 9/3")
            s_disability = st.selectbox("Dạng khuyết tật:", ["Khuyết tật trí tuệ nhẹ", "Khuyết tật vận động tay", "Khó khăn ngôn ngữ"])
            s_note = st.text_area("Ghi chú đặc điểm sư phạm:")

            submit_student = st.form_submit_button("💾 Lưu hồ sơ học sinh")
            if submit_student:
                if s_name and s_school and s_class:
                    st.session_state.student_profiles.append({
                        "name": s_name.strip(),
                        "province": s_province.strip() or st.session_state.province,
                        "ward": s_ward.strip() or st.session_state.ward,
                        "school": s_school.strip(),
                        "class": s_class.strip(),
                        "disability": s_disability,
                        "note": s_note.strip()
                    })
                    save_persistent_store()
                    play_feedback_sound("correct", trigger_id="saved_stu")
                    st.success(f"🎉 Đã lưu vĩnh viễn hồ sơ em **{s_name.strip()}** ({s_class.strip()}, {s_school.strip()})!")
                else:
                    st.warning("⚠️ Vui lòng điền đầy đủ Họ tên, Trường và Lớp!")

        st.markdown("### 📚 Danh sách học sinh đang quản lý (hỗ trợ chỉnh sửa & xóa trực tiếp):")
        if len(st.session_state.student_profiles) == 0:
            st.info("ℹ️ Chưa có hồ sơ học sinh nào. Vui lòng thêm mới phía trên.")
        else:
            disability_options = ["Khuyết tật trí tuệ nhẹ", "Khuyết tật vận động tay", "Khó khăn ngôn ngữ"]
            for s_idx, stu in enumerate(list(st.session_state.student_profiles)):
                with st.expander(f"👤 {stu['name']} - {stu['class']} - {stu['school']} ({stu['disability']})"):
                    st.write(f"📍 **Khu vực:** {stu.get('ward', '')}, {stu.get('province', '')}")
                    st.write(f"🏫 **Trường:** {stu['school']} | 📚 **Lớp:** {stu['class']}")
                    st.write(f"🧩 **Dạng khuyết tật:** {stu['disability']}")
                    st.write(f"📝 **Ghi chú:** {stu['note']}")

                    st.markdown("##### ✏️ Chỉnh sửa hồ sơ học sinh:")
                    with st.form(f"edit_student_form_{s_idx}"):
                        e_name = st.text_input("Họ và tên học sinh:", value=stu.get("name", ""), key=f"ename_{s_idx}")
                        ec1, ec2 = st.columns(2)
                        with ec1:
                            e_province = st.text_input("Tỉnh / Thành phố:", value=stu.get("province", ""), key=f"eprov_{s_idx}")
                        with ec2:
                            e_ward = st.text_input("Phường / Xã / Thị trấn:", value=stu.get("ward", ""), key=f"eward_{s_idx}")
                        ec3, ec4 = st.columns(2)
                        with ec3:
                            e_school = st.text_input("Tên trường THCS:", value=stu.get("school", ""), key=f"esch_{s_idx}")
                        with ec4:
                            e_class = st.text_input("Lớp học:", value=stu.get("class", ""), key=f"ecls_{s_idx}")

                        cur_dis = stu.get("disability", disability_options[0])
                        dis_idx = disability_options.index(cur_dis) if cur_dis in disability_options else 0
                        e_disability = st.selectbox("Dạng khuyết tật:", disability_options, index=dis_idx, key=f"edis_{s_idx}")
                        e_note = st.text_area("Ghi chú đặc điểm sư phạm:", value=stu.get("note", ""), key=f"enote_{s_idx}")

                        col_act1, col_act2 = st.columns(2)
                        with col_act1:
                            save_edit = st.form_submit_button("💾 Lưu chỉnh sửa", use_container_width=True)
                        with col_act2:
                            delete_stu = st.form_submit_button("🗑️ Xóa hồ sơ", use_container_width=True)

                        if save_edit:
                            if not e_name.strip() or not e_school.strip() or not e_class.strip():
                                st.warning("⚠️ Họ tên, Trường và Lớp không được để trống!")
                            else:
                                st.session_state.student_profiles[s_idx] = {
                                    "name": e_name.strip(),
                                    "province": e_province.strip() or st.session_state.province,
                                    "ward": e_ward.strip() or st.session_state.ward,
                                    "school": e_school.strip(),
                                    "class": e_class.strip(),
                                    "disability": e_disability,
                                    "note": e_note.strip()
                                }
                                save_persistent_store()
                                play_feedback_sound("correct", trigger_id=f"edited_stu_{s_idx}")
                                st.success(f"✅ Đã cập nhật hồ sơ học sinh **{e_name.strip()}** và lưu vào file!")
                                st.rerun()

                        if delete_stu:
                            removed_name = st.session_state.student_profiles[s_idx].get("name", "")
                            st.session_state.student_profiles.pop(s_idx)
                            save_persistent_store()
                            play_feedback_sound("wrong", trigger_id=f"del_stu_{s_idx}")
                            st.success(f"🗑️ Đã xóa hồ sơ học sinh **{removed_name}** và cập nhật file lưu trữ!")
                            st.rerun()

    with tab2:
        st.subheader("🧩 Tạo bài kiểm tra cho học sinh: chọn dạng, số câu, chọn lớp & học sinh cùng khu vực/trường")

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            exam_subject = st.selectbox(
                "Chọn môn học kiểm tra:",
                ["Môn Toán", "Môn Ngữ Văn", "Môn Khoa học Tự nhiên", "Môn Lịch sử & Địa lý", "Môn Tin học"],
                key="exam_sub"
            )
        with col_g2:
            exam_topic = st.text_input(
                "Chủ đề / Nội dung bài kiểm tra:",
                placeholder="Ví dụ: Phương trình bậc nhất, Quang hợp, Thuật toán..."
            )

        teacher_specs = render_question_builder_controls("teacher_exam_builder")

        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            if st.button("🤖 AI Sinh câu hỏi theo lựa chọn", use_container_width=True):
                if not teacher_specs:
                    st.warning("⚠️ Vui lòng chọn ít nhất 1 dạng câu hỏi và số câu lớn hơn 0!")
                else:
                    with st.spinner("🤖 AI đang tạo câu hỏi theo dạng và số lượng đã chọn..."):
                        try:
                            topic_clean = proofread_and_normalize_user_input(exam_topic) or f"Kiến thức trọng tâm {exam_subject}"
                            st.session_state.teacher_draft_questions = generate_ai_custom_questions(
                                exam_subject,
                                topic_clean,
                                teacher_specs
                            )
                            play_feedback_sound("correct", trigger_id="gen_exam_ai")
                            st.success(f"🎉 Đã tạo {len(st.session_state.teacher_draft_questions)} câu hỏi kiểm tra!")
                        except Exception:
                            topic_clean = proofread_and_normalize_user_input(exam_topic) or f"Kiến thức trọng tâm {exam_subject}"
                            st.session_state.teacher_draft_questions = get_sample_custom_questions(
                                exam_subject,
                                topic_clean,
                                teacher_specs
                            )
                            st.info("✨ Đã nạp bộ câu hỏi chuẩn theo lựa chọn của thầy/cô!")

        with c_btn2:
            if st.button("⚡ Nạp mẫu câu hỏi theo lựa chọn", use_container_width=True):
                if not teacher_specs:
                    st.warning("⚠️ Vui lòng chọn ít nhất 1 dạng câu hỏi và số câu lớn hơn 0!")
                else:
                    topic_clean = proofread_and_normalize_user_input(exam_topic) or f"Kiến thức trọng tâm {exam_subject}"
                    st.session_state.teacher_draft_questions = get_sample_custom_questions(
                        exam_subject,
                        topic_clean,
                        teacher_specs
                    )
                    st.success(f"✅ Đã nạp {len(st.session_state.teacher_draft_questions)} câu hỏi mẫu!")

        if st.session_state.teacher_draft_questions:
            st.markdown("#### 👀 Xem trước bộ câu hỏi đã tạo (bản giao cho học sinh sẽ ẩn hoàn toàn đáp án & giải thích):")
            render_interactive_question_list(
                st.session_state.teacher_draft_questions,
                ns_key="teacher_preview",
                student_name="",
                subject=exam_subject,
                exam_title=exam_topic or exam_subject,
                hide_answers=False
            )

            st.markdown("#### 📤 Chọn học sinh trong hồ sơ & Giao bài (không kèm đáp án)")
            if len(st.session_state.student_profiles) == 0:
                st.warning("⚠️ Chưa có học sinh nào trong hồ sơ! Vui lòng thêm học sinh ở tab 'Danh sách Hồ sơ Học sinh' trước khi giao bài.")
            else:
                gv_school = st.session_state.school.strip().lower()
                gv_ward = st.session_state.ward.strip().lower()
                gv_province = st.session_state.province.strip().lower()

                only_same_unit = st.checkbox(
                    f"📍 Chỉ hiển thị học sinh cùng Trường ({st.session_state.school}), Phường/Xã ({st.session_state.ward}), Tỉnh/TP ({st.session_state.province})",
                    value=True,
                    key="filter_same_unit_cb"
                )

                filtered_students = []
                for s in st.session_state.student_profiles:
                    s_sch = (s.get("school") or "").strip().lower()
                    s_wrd = (s.get("ward") or "").strip().lower()
                    s_prv = (s.get("province") or "").strip().lower()
                    same_match = (
                        (not gv_school or s_sch == gv_school)
                        and (not gv_ward or s_wrd == gv_ward)
                        and (not gv_province or s_prv == gv_province)
                    )
                    if not only_same_unit or same_match:
                        filtered_students.append(s)

                if len(filtered_students) == 0:
                    st.info("ℹ️ Hiện chưa có học sinh nào khớp chính xác Trường/Phường-Xã/Tỉnh-TP của giáo viên. Bỏ chọn ô lọc phía trên để xem toàn bộ hồ sơ đã lưu.")
                else:
                    available_classes = sorted(list({s.get("class", "").strip() for s in filtered_students if s.get("class", "").strip()}))
                    class_filter_options = ["Tất cả các lớp"] + available_classes
                    selected_class = st.selectbox(
                        "1. Chọn lớp muốn giao bài:",
                        options=class_filter_options,
                        key="select_class_for_exam"
                    )

                    class_students = (
                        filtered_students
                        if selected_class == "Tất cả các lớp"
                        else [s for s in filtered_students if s.get("class", "").strip() == selected_class]
                    )

                    student_options = {
                        f"{s['name']} — {s['class']} ({s['disability']})": s["name"]
                        for s in class_students
                    }

                    with st.form("assign_multi_exam_form"):
                        selected_labels = st.multiselect(
                            "2. Chọn tên học sinh muốn giao bài:",
                            options=list(student_options.keys()),
                            default=list(student_options.keys())
                        )
                        exam_title_inp = st.text_input(
                            "Tên bài kiểm tra / bài tập:",
                            value=exam_topic.strip() or f"Kiểm tra {exam_subject}"
                        )
                        submit_exam = st.form_submit_button("🚀 Giao bài cho học sinh đã chọn (không đáp án)", use_container_width=True)

                        if submit_exam:
                            if not selected_labels:
                                st.warning("⚠️ Vui lòng chọn ít nhất 1 học sinh để giao bài!")
                            else:
                                clean_title_val = proofread_and_normalize_user_input(exam_title_inp) or f"Kiểm tra {exam_subject}"
                                student_safe_questions = [
                                    sanitize_question_for_student(q)
                                    for q in st.session_state.teacher_draft_questions
                                ]
                                for lbl in selected_labels:
                                    stu_name = student_options[lbl]
                                    st.session_state.assigned_tasks.append({
                                        "task_id": f"task_{int(time.time() * 1000)}_{stu_name}",
                                        "student": stu_name,
                                        "subject": exam_subject,
                                        "title": clean_title_val,
                                        "content": f"Bài kiểm tra tương tác ({len(student_safe_questions)} câu) - {clean_title_val}",
                                        "questions": student_safe_questions,
                                        "hide_answers": True,
                                        "status": "Chưa nộp"
                                    })
                                save_persistent_store()
                                play_feedback_sound("correct", trigger_id="sent_exam")
                                sent_names = ", ".join(student_options[lbl] for lbl in selected_labels)
                                st.success(f"✅ Đã giao bài **{clean_title_val}** (ẩn hoàn toàn đáp án) cho: **{sent_names}**!")

            st.markdown("#### 📥 Xuất file bộ câu hỏi vừa tạo (tự động rà soát chính tả, không ký tự lạ)")
            exp_col1, exp_col2 = st.columns(2)
            with exp_col1:
                no_ans_text = format_questions_for_export(st.session_state.teacher_draft_questions, include_answers=False)
                export_clean_file(
                    title_text=f"Đề bài - {exam_topic or exam_subject}",
                    raw_body=no_ans_text,
                    base_filename="De_bai_khong_dap_an",
                    label="📄 Xuất đề bài không đáp án",
                    key="dl_no_ans_clean",
                    proofread_user_input=True
                )
            with exp_col2:
                with_ans_text = format_questions_for_export(st.session_state.teacher_draft_questions, include_answers=True)
                export_clean_file(
                    title_text=f"Đề bài và đáp án - {exam_topic or exam_subject}",
                    raw_body=with_ans_text,
                    base_filename="De_bai_co_dap_an",
                    label="📑 Xuất đề kèm đáp án và giải thích",
                    key="dl_with_ans_clean",
                    proofread_user_input=True
                )

        st.markdown("### 📋 Lịch sử các bài kiểm tra / bài tập đã giao:")
        if len(st.session_state.assigned_tasks) == 0:
            st.info("ℹ️ Chưa có bài tập/kiểm tra nào được giao.")
        else:
            for t in st.session_state.assigned_tasks:
                q_count = len(t.get("questions", []))
                extra_info = f" | 🧩 {q_count} câu (ẩn đáp án khi làm)" if q_count > 0 else ""
                st.info(
                    f"👤 **Học sinh:** {t['student']} | 📖 **Môn:** {t['subject']} | 📌 **Trạng thái:** {t['status']}{extra_info}\n\n"
                    f"**Nội dung:** {t['content']}"
                )

    with tab3:
        st.subheader("📈 Tổng hợp lỗi sai từng câu, phạm trù, lý do sai, cách làm đúng dạng & giao bài tương tự nhiều lần")
        if len(st.session_state.reports) == 0:
            st.info("ℹ️ Chưa có báo cáo tiến bộ & lỗi sai nào từ học sinh gửi về.")
        else:
            for r_idx, rep in enumerate(st.session_state.reports):
                stu_name = rep["student"]
                subj_name = rep["subject"]
                err_items = rep.get("error_items", [])
                remedial_qs = rep.get("remedial_questions", [])
                send_cnt = rep.get("remedial_send_count", 0)

                st.markdown('<div class="visual-card rose">', unsafe_allow_html=True)
                st.markdown(f"### 👤 Học sinh: **{stu_name}** | 📖 Môn: **{subj_name}**")
                st.markdown(
                    f"📊 **Tiến bộ:** Đã làm **{rep['total_learned']}** câu | "
                    f"Đúng: **{rep['correct_count']}** câu | "
                    f"Sai: **{rep['wrong_count']}** câu | "
                    f"Đã giao luyện lại: **{send_cnt}** lượt"
                )

                if err_items:
                    st.markdown("#### 🔍 Chi tiết từng câu sai:")
                    for err in err_items:
                        st.markdown(
                            f"- **Câu {err['q_number']}:** {err['question_text']}\n"
                            f"  - **Phạm trù kiến thức:** `{err['category']}`\n"
                            f"  - **Em trả lời:** `{err['user_answer']}` | **Đáp án chuẩn:** `{err['correct_answer']}`\n"
                            f"  - **Lý do sai:** {err['reason']}\n"
                            f"  - **Cách để học sinh làm đúng dạng này:** {err['how_to_fix']}"
                        )

                    if remedial_qs:
                        st.markdown("#### 🧩 Bộ bài tập tương tự tự động thiết kế theo từng câu sai:")
                        for rq_i, rq in enumerate(remedial_qs, start=1):
                            st.write(f"{rq_i}. *[{rq.get('category', subj_name)}]* {rq.get('question', '')}")

                        if st.button(
                            f"🔁 Giao bài tập tương tự cho {stu_name} làm lại (Lượt #{send_cnt + 1})",
                            key=f"send_remedial_{r_idx}_{stu_name}",
                            use_container_width=True
                        ):
                            next_attempt = send_cnt + 1
                            fresh_remedial_for_student = []
                            for err in err_items:
                                raw_q = err.get("raw_question") or err.get("similar_question") or {}
                                sim_q = build_similar_question_for_error(
                                    q=raw_q,
                                    q_number=err.get("q_number", 1),
                                    category=err.get("category", subj_name),
                                    subject=subj_name,
                                    attempt_num=next_attempt,
                                )
                                fresh_remedial_for_student.append(sanitize_question_for_student(sim_q))

                            rep["remedial_send_count"] = next_attempt
                            st.session_state.assigned_tasks.append({
                                "task_id": f"remedial_{r_idx}_{next_attempt}_{int(time.time() * 1000)}",
                                "student": stu_name,
                                "subject": subj_name,
                                "title": f"Luyện lại dạng sai Lượt #{next_attempt} ({subj_name})",
                                "content": f"Bài tập tương tự Lượt #{next_attempt} ({len(fresh_remedial_for_student)} câu) rèn lại các phạm trù em làm chưa đúng.",
                                "questions": fresh_remedial_for_student,
                                "hide_answers": True,
                                "status": "Chưa nộp"
                            })
                            save_persistent_store()
                            play_feedback_sound("correct", trigger_id=f"remedial_sent_{r_idx}_{next_attempt}")
                            st.success(f"✅ Đã giao bài tập tương tự Lượt #{next_attempt} (không đáp án) cho học sinh **{stu_name}**! Thầy/cô có thể tiếp tục bấm nút để giao thêm các lượt luyện lại khác.")

                export_clean_file(
                    title_text=f"Báo cáo lỗi sai và bài tập tương tự - {stu_name}",
                    raw_body=format_error_report_for_export(rep),
                    base_filename=f"Bao_cao_loi_sai_{stu_name}_{r_idx + 1}",
                    label=f"📥 Xuất báo cáo lỗi sai của {stu_name}",
                    key=f"dl_err_rep_{r_idx}_{stu_name}",
                    proofread_user_input=True
                )

                st.markdown('</div>', unsafe_allow_html=True)
                st.markdown("---")


@st.fragment
def render_student_learning_section():
    b_col1, b_col2 = st.columns([3.0, 1.0])
    with b_col1:
        st.markdown("""
            <div class="hero-banner">
                <h2>🎮 KHO HỌC LIỆU TƯƠNG TÁC & TẠO BÀI ÔN TẬP / KIỂM TRA</h2>
                <p style="font-size: 19px; margin-top: 6px;">📖 Tóm tắt dễ hiểu • 🧩 Chọn dạng Trắc nghiệm/Tự luận & số câu • 🔊 Âm thanh chúc mừng</p>
            </div>
        """, unsafe_allow_html=True)
    with b_col2:
        st.markdown(get_illustration_svg("multi_choice"), unsafe_allow_html=True)

    tab_sum, tab_quiz, tab_multi = st.tabs([
        "📖 Tóm tắt Lý thuyết Tự học",
        "✍️ Tạo Bài Ôn Tập / Kiểm Tra từ Văn bản / File / Ảnh",
        "🧩 Luyện Tập Tương Tác (Có Âm Thanh Phản Hồi)"
    ])

    with tab_sum:
        st.write("Em có thể nhập tên bài học hoặc đoạn lý thuyết SGK mà em muốn tìm hiểu. Hệ thống AI sẽ tự động cô đọng lại thành các ý chính ngắn gọn, súc tích và dễ nhớ nhất dành riêng cho em.")

        with st.form("student_learning_form"):
            topic_input = st.text_area("Nhập tên bài học hoặc nội dung lý thuyết em muốn tóm tắt:", placeholder="Ví dụ: Bài 5: Định lý Pytago trong môn Toán, hoặc nội dung về Quang hợp ở thực vật...")
            subject_choice = st.selectbox("Chọn môn học:", ["Môn Toán", "Môn Ngữ Văn", "Môn Khoa học Tự nhiên", "Môn Lịch sử & Địa lý", "Môn Tin học"])
            submit_learn = st.form_submit_button("✨ Tóm tắt ngắn gọn & Dễ hiểu")

            if submit_learn:
                if not topic_input.strip():
                    st.warning("⚠️ Vui lòng nhập tên bài học hoặc nội dung lý thuyết em muốn học!")
                else:
                    with st.spinner("✨ AI đang cô đọng kiến thức và soạn tóm tắt ngắn gọn cho em..."):
                        try:
                            topic_clean = proofread_and_normalize_user_input(topic_input)
                            prompt_summary = f"""
                            Bạn là trợ lý học tập thân thiện, chuyên hỗ trợ học sinh THCS. Hãy tóm tắt nội dung bài học hoặc lý thuyết sau đây về chủ đề '{topic_clean}' (thuộc {subject_choice}).

                            Yêu cầu trình bày bắt buộc:
1. **🎯 Cốt lõi bài học:** Tóm tắt trọng tâm bằng 3 đến 4 ý chính ngắn gọn nhất (dùng gạch đầu dòng).
2. **💡 Giải thích dễ hiểu:** Dùng câu từ gần gũi, ví dụ thực tế đơn giản để học sinh dễ hình dung.
3. **📝 Từ khóa cần nhớ:** Liệt kê 3 từ khóa quan trọng nhất mà em cần ghi nhớ trong bài này.

                            Trình bày bằng định dạng Markdown sinh động, rõ ràng, viết tiếng Việt chuẩn chính tả.
                            """
                            summary_result = call_ai_cached(prompt_summary)
                            play_feedback_sound("correct", trigger_id="sum_done")
                            st.success("🎉 Đã tóm tắt xong nội dung học tập dành riêng cho em!")
                            st.markdown('<div class="visual-card green">', unsafe_allow_html=True)
                            st.markdown(summary_result)
                            st.markdown('</div>', unsafe_allow_html=True)
                            export_clean_file(
                                title_text=f"Tóm tắt lý thuyết - {subject_choice}",
                                raw_body=summary_result,
                                base_filename=f"Tom_tat_{subject_choice}",
                                label="📥 Xuất bản tóm tắt",
                                key="dl_summary_clean",
                                proofread_user_input=True
                            )

                        except Exception as e:
                            st.error(f"❌ Lỗi khi kết nối AI: {e}")

    with tab_quiz:
        st.write("Em có thể nhập nội dung bài học, tải **file bài học** hoặc tải **hình ảnh bài học**, chọn **dạng câu hỏi** và **số lượng câu** để tạo bài ôn tập / bài kiểm tra kèm **giải thích từng câu**.")

        quiz_subject = st.selectbox(
            "Chọn môn học:",
            ["Môn Toán", "Môn Ngữ Văn", "Môn Khoa học Tự nhiên", "Môn Lịch sử & Địa lý", "Môn Tin học"],
            key="quiz_sub"
        )
        lesson_text_input = st.text_area(
            "1. Nhập nội dung bài học / chủ đề cần ôn tập hoặc kiểm tra:",
            placeholder="Ví dụ: Tế bào nhân thực, Phương trình bậc nhất một ẩn, hoặc dán nội dung bài học..."
        )
        lesson_file = st.file_uploader(
            "2. Hoặc đính kèm file bài học (.txt, .docx, .pdf):",
            type=["txt", "docx", "pdf"],
            key="quiz_file"
        )
        lesson_image = st.file_uploader(
            "3. Hoặc tải lên hình ảnh bài học / trang vở / SGK (.png, .jpg, .jpeg, .webp):",
            type=["png", "jpg", "jpeg", "webp"],
            key="quiz_img"
        )

        quiz_specs = render_question_builder_controls("student_review_builder")

        if st.button("📝 Tạo Bài Ôn Tập / Kiểm Tra & Giải thích chi tiết", use_container_width=True):
            extracted_file_text = extract_uploaded_doc_text(lesson_file) if lesson_file else ""
            has_any_input = bool(lesson_text_input.strip() or extracted_file_text.strip() or lesson_image is not None)

            if not has_any_input:
                st.warning("⚠️ Vui lòng nhập nội dung bài học, tải file bài học hoặc tải hình ảnh bài học!")
            elif not quiz_specs:
                st.warning("⚠️ Vui lòng chọn ít nhất 1 dạng câu hỏi và số câu lớn hơn 0!")
            else:
                with st.spinner("📝 AI đang đọc tài liệu/hình ảnh và soạn bài ôn tập/kiểm tra theo lựa chọn của em..."):
                    try:
                        text_clean = proofread_and_normalize_user_input(lesson_text_input)
                        req_lines = "\n".join(f"- {cnt} câu dạng '{qtype}'" for qtype, cnt in quiz_specs)
                        prompt_quiz = f"""
                        Bạn là gia sư THCS tận tâm, dễ hiểu. Dựa trên nội dung bài học môn {quiz_subject} dưới đây (bao gồm văn bản, nội dung file, và/hoặc hình ảnh bài học đính kèm), hãy tạo bài ôn tập / kiểm tra cho học sinh theo đúng số lượng và dạng sau:

                        {req_lines}

                        Nội dung nhập: {text_clean}
                        Nội dung từ file: {extracted_file_text}

                        Yêu cầu bắt buộc:
                        - Tuân thủ chính xác số lượng từng dạng câu hỏi đã liệt kê ở trên.
                        - Với trắc nghiệm Đúng/Sai: đưa ra phát biểu rõ ràng để học sinh xác định Đúng hay Sai.
                        - Với trắc nghiệm Nhiều lựa chọn: có các phương án để chọn nhiều ý đúng.
                        - Với trắc nghiệm Trả lời ngắn: câu hỏi điền từ/cụm từ hoặc đáp án ngắn gọn.
                        - Với Tự luận: câu hỏi tự luận từ cơ bản đến vận dụng nhẹ.
                        - Ghi rõ **Đáp án đúng / Hướng dẫn giải** và **Giải thích chi tiết từng câu** bằng tiếng Việt chuẩn chính tả, dễ hiểu cho học sinh THCS.
                        - Trình bày rõ ràng bằng Markdown.
                        """

                        if lesson_image is not None:
                            img_bytes = lesson_image.getvalue()
                            mime_img = lesson_image.type or "image/jpeg"
                            quiz_out = call_ai_multimodal(prompt_quiz, image_bytes=img_bytes, mime_type=mime_img)
                        else:
                            quiz_out = call_ai_cached(prompt_quiz)

                        st.session_state.practice_quiz_result = quiz_out
                        play_feedback_sound("correct", trigger_id="quiz_gen_ok")
                        st.success("🎉 Đã tạo xong bài ôn tập / kiểm tra và giải thích chi tiết từng câu!")
                    except Exception as e:
                        st.error(f"❌ Lỗi khi tạo bài ôn tập / kiểm tra: {e}")

        if st.session_state.practice_quiz_result:
            st.markdown('<div class="visual-card amber">', unsafe_allow_html=True)
            st.markdown(st.session_state.practice_quiz_result)
            st.markdown('</div>', unsafe_allow_html=True)
            export_clean_file(
                title_text=f"Bài ôn tập và kiểm tra - {quiz_subject}",
                raw_body=st.session_state.practice_quiz_result,
                base_filename=f"Bai_on_tap_{quiz_subject}",
                label="📥 Xuất bài ôn tập",
                key="dl_student_quiz_clean",
                proofread_user_input=True
            )

    with tab_multi:
        st.write("Chọn dạng **Trắc nghiệm** (Đúng/Sai, Nhiều lựa chọn, Trả lời ngắn) hoặc **Tự luận**, chọn **số câu** và làm bài tương tác trực tiếp có **âm thanh phản hồi**!")

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            prac_sub = st.selectbox(
                "Chọn môn luyện tập:",
                ["Môn Toán", "Môn Ngữ Văn", "Môn Khoa học Tự nhiên", "Môn Lịch sử & Địa lý", "Môn Tin học"],
                key="prac_sub_4d"
            )
        with col_m2:
            prac_topic = st.text_input(
                "1. Nhập chủ đề hoặc nội dung bài học em muốn tạo câu hỏi:",
                placeholder="Ví dụ: Định lý Pytago, Quang hợp, Thuật toán..."
            )

        prac_file = st.file_uploader(
            "2. Tải file nội dung bài học (.txt, .docx, .pdf):",
            type=["txt", "docx", "pdf"],
            key="prac_file_4d"
        )
        prac_img = st.file_uploader(
            "3. Tải hình ảnh bài học / trang vở / SGK (.png, .jpg, .jpeg, .webp):",
            type=["png", "jpg", "jpeg", "webp"],
            key="prac_img_4d"
        )

        prac_specs = render_question_builder_controls("student_interactive_builder")

        b1, b2 = st.columns(2)
        with b1:
            if st.button("🪄 Tạo bài luyện tập theo lựa chọn", use_container_width=True):
                if not prac_specs:
                    st.warning("⚠️ Vui lòng chọn ít nhất 1 dạng câu hỏi và số câu lớn hơn 0!")
                else:
                    extracted_prac_file = extract_uploaded_doc_text(prac_file) if prac_file else ""
                    with st.spinner("🪄 Đang đọc nội dung/file/hình ảnh và chuẩn bị bộ câu hỏi..."):
                        try:
                            prac_topic_clean = proofread_and_normalize_user_input(prac_topic) or f"Ôn tập {prac_sub}"
                            img_b = prac_img.getvalue() if prac_img else None
                            m_type = (prac_img.type or "image/jpeg") if prac_img else "image/jpeg"
                            st.session_state.student_practice_questions = generate_ai_custom_questions(
                                prac_sub,
                                prac_topic_clean,
                                prac_specs,
                                extra_file_text=extracted_prac_file,
                                image_bytes=img_b,
                                mime_type=m_type
                            )
                        except Exception:
                            prac_topic_clean = proofread_and_normalize_user_input(prac_topic) or f"Ôn tập {prac_sub}"
                            st.session_state.student_practice_questions = get_sample_custom_questions(
                                prac_sub,
                                prac_topic_clean,
                                prac_specs
                            )
                        play_feedback_sound("correct", trigger_id="prac_ready")
                        st.success(f"🎉 Đã sẵn sàng {len(st.session_state.student_practice_questions)} câu hỏi! Bắt đầu làm bài nhé!")

        with b2:
            if st.button("🎲 Tải nhanh bộ câu hỏi mẫu theo lựa chọn", use_container_width=True):
                if not prac_specs:
                    st.warning("⚠️ Vui lòng chọn ít nhất 1 dạng câu hỏi và số câu lớn hơn 0!")
                else:
                    prac_topic_clean = proofread_and_normalize_user_input(prac_topic) or f"Ôn tập {prac_sub}"
                    st.session_state.student_practice_questions = get_sample_custom_questions(
                        prac_sub,
                        prac_topic_clean,
                        prac_specs
                    )
                    play_feedback_sound("correct", trigger_id="prac_sample_ready")
                    st.success(f"✅ Đã tải {len(st.session_state.student_practice_questions)} câu hỏi mẫu!")

        if st.session_state.student_practice_questions:
            render_interactive_question_list(
                st.session_state.student_practice_questions,
                ns_key="student_practice_4d",
                student_name=st.session_state.username,
                subject=prac_sub,
                exam_title=prac_topic or f"Luyện tập {prac_sub}",
                hide_answers=False
            )
            export_clean_file(
                title_text=f"Bộ câu hỏi luyện tập - {prac_topic or prac_sub}",
                raw_body=format_questions_for_export(st.session_state.student_practice_questions, include_answers=True),
                base_filename=f"Bo_cau_hoi_luyen_tap_{prac_sub}",
                label="📥 Xuất bộ câu hỏi luyện tập",
                key="dl_student_prac_qs_clean",
                proofread_user_input=True
            )


@st.fragment
def render_student_tasks_section():
    b_col1, b_col2 = st.columns([3.0, 1.0])
    with b_col1:
        st.markdown("""
            <div class="hero-banner">
                <h2>📝 LÀM BÀI TẬP & BÀI KIỂM TRA GIÁO VIÊN GỬI</h2>
                <p style="font-size: 19px; margin-top: 6px;">🔒 Mỗi câu chỉ chọn 1 lần & ẩn đáp án • 📊 Báo cáo tiến bộ lưu bền vững sau khi đăng xuất/đăng nhập lại</p>
            </div>
        """, unsafe_allow_html=True)
    with b_col2:
        st.markdown(get_illustration_svg("short_answer"), unsafe_allow_html=True)

    tab_hs1, tab_hs2 = st.tabs(["1️⃣ Làm bài tập / kiểm tra & Xem báo cáo lỗi sai", "2️⃣ Đánh giá tiến bộ cá nhân"])

    with tab_hs1:
        st.write("Dưới đây là các bài tập/bài kiểm tra giáo viên đã giao riêng cho em (không hiển thị đáp án, mỗi câu chỉ chọn 1 lần). Sau khi nộp, hệ thống sẽ hiện báo cáo lỗi sai & tiến bộ của em và lưu cho cả học sinh và giáo viên.")

        my_tasks = [t for t in st.session_state.assigned_tasks if t["student"].lower() == st.session_state.username.lower()]

        if len(my_tasks) == 0:
            st.info(f"ℹ️ Hiện tại không có bài tập nào được giao riêng cho tên tài khoản **{st.session_state.username}**. (Lưu ý: Tên đăng nhập của em phải khớp với tên giáo viên đã lưu trong hồ sơ).")
        else:
            for idx, task in enumerate(my_tasks):
                task_uid = task.get("task_id") or f"task_{idx}"
                with st.expander(f"📚 {task.get('title', 'Bài tập')} - Môn {task['subject']} | Trạng thái: {task['status']}", expanded=True):
                    st.write(f"**Yêu cầu của giáo viên:** {task['content']}")

                    if task.get("questions"):
                        def _mark_task_done(t_ref=task):
                            t_ref["status"] = "Đã hoàn thành"
                            save_persistent_store()

                        render_interactive_question_list(
                            task["questions"],
                            ns_key=f"assigned_exam_{task_uid}",
                            student_name=st.session_state.username,
                            subject=task["subject"],
                            exam_title=task.get("title", task["subject"]),
                            hide_answers=True,
                            show_report_on_finish=True,
                            report_state_key=f"post_report_{task_uid}",
                            on_submitted_callback=_mark_task_done
                        )
                    else:
                        essay_lock_key = f"essay_locked_{task_uid}"
                        is_essay_locked = bool(st.session_state.get(essay_lock_key, False))

                        student_ans = st.text_area(
                            "Bài làm / Đáp án của em:",
                            key=f"ans_{task_uid}",
                            disabled=is_essay_locked,
                            placeholder="Nhập đáp án hoặc câu trả lời của em vào đây..."
                        )
                        is_correct_input = st.selectbox(
                            "Em tự đánh giá bài làm của mình thế nào?",
                            ["Chưa rõ / Cần kiểm tra lại", "Em chắc chắn đúng", "Em cảm thấy còn phân vân / có thể sai"],
                            key=f"eval_{task_uid}",
                            disabled=is_essay_locked
                        )

                        if st.button(
                            "🔒 Đã khóa bài nộp" if is_essay_locked else "📤 Nộp bài & Hiện báo cáo lỗi sai",
                            key=f"btn_{task_uid}",
                            disabled=is_essay_locked
                        ):
                            if not student_ans.strip():
                                st.warning("⚠️ Vui lòng nhập câu trả lời trước khi nộp!")
                            else:
                                st.session_state[essay_lock_key] = True
                                task["status"] = "Đã nộp"
                                is_wrong = ("phân vân" in is_correct_input or "Chưa rõ" in is_correct_input)
                                wrong_count = 1 if is_wrong else 0
                                correct_count = 0 if wrong_count > 0 else 1

                                play_feedback_sound("correct" if wrong_count == 0 else "wrong", trigger_id=f"sub_task_{task_uid}")
                                st.success("🎉 Nộp bài thành công! Báo cáo lỗi sai & tiến bộ đã được lưu và gửi về giáo viên.")

                                clean_student_ans = proofread_and_normalize_user_input(student_ans)
                                error_items = []
                                mock_q = {
                                    "id": f"{task_uid}_q1",
                                    "type": "essay",
                                    "type_label": "📝 Bài tập tự luận / vận dụng",
                                    "category": f"{task['subject']} — {task.get('title', 'Bài tập giao')}",
                                    "question": task["content"],
                                    "sample_answer": "Trình bày đủ các bước giải và nêu ví dụ hoặc kết quả kiểm tra lại.",
                                    "explanation": "Kiểm tra lại từng bước biến đổi và kết luận trước khi nộp."
                                }
                                if is_wrong:
                                    error_items.append(
                                        build_error_analysis_and_remedial(
                                            q=mock_q,
                                            q_number=1,
                                            user_val=clean_student_ans,
                                            subject=task["subject"],
                                            exam_title=task.get("title", task["subject"]),
                                        )
                                    )

                                rep_obj = submit_error_report_to_teacher(
                                    student_name=st.session_state.username,
                                    subject=task["subject"],
                                    exam_title=task.get("title", task["subject"]),
                                    total_q=1,
                                    correct_q=correct_count,
                                    error_items=error_items,
                                )
                                st.session_state[f"post_report_{task_uid}"] = rep_obj
                                save_persistent_store()
                                st.rerun()

                        if st.session_state.get(f"post_report_{task_uid}"):
                            render_student_post_submit_report(st.session_state[f"post_report_{task_uid}"])

    with tab_hs2:
        st.markdown("## 📈 Đánh giá Tiến bộ Cá nhân")
        st.write("Theo dõi tổng quan số câu em đã học, số câu làm đúng và số câu sai (được lưu lại kể cả khi đăng xuất/đăng nhập lại).")

        my_prog = st.session_state.student_progress.get(st.session_state.username, {"total": 0, "correct": 0, "wrong": 0})
        my_reports = [r for r in st.session_state.reports if r.get("student", "").lower() == st.session_state.username.lower()]

        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            st.metric(label="📚 Tổng số câu đã học/làm", value=my_prog["total"])
        with col_p2:
            st.metric(label="✅ Số câu làm đúng", value=my_prog["correct"])
        with col_p3:
            st.metric(label="❌ Số câu làm sai", value=my_prog["wrong"])

        if my_prog["total"] > 0:
            accuracy_rate = (my_prog["correct"] / my_prog["total"]) * 100
            st.progress(accuracy_rate / 100)
            st.info(f"🌟 Tỷ lệ làm đúng của em đạt: **{accuracy_rate:.1f}%**. Hãy tiếp tục phát huy nhé!")
        else:
            st.warning("ℹ️ Em chưa hoàn thành bài tập nào. Hãy sang tab 1 để nhận và làm bài tập từ giáo viên!")

        if my_reports:
            st.markdown("### 🗂️ Lịch sử báo cáo tiến bộ & lỗi sai đã lưu của em:")
            for r_idx, rep in enumerate(reversed(my_reports)):
                st.markdown(f"**Lần #{len(my_reports) - r_idx} — {rep.get('exam_title', rep.get('subject', ''))} ({rep.get('subject', '')}):** Đúng **{rep.get('correct_count', 0)}/{rep.get('total_learned', 0)}** câu, Sai **{rep.get('wrong_count', 0)}** câu.")
                for err in rep.get("error_items", []):
                    st.write(f"  - Câu {err.get('q_number')}: *{err.get('category')}* — Lý do: {err.get('reason')} | Cách làm đúng: {err.get('how_to_fix')}")


# --- GỌI CÁC PHÂN HỆ DỰA TRÊN ROLE VÀ MENU ---
if st.session_state.role == "Giáo viên":
    if menu == "1. Trạm AI: Lập Kế hoạch IEP & Tối ưu Lý thuyết":
        render_teacher_iep_section()
    else:
        render_teacher_management_section()
else:
    if menu == "1. Kho học liệu tương tác (Tóm tắt Lý thuyết Tự học)":
        render_student_learning_section()
    else:
        render_student_tasks_section()