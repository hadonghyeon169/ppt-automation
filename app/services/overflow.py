"""
넘침(overflow) 방지 보정 — 스킬 문서 "넘침 방지 및 레이아웃 보정" 절 이식.

배포 환경에 실제 폰트(Noto Sans, Sarabun 등)가 없을 수 있으므로, PIL의 기본
DejaVu Sans로 근사 폭을 구해 "넘치는지 여부"와 "몇 배 넘치는지"를 판단한다.
완벽한 폭 계산이 목적이 아니라, wrap="none" 도형의 cx(너비)를 얼마나 늘릴지
또는 wrap="square"로 전환할지를 정하기 위한 근사치다.
"""
import os
from PIL import ImageFont

EMU_PER_INCH = 914400
EMU_PER_PT = EMU_PER_INCH / 72.0

_FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]


def _find_fallback_font():
    for p in _FONT_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


_FALLBACK_FONT_PATH = _find_fallback_font()


def estimate_text_width_emu(text, font_size_pt):
    """DejaVu Sans 근사치로 텍스트 폭을 EMU 단위로 추정."""
    if not text:
        return 0
    if _FALLBACK_FONT_PATH:
        try:
            px_size = max(int(font_size_pt * 4 / 3), 1)  # pt -> px 근사 (96dpi 가정)
            font = ImageFont.truetype(_FALLBACK_FONT_PATH, px_size)
            bbox = font.getbbox(text)
            width_px = bbox[2] - bbox[0]
            width_pt = width_px * 72.0 / 96.0
            return int(width_pt * EMU_PER_PT)
        except Exception:
            pass
    # 폴백: 평균 문자폭을 font_size의 0.55배로 근사
    avg_char_width_pt = font_size_pt * 0.55
    return int(len(text) * avg_char_width_pt * EMU_PER_PT)


def compute_overflow_ratio(text, font_size_pt, shape_cx_emu):
    if not shape_cx_emu:
        return 1.0
    width = estimate_text_width_emu(text, font_size_pt)
    return width / shape_cx_emu


LINE_SPACING_FACTOR = 1.25  # 줄간격 근사치 (폰트 크기 대비 배수)
USABLE_WIDTH_RATIO = 0.94   # 도형 내부 여백(lIns/rIns) 근사 차감
USABLE_HEIGHT_RATIO = 0.88  # 도형 내부 여백(tIns/bIns) + 줄바꿈 근사오차 대비 차감


def estimate_line_count(text, font_size_pt, box_cx_emu):
    """word-wrap 도형에서 텍스트가 실제로 몇 줄로 감길지 근사한다.
    명시적 개행(\\n)은 각각 최소 한 줄로 세고, 각 줄 안에서는 폭 기준으로 자동
    줄바꿈되는 횟수를 ceil(텍스트폭 / 도형폭)으로 근사한다."""
    if not text:
        return 1
    usable_cx = max(int(box_cx_emu * USABLE_WIDTH_RATIO), 1)
    total_lines = 0
    for raw_line in text.split("\n"):
        if not raw_line.strip():
            total_lines += 1
            continue
        width = estimate_text_width_emu(raw_line, font_size_pt)
        total_lines += max(1, -(-width // usable_cx))  # ceil division
    return max(1, total_lines)


def estimate_max_word_width_emu(text, font_size_pt):
    """텍스트 안에서 가장 긴 '단어'(공백 기준 토큰)의 폭을 추정한다.
    word-wrap은 단어 경계에서만 줄바꿈되므로, 특정 단어 하나의 폭이 박스 폭보다
    크면 줄 수를 아무리 잘 배분해도 그 단어가 있는 줄은 박스 옆으로 삐져나간다.
    estimate_line_count()는 "총 텍스트 폭 / 박스 폭"을 올림 나눗셈해서 줄 수만
    근사하기 때문에, 이런 긴 단어 하나 때문에 생기는 가로 넘침은 잡아내지 못한다
    (총 높이는 맞아 보이지만 특정 줄만 옆으로 튀어나오는 경우)."""
    max_width = 0
    for raw_line in (text or "").split("\n"):
        for word in raw_line.split():
            w = estimate_text_width_emu(word, font_size_pt)
            if w > max_width:
                max_width = w
    return max_width


def fit_font_size_to_box(text_or_lines, requested_font_size_pt, box_cx_emu, box_cy_emu,
                          min_font_size_pt=10, line_spacing=LINE_SPACING_FACTOR):
    """word-wrap 도형(wrap != "none")용: 도형 실제 높이(cy)에 텍스트가 들어가고,
    동시에 가장 긴 단어 하나가 박스 폭을 벗어나지 않을 때까지 폰트 크기를 1pt씩
    낮춘다. AI가 제안한 force_font_size_pt는 이 PPT 템플릿의 실제 도형 크기를
    모른 채 나온 "권장값"일 뿐이므로, 여기서 실제 크기 기준으로 최종 검증/보정한다.

    높이만 보고 폭의 개별 단어 넘침을 확인하지 않으면, 총 줄 수 계산상으로는
    "들어간다"고 나와도 긴 단어 하나가 있는 줄만 박스 옆으로 삐져나가는 경우를
    info로 잘못 분류하게 된다(실제 사례: 러시아어 긴 단어가 포함된 도형에서
    글자가 상자 옆/밖으로 넘어가는데도 자동 축소가 "성공"으로 처리됨).

    text_or_lines: 문자열 하나 또는 문단 리스트(문단은 줄바꿈으로 취급).
    반환: (fitted_font_size_pt, overflow_unresolved: bool)
    overflow_unresolved=True면 min_font_size_pt까지 낮춰도 (높이 또는 폭이) 넘친다는
    뜻 (수동 확인 필요)."""
    if not box_cx_emu or not box_cy_emu:
        return requested_font_size_pt, False
    if isinstance(text_or_lines, (list, tuple)):
        full_text = "\n".join(t for t in text_or_lines if t)
    else:
        full_text = text_or_lines or ""
    if not full_text.strip():
        return requested_font_size_pt, False

    usable_cy = max(int(box_cy_emu * USABLE_HEIGHT_RATIO), 1)
    usable_cx = max(int(box_cx_emu * USABLE_WIDTH_RATIO), 1)
    size = requested_font_size_pt
    while size >= min_font_size_pt:
        lines = estimate_line_count(full_text, size, box_cx_emu)
        est_height = lines * size * line_spacing * EMU_PER_PT
        word_width = estimate_max_word_width_emu(full_text, size)
        if est_height <= usable_cy and word_width <= usable_cx:
            return size, False
        size -= 1
    return min_font_size_pt, True


def fit_font_size_to_width(text, requested_font_size_pt, max_cx_emu, min_font_size_pt=10):
    """wrap="none"(줄바꿈 없음, 한 줄) 도형용: 박스를 슬라이드 폭 한도(max_cx_emu)까지
    넓혀도 텍스트가 다 안 들어가면 폰트를 줄인다. suggest_independent_shape_resize는
    박스를 최대 max_width_ratio(기본 92%)까지만 넓히므로, 그래도 넘치는 긴 번역문은
    이전엔 아무 보정 없이 슬라이드 밖으로 삐져나갔다."""
    if not text or not max_cx_emu:
        return requested_font_size_pt, False
    size = requested_font_size_pt
    while size >= min_font_size_pt:
        width = estimate_text_width_emu(text, size)
        if width <= max_cx_emu:
            return size, False
        size -= 1
    return min_font_size_pt, True


def suggest_independent_shape_resize(text, font_size_pt, xfrm, align, slide_width_emu, max_width_ratio=0.92):
    """독립적인 제목/안내문 도형: 폭을 늘리고 정렬 기준으로 x를 재조정."""
    est_width = estimate_text_width_emu(text, font_size_pt)
    max_cx = int(slide_width_emu * max_width_ratio)
    new_cx = min(max(est_width, xfrm['cx']), max_cx)

    old_x, old_cx = xfrm['x'], xfrm['cx']
    if align == 'ctr':
        center = old_x + old_cx / 2
        new_x = int(center - new_cx / 2)
    elif align == 'r':
        right = old_x + old_cx
        new_x = int(right - new_cx)
    else:  # 좌측 정렬 기본값
        new_x = old_x

    new_x = max(0, min(new_x, slide_width_emu - new_cx))
    return {"x": new_x, "cx": new_cx}
