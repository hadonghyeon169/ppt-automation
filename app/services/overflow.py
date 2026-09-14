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
