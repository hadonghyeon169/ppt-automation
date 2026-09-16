"""
넘침(overflow) 방지 보정 — 스킬 문서 "넘침 방지 및 레이아웃 보정" 절 이식.

배포 환경에 실제 폰트(Noto Sans, Sarabun 등)가 없을 수 있으므로, PIL의 기본
DejaVu Sans로 근사 폭을 구해 "넘치는지 여부"와 "몇 배 넘치는지"를 판단한다.
완벽한 폭 계산이 목적이 아니라, wrap="none" 도형의 cx(너비)를 얼마나 늘릴지
또는 wrap="square"로 전환할지를 정하기 위한 근사치다.
"""
import os
import re
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


def compute_overflow_ratio(text, font_size_pt, shape_cx_emu, insets_emu=0):
    """도형 cx(외곽 폭) 전체가 아니라, 텍스트 상자 내부 여백(lIns+rIns=insets_emu)을
    뺀 실사용 가능 폭 기준으로 넘침 비율을 계산한다. insets_emu를 빼지 않으면
    실제로는 넘치는 텍스트(예: 좌우 여백이 각 0.15in인 도형)도 비율이 1.05 기준선
    바로 아래로 나와 넘침 보정이 걸리지 않는 문제가 있었다."""
    usable_cx = shape_cx_emu - insets_emu
    if not usable_cx or usable_cx <= 0:
        return 1.0
    width = estimate_text_width_emu(text, font_size_pt)
    return width / usable_cx


LINE_SPACING_FACTOR = 1.25  # 줄간격 근사치 (폰트 크기 대비 배수)
USABLE_WIDTH_RATIO = 0.94   # 도형 내부 여백(lIns/rIns) 근사 차감
USABLE_HEIGHT_RATIO = 0.88  # 도형 내부 여백(tIns/bIns) + 줄바꿈 근사오차 대비 차감


def estimate_line_count(text, font_size_pt, box_cx_emu, usable_cx_emu=None):
    """word-wrap 도형에서 텍스트가 실제로 몇 줄로 감길지 근사한다.
    명시적 개행(\\n)은 각각 최소 한 줄로 세고, 각 줄 안에서는 폭 기준으로 자동
    줄바꿈되는 횟수를 ceil(텍스트폭 / 도형폭)으로 근사한다.
    usable_cx_emu를 넘기면 고정 비율(USABLE_WIDTH_RATIO) 대신 그 값을 실사용
    가능 폭으로 쓴다 (도형의 실제 lIns/rIns 기반 — fit_font_size_to_box 참고)."""
    if not text:
        return 1
    usable_cx = max(int(usable_cx_emu if usable_cx_emu is not None else box_cx_emu * USABLE_WIDTH_RATIO), 1)
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
                          min_font_size_pt=10, line_spacing=LINE_SPACING_FACTOR,
                          insets_lr_emu=None, insets_tb_emu=None):
    """word-wrap 도형(wrap != "none")용: 도형 실제 높이(cy)에 텍스트가 들어가고,
    동시에 가장 긴 단어 하나가 박스 폭을 벗어나지 않을 때까지 폰트 크기를 1pt씩
    낮춘다. AI가 제안한 force_font_size_pt는 이 PPT 템플릿의 실제 도형 크기를
    모른 채 나온 "권장값"일 뿐이므로, 여기서 실제 크기 기준으로 최종 검증/보정한다.

    높이만 보고 폭의 개별 단어 넘침을 확인하지 않으면, 총 줄 수 계산상으로는
    "들어간다"고 나와도 긴 단어 하나가 있는 줄만 박스 옆으로 삐져나가는 경우를
    info로 잘못 분류하게 된다(실제 사례: 러시아어 긴 단어가 포함된 도형에서
    글자가 상자 옆/밖으로 넘어가는데도 자동 축소가 "성공"으로 처리됨).

    insets_lr_emu/insets_tb_emu(도형의 실제 lIns+rIns, tIns+bIns)를 넘기면 그 값으로
    실사용 가능 폭/높이를 정확히 계산하고, 넘기지 않으면 기존처럼 고정 비율
    (USABLE_WIDTH_RATIO/USABLE_HEIGHT_RATIO)로 근사한다. 고정 비율은 실제보다
    여유 있게 잡히는 도형(예: 좌우 여백이 각 0.15in로 큰 도형)에서 "이 폰트면
    들어간다"고 오판해, 실제로는 2줄로 감긴 텍스트가 도형 아래로 삐져나와 다음
    도형과 겹쳐 보이는 문제가 있었다(실사례: "이 건물은 높은 편이에요" 러시아어
    번역 도형 — 폭 2.57in에 18pt 문장이 들어가려면 2줄이 필요한데, 근사 비율로는
    "1줄로 들어간다"고 잘못 판단해 폰트를 줄이지 않고 그대로 둠).

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

    if insets_tb_emu is not None:
        usable_cy = max(box_cy_emu - insets_tb_emu, 1)
    else:
        usable_cy = max(int(box_cy_emu * USABLE_HEIGHT_RATIO), 1)
    if insets_lr_emu is not None:
        usable_cx = max(box_cx_emu - insets_lr_emu, 1)
    else:
        usable_cx = max(int(box_cx_emu * USABLE_WIDTH_RATIO), 1)
    size = requested_font_size_pt
    while size >= min_font_size_pt:
        lines = estimate_line_count(full_text, size, box_cx_emu, usable_cx_emu=usable_cx)
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


def split_text_for_width(text, font_size_pt, max_cx_emu, max_lines=2):
    """한 줄로 돼 있는 텍스트가 max_cx_emu 폭을 넘칠 때, 박스 크기를 바꾸지 않고
    max_lines줄 이내로 줄바꿈해서 각 줄이 폭 안에 들어가도록 분할을 시도한다.

    그룹으로 묶인 wrap="none" 도형(예: 말풍선, 문법 템플릿 "예문" 박스)은 박스를
    가로로 넓히면 그룹 전체 좌표가 틀어지므로 자동 확장을 하지 않는데, 그러면 폰트를
    최소치까지 줄여도 넘치는 긴 번역문은 그동안 아무 보정 없이 옆으로 삐져나갔다.
    분할 지점은 쉼표(문장이 자연스럽게 끊기는 지점)를 우선하고, 쉼표가 없거나
    분할 후에도 안 맞으면 공백(단어 경계)에서 전체 폭이 최대한 고르게 나뉘도록
    고른다. 어떻게 나눠도 어느 한 줄이 폭을 넘으면(예: 쉼표/공백이 아예 없는
    긴 단어 하나) None을 반환해 호출자가 기존처럼 "수동 확인" 경고로 폴백하게 한다.

    반환: 줄 리스트(성공, 원래 한 줄로 충분하면 길이 1) 또는 None(분할해도 안 맞음)."""
    if not text:
        return None
    if estimate_text_width_emu(text, font_size_pt) <= max_cx_emu:
        return [text]
    if max_lines < 2:
        return None

    def _best_split(points):
        candidates = list(points)
        if not candidates:
            return None
        total_w = estimate_text_width_emu(text, font_size_pt)
        target = total_w / max_lines
        best_idx = min(
            candidates,
            key=lambda i: abs(estimate_text_width_emu(text[:i], font_size_pt) - target),
        )
        first, rest = text[:best_idx].strip(), text[best_idx:].strip()
        if not first or not rest:
            return None
        rest_lines = split_text_for_width(rest, font_size_pt, max_cx_emu, max_lines - 1)
        if rest_lines is None:
            return None
        lines = [first] + rest_lines
        if all(estimate_text_width_emu(line, font_size_pt) <= max_cx_emu for line in lines):
            return lines
        return None

    comma_points = (m.end() for m in re.finditer(r'[,，、]\s*', text))
    result = _best_split(comma_points)
    if result:
        return result
    space_points = (m.start() for m in re.finditer(r'\s+', text))
    return _best_split(space_points)


def split_at_best_comma(text, font_size_pt, usable_cx_emu=None):
    """쉼표(,，、)가 있으면 그 지점에서 정확히 2줄로 나눈다.

    split_text_for_width와 달리 "박스 폭을 넘는지" 여부와 무관하게 쉼표가 있으면
    항상 분할 대상으로 본다 — 말풍선 캡션 도형(이름="번역", wrap="square") 전용
    규칙이다. 사용자 확인 사항: 4번/26번 슬라이드 같은 캐릭터+말풍선 슬라이드는
    번역하면 문장이 거의 항상 원문보다 길어지는데, 쉼표가 있으면 한 줄에 다
    들어가는 짧은 번역문이라도 PowerPoint 자동 워드랩(임의의 단어 경계에서 잘림)에
    맡기지 않고 그 쉼표 지점에서 항상 2줄로 끊어야 자연스럽다고 확인받았다.

    usable_cx_emu(도형의 실사용 가능 폭)를 넘기면, 두 줄 다 그 폭 안에 들어가는
    쉼표 지점 중에서 가장 균형 잡힌 곳을 고른다 — 실제 사례: "그냥 총 텍스트 폭의
    절반"으로만 나누면(이전 버전) 박스가 좁을 때 나뉜 줄 하나가 여전히 박스보다
    넓어서, 뒤이은 fit_font_size_to_box가 "결국 3~4줄로 더 감긴다"고 보고 필요
    이상으로 폰트를 작게 줄이는 문제가 있었다. 두 줄이 실제로 박스 폭에 맞는
    지점을 우선 찾으면 폰트를 불필요하게 줄이지 않고도 정확히 2줄로 끝난다.
    둘 다 폭에 맞는 지점이 하나도 없으면(문장 자체가 워낙 길 때) 차선책으로
    "더 넓은 쪽 줄의 폭이 가장 작아지는" 지점을 고른다 — 이 경우는 fit_font_size_to_box가
    추가로 폰트를 줄이는 게 맞다(실제로 더 긴 줄이 있으니까).
    쉼표가 없거나 분할 결과 어느 한쪽이 빈 문자열이면 None을 반환한다."""
    if not text:
        return None
    comma_points = [m.end() for m in re.finditer(r'[,，、]\s*', text)]
    if not comma_points:
        return None

    def widths(i):
        return (
            estimate_text_width_emu(text[:i], font_size_pt),
            estimate_text_width_emu(text[i:], font_size_pt),
        )

    if usable_cx_emu:
        fitting = [i for i in comma_points if all(w <= usable_cx_emu for w in widths(i))]
        if fitting:
            total_w = estimate_text_width_emu(text, font_size_pt)
            target = total_w / 2
            best_idx = min(fitting, key=lambda i: abs(widths(i)[0] - target))
        else:
            # 어느 지점을 골라도 한쪽은 박스보다 넓다 — 더 넓은 쪽을 최소화한다.
            best_idx = min(comma_points, key=lambda i: max(widths(i)))
    else:
        total_w = estimate_text_width_emu(text, font_size_pt)
        target = total_w / 2
        best_idx = min(comma_points, key=lambda i: abs(widths(i)[0] - target))

    first, rest = text[:best_idx].strip(), text[best_idx:].strip()
    if not first or not rest:
        return None
    return [first, rest]


def suggest_independent_shape_resize(text, font_size_pt, xfrm, align, slide_width_emu,
                                      insets_emu=0, max_width_ratio=0.92):
    """독립적인 제목/안내문 도형: 폭을 늘리고 정렬 기준으로 x를 재조정.

    est_width는 텍스트 상자 "내부"에 필요한 폭이므로, 도형의 외곽 cx를 정할 때는
    좌우 내부 여백(insets_emu)을 다시 더해줘야 한다 — 그렇지 않으면 여백만큼
    항상 조금씩 모자라게 넓혀서 여전히 넘칠 수 있다."""
    est_width = estimate_text_width_emu(text, font_size_pt)
    max_cx = int(slide_width_emu * max_width_ratio)
    new_cx = min(max(est_width + insets_emu, xfrm['cx']), max_cx)

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
