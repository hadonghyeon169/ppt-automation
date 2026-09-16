"""
번역 계획(plan)을 실제 PPTX 파일에 반영하는 오케스트레이터.
1) 번역 적용 (apply_shape_level / restore_shape_translation / apply_multi_paragraph)
2) 넘침(overflow) 보정 (wrap="none" 도형 폭 재계산 또는 wrap="square" 전환)
3) 새 파일명으로 저장 (원본 보존)
"""
import logging
import os
import unicodedata
from pptx import Presentation

from . import ppt_xml_ops as ops
from . import overflow as ov
from .slide_extractor import AUDIO_REQUIRED_MARKERS

logger = logging.getLogger(__name__)


ADJACENCY_GAP_EMU_RATIO = 0.06  # 슬라이드 폭 대비, 이 이내로 다른 도형이 붙어있으면 "인접 쌍"으로 간주

# 문법 구조 다이어그램에 쓰이는 품사 줄임말 라벨. "받침"과 마찬가지로 어떤 목표 언어로도
# 번역하지 않고 원문 그대로 유지해야 한다 — 실제 샘플("초급2 12강")에서 "동"/"형" 라벨이
# 12/14/15/16/17/18/19/20/21/22/23/24/37/39/40/41/42/43번 슬라이드에 동일한 도형
# (shape_name="모서리가 둥근 직사각형 5")으로 반복해서 나오는데, 번역은 3슬라이드씩 서로
# 독립된 AI 호출로 처리되다 보니 배치마다 판단이 달라져(예: 12번은 원문 유지, 37번은
# "Глагол"로 번역) 같은 파일 안에서 라벨이 통일되지 않는 문제가 있었다. skill_prompt.py의
# "받침" 규칙처럼 프롬프트 지시만으로는 배치 간 일관성을 보장할 수 없으므로, 여기서
# 코드 레벨로 강제 적용한다 (AI가 어떤 판단을 내렸든 무시하고 항상 원문 유지).
POS_ABBREVIATION_LABELS = {"동", "명", "형", "부", "관", "감", "조"}


def _boxes_horizontally_adjacent(a, b, slide_width_emu):
    """a의 오른쪽(또는 왼쪽) 근처에 b가 있고 y축으로 겹치는지 근사 판단."""
    gap = slide_width_emu * ADJACENCY_GAP_EMU_RATIO
    y_overlap = not (a['y'] + a['cy'] < b['y'] or b['y'] + b['cy'] < a['y'])
    if not y_overlap:
        return False
    right_gap = b['x'] - (a['x'] + a['cx'])
    left_gap = a['x'] - (b['x'] + b['cx'])
    return (0 <= right_gap <= gap) or (0 <= left_gap <= gap)


# ── "예문" 번호 예제 번역 재배치 ────────────────────────────────────────────
# 문법 설명 "예문" 슬라이드(AUDIO_REQUIRED_MARKERS로 식별되는 13번/38번 유형)는
# "1)...2)...3)..." 형태로 번호 매겨진 한국어 예문이 한 도형(여러 문단, lnSpc 300%로
# 문단마다 넉넉한 세로 "슬롯"을 가짐)에 들어있고, 각 문단에 대응하는 번역이 별도
# 도형("번역")으로 템플릿에 고정 배치되어 있다. 기존엔 이 번역 도형들이 한국어
# 문단 오른쪽 멀리에 있어, 번역문이 길어지면 카드 밖으로 넘치거나(13번 사례) 다음
# 예문과 위아래로 겹쳤다(38번 사례). 사용자가 AskUserQuestion으로 확정한 방식대로,
# 각 번역을 대응 한국어 문단 "바로 아래"로 옮기고, 한 항목이 자기 슬롯에 다 안
# 들어가면 그 초과분만큼 이후 항목들을 함께 아래로 민다.
EXAMPLE_TRANSLATION_GAP_EMU = int(0.05 * ov.EMU_PER_INCH)
# 300%(lnSpc spcPct=300000)는 문단 "슬롯" 전체 간격이지, 실제 한국어 텍스트 한
# 줄의 렌더링 높이가 아니다 — 슬롯 안에서 번역이 들어갈 공간을 구하려면 실제
# 한 줄 높이만 빼야 하므로, 일반적인 줄간격 배수(1.3배)를 따로 쓴다.
EXAMPLE_KR_LINE_HEIGHT_FACTOR = 1.3


def _find_example_translation_groups(extracted_slide):
    """"예문" 슬라이드에서, 여러 문단짜리 한국어 예문 컨테이너 도형과 그 옆에 개별
    배치된 번역("번역"으로 이름 붙은) 도형들을 쌍으로 묶는다. 컨테이너 문단 수와
    번역 도형 개수가 정확히 일치하고, 번역 도형들이 컨테이너의 세로 범위 안에
    걸쳐 있는 경우만 그룹으로 인정한다(오탐 방지) — 실제 관찰된 구조(초급2 12강
    13/38번 슬라이드)에서 템플릿이 각 한국어 문단과 번역 도형을 y좌표로 나란히
    배치해두는 패턴에 기반한다."""
    # 문단 수 조건을 3개 이상 + 모든 문단이 비어있지 않음으로 좁힌다. 같은 "예문"
    # 슬라이드 안에는 "정의"/"구조" 문법 설명 상자도 함께 있는데, 이들은 문단이
    # 2개뿐이고(그중 하나는 라벨만 있거나 비어있는 경우가 많음: 예) ['정의', ''])
    # 번역 도형 개수도 우연히 일치할 수 있어(실제로 2문단 규칙 설명 상자 하나가
    # 그런 경우였다) 잘못 걸려들 위험이 있었다. 실제 두 "예문" 상자(13/38번)는
    # 항상 3문단이고 각 문단이 완전한 예문 문장이라 이 조건으로 정확히 구분된다.
    shapes = extracted_slide["shapes"]
    containers = [
        s for s in shapes
        if s["paragraph_count"] >= 3
        and all((p or "").strip() for p in s.get("paragraph_texts") or [])
        and "번역" not in (s["shape_name"] or "")
        and s["has_korean"]
        and s["xfrm_emu"]
        and not s.get("is_grouped")
    ]
    translation_shapes = [
        s for s in shapes
        if (s["shape_name"] or "") == "번역"
        and s["xfrm_emu"]
        and not s.get("is_grouped")
    ]
    groups = []
    used_ids = set()
    for container in containers:
        n = container["paragraph_count"]
        c = container["xfrm_emu"]
        c_y0, c_y1 = c["y"], c["y"] + c["cy"]
        # y축 겹침만으로는 슬라이드 반대편(왼쪽)에 있는, 전혀 다른 상자(예: "정의"/
        # "구조" 규칙 설명 상자의 번역)까지 끌려들어올 수 있다(실제 겪음: 왼쪽에
        # 있는 번역 도형이 우연히 이 컨테이너의 y범위와 겹쳐 후보에 잡힘) — 템플릿상
        # 번역 도형은 항상 한국어 컨테이너의 왼쪽 경계보다 오른쪽에 있으므로 x 조건도
        # 함께 건다.
        candidates = [
            s for s in translation_shapes
            if s["shape_id"] not in used_ids
            and s["xfrm_emu"]["x"] >= c["x"]
            and (s["xfrm_emu"]["y"] + s["xfrm_emu"]["cy"]) > c_y0
            and s["xfrm_emu"]["y"] < c_y1
        ]
        if len(candidates) != n:
            continue  # 개수가 안 맞으면 오탐 가능성이 있으니 건드리지 않고 건너뜀
        candidates.sort(key=lambda s: s["xfrm_emu"]["y"])
        for s in candidates:
            used_ids.add(s["shape_id"])
        groups.append({"container": container, "translations": candidates})
    return groups


def _reposition_example_translations(extracted_slide, translated_by_shape_id, slide_idx, review_flags):
    """예문 슬라이드에서 각 번역 도형을 대응하는 한국어 문단 바로 아래로 옮기고,
    슬롯을 넘치는 항목이 있으면 그 초과분만큼 이후 항목을 아래로 민다.
    반환: 재배치에 성공한 shape_id 집합 (이후 일반 넘침 보정 루프에서 제외해야 함)."""
    repositioned_ids = set()
    combined_text = "".join(s["full_text"] for s in extracted_slide["shapes"])
    if not all(marker in combined_text for marker in AUDIO_REQUIRED_MARKERS):
        # 일반 슬라이드에 영향을 주지 않도록, 확실한 "예문" 슬라이드에서만 동작한다.
        return repositioned_ids

    for group in _find_example_translation_groups(extracted_slide):
        container = group["container"]
        entries = []
        ok = True
        for t_meta in group["translations"]:
            entry = translated_by_shape_id.get(t_meta["shape_id"])
            if not entry:
                ok = False
                break
            entries.append((t_meta, entry))
        if not ok or not entries:
            # 그룹 중 하나라도 이번에 번역되지 않았으면(스킵/누락 등) 안전하게
            # 기존 배치를 그대로 둔다 — 어설프게 절반만 옮기지 않는다.
            continue

        c_xfrm = container["xfrm_emu"]
        kr_font_pt = container["font_sizes_pt"][0] if container.get("font_sizes_pt") else 27
        kr_line_h_emu = kr_font_pt * EXAMPLE_KR_LINE_HEIGHT_FACTOR * ov.EMU_PER_PT
        # 컨테이너 높이를 문단 수로 균등분할해서 슬롯 시작점을 역산하면 실제와 어긋난다
        # — 300% 줄간격 문단은 렌더러(LibreOffice로 실측/렌더 검증함)가 그 여백
        # 대부분을 텍스트 "앞"에 두는 것으로 관찰되어, 한국어 텍스트 자체가 슬롯
        # 맨 위가 아니라 슬롯 중간 즈음에서 시작한다. 반면 템플릿이 원래 배치해둔
        # 번역 도형의 y좌표는 "이 문단과 같은 행"이라는 실측 기준점이므로(사용자가
        # 준 참고 파일에서 확인: 각 번역이 대응 한국어 문단과 같은 행에 나란히
        # 배치되어 있었음), 그 y좌표를 그대로 기준점으로 삼아 그 아래로 내린다.
        row_anchors = [t_meta["xfrm_emu"]["y"] for t_meta, _ in entries]

        # 실사용자 파일로 검증하다 실제로 겪은 버그: 컨테이너를 문단 수로 나눈
        # "슬롯" 경계를 폰트 크기 상한으로 그대로 썼더니, 마지막 항목은 컨테이너
        # 바닥까지 남는 공간이 애초에 한 줄 분량(템플릿이 원래 "옆에 붙여서" 배치
        # 했을 때 쓰던 공간)밖에 없어서 거의 항상 최소 폰트(10pt)까지 짓눌리고,
        # 그마저도 카드 밖으로 살짝 넘쳤다(사용자가 첨부한 실제 파일에서
        # 확인: rPr sz=1000, cy가 컨테이너 바닥을 넘어감). 예문 번역은 거의 항상
        # 한 줄짜리 짧은 문장이므로, 슬롯 경계에 얽매이지 않고 "여유 있는" 고정
        # 상한만 두어(비정상적으로 긴 번역에 대한 안전장치 목적) 원래 폰트 크기를
        # 최대한 유지한다. 각 항목의 실제 위치는 "자기 행(row_anchor) 기준 위치"와
        # "바로 앞 항목이 끝난 위치" 중 더 아래쪽을 쓰는 식으로 순차적으로 정해서,
        # 앞 항목이 짧으면(거의 항상 그렇다) 남는 공간이 뒤 항목에 자연스럽게
        # 재분배되고, 앞 항목이 길 때만 실제로 필요한 만큼만 뒤 항목을 민다.
        GENEROUS_FIT_CEILING_EMU = int(1.1 * ov.EMU_PER_INCH)
        container_bottom = c_xfrm["y"] + c_xfrm["cy"]
        cumulative_bottom = None

        for i, (t_meta, (sp, _meta, final_text, applied_font_size)) in enumerate(entries):
            desired_top = row_anchors[i] + kr_line_h_emu + EXAMPLE_TRANSLATION_GAP_EMU
            if cumulative_bottom is None:
                new_y = desired_top
            else:
                new_y = max(desired_top, cumulative_bottom + EXAMPLE_TRANSLATION_GAP_EMU)
            pushed_down = new_y - desired_top

            new_x, new_cx = c_xfrm["x"], c_xfrm["cx"]
            requested_pt = applied_font_size or (
                t_meta["font_sizes_pt"][0] if t_meta.get("font_sizes_pt") else 20
            )

            fitted_pt, overflow_unresolved = ov.fit_font_size_to_box(
                final_text, requested_pt, new_cx, GENEROUS_FIT_CEILING_EMU,
                insets_lr_emu=t_meta.get("insets_lr_emu"),
                insets_tb_emu=t_meta.get("insets_tb_emu"),
            )
            usable_cx = (
                new_cx - t_meta["insets_lr_emu"]
                if t_meta.get("insets_lr_emu") is not None
                else int(new_cx * ov.USABLE_WIDTH_RATIO)
            )
            lines = ov.estimate_line_count(final_text, fitted_pt, new_cx, usable_cx_emu=max(usable_cx, 1))
            needed_h = lines * fitted_pt * ov.LINE_SPACING_FACTOR * ov.EMU_PER_PT

            try:
                ops.set_body_pr_wrap_square_autofit(sp)
                ops.set_shape_xfrm(
                    sp, x=int(new_x), y=int(new_y), cx=int(new_cx), cy=int(needed_h)
                )
                if fitted_pt != requested_pt:
                    ops.force_font_size(sp, int(fitted_pt * 100))
            except Exception as e:
                review_flags.append({
                    "slide_index": slide_idx, "shape_name": t_meta["shape_name"],
                    "source_text": t_meta["full_text"], "translated_text": final_text,
                    "issue": f"예문 번역 재배치 중 오류 (수동 확인 필요): {e}", "severity": "warning",
                })
                continue

            repositioned_ids.add(t_meta["shape_id"])
            cumulative_bottom = new_y + needed_h
            went_past_container = cumulative_bottom > container_bottom
            if pushed_down > EXAMPLE_TRANSLATION_GAP_EMU or overflow_unresolved or went_past_container:
                review_flags.append({
                    "slide_index": slide_idx, "shape_name": t_meta["shape_name"],
                    "source_text": t_meta["full_text"], "translated_text": final_text,
                    "issue": "예문 번역이 길어 위치가 아래로 밀렸거나 카드 경계를 살짝 넘었을 "
                             "수 있습니다 — 확인해주세요.",
                    "severity": "warning",
                })

    return repositioned_ids


def apply_translation_plan(pptx_path, extracted, plan_by_shape, lang_code, lang_meta, out_path):
    prs = Presentation(pptx_path)
    slide_width = extracted["slide_width_emu"]

    applied_records = []
    review_flags = []
    untranslated_slides = set()  # 원문이 그대로 남은(번역 누락) 슬라이드 — 부분 재번역 대상

    for slide_idx, slide in enumerate(prs.slides):
        sp_list = ops.iter_all_shapes(slide)
        extracted_slide = extracted["slides"][slide_idx]
        extracted_by_id = {sh["shape_id"]: sh for sh in extracted_slide["shapes"]}

        # 말풍선(캐릭터 대사) 슬라이드 여부 — PowerPoint가 자동으로 붙이는 도형
        # 이름이 "말풍선: 모서리가 둥근 사각형 N" 형태로 저장되므로 "말풍선"으로
        # 시작하는지만 보면 된다. 이 슬라이드에서 이름이 "번역"인 wrap="square"
        # 캡션 도형에는 아래에서 쉼표 기준 강제 줄바꿈을 적용한다.
        bubble_caption_slide = any(
            (s["shape_name"] or "").startswith("말풍선") for s in extracted_slide["shapes"]
        )

        # 이번 슬라이드에서 실제로 번역이 적용된 도형들 (넘침 보정 대상)
        translated_sp_on_slide = []

        for shape_idx, sp in enumerate(sp_list):
            # slide_extractor와 동일한 방식(도형 고유 XML id 기반)으로 shape_id를 계산해야
            # plan_by_shape/extracted_by_id와 정확히 매칭된다 (순번 기반이면 이후
            # cleanup_duplicate_shapes가 도형을 지울 때 뒤따르는 도형들의 순번이 밀려서
            # 스냅샷과 어긋난다 — 실제로 겪은 버그).
            cnvpr_id = ops.get_shape_cnvpr_id(sp)
            stable_id = cnvpr_id if cnvpr_id is not None else f"x{shape_idx}"
            shape_id = f"s{slide_idx}_{stable_id}"
            plan = plan_by_shape.get(shape_id)
            meta = extracted_by_id.get(shape_id)
            if not meta:
                continue
            if (meta["full_text"] or "").strip() in POS_ABBREVIATION_LABELS:
                # 품사 줄임말 라벨 — AI의 판단(번역/스킵 여부, plan 유무)과 무관하게 항상
                # 원문 유지. "누락" 에러로도 잡히지 않도록 plan 존재 여부 확인보다 먼저 처리.
                review_flags.append({
                    "slide_index": slide_idx, "shape_name": meta["shape_name"],
                    "source_text": meta["full_text"], "translated_text": meta["full_text"],
                    "issue": "품사 줄임말 라벨(동/명/형/부 등)은 규칙에 따라 번역하지 않고 원문을 그대로 유지했습니다.",
                    "severity": "info",
                })
                continue
            if not plan:
                # LLM 응답에 이 도형이 아예 없었다 (배치 자체가 스킵됐거나, 응답이 잘렸거나,
                # 모델이 그냥 빠뜨림). action="skip"과 달리 이건 "의도적 판단"이 아니라
                # 누락이므로 원문이 그대로 남아있다는 걸 반드시 눈에 띄게 표시해야 한다
                # (전에는 여기서 조용히 continue만 해서, 번역이 하나도 안 된 슬라이드가
                # 검수 화면에 아무 표시 없이 원문 그대로 섞여 나가는 사고가 있었다).
                review_flags.append({
                    "slide_index": slide_idx, "shape_name": meta["shape_name"],
                    "source_text": meta["full_text"], "translated_text": None,
                    "issue": "AI 번역 응답에 이 도형이 누락되어 원문이 그대로 남아 있습니다.",
                    "severity": "error",
                })
                untranslated_slides.add(slide_idx)
                continue
            if plan.get("action") != "translate":
                continue  # 모델이 의도적으로 skip 처리(지시문/페이지번호 등) — 원문 유지가 맞음

            translated_text = plan.get("translated_text") or ""
            translated_paragraphs = plan.get("translated_paragraphs")
            if not translated_text and not translated_paragraphs:
                # action은 "translate"인데 실제 번역 결과가 비어있는 경우 — 이것도 원문이
                # 그대로 남으므로 누락과 동일하게 취급한다.
                review_flags.append({
                    "slide_index": slide_idx, "shape_name": meta["shape_name"],
                    "source_text": meta["full_text"], "translated_text": None,
                    "issue": "AI가 번역하겠다고 표시했지만 번역 결과가 비어 있어 원문이 그대로 남아 있습니다.",
                    "severity": "error",
                })
                untranslated_slides.add(slide_idx)
                continue
            bold_category = plan.get("bold_category", "plain")
            force_bold = True if bold_category == "mixed_bold" else False
            force_sz_pt = plan.get("force_font_size_pt")
            is_white = meta["is_white_text"]

            # ── 폰트 크기 사전 검증(모든 도형) ──────────────────────────────
            # force_font_size_pt(언어별 정밀 규칙 또는 AI 판단)는 이 PPT 템플릿의
            # 실제 도형 크기를 모르는 상태에서 나온 "권장값"일 뿐이다. wrap="square"
            # (또는 wrap 속성 없음 = 기본값 square) 도형은 실제 cy(높이) 기준으로,
            # wrap="none"(줄바꿈 없는 한 줄 도형)은 슬라이드 폭 한도 기준으로 텍스트가
            # 들어가는지 검증하고, 넘치면 폰트를 줄인다. wrap="none"은 이후 박스 자체를
            # 넓히는 보정도 받지만 그마저도 슬라이드 폭의 92%까지만 넓어지므로, 아주 긴
            # 번역문은 폰트 축소 없이는 여전히 슬라이드 밖으로 삐져나갈 수 있었다.
            requested_pt = force_sz_pt or (meta["font_sizes_pt"][0] if meta["font_sizes_pt"] else 18)

            # ── 말풍선 캡션 도형 전용: 쉼표 기준 강제 줄바꿈 ──────────────────
            # 사용자 확인 사항: 4번/26번 같은 캐릭터+말풍선 슬라이드는 번역하면
            # 문장이 거의 항상 길어지는데, 쉼표가 있으면 한 줄에 다 들어가는
            # 경우라도 항상 그 지점에서 2줄로 나눠야 한다 — PowerPoint 자동
            # 워드랩(임의의 단어 경계에서 잘림)에 맡기지 않기 위함. 아래
            # fit_font_size_to_box 호출보다 먼저 텍스트에 "\n"을 넣어야
            # 폰트/높이 계산이 실제 2줄 레이아웃 기준으로 정확히 이루어진다.
            if (
                bubble_caption_slide
                and meta["shape_name"] == "번역"
                and meta["wrap"] == "square"
                and not translated_paragraphs
                and translated_text
            ):
                usable_cx_for_split = None
                if meta["xfrm_emu"]:
                    insets_lr = meta.get("insets_lr_emu")
                    box_cx = meta["xfrm_emu"]["cx"]
                    usable_cx_for_split = max(
                        box_cx - insets_lr if insets_lr is not None else int(box_cx * ov.USABLE_WIDTH_RATIO),
                        1,
                    )
                comma_lines = ov.split_at_best_comma(translated_text, requested_pt, usable_cx_for_split)
                if comma_lines:
                    translated_text = "\n".join(comma_lines)
                    review_flags.append({
                        "slide_index": slide_idx, "shape_name": meta["shape_name"],
                        "source_text": meta["full_text"], "translated_text": translated_text,
                        "issue": "말풍선 캡션 도형 — 쉼표 지점에서 2줄로 자동 줄바꿈했습니다.",
                        "severity": "info",
                    })

            fit_note = None
            if meta["xfrm_emu"] and meta["wrap"] != "none":
                text_for_fit = (
                    translated_paragraphs
                    if (meta["paragraph_count"] > 1 and translated_paragraphs)
                    else translated_text
                )
                fitted_pt, overflow_unresolved = ov.fit_font_size_to_box(
                    text_for_fit, requested_pt,
                    meta["xfrm_emu"]["cx"], meta["xfrm_emu"]["cy"],
                    insets_lr_emu=meta.get("insets_lr_emu"),
                    insets_tb_emu=meta.get("insets_tb_emu"),
                )
                if fitted_pt < requested_pt:
                    if overflow_unresolved:
                        fit_note = (
                            f"번역 텍스트가 도형 크기에 비해 너무 깁니다 (최소 {fitted_pt}pt로 "
                            f"줄여도 넘칠 수 있음 — 텍스트를 줄이거나 도형 크기를 수동으로 확인해주세요)."
                        )
                    else:
                        fit_note = f"도형 크기에 맞춰 폰트를 {requested_pt}pt → {fitted_pt}pt로 자동 축소했습니다."
                    force_sz_pt = fitted_pt
            elif meta["xfrm_emu"] and meta["wrap"] == "none":
                max_cx = int(slide_width * 0.92)
                fitted_pt, overflow_unresolved = ov.fit_font_size_to_width(
                    translated_text, requested_pt, max_cx,
                )
                if fitted_pt < requested_pt:
                    if overflow_unresolved:
                        fit_note = (
                            f"번역 텍스트가 슬라이드 폭에 비해 너무 깁니다 (최소 {fitted_pt}pt로 "
                            f"줄여도 넘칠 수 있음 — 텍스트를 줄이거나 도형을 수동으로 확인해주세요)."
                        )
                    else:
                        fit_note = f"도형 폭에 맞춰 폰트를 {requested_pt}pt → {fitted_pt}pt로 자동 축소했습니다."
                    force_sz_pt = fitted_pt

            force_sz = int(force_sz_pt * 100) if force_sz_pt else None
            # force_color: AI가 언어별 정밀 스타일 규칙에 따라 "white"/"black"을 지정한
            # 경우에만 채워진다. null이면 기존처럼 원본 도형 색을 그대로 유지한다.
            raw_force_color = plan.get("force_color")
            force_color = raw_force_color if raw_force_color in ("white", "black") else None

            try:
                if meta["paragraph_count"] > 1 and translated_paragraphs:
                    ops.apply_multi_paragraph(
                        sp, translated_paragraphs, lang_code, lang_meta["font"],
                        is_complex=lang_meta["complex_script"], force_sz=force_sz,
                        force_color=force_color,
                    )
                elif is_white:
                    ops.restore_shape_translation(
                        sp, translated_text, lang_code, lang_meta["font"],
                        is_complex=lang_meta["complex_script"], force_sz=force_sz, force_bold=force_bold,
                        force_color=force_color,
                    )
                else:
                    ops.apply_shape_level(
                        sp, translated_text, lang_code, lang_meta["font"],
                        is_complex=lang_meta["complex_script"], force_sz=force_sz, force_bold=force_bold,
                        force_color=force_color,
                    )
            except Exception as e:
                review_flags.append({
                    "slide_index": slide_idx, "shape_name": meta["shape_name"],
                    "source_text": meta["full_text"], "translated_text": translated_text,
                    "issue": f"XML 적용 중 오류: {e}", "severity": "error",
                })
                continue

            applied_records.append({
                "slide_index": slide_idx,
                "shape_id": shape_id,
                "shape_name": meta["shape_name"],
                "shape_type": meta.get("heuristic_type"),
                "source_korean": meta["full_text"] if meta["has_korean"] else None,
                "source_chinese": meta["full_text"] if meta["has_chinese"] else None,
                "translated_text": translated_text or " / ".join(translated_paragraphs or []),
                "is_white_text": is_white,
                "is_bold": force_bold,
                "font_size": force_sz_pt,
            })

            if plan.get("needs_review"):
                note_text = plan.get("note") or "모델이 검수가 필요하다고 표시함"
                # 저장 시점(오래전 실행분 등)에 따라 유니코드 정규화 형태가 다를 수 있어
                # NFC로 통일한다 — 그렇지 않으면 겉보기엔 동일한 "추정"/"말풍선" 같은
                # 한글 문자열이 코드포인트 단위 부분 문자열 매칭에서 실패할 수 있다.
                note_text = unicodedata.normalize("NFC", note_text)
                model_severity = plan.get("review_severity")
                logger.info(
                    "needs_review note (shape=%s, model_severity=%r): %r",
                    shape_id, model_severity, note_text,
                )
                if model_severity in ("info", "warning"):
                    # 최신 프롬프트는 모델이 직접 info/warning을 구분해서 내려준다 —
                    # note 문구를 추측할 필요 없이 이 값을 그대로 신뢰한다.
                    severity = model_severity
                else:
                    # review_severity 필드가 없는 예전 저장된 plan(재검사 등)과의 호환용
                    # 폴백. 러시아어 말풍선 규칙(캐릭터 성별에 따라 32/24pt)은 소스
                    # 텍스트만으로 캐릭터를 특정할 수 없어 거의 항상 needs_review로
                    # 표시되는데, note 문구가 "여성 캐릭터 여부 미확인", "캐릭터 유무
                    # 확인 필요", "말풍선으로 추정되어 32pt 적용", "여캐릭터", "성별
                    # 확인 불가" 등 매번 다르게 나와서 특정 문구 하나만 걸러서는 놓치는
                    # 경우가 많았다. 폰트는 이미 규칙대로 적용된 상태이므로 이런 자동
                    # 판단 참고 노트는 info로, 그 외(번역 품질 의심, 위치 조정 필요 등
                    # "말풍선"/"캐릭터" 언급이 없는 사유)는 warning으로 유지한다.
                    UNCERTAIN_KEYWORDS = (
                        "추정", "불명확", "불확실", "미확인", "확인 필요", "확인 불가", "여부", "임의 적용",
                    )
                    looks_like_font_heuristic_note = (
                        ("말풍선" in note_text or "캐릭터" in note_text)
                        and any(kw in note_text for kw in UNCERTAIN_KEYWORDS)
                    )
                    severity = "info" if looks_like_font_heuristic_note else "warning"
                    logger.info(
                        "fallback classification (shape=%s): matched=%s -> severity=%s",
                        shape_id, looks_like_font_heuristic_note, severity,
                    )
                review_flags.append({
                    "slide_index": slide_idx, "shape_name": meta["shape_name"],
                    "source_text": meta["full_text"], "translated_text": translated_text,
                    "issue": note_text,
                    "severity": severity,
                })

            if fit_note:
                review_flags.append({
                    "slide_index": slide_idx, "shape_name": meta["shape_name"],
                    "source_text": meta["full_text"], "translated_text": translated_text,
                    "issue": fit_note,
                    "severity": "warning" if "확인해주세요" in fit_note else "info",
                })

            applied_font_size = force_sz_pt or requested_pt
            translated_sp_on_slide.append(
                (sp, meta, translated_text or " ".join(translated_paragraphs or []), applied_font_size)
            )

        # ── "예문" 번역 재배치 (넘침 보정보다 먼저) ───────────────────────
        # 대상 도형은 여기서 위치/폭/폰트가 이미 최종 확정되므로, 아래 일반 넘침
        # 보정 루프에서는 건너뛰어야 한다(중복 처리 시 방금 잡은 배치가 다시
        # 틀어질 수 있음).
        translated_by_shape_id = {
            m["shape_id"]: (sp, m, txt, fs) for sp, m, txt, fs in translated_sp_on_slide
        }
        repositioned_ids = _reposition_example_translations(
            extracted_slide, translated_by_shape_id, slide_idx, review_flags
        )

        # ── 넘침(overflow) 보정 ──────────────────────────────────────────
        all_boxes = [sh["xfrm_emu"] for sh in extracted_slide["shapes"] if sh["xfrm_emu"]]
        for sp, meta, final_text, font_size in translated_sp_on_slide:
            if meta["shape_id"] in repositioned_ids:
                continue
            xfrm = meta["xfrm_emu"]
            wrap = meta["wrap"]
            if not xfrm or wrap != "none":
                continue
            # font_size는 위 사전 검증 단계에서 이미 적용된(축소됐을 수 있는) 실제 폰트
            # 크기다 — 원본 크기를 다시 쓰면 축소를 반영 못 하고 박스만 과도하게 넓히게 된다.
            # insets_lr_emu(좌우 내부 여백)를 빼지 않으면 실제로는 넘치는 도형도 비율이
            # 1.05 기준선 바로 아래로 나와 보정이 안 걸리는 문제가 있었다(실사례:
            # "예문" 번역 도형들이 좌우 여백 각 0.15in인데 이를 무시하면 ratio
            # 0.98~0.99로 계산되어 통과했지만, 실사용 폭 기준으로는 1.05~1.06으로 넘침).
            insets_emu = meta.get("insets_lr_emu", 0) or 0
            ratio = ov.compute_overflow_ratio(final_text, font_size, xfrm["cx"], insets_emu)
            if ratio <= 1.05:
                continue  # 넘치지 않음

            others = [b for b in all_boxes if b and b != xfrm]
            has_adjacent = any(_boxes_horizontally_adjacent(xfrm, b, slide_width) for b in others)

            try:
                if has_adjacent:
                    ops.set_body_pr_wrap_square_autofit(sp)
                elif meta.get("is_grouped"):
                    # 그룹(p:grpSp) 안에 있는 도형은 xfrm이 그룹의 자식 좌표계 기준이라,
                    # 여기서 계산한 절대 좌표로 그대로 덮어쓰면 그룹 전체가 틀어진다.
                    # 폰트 축소는 이미 위 사전 검증 단계에서 적용됐으므로 가로로는 더
                    # 넓힐 수 없다 — 대신 쉼표/단어 경계에서 최대 2줄로 나눠 세로
                    # 방향으로 배치해본다(말풍선, 문법 템플릿 "예문" 박스 등에서
                    # 실제로 반복 관찰된 옆으로 삐져나가는 문제에 대한 대응).
                    # 도형 높이(cy)까지 늘어난 줄 수가 실제로 들어가는지는 이 시점에
                    # 확인할 수 없으므로 여전히 warning으로 남겨 사람이 확인하게 한다.
                    usable_cx = max(xfrm["cx"] - insets_emu, 1)
                    split_lines = ov.split_text_for_width(final_text, font_size, usable_cx, max_lines=2)
                    if split_lines and len(split_lines) > 1:
                        ops.force_multiline(sp, split_lines)
                        review_flags.append({
                            "slide_index": slide_idx, "shape_name": meta["shape_name"],
                            "source_text": meta["full_text"], "translated_text": final_text,
                            "issue": f"그룹 도형이라 박스를 가로로 넓힐 수 없어 텍스트를 "
                                     f"{len(split_lines)}줄로 자동 줄바꿈했습니다 — 도형 높이 안에 "
                                     "다 들어가는지 확인해주세요.",
                            "severity": "warning",
                        })
                    else:
                        review_flags.append({
                            "slide_index": slide_idx, "shape_name": meta["shape_name"],
                            "source_text": meta["full_text"], "translated_text": final_text,
                            "issue": "그룹으로 묶인 도형이라 박스 자동 확장을 건너뛰었고, "
                                     "줄바꿈으로도 폭에 맞출 수 없었습니다 (쉼표/공백 구분이 "
                                     "없는 긴 단어 등) — 넘치는지 수동으로 확인해주세요.",
                            "severity": "warning",
                        })
                else:
                    new_geo = ov.suggest_independent_shape_resize(
                        final_text, font_size, xfrm, meta["align"], slide_width,
                        insets_emu=insets_emu,
                    )
                    ops.set_shape_xfrm(sp, x=new_geo["x"], cx=new_geo["cx"])
            except Exception as e:
                review_flags.append({
                    "slide_index": slide_idx, "shape_name": meta["shape_name"],
                    "source_text": meta["full_text"], "translated_text": final_text,
                    "issue": f"넘침 보정 중 오류 (수동 확인 필요): {e}", "severity": "warning",
                })

    removed = ops.cleanup_duplicate_shapes(prs)
    if removed:
        review_flags.append({
            "slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
            "issue": f"중복/잔존 도형 {removed}개를 자동 정리했습니다. 렌더링 결과를 확인해주세요.",
            "severity": "info",
        })

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    prs.save(out_path)
    # 부분 재번역 대상: 슬라이드 하나당 1개짜리 "배치"로 반환한다 (기존 실패-배치
    # 재시도 인프라(failed_batches_json)가 [[slide_idx, ...], ...] 형태를 그대로
    # 재사용하므로, 여기서는 슬라이드 단위로 세분화해 이미 성공한 슬라이드까지
    # 다시 부르지 않게 한다).
    untranslated_batches = [[idx] for idx in sorted(untranslated_slides)]
    return applied_records, review_flags, untranslated_batches


def make_output_filename(original_filename, lang_meta, project_name=None):
    """"PPT 강의 제작 인수인계서" 1-3 저장 규칙([국적] 초급n n과 과제목)을 최대한 따른다.
    프로젝트 이름을 그 형식(예: '초급1 3과 자기소개')으로 입력해두면 그대로 재현되고,
    비워두면 원본 파일명을 대신 사용한다."""
    base = (project_name or "").strip() or os.path.splitext(original_filename)[0]
    code = lang_meta.get("country_code") or lang_meta.get("label", "")
    return f"[{code}] {base}.pptx"
