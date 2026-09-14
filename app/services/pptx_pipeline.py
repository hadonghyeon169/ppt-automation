"""
번역 계획(plan)을 실제 PPTX 파일에 반영하는 오케스트레이터.
1) 번역 적용 (apply_shape_level / restore_shape_translation / apply_multi_paragraph)
2) 넘침(overflow) 보정 (wrap="none" 도형 폭 재계산 또는 wrap="square" 전환)
3) 새 파일명으로 저장 (원본 보존)
"""
import os
from pptx import Presentation

from . import ppt_xml_ops as ops
from . import overflow as ov


ADJACENCY_GAP_EMU_RATIO = 0.06  # 슬라이드 폭 대비, 이 이내로 다른 도형이 붙어있으면 "인접 쌍"으로 간주


def _boxes_horizontally_adjacent(a, b, slide_width_emu):
    """a의 오른쪽(또는 왼쪽) 근처에 b가 있고 y축으로 겹치는지 근사 판단."""
    gap = slide_width_emu * ADJACENCY_GAP_EMU_RATIO
    y_overlap = not (a['y'] + a['cy'] < b['y'] or b['y'] + b['cy'] < a['y'])
    if not y_overlap:
        return False
    right_gap = b['x'] - (a['x'] + a['cx'])
    left_gap = a['x'] - (b['x'] + b['cx'])
    return (0 <= right_gap <= gap) or (0 <= left_gap <= gap)


def apply_translation_plan(pptx_path, extracted, plan_by_shape, lang_code, lang_meta, out_path):
    prs = Presentation(pptx_path)
    slide_width = extracted["slide_width_emu"]

    applied_records = []
    review_flags = []

    for slide_idx, slide in enumerate(prs.slides):
        sp_list = ops.iter_all_shapes(slide)
        extracted_slide = extracted["slides"][slide_idx]
        extracted_by_id = {sh["shape_id"]: sh for sh in extracted_slide["shapes"]}

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
            if not plan or not meta:
                continue
            if plan.get("action") != "translate":
                continue

            translated_text = plan.get("translated_text") or ""
            translated_paragraphs = plan.get("translated_paragraphs")
            bold_category = plan.get("bold_category", "plain")
            force_bold = True if bold_category == "mixed_bold" else False
            force_sz_pt = plan.get("force_font_size_pt")
            is_white = meta["is_white_text"]

            # ── 폰트 크기 사전 검증(word-wrap 도형만) ──────────────────────
            # force_font_size_pt(언어별 정밀 규칙 또는 AI 판단)는 이 PPT 템플릿의
            # 실제 도형 크기를 모르는 상태에서 나온 "권장값"일 뿐이다. wrap="square"
            # (또는 wrap 속성 없음 = 기본값 square) 도형은 실제 cy(높이) 기준으로
            # 텍스트가 들어가는지 검증하고, 넘치면 폰트를 줄인다. 그동안 이 검증이
            # 전혀 없어서(예전엔 wrap="none" 도형만 넘침 보정) 말풍선/라벨/본문
            # 텍스트가 도형을 넘어가는 문제(텍스트가 너무 크거나 길어서 안 맞음)가
            # 있었다.
            requested_pt = force_sz_pt or (meta["font_sizes_pt"][0] if meta["font_sizes_pt"] else 18)
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
                review_flags.append({
                    "slide_index": slide_idx, "shape_name": meta["shape_name"],
                    "source_text": meta["full_text"], "translated_text": translated_text,
                    "issue": plan.get("note") or "모델이 검수가 필요하다고 표시함",
                    "severity": "warning",
                })

            if fit_note:
                review_flags.append({
                    "slide_index": slide_idx, "shape_name": meta["shape_name"],
                    "source_text": meta["full_text"], "translated_text": translated_text,
                    "issue": fit_note,
                    "severity": "warning" if "확인해주세요" in fit_note else "info",
                })

            translated_sp_on_slide.append((sp, meta, translated_text or " ".join(translated_paragraphs or [])))

        # ── 넘침(overflow) 보정 ──────────────────────────────────────────
        all_boxes = [sh["xfrm_emu"] for sh in extracted_slide["shapes"] if sh["xfrm_emu"]]
        for sp, meta, final_text in translated_sp_on_slide:
            xfrm = meta["xfrm_emu"]
            wrap = meta["wrap"]
            if not xfrm or wrap != "none":
                continue
            font_size = (meta["font_sizes_pt"][0] if meta["font_sizes_pt"] else 18)
            ratio = ov.compute_overflow_ratio(final_text, font_size, xfrm["cx"])
            if ratio <= 1.05:
                continue  # 넘치지 않음

            others = [b for b in all_boxes if b and b != xfrm]
            has_adjacent = any(_boxes_horizontally_adjacent(xfrm, b, slide_width) for b in others)

            try:
                if has_adjacent:
                    ops.set_body_pr_wrap_square_autofit(sp)
                else:
                    new_geo = ov.suggest_independent_shape_resize(
                        final_text, font_size, xfrm, meta["align"], slide_width,
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
    return applied_records, review_flags


def make_output_filename(original_filename, lang_meta, project_name=None):
    """"PPT 강의 제작 인수인계서" 1-3 저장 규칙([국적] 초급n n과 과제목)을 최대한 따른다.
    프로젝트 이름을 그 형식(예: '초급1 3과 자기소개')으로 입력해두면 그대로 재현되고,
    비워두면 원본 파일명을 대신 사용한다."""
    base = (project_name or "").strip() or os.path.splitext(original_filename)[0]
    code = lang_meta.get("country_code") or lang_meta.get("label", "")
    return f"[{code}] {base}.pptx"
