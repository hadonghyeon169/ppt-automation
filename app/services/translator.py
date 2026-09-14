"""
번역 계획(translation plan) 생성 오케스트레이터.
슬라이드를 배치로 묶어 Claude API에 보내고, 도형별 번역 결정을 받아온다.
실제 XML 반영은 pptx_pipeline.py가 담당한다.
"""
import json
import logging

from . import llm_clients
from .skill_prompt import build_system_prompt
from .slide_extractor import classify_shape_type

logger = logging.getLogger(__name__)

SLIDES_PER_BATCH = 3
# 4에서 3으로 낮췄다: 언어별 정밀 스타일 규칙(force_color/force_font_size_pt 등)이
# 늘면서 shape당 출력 JSON이 길어져, 도형이 많은 슬라이드가 낀 배치에서
# max_tokens(16000)를 넘겨 응답이 잘리는 경우가 있었다 (Unterminated string 오류).


def _slide_to_prompt_payload(slide):
    shapes = []
    for sh in slide["shapes"]:
        shapes.append({
            "shape_id": sh["shape_id"],
            "shape_name": sh["shape_name"],
            "full_text": sh["full_text"],
            "paragraph_texts": sh["paragraph_texts"] if sh["paragraph_count"] > 1 else None,
            "has_chinese": sh["has_chinese"],
            "has_korean": sh["has_korean"],
            "heuristic_type": classify_shape_type(sh),
        })
    return {"slide_index": slide["slide_index"], "shapes": shapes}


def build_reference_hint(reference_extracted, slide_indices, max_chars=3000):
    if not reference_extracted:
        return None
    chunks = []
    for si in slide_indices:
        if si >= len(reference_extracted["slides"]):
            continue
        slide = reference_extracted["slides"][si]
        texts = [f'- {sh["shape_name"]}: {sh["full_text"]}' for sh in slide["shapes"] if sh["full_text"]]
        if texts:
            chunks.append(f"[슬라이드 {si + 1}]\n" + "\n".join(texts))
    hint = "\n\n".join(chunks)
    return hint[:max_chars] if hint else None


def plan_translation(api_key, model, extracted, lang_code, lang_meta, reference_extracted=None,
                      progress_cb=None):
    """전체 프레젠테이션에 대해 슬라이드 배치 단위로 Claude를 호출하고,
    shape_id -> plan(dict) 매핑을 반환한다."""
    slides = extracted["slides"]
    plan_by_shape = {}
    errors = []

    batches = [slides[i:i + SLIDES_PER_BATCH] for i in range(0, len(slides), SLIDES_PER_BATCH)]
    for bi, batch in enumerate(batches):
        slide_indices = [s["slide_index"] for s in batch]
        payload_slides = [_slide_to_prompt_payload(s) for s in batch]
        # 번역 대상 도형이 하나도 없는 배치는 API 호출을 건너뛴다
        if not any(sh["has_chinese"] or ("번역" in (sh["shape_name"] or "")) for s in payload_slides for sh in s["shapes"]):
            if progress_cb:
                progress_cb(bi + 1, len(batches), skipped=True)
            continue

        ref_hint = build_reference_hint(reference_extracted, slide_indices)
        system_prompt = build_system_prompt(
            lang_meta["label"], lang_meta["font"], lang_code, lang_meta["complex_script"],
            reference_hint=ref_hint,
        )
        user_content = json.dumps({"slides": payload_slides}, ensure_ascii=False)

        try:
            raw = llm_clients.call_claude(api_key, model, system_prompt, user_content)
            data = llm_clients.extract_json(raw)
            for item in data.get("shapes", []):
                sid = item.get("shape_id")
                if sid:
                    plan_by_shape[sid] = item
        except Exception as e:
            logger.exception("translation batch failed")
            errors.append({"slides": slide_indices, "error": str(e)})

        if progress_cb:
            progress_cb(bi + 1, len(batches))

    return plan_by_shape, errors
