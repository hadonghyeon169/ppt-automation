"""
번역 계획(translation plan) 생성 오케스트레이터.
슬라이드를 배치로 묶어 Claude API에 보내고, 도형별 번역 결정을 받아온다.
실제 XML 반영은 pptx_pipeline.py가 담당한다.
"""
import json
import logging
import re

from . import llm_clients
from .skill_prompt import build_system_prompt
from .slide_extractor import classify_shape_type

logger = logging.getLogger(__name__)

SLIDES_PER_BATCH = 3
# 4에서 3으로 낮췄다: 언어별 정밀 스타일 규칙(force_color/force_font_size_pt 등)이
# 늘면서 shape당 출력 JSON이 길어져, 도형이 많은 슬라이드가 낀 배치에서
# max_tokens(16000)를 넘겨 응답이 잘리는 경우가 있었다 (Unterminated string 오류).

# 완성본(참고 PPT) 대조 시 매칭 기준 — 단어 단위 자카드 유사도가 이 값 미만이면
# "같은 내용의 슬라이드"로 보지 않고 힌트에서 제외한다.
# 처음에는 문자 단위 difflib.SequenceMatcher.ratio()로 구현했는데, "정의/구조/예문"
# 같은 공통 라벨 단어나 조사("을/를/이/가" 등)가 겹치는 것만으로도 서로 전혀 다른
# 문법 슬라이드끼리 ratio 0.5~0.8까지 나와 잘못 매칭되는 문제가 실측으로 확인됐다
# (예: "동사 + -았/었- + -어요" vs "형용사 + -(으)ㄴ 것 같다"처럼 완전히 다른 문법인데도
# 라벨/기호 겹침만으로 높은 유사도가 나옴). 단어(어절) 단위 자카드로 바꾸면 이런 오탐이
# 0.1~0.2대로 확실히 낮게 나오고, 실제로 같은 문구인 "책을 펴주세요!", 강의 소개 문단
# 등은 0.6 이상으로 뚜렷이 구분됨 — 그래서 자카드 기준 0.45를 임계값으로 잡았다.
# 각 강의 고유 문법 설명(정의/구조/예문)처럼 강마다 내용 자체가 다른 슬라이드는 애초에
# 참고 PPT에 "같은 내용"이 없으므로 매칭되지 않는 것이 정상이다 (그런 슬라이드의
# 폰트 크기/색상 규칙은 skill_prompt.py의 고정 규칙이 담당한다).
REFERENCE_MATCH_MIN_RATIO = 0.45

_TOKEN_SPLIT_RE = re.compile(r"[\s,.!?;:()\[\]{}'\"~\-–—·•、。！？：；]+")


def _tokenize_for_match(text):
    return {t for t in _TOKEN_SPLIT_RE.split(text) if len(t) >= 2}


def _word_jaccard(text_a, text_b):
    ta, tb = _tokenize_for_match(text_a), _tokenize_for_match(text_b)
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / len(ta | tb)


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


def _korean_signature(slide):
    """슬라이드를 대표하는 "순수 한국어 원문" 문자열 — 참고 PPT와 매칭할 때 쓰는 키.
    has_korean=True이면서 has_chinese=False인 도형(번역 대상이 아닌 한국어 원문 도형)의
    텍스트만 모은다. 번역 대상 도형(번역문/중국어)은 강마다 다르므로 매칭 기준에서
    제외 — 한국어 원문이야말로 "같은 슬라이드인지"를 가장 안정적으로 판별하는 값이다
    (완성본 PPT도 핵심 규칙상 한국어 원문 도형은 절대 수정하지 않으므로 그대로 남아있음)."""
    parts = [
        sh["full_text"] for sh in slide["shapes"]
        if sh.get("has_korean") and not sh.get("has_chinese") and sh["full_text"]
    ]
    return "\n".join(parts)


def _find_best_reference_slides(reference_extracted, current_slide, top_n=1,
                                 min_ratio=REFERENCE_MATCH_MIN_RATIO):
    """current_slide(지금 번역 중인 슬라이드)와 한국어 원문 내용이 가장 비슷한 참고 PPT
    슬라이드를 찾는다.

    예전에는 슬라이드 번호를 그대로 1:1로 맞춰서 대조했는데, 강마다 슬라이드 수가
    다르면(예: 어떤 과는 40장, 어떤 과는 50장) 번호가 어긋나 전혀 무관한 슬라이드끼리
    비교하는 문제가 있었다. 슬라이드 번호가 아니라 한국어 원문 텍스트의 유사도로
    "내용이 같거나 비슷한 슬라이드"를 찾도록 바꿨다 — 표지, 강의 소개, "책을 펴주세요!",
    "사용 문법"/"사용 단어" 라벨, 마무리 인사말처럼 강마다 반복되는 공통 슬라이드는
    번호가 달라도 한국어 텍스트가 거의 동일하므로 정확히 매칭된다."""
    sig = _korean_signature(current_slide)
    if not sig.strip():
        return []
    # 참고 PPT 슬라이드별 서명은 배치 호출마다 반복 계산하면 낭비이므로 최초 1회만
    # 계산해서 reference_extracted 딕셔너리 자체에 캐싱해둔다 (같은 실행 동안 재사용).
    cache = reference_extracted.setdefault("_korean_sig_cache", {})
    scored = []
    for ref_slide in reference_extracted["slides"]:
        ref_idx = ref_slide["slide_index"]
        ref_sig = cache.get(ref_idx)
        if ref_sig is None:
            ref_sig = _korean_signature(ref_slide)
            cache[ref_idx] = ref_sig
        if not ref_sig.strip():
            continue
        ratio = _word_jaccard(sig, ref_sig)
        if ratio >= min_ratio:
            scored.append((ratio, ref_slide))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [s for _, s in scored[:top_n]]


def build_reference_hint(reference_extracted, extracted, slide_indices, max_chars=3000):
    """참고 PPT(같은 시리즈의 완성본)에서 지금 배치의 슬라이드들과 내용이 비슷한
    슬라이드를 찾아 그 실제 완성본 텍스트(번역 결과 포함)를 LLM 힌트로 만든다.
    매칭은 슬라이드 번호가 아니라 한국어 원문 내용 유사도 기준(_find_best_reference_slides
    참고) — 강마다 페이지 수가 달라도 올바르게 대응된다."""
    if not reference_extracted:
        return None
    slides_by_index = {s["slide_index"]: s for s in extracted["slides"]}
    chunks = []
    seen_ref_indices = set()
    for si in slide_indices:
        current_slide = slides_by_index.get(si)
        if not current_slide:
            continue
        for ref_slide in _find_best_reference_slides(reference_extracted, current_slide):
            ref_idx = ref_slide["slide_index"]
            if ref_idx in seen_ref_indices:
                continue
            seen_ref_indices.add(ref_idx)
            texts = [f'- {sh["shape_name"]}: {sh["full_text"]}' for sh in ref_slide["shapes"] if sh["full_text"]]
            if texts:
                chunks.append(
                    f"[참고(완성본) 슬라이드 {ref_idx + 1} — 현재 슬라이드 {si + 1}과 내용 유사]\n"
                    + "\n".join(texts)
                )
    hint = "\n\n".join(chunks)
    return hint[:max_chars] if hint else None


def _run_batch(api_key, model, extracted, lang_code, lang_meta, slide_indices, reference_extracted=None):
    """지정된 슬라이드 인덱스 목록(배치 1개)에 대해 Claude를 1회 호출하고 shapes 배열을
    반환한다. plan_translation(전체 실행)과 retry_failed_batches(부분 재시도)가 공유하는
    핵심 호출 로직 — 배치를 어떻게 나누고 실패를 어떻게 다루는지는 호출자가 결정한다."""
    slides_by_index = {s["slide_index"]: s for s in extracted["slides"]}
    batch = [slides_by_index[i] for i in slide_indices if i in slides_by_index]
    payload_slides = [_slide_to_prompt_payload(s) for s in batch]

    ref_hint = build_reference_hint(reference_extracted, extracted, slide_indices)
    system_prompt = build_system_prompt(
        lang_meta["label"], lang_meta["font"], lang_code, lang_meta["complex_script"],
        reference_hint=ref_hint,
    )
    user_content = json.dumps({"slides": payload_slides}, ensure_ascii=False)

    raw = llm_clients.call_claude(api_key, model, system_prompt, user_content)
    data = llm_clients.extract_json(raw)
    return data.get("shapes", [])


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
        # 텍스트가 있는 도형이 하나도 없는(완전히 빈) 배치만 API 호출을 건너뛴다.
        # 예전에는 has_chinese/"번역" 이름 조건으로도 걸러냈는데, 이 조건에 안 걸리는
        # "순수 한국어 도형(원문 예문, 말풍선, 정의 박스 등)"이 실제로는 번역 대상인
        # 경우가 많아서 해당 슬라이드가 통째로 한 번도 LLM에 보내지지 않고 원문 그대로
        # 남는 사고가 있었다 (검수 단계에서도 플래그 없이 조용히 넘어감). 텍스트가
        # 하나라도 있으면 항상 LLM 판단(translate/skip)을 받도록 바꾼다.
        if not any(s["shapes"] for s in payload_slides):
            if progress_cb:
                progress_cb(bi + 1, len(batches), skipped=True)
            continue

        try:
            items = _run_batch(api_key, model, extracted, lang_code, lang_meta, slide_indices, reference_extracted)
            for item in items:
                sid = item.get("shape_id")
                if sid:
                    plan_by_shape[sid] = item
        except Exception as e:
            logger.exception("translation batch failed")
            errors.append({"slides": slide_indices, "error": str(e)})

        if progress_cb:
            progress_cb(bi + 1, len(batches))

    return plan_by_shape, errors


def retry_failed_batches(api_key, model, extracted, lang_code, lang_meta, failed_batches,
                          reference_extracted=None, progress_cb=None):
    """failed_batches: [[slide_idx, ...], ...] — 직전 실행에서 실패했던 배치들만 다시 호출한다.
    성공했던 배치는 API를 다시 호출하지 않으므로 토큰 낭비가 없다.
    반환: (plan_updates: shape_id -> plan dict, errors: 이번에도 실패한 배치 목록, 이전과 동일 형식)"""
    plan_updates = {}
    errors = []
    for bi, slide_indices in enumerate(failed_batches):
        try:
            items = _run_batch(api_key, model, extracted, lang_code, lang_meta, slide_indices, reference_extracted)
            for item in items:
                sid = item.get("shape_id")
                if sid:
                    plan_updates[sid] = item
        except Exception as e:
            logger.exception("translation batch retry failed")
            errors.append({"slides": slide_indices, "error": str(e)})
        if progress_cb:
            progress_cb(bi + 1, len(failed_batches))
    return plan_updates, errors
