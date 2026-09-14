"""
PPTX 파일을 열어 슬라이드/도형 구조를 JSON으로 직렬화한다.
이 구조는 (1) DB에 감사(audit) 기록으로 저장되고, (2) 번역 계획을 세우는 LLM 호출의
입력으로 사용된다. 실제 어떤 도형을 어떻게 번역할지의 "판단"은 여기서 하지 않고
원문/서식 사실관계만 뽑아낸다 (판단은 translator.py + culturefi 스킬 규칙 프롬프트가 담당).
"""
from pptx import Presentation
from pptx.oxml.ns import qn
from . import ppt_xml_ops as ops


def extract_presentation(pptx_path):
    prs = Presentation(pptx_path)
    slide_w, slide_h = prs.slide_width, prs.slide_height
    slides_data = []
    for slide_idx, slide in enumerate(prs.slides):
        shapes_data = []
        for shape_idx, sp in enumerate(ops.iter_all_shapes(slide)):
            full_text = ops.get_shape_full_text(sp)
            if not full_text:
                continue
            name = ops.get_shape_name(sp)
            xfrm = ops.get_shape_xfrm(sp)
            wrap = ops.get_body_pr_wrap(sp)
            align = ops.get_paragraph_align(sp)
            is_white = ops.has_bg_color(sp)
            runs = list(sp.iter(qn('a:r')))
            bold_flags = []
            sizes = []
            for r in runs:
                rPr = r.find(qn('a:rPr'))
                if rPr is not None:
                    b = rPr.get('b')
                    bold_flags.append(b == '1')
                    sz = rPr.get('sz')
                    if sz:
                        sizes.append(int(sz) / 100.0)
            txBody = sp.find(qn('p:txBody'))
            para_count = len(txBody.findall(qn('a:p'))) if txBody is not None else 0
            para_texts = []
            if txBody is not None:
                for p in txBody.findall(qn('a:p')):
                    ptext = "".join((t.text or "") for t in p.iter(qn('a:t')))
                    para_texts.append(ptext)

            shapes_data.append({
                "shape_id": f"s{slide_idx}_{shape_idx}",
                "shape_name": name,
                "full_text": full_text,
                "paragraph_texts": para_texts,
                "paragraph_count": para_count,
                "has_chinese": ops.has_chinese(full_text),
                "has_korean": ops.has_korean(full_text),
                "is_white_text": is_white,
                "bold_flags": bold_flags,
                "font_sizes_pt": sizes,
                "wrap": wrap,
                "align": align,
                "xfrm_emu": xfrm,
            })
        slides_data.append({
            "slide_index": slide_idx,
            "shapes": shapes_data,
        })
    return {
        "slide_width_emu": slide_w,
        "slide_height_emu": slide_h,
        "slide_count": len(slides_data),
        "slides": slides_data,
    }


def classify_shape_type(shape):
    """스킬 문서 "PPT 파일 유형" 절의 A/B/C 판단을 1차로 근사.
    최종 판단(특히 B유형의 "대응 중국어 도형 존재 여부")은 슬라이드 전체 컨텍스트가
    필요하므로 translator.py의 LLM 판단 단계에서 보정된다. 여기서는 힌트만 제공."""
    name = shape["shape_name"] or ""
    has_zh = shape["has_chinese"]
    has_ko = shape["has_korean"]
    name_has_translation_marker = "번역" in name

    if name_has_translation_marker and has_zh:
        return "A"
    if name_has_translation_marker and has_ko and not has_zh:
        return "B"
    if not name_has_translation_marker and has_zh:
        return "C"
    return "N"  # 번역 대상 아님 (한국어 원문 등)
