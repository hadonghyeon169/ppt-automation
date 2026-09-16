"""
PPTX 파일을 열어 슬라이드/도형 구조를 JSON으로 직렬화한다.
이 구조는 (1) DB에 감사(audit) 기록으로 저장되고, (2) 번역 계획을 세우는 LLM 호출의
입력으로 사용된다. 실제 어떤 도형을 어떻게 번역할지의 "판단"은 여기서 하지 않고
원문/서식 사실관계만 뽑아낸다 (판단은 translator.py + culturefi 스킬 규칙 프롬프트가 담당).
"""
from pptx import Presentation
from pptx.oxml.ns import qn
from . import ppt_xml_ops as ops

# "13번/38번 슬라이드"처럼 문법 설명 예문 박스가 있는 슬라이드에만 음성을 넣기 위한
# 판별 마커. 사용자가 확인해준 "초급2 12강" 샘플(44개 슬라이드 전체)에서 이 세 단어가
# 모두 함께 나타나는 슬라이드는 정확히 13번/38번 두 곳뿐이었다(다른 슬라이드는 셋 중
# 어느 것도 없었음 — 오탐/누락 없음). CultureFi 문법 설명 템플릿의 "정의/구조/예문"
# 라벨 3종 세트로 보이며, 이 라벨들은 번역 대상 도형(이름에 "번역"이 붙는 도형)이
# 아니라 순수 한국어 라벨 도형("Rect 0" 등)에 들어있어 번역 후에도 원문 그대로
# 남아있다 — 그래서 슬라이드 번호가 다른 다른 PPT 파일에서도 이 문자열로 안정적으로
# 같은 유형의 슬라이드를 찾아낼 수 있다 (번호가 아니라 내용/구조 기반 판단).
AUDIO_REQUIRED_MARKERS = ("예문", "정의", "구조")


def get_audio_required_slide_indices(extracted):
    """extract_presentation() 결과에서 음성이 필요한(문법 설명 예문) 슬라이드의
    slide_index 목록을 반환한다. 판단 기준: 슬라이드 안의 모든 도형 텍스트를 합쳤을 때
    AUDIO_REQUIRED_MARKERS의 세 마커가 전부 함께 나타나는 슬라이드."""
    indices = []
    for slide in extracted["slides"]:
        combined = "".join(s["full_text"] for s in slide["shapes"])
        if all(marker in combined for marker in AUDIO_REQUIRED_MARKERS):
            indices.append(slide["slide_index"])
    return indices


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
            cnvpr_id = ops.get_shape_cnvpr_id(sp)
            # 그룹(p:grpSp) 안에 중첩된 도형은 로컬 xfrm이 그룹의 자식 좌표계 기준이라
            # 그대로 쓰면 실제 슬라이드 위치/크기와 다르다 — 문법 템플릿 박스나
            # 캐릭터+말풍선 조합처럼 그룹으로 묶인 도형에서 넘침 판정/폰트 계산이
            # 어긋나는 원인이었으므로 절대 좌표로 변환해서 사용한다.
            xfrm = ops.get_shape_absolute_xfrm(sp)
            is_grouped = ops.shape_is_grouped(sp)
            wrap = ops.get_body_pr_wrap(sp)
            align = ops.get_paragraph_align(sp)
            # wrap="none" 도형의 넘침 판정에서 텍스트 상자 내부 여백(lIns+rIns)을
            # 빼지 않아 실제로 넘치는 번역문을 "안 넘침"으로 오판하던 문제가 있었다
            # (자세한 배경은 ppt_xml_ops.get_body_pr_insets 참고) — 여기서 함께 뽑아
            # pptx_pipeline.py의 넘침 보정 계산에 넘겨준다.
            insets_lr_emu = ops.get_body_pr_insets(sp)
            # wrap="square"(줄바꿈) 도형의 높이 적합 판정(fit_font_size_to_box)에서도
            # 같은 이유로 실제 상하 여백(tIns+bIns)이 필요해 함께 뽑는다 — 자세한
            # 배경은 ppt_xml_ops.get_body_pr_insets_tb 참고.
            insets_tb_emu = ops.get_body_pr_insets_tb(sp)
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

            # 도형 식별자: 슬라이드 내 순번이 아니라 도형 고유 XML id(cNvPr/@id) 기반으로 만든다.
            # 순번 기반이면 나중에 다른 도형이 삭제될 때(예: cleanup_duplicate_shapes) 뒤따르는
            # 도형들의 shape_id가 통째로 밀려서, 번역 승인 시점 스냅샷과 최종 검수 시점 재추출
            # 결과가 어긋나는 문제가 있었다 (표면상 "텍스트 불일치"/"도형을 찾을 수 없음"으로 보임).
            stable_id = cnvpr_id if cnvpr_id is not None else f"x{shape_idx}"
            shapes_data.append({
                "shape_id": f"s{slide_idx}_{stable_id}",
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
                "is_grouped": is_grouped,
                "insets_lr_emu": insets_lr_emu,
                "insets_tb_emu": insets_tb_emu,
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
