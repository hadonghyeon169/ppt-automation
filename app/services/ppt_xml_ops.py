"""
culturefi-ppt-translation 스킬 문서의 핵심 함수들을 그대로 이식한 저수준 OOXML 조작 모듈.
(스킬 문서 "핵심 함수" / "볼드체 규칙" / "흰색 도형 감지" / "도형 탐색 방법" 절 참고)

이 모듈은 판단(무엇을 어떻게 번역할지)은 하지 않고, 이미 결정된 번역문·서식 값을
실제 PPTX XML에 정확히 반영하는 역할만 담당한다. 판단은 slide_extractor.py /
translator.py 쪽에서 이루어진다.
"""
from lxml import etree
from pptx.oxml.ns import qn


def has_chinese(text: str) -> bool:
    if not text:
        return False
    return any('一' <= c <= '鿿' for c in text)


def has_korean(text: str) -> bool:
    if not text:
        return False
    return any('가' <= c <= '힣' for c in text)


def get_shape_full_text(sp_elem) -> str:
    return "".join((t.text or "") for t in sp_elem.iter(qn('a:t'))).strip()


def has_bg_color(sp_elem) -> bool:
    """흰색(또는 배경1/밝은1 스킴) 텍스트가 적용된 도형인지 감지.
    스킬 문서의 has_bg_color()를 그대로 따름 — 이름과 달리 '텍스트가 흰색인가'를 검사한다."""
    for run in sp_elem.iter(qn('a:r')):
        rPr = run.find(qn('a:rPr'))
        if rPr is None:
            continue
        sf = rPr.find(qn('a:solidFill'))
        if sf is None:
            continue
        srgb = sf.find(qn('a:srgbClr'))
        if srgb is not None and srgb.get('val', '').upper() == 'FFFFFF':
            return True
        scheme = sf.find(qn('a:schemeClr'))
        if scheme is not None and scheme.get('val', '') in ('bg1', 'bg2', 'lt1', 'lt2'):
            return True
    return False


def iter_all_shapes(slide):
    """GroupShape 내부까지 포함해 슬라이드의 모든 p:sp 도형을 순회한다."""
    return list(slide._element.iter(qn('p:sp')))


def get_shape_name(sp_elem) -> str:
    nv = sp_elem.find(qn('p:nvSpPr'))
    if nv is None:
        return ""
    cnv = nv.find(qn('p:cNvPr'))
    if cnv is None:
        return ""
    return cnv.get('name', '')


def apply_target_font_to_run(run_elem, lang_code, font_name, is_complex_script=False, force_bold=None):
    """스킬 문서 "볼드체 규칙" 절의 apply_target_font_to_run()을 그대로 이식.
    - 한국어가 섞인 텍스트(번역 후 기준) → 볼드 강제 on, 폰트는 원본 유지
    - 순수 목표 언어 텍스트 → 볼드 off, 목표 폰트 적용
    force_bold가 명시되면 그 값을 우선한다 (특수 케이스 대응용).
    """
    rPr = run_elem.find(qn('a:rPr'))
    if rPr is None:
        rPr = etree.SubElement(run_elem, qn('a:rPr'))
        run_elem.insert(0, rPr)
    rPr.set('lang', lang_code)

    t = run_elem.find(qn('a:t'))
    text = t.text if t is not None else ''
    mixed_with_korean = bool(text and has_korean(text))

    if mixed_with_korean:
        rPr.set('b', '1' if force_bold is None else ('1' if force_bold else '0'))
        return

    rPr.set('b', '0' if force_bold is None else ('1' if force_bold else '0'))
    if is_complex_script:
        _set_or_create(rPr, qn('a:cs'), font_name, extra={'charset': '0'})
        for tag in (qn('a:latin'), qn('a:ea')):
            elem = rPr.find(tag)
            if elem is not None:
                elem.set('typeface', 'Noto Sans KR')
    else:
        for tag in (qn('a:latin'), qn('a:ea')):
            _set_or_create(rPr, tag, font_name)


def _set_or_create(rPr, tag, typeface, extra=None):
    elem = rPr.find(tag)
    if elem is None:
        elem = etree.SubElement(rPr, tag)
    elem.set('typeface', typeface)
    if extra:
        for k, v in extra.items():
            elem.set(k, v)


def apply_shape_level(sp_elem, translated_text, lang_code, font_name, is_complex=False,
                       force_sz=None, force_bold=None):
    """일반(흰색 아님) 도형 번역 적용. 여러 run은 첫 run으로 합치고 나머지는 비운다."""
    all_runs = list(sp_elem.iter(qn('a:r')))
    if not all_runs:
        return
    t = all_runs[0].find(qn('a:t'))
    if t is not None:
        t.text = translated_text
    apply_target_font_to_run(all_runs[0], lang_code, font_name, is_complex, force_bold=force_bold)
    if force_sz is not None:
        rPr = all_runs[0].find(qn('a:rPr'))
        if rPr is not None:
            rPr.set('sz', str(int(force_sz)))
    rPr = all_runs[0].find(qn('a:rPr'))
    if rPr is not None:
        for sf in list(rPr.findall(qn('a:solidFill'))):
            rPr.remove(sf)
    for run in all_runs[1:]:
        t = run.find(qn('a:t'))
        if t is not None:
            t.text = ''


def restore_shape_translation(sp_elem, translated_text, lang_code, font_name, is_complex=False,
                               force_sz=None, force_bold=None):
    """흰색(배경색 강조) 도형 번역 적용 — solidFill을 건드리지 않아 색상을 보존한다."""
    all_runs = list(sp_elem.iter(qn('a:r')))
    if not all_runs:
        return
    t = all_runs[0].find(qn('a:t'))
    if t is not None:
        t.text = translated_text
    apply_target_font_to_run(all_runs[0], lang_code, font_name, is_complex, force_bold=force_bold)
    if force_sz is not None:
        rPr = all_runs[0].find(qn('a:rPr'))
        if rPr is not None:
            rPr.set('sz', str(int(force_sz)))
    for run in all_runs[1:]:
        t = run.find(qn('a:t'))
        if t is not None:
            t.text = ''


def apply_multi_paragraph(sp_elem, line_translations, lang_code, font_name, is_complex=False, force_sz=None):
    """여러 <a:p> 문단으로 나뉜 도형(예: 강의 소개) 번역 — 문단별로 번역문을 배치."""
    txBody = sp_elem.find(qn('p:txBody'))
    if txBody is None:
        return
    paras = txBody.findall(qn('a:p'))
    for i, p in enumerate(paras):
        runs = p.findall(qn('a:r'))
        if not runs:
            continue
        line_text = line_translations[i] if i < len(line_translations) else ''
        t = runs[0].find(qn('a:t'))
        if t is not None:
            t.text = line_text
        apply_target_font_to_run(runs[0], lang_code, font_name, is_complex)
        if force_sz is not None:
            rPr = runs[0].find(qn('a:rPr'))
            if rPr is not None:
                rPr.set('sz', str(int(force_sz)))
        for run in runs[1:]:
            t = run.find(qn('a:t'))
            if t is not None:
                t.text = ''


def force_run_color(rPr, hex_color):
    """텍스트 색상 강제 고정. OOXML 스키마 순서(ln → fill → effectLst → ... )를 지켜 <a:ln> 뒤에 삽입."""
    for sf in list(rPr.findall(qn('a:solidFill'))):
        rPr.remove(sf)
    sf = etree.SubElement(rPr, qn('a:solidFill'))
    srgb = etree.SubElement(sf, qn('a:srgbClr'))
    srgb.set('val', hex_color)
    rPr.remove(sf)
    insert_idx = 0
    ln = rPr.find(qn('a:ln'))
    if ln is not None:
        insert_idx = list(rPr).index(ln) + 1
    rPr.insert(insert_idx, sf)


def get_body_pr_wrap(sp_elem):
    txBody = sp_elem.find(qn('p:txBody'))
    if txBody is None:
        return None
    bodyPr = txBody.find(qn('a:bodyPr'))
    if bodyPr is None:
        return None
    return bodyPr.get('wrap')


def set_body_pr_wrap_square_autofit(sp_elem):
    """인접 라벨-내용 쌍 도형용: wrap을 square로 바꾸고 높이만 늘어나게(spAutoFit) 설정."""
    txBody = sp_elem.find(qn('p:txBody'))
    if txBody is None:
        return
    bodyPr = txBody.find(qn('a:bodyPr'))
    if bodyPr is None:
        bodyPr = etree.SubElement(txBody, qn('a:bodyPr'))
    bodyPr.set('wrap', 'square')
    for tag in (qn('a:noAutofit'), qn('a:normAutofit')):
        el = bodyPr.find(tag)
        if el is not None:
            bodyPr.remove(el)
    if bodyPr.find(qn('a:spAutoFit')) is None:
        etree.SubElement(bodyPr, qn('a:spAutoFit'))


def get_shape_xfrm(sp_elem):
    """spPr/xfrm 의 off(x,y), ext(cx,cy)를 EMU 단위로 반환. 없으면 None."""
    spPr = sp_elem.find(qn('p:spPr'))
    if spPr is None:
        return None
    xfrm = spPr.find(qn('a:xfrm'))
    if xfrm is None:
        return None
    off = xfrm.find(qn('a:off'))
    ext = xfrm.find(qn('a:ext'))
    if off is None or ext is None:
        return None
    return {
        'x': int(off.get('x')), 'y': int(off.get('y')),
        'cx': int(ext.get('cx')), 'cy': int(ext.get('cy')),
    }


def set_shape_xfrm(sp_elem, x=None, y=None, cx=None, cy=None):
    spPr = sp_elem.find(qn('p:spPr'))
    if spPr is None:
        return
    xfrm = spPr.find(qn('a:xfrm'))
    if xfrm is None:
        return
    off = xfrm.find(qn('a:off'))
    ext = xfrm.find(qn('a:ext'))
    if off is not None:
        if x is not None:
            off.set('x', str(int(x)))
        if y is not None:
            off.set('y', str(int(y)))
    if ext is not None:
        if cx is not None:
            ext.set('cx', str(int(cx)))
        if cy is not None:
            ext.set('cy', str(int(cy)))


def get_paragraph_align(sp_elem):
    txBody = sp_elem.find(qn('p:txBody'))
    if txBody is None:
        return None
    p = txBody.find(qn('a:p'))
    if p is None:
        return None
    pPr = p.find(qn('a:pPr'))
    if pPr is None:
        return None
    return pPr.get('algn')


def cleanup_duplicate_shapes(prs):
    """중복 도형 정리 (오류 대응) — 스킬 문서 "중복 도형 정리" 절 그대로."""
    removed = 0
    for slide in prs.slides:
        spTree = slide._element.find('.//' + qn('p:spTree'))
        seen = set()
        to_remove = []
        for sp in spTree.findall(qn('p:sp')):
            full = get_shape_full_text(sp)
            if not full or not has_bg_color(sp):
                continue
            if has_chinese(full):
                to_remove.append(sp)
            else:
                if full in seen:
                    to_remove.append(sp)
                else:
                    seen.add(full)
        for sp in to_remove:
            spTree.remove(sp)
            removed += 1
    return removed
