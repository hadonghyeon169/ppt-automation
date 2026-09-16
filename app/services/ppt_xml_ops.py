"""
culturefi-ppt-translation 스킬 문서의 핵심 함수들을 그대로 이식한 저수준 OOXML 조작 모듈.
(스킬 문서 "핵심 함수" / "볼드체 규칙" / "흰색 도형 감지" / "도형 탐색 방법" 절 참고)

이 모듈은 판단(무엇을 어떻게 번역할지)은 하지 않고, 이미 결정된 번역문·서식 값을
실제 PPTX XML에 정확히 반영하는 역할만 담당한다. 판단은 slide_extractor.py /
translator.py 쪽에서 이루어진다.
"""
import copy
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


def get_shape_cnvpr_id(sp_elem):
    """도형의 실제 OOXML id(p:nvSpPr/p:cNvPr/@id)를 반환. PowerPoint가 슬라이드 안에서
    도형마다 고유하게 부여하는 값이라, 슬라이드 내 다른 도형이 삭제/추가되어도 바뀌지 않는다
    (순번 기반 식별과 달리 안정적). 없으면 None."""
    nv = sp_elem.find(qn('p:nvSpPr'))
    if nv is None:
        return None
    cnv = nv.find(qn('p:cNvPr'))
    if cnv is None:
        return None
    return cnv.get('id')


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
        # 한국어 문법 용어를 의도적으로 원어 유지하면서 나머지는 목표 언어로 번역한
        # "혼합" 런은, 예전엔 여기서 그냥 return해서 a:latin(라틴/키릴 문자용 폰트
        # 슬롯)을 전혀 지정하지 않았다. 원래 순수 한국어 전용이던 도형은 a:ea(동아시아
        # 폰트)만 명시돼 있고 a:latin은 테마 기본값에 의존하는 경우가 흔한데, 번역 후
        # 키릴/라틴 문자가 이 런에 처음 섞여 들어가면서 그동안 방치돼 있던 테마 기본
        # 라틴 폰트가 그대로 노출되어(디자인과 전혀 다른 서체로 렌더링됨) 눈에 띄게
        # 어긋나 보이는 문제가 있었다. PowerPoint/LibreOffice는 유니코드 스크립트별로
        # a:latin/a:ea/a:cs 중 알맞은 슬롯을 문자 단위로 자동 선택해 렌더링하므로,
        # a:ea(한국어 부분)는 그대로 둔 채 a:latin(cs 언어면 a:cs)만 목표 폰트로
        # 맞춰주면 한 런 안에서도 한국어는 기존 폰트, 나머지는 목표 폰트로 올바르게
        # 섞여 렌더링된다.
        if is_complex_script:
            _set_or_create(rPr, qn('a:cs'), font_name, extra={'charset': '0'})
        else:
            _set_or_create(rPr, qn('a:latin'), font_name)
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


def _append_extra_lines(base_run, extra_lines):
    """base_run은 이미 텍스트/폰트/크기/색상까지 다 적용된 상태의 <a:r> — 이걸 그대로
    복제해 각 추가 줄을 <a:br/> + 복제된 run으로 원래 run 뒤에 순서대로 이어붙인다.

    OOXML은 <a:t> 텍스트 안의 리터럴 개행 문자("\\n")를 줄바꿈으로 렌더링하지
    않는다 — 반드시 <a:br/> 요소가 있어야 실제로 줄이 바뀐다. 그런데 번역 모델은
    "예문"/"구조" 박스처럼 원래 한 문단 안에 여러 줄(수동 줄바꿈)이 들어있던 도형을
    번역할 때 translated_text 하나에 줄 구분을 "\\n"으로만 표시해 돌려주는 경우가
    있는데, 예전엔 이 문자열을 그대로 <a:t>에 밀어넣기만 해서 렌더링 시 여러 줄이
    아니라 한 줄로 이어져 도형 옆(또는 밖)으로 삐져나가는 원인이 됐다 (실제 사례:
    문법 템플릿 "예문" 박스 — 번역문이 1)/2)/3) 구분 없이 한 줄로 이어짐).
    fit_font_size_to_box()의 줄 수 계산은 애초에 "\\n" = 새 줄"이라고 가정하고
    높이를 추정하므로, 여기서 실제로 그 가정대로 <a:br/>을 만들어주면 넘침 계산과
    실제 렌더링 결과가 다시 일치하게 된다."""
    anchor = base_run
    for line in extra_lines:
        br = etree.Element(qn('a:br'))
        anchor.addnext(br)
        anchor = br
        new_run = copy.deepcopy(base_run)
        new_t = new_run.find(qn('a:t'))
        if new_t is None:
            new_t = etree.SubElement(new_run, qn('a:t'))
        new_t.text = line
        anchor.addnext(new_run)
        anchor = new_run


def apply_shape_level(sp_elem, translated_text, lang_code, font_name, is_complex=False,
                       force_sz=None, force_bold=None, force_color=None):
    """일반(흰색 아님) 도형 번역 적용. 여러 run은 첫 run으로 합치고 나머지는 비운다.
    force_color("white"|"black"|None): 언어별 정밀 스타일 규칙이 색을 강제 지정한 경우만
    사용. None이면 기존 동작(테마 기본색을 따르도록 solidFill 제거) 그대로."""
    all_runs = list(sp_elem.iter(qn('a:r')))
    if not all_runs:
        return
    lines = (translated_text or "").split("\n")
    t = all_runs[0].find(qn('a:t'))
    if t is not None:
        t.text = lines[0]
    apply_target_font_to_run(all_runs[0], lang_code, font_name, is_complex, force_bold=force_bold)
    if force_sz is not None:
        rPr = all_runs[0].find(qn('a:rPr'))
        if rPr is not None:
            rPr.set('sz', str(int(force_sz)))
    rPr = all_runs[0].find(qn('a:rPr'))
    if rPr is not None:
        if force_color in ('white', 'black'):
            force_run_color(rPr, 'FFFFFF' if force_color == 'white' else '000000')
        else:
            for sf in list(rPr.findall(qn('a:solidFill'))):
                rPr.remove(sf)
    for run in all_runs[1:]:
        t = run.find(qn('a:t'))
        if t is not None:
            t.text = ''
    if len(lines) > 1:
        _append_extra_lines(all_runs[0], lines[1:])


def restore_shape_translation(sp_elem, translated_text, lang_code, font_name, is_complex=False,
                               force_sz=None, force_bold=None, force_color=None):
    """흰색(배경색 강조) 도형 번역 적용 — 기본적으로 solidFill을 건드리지 않아 색상을
    보존한다. force_color가 명시된 경우에만("white"|"black") 색을 명시적으로 덮어쓴다
    (예: 원래 흰색이었지만 정밀 규칙상 검정으로 바꿔야 하는 라벨)."""
    all_runs = list(sp_elem.iter(qn('a:r')))
    if not all_runs:
        return
    lines = (translated_text or "").split("\n")
    t = all_runs[0].find(qn('a:t'))
    if t is not None:
        t.text = lines[0]
    apply_target_font_to_run(all_runs[0], lang_code, font_name, is_complex, force_bold=force_bold)
    if force_sz is not None:
        rPr = all_runs[0].find(qn('a:rPr'))
        if rPr is not None:
            rPr.set('sz', str(int(force_sz)))
    if force_color in ('white', 'black'):
        rPr = all_runs[0].find(qn('a:rPr'))
        if rPr is not None:
            force_run_color(rPr, 'FFFFFF' if force_color == 'white' else '000000')
    for run in all_runs[1:]:
        t = run.find(qn('a:t'))
        if t is not None:
            t.text = ''
    if len(lines) > 1:
        _append_extra_lines(all_runs[0], lines[1:])


def apply_multi_paragraph(sp_elem, line_translations, lang_code, font_name, is_complex=False,
                           force_sz=None, force_color=None):
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
        if force_color in ('white', 'black'):
            rPr = runs[0].find(qn('a:rPr'))
            if rPr is not None:
                force_run_color(rPr, 'FFFFFF' if force_color == 'white' else '000000')
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
    """spPr/xfrm 의 off(x,y), ext(cx,cy)를 EMU 단위로 반환. 없으면 None.
    주의: 이 도형이 그룹(p:grpSp) 안에 있으면 이 값은 슬라이드 절대 좌표가 아니라
    그 그룹의 "자식 좌표계"(chOff/chExt) 기준이다 — 넘침/폰트 계산에는 반드시
    get_shape_absolute_xfrm()을 사용해야 한다."""
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


def _group_xfrm(grp_elem):
    """p:grpSp의 grpSpPr/xfrm에서 off/ext(부모 좌표계 기준 위치·크기)와
    chOff/chExt(자식 도형들이 사용하는 내부 좌표계)를 모두 읽는다. 하나라도
    없으면 이 그룹은 변환을 적용할 수 없으므로 None을 반환한다."""
    grpSpPr = grp_elem.find(qn('p:grpSpPr'))
    if grpSpPr is None:
        return None
    xfrm = grpSpPr.find(qn('a:xfrm'))
    if xfrm is None:
        return None
    off = xfrm.find(qn('a:off'))
    ext = xfrm.find(qn('a:ext'))
    chOff = xfrm.find(qn('a:chOff'))
    chExt = xfrm.find(qn('a:chExt'))
    if off is None or ext is None or chOff is None or chExt is None:
        return None
    return {
        'x': int(off.get('x')), 'y': int(off.get('y')),
        'cx': int(ext.get('cx')), 'cy': int(ext.get('cy')),
        'chx': int(chOff.get('x')), 'chy': int(chOff.get('y')),
        'chcx': int(chExt.get('cx')), 'chcy': int(chExt.get('cy')),
    }


def shape_is_grouped(sp_elem):
    """이 도형이 하나 이상의 p:grpSp 안에 중첩되어 있는지 여부."""
    node = sp_elem.getparent()
    while node is not None:
        if node.tag == qn('p:grpSp'):
            return True
        if node.tag == qn('p:spTree'):
            return False
        node = node.getparent()
    return False


def get_shape_absolute_xfrm(sp_elem):
    """도형의 슬라이드 절대 좌표/크기(EMU)를 반환한다. 그룹(p:grpSp)에 중첩된
    도형은 spPr/xfrm이 그룹의 자식 좌표계(chOff/chExt) 기준이라 그대로 쓰면
    실제 위치·크기와 전혀 다를 수 있다 — 문법 템플릿 박스, 캐릭터+말풍선 조합처럼
    그룹으로 묶인 도형에서 넘침 감지/폰트 크기 계산이 어긋나는 원인이었다.
    조상 그룹들을 안쪽에서 바깥쪽 순서로 순회하며 각 그룹의 (off,ext,chOff,chExt)로
    좌표를 슬라이드 절대 좌표계까지 누적 변환한다."""
    local = get_shape_xfrm(sp_elem)
    if local is None:
        return None
    x, y, cx, cy = local['x'], local['y'], local['cx'], local['cy']

    node = sp_elem.getparent()
    while node is not None:
        if node.tag == qn('p:grpSp'):
            g = _group_xfrm(node)
            if g is None:
                break  # 이 그룹에 좌표 정보가 없으면 더 이상 보정할 수 없음
            scale_x = (g['cx'] / g['chcx']) if g['chcx'] else 1.0
            scale_y = (g['cy'] / g['chcy']) if g['chcy'] else 1.0
            x = g['x'] + (x - g['chx']) * scale_x
            y = g['y'] + (y - g['chy']) * scale_y
            cx = cx * scale_x
            cy = cy * scale_y
        if node.tag == qn('p:spTree'):
            break
        node = node.getparent()

    return {'x': int(x), 'y': int(y), 'cx': int(cx), 'cy': int(cy)}


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


AUDIO_ICON_SIZE_EMU = 365760  # 0.4in — 슬라이드에 삽입되는 오디오 아이콘 크기


def embed_autoplay_audio(slide, audio_path, mime_type, slide_width_emu, slide_height_emu):
    """슬라이드에 오디오 파일을 삽입하고, 슬라이드 진입 시 자동 재생되도록
    <p:timing> 타이밍 트리를 추가한다 (PowerPoint에서 "시작: 자동 실행"으로 오디오를
    넣었을 때와 동일한 구조). python-pptx의 add_movie()는 미디어 파트/관계 생성까지만
    해주고 자동재생 타이밍은 만들어주지 않으므로 이 함수에서 보강한다."""
    slide_w, slide_h = slide_width_emu, slide_height_emu
    size = AUDIO_ICON_SIZE_EMU
    margin = 91440  # 0.1in
    left = slide_w - size - margin
    top = slide_h - size - margin

    movie_shape = slide.shapes.add_movie(audio_path, left, top, size, size, mime_type=mime_type)
    sp_elem = movie_shape._element

    # add_movie()는 <a:videoFile>을 쓰는데, 오디오 파트이므로 스펙에 맞게 <a:audioFile>로 교체한다
    # (PowerPoint/LibreOffice 모두 관대하게 처리하지만 정확한 태그를 쓰는 게 안전하다).
    nvPr = sp_elem.find(qn('p:nvPicPr') + '/' + qn('p:nvPr'))
    if nvPr is not None:
        video_file = nvPr.find(qn('a:videoFile'))
        if video_file is not None:
            video_file.tag = qn('a:audioFile')

    shape_id = sp_elem.find(qn('p:nvPicPr') + '/' + qn('p:cNvPr')).get('id')
    _append_autoplay_timing(slide._element, shape_id)
    return movie_shape


def _append_autoplay_timing(sld_elem, spid):
    """슬라이드 진입 즉시 spid로 지정된 미디어를 1회 자동 재생하는 <p:timing> 트리 삽입.
    CT_Slide 스키마 순서(cSld, clrMapOvr?, transition?, timing?, extLst?)를 지켜
    올바른 위치에 넣는다."""
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    timing_xml = f"""
<p:timing xmlns:p="{P_NS}">
  <p:tnLst>
    <p:par>
      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
        <p:childTnLst>
          <p:seq concurrent="1" nextAc="seek">
            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
              <p:childTnLst>
                <p:par>
                  <p:cTn id="3" fill="hold">
                    <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
                    <p:childTnLst>
                      <p:par>
                        <p:cTn id="4" fill="hold">
                          <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                          <p:childTnLst>
                            <p:par>
                              <p:cTn id="5" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">
                                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                                <p:childTnLst>
                                  <p:cmd type="call" cmd="playFrom(0.0)">
                                    <p:cBhvr>
                                      <p:cTn id="6" dur="indefinite" fill="hold"/>
                                      <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
                                    </p:cBhvr>
                                  </p:cmd>
                                </p:childTnLst>
                              </p:cTn>
                            </p:par>
                          </p:childTnLst>
                        </p:cTn>
                      </p:par>
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
            <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
            <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
          </p:seq>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
  <p:bldLst><p:bldMedia spid="{spid}"/></p:bldLst>
</p:timing>
""".strip()
    new_timing = etree.fromstring(timing_xml.encode('utf-8'))

    # 기존 timing이 있으면(여러 오디오를 순차 삽입하는 경우 등) 제거하고 새로 넣는다
    # — 이 앱은 슬라이드당 오디오 1개만 다루므로 항상 교체가 맞다.
    old_timing = sld_elem.find(qn('p:timing'))
    if old_timing is not None:
        sld_elem.remove(old_timing)

    insert_after_tags = [qn('p:cSld'), qn('p:clrMapOvr'), qn('p:transition')]
    insert_idx = 0
    for i, child in enumerate(sld_elem):
        if child.tag in insert_after_tags:
            insert_idx = i + 1
    sld_elem.insert(insert_idx, new_timing)


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
