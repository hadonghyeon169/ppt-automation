"""
Typecast로 생성된 슬라이드별 오디오 파일을 실제 PPTX에 삽입하는 단계.

이전까지는 TTS 오디오가 디스크(project audio 폴더)에만 저장되고 PPTX 파일에는
전혀 반영되지 않았다 (검수 화면에서 재생 길이/전사 비교용으로만 쓰임). 이 모듈이
그 빠진 단계를 채운다: 번역이 반영된 PPTX를 열어 슬라이드별로 대응하는 오디오를
삽입하고, 슬라이드 진입 시 자동 재생되도록 설정한 뒤 별도 파일로 저장한다
(원본 translated 파일은 건드리지 않는다).
"""
import os
from pptx import Presentation

from . import ppt_xml_ops as ops

_MIME_BY_EXT = {
    ".mp3": "audio/mpeg",
    ".wav": "audio/x-wav",
    ".m4a": "audio/mp4",
    ".ogg": "audio/ogg",
}


def _mime_for_path(path):
    ext = os.path.splitext(path)[1].lower()
    return _MIME_BY_EXT.get(ext, "audio/mpeg")


def embed_audio_into_pptx(translated_pptx_path, audio_assets, out_path):
    """audio_assets: repo.list_audio_assets() 결과 (slide_index, status, file_path 포함).
    status=='done'이고 file_path가 실제 존재하는 것만 삽입한다.
    반환: (embedded_count, skipped) — skipped는 [{slide_index, reason}] 형태."""
    prs = Presentation(translated_pptx_path)
    slide_w, slide_h = prs.slide_width, prs.slide_height

    by_slide = {a["slide_index"]: a for a in audio_assets}
    embedded = 0
    skipped = []

    for slide_idx, slide in enumerate(prs.slides):
        asset = by_slide.get(slide_idx)
        if not asset:
            continue
        if asset["status"] != "done":
            if asset["status"] == "error":
                skipped.append({"slide_index": slide_idx, "reason": f"음성 생성 실패: {asset.get('error_message')}"})
            continue
        file_path = asset.get("file_path")
        if not file_path or not os.path.exists(file_path):
            if asset.get("script_text"):  # 스크립트가 있었는데 파일이 없으면 진짜 문제
                skipped.append({"slide_index": slide_idx, "reason": "오디오 파일을 찾을 수 없음"})
            continue

        try:
            ops.embed_autoplay_audio(slide, file_path, _mime_for_path(file_path), slide_w, slide_h)
            embedded += 1
        except Exception as e:
            skipped.append({"slide_index": slide_idx, "reason": f"오디오 삽입 중 오류: {e}"})

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    prs.save(out_path)
    return embedded, skipped
