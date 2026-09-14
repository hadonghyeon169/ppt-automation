"""
3단계(번역 검수) / 5단계(음성·최종 PPT 검수) 자동 비교 로직.
"AI 자동 비교 + 최종 사람 확인" 방식으로 동작 — 여기서 나온 flag들은 review_flags 테이블에
저장되고, 사람이 화면에서 최종 승인 버튼을 눌러야 다음 단계로 진행된다.
"""
import json
import io
import requests

from . import llm_clients
from .ppt_xml_ops import has_chinese, has_korean

QA_SYSTEM_PROMPT = """\
당신은 한국어→목표언어 번역 품질 검수자입니다. 컬처파이 한국어학원의 KIIP 강의 PPT
번역 결과를 검수합니다. 아래 항목 쌍(한국어 원문, 번역문) 목록을 보고 각 항목에 대해:
- 의미가 정확히 전달되었는지 (문법 설명이라면 특히 정확해야 함)
- 원문에 없는 내용이 추가되거나 빠지지 않았는지
- 고유명사가 임의로 번역되지 않았는지
확인하고, 문제가 있는 항목만 골라 이유와 함께 보고하세요. 문제 없으면 배열에 포함하지 마세요.
반드시 JSON만 출력하세요:
{"issues": [{"shape_id": "...", "severity": "warning"|"error", "issue": "<한국어로 구체적 이유>"}]}
"""


def run_translation_qa(api_key, model, records, batch_size=20):
    """records: [{shape_id, slide_index, shape_name, source_korean, translated_text}, ...]"""
    flags = []
    checkable = [r for r in records if r.get("source_korean") and r.get("translated_text")]
    for i in range(0, len(checkable), batch_size):
        batch = checkable[i:i + batch_size]
        payload = [
            {"shape_id": r["shape_id"], "korean": r["source_korean"], "translated": r["translated_text"]}
            for r in batch
        ]
        try:
            raw = llm_clients.call_gpt(api_key, model, QA_SYSTEM_PROMPT, json.dumps(payload, ensure_ascii=False))
            data = llm_clients.extract_json(raw)
            by_id = {r["shape_id"]: r for r in batch}
            for issue in data.get("issues", []):
                r = by_id.get(issue.get("shape_id"))
                if not r:
                    continue
                flags.append({
                    "slide_index": r["slide_index"], "shape_name": r["shape_name"],
                    "source_text": r["source_korean"], "translated_text": r["translated_text"],
                    "issue": issue.get("issue", ""), "severity": issue.get("severity", "warning"),
                })
        except Exception as e:
            flags.append({
                "slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                "issue": f"AI 검수 호출 실패 (배치 {i // batch_size + 1}): {e}", "severity": "warning",
            })
    return flags


def compare_text_snapshots(current_extracted, snapshot_records):
    """번역 승인 시점 스냅샷(snapshot_records, DB의 slide_texts)과 현재 파일의 텍스트를 비교.
    5단계 "지금 만들어진 PPT와 비교 후 100% 일치" 요구사항의 결정론적 구현체.
    """
    current_map = {}
    for slide in current_extracted["slides"]:
        for sh in slide["shapes"]:
            current_map[sh["shape_id"]] = sh["full_text"]

    mismatches = []
    for rec in snapshot_records:
        sid = rec["shape_id"]
        expected = rec["translated_text"]
        actual = current_map.get(sid)
        if actual is None:
            mismatches.append({
                "shape_id": sid, "slide_index": rec["slide_index"], "shape_name": rec["shape_name"],
                "expected": expected, "actual": None, "reason": "도형을 찾을 수 없음 (삭제되었거나 구조 변경됨)",
            })
        elif actual.strip() != (expected or "").strip():
            mismatches.append({
                "shape_id": sid, "slide_index": rec["slide_index"], "shape_name": rec["shape_name"],
                "expected": expected, "actual": actual, "reason": "텍스트 불일치",
            })
    match_ratio = 1.0 - (len(mismatches) / len(snapshot_records)) if snapshot_records else 1.0
    return mismatches, match_ratio


def transcribe_audio_openai(api_key, file_path, model="whisper-1", language_hint=None):
    """OpenAI Whisper로 오디오를 텍스트로 변환 (5단계 음성 검수 보조용, 선택 기능)."""
    if not api_key:
        raise RuntimeError("OpenAI API 키가 설정되어 있지 않습니다.")
    url = "https://api.openai.com/v1/audio/transcriptions"
    headers = {"Authorization": f"Bearer {api_key}"}
    with open(file_path, "rb") as f:
        files = {"file": (file_path.split("/")[-1], f, "audio/mpeg")}
        data = {"model": model}
        if language_hint:
            data["language"] = language_hint
        resp = requests.post(url, headers=headers, files=files, data=data, timeout=180)
    if resp.status_code != 200:
        raise RuntimeError(f"오디오 전사 실패 ({resp.status_code}): {resp.text[:300]}")
    return resp.json().get("text", "")


def estimate_speaking_seconds(text, chars_per_minute=380):
    """언어에 따라 편차가 크므로 매우 대략적인 근사치 — 극단적 불일치(무음/누락) 감지용."""
    if not text:
        return 0.0
    return max(len(text) / chars_per_minute * 60.0, 0.5)
