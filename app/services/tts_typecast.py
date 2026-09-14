"""
타입캐스트(Typecast) TTS REST API 클라이언트.
공식 문서 기준 (2026-09 확인):
  - GET  https://api.typecast.ai/v1/voices        (보이스 목록)
  - POST https://api.typecast.ai/v1/text-to-speech (음성 합성, 응답은 raw 오디오 바이트)
  - 인증 헤더: X-API-KEY
  - text는 1~2000자 제한 → 슬라이드 스크립트가 길면 호출 측에서 분할해야 한다.
"""
import requests

BASE_URL = "https://api.typecast.ai/v1"
MAX_TEXT_LEN = 2000


class TypecastError(Exception):
    pass


def list_voices(api_key, model=None):
    if not api_key:
        raise TypecastError("타입캐스트 API 키가 설정되어 있지 않습니다.")
    headers = {"X-API-KEY": api_key}
    params = {"model": model} if model else {}
    resp = requests.get(f"{BASE_URL}/voices", headers=headers, params=params, timeout=30)
    if resp.status_code != 200:
        raise TypecastError(f"보이스 목록 조회 실패 ({resp.status_code}): {resp.text[:300]}")
    return resp.json()


def synthesize(api_key, voice_id, text, model="ssfm-v30", language=None, audio_format="mp3",
               emotion_type=None, emotion_intensity=None, out_path=None):
    if not api_key:
        raise TypecastError("타입캐스트 API 키가 설정되어 있지 않습니다.")
    if not text:
        raise TypecastError("합성할 텍스트가 비어 있습니다.")
    if len(text) > MAX_TEXT_LEN:
        raise TypecastError(f"텍스트가 {MAX_TEXT_LEN}자를 초과합니다 ({len(text)}자). 분할이 필요합니다.")

    headers = {"X-API-KEY": api_key, "Content-Type": "application/json"}
    body = {
        "voice_id": voice_id,
        "text": text,
        "model": model,
        "output": {"audio_format": audio_format},
    }
    if language:
        body["language"] = language
    if emotion_type:
        body["prompt"] = {"emotion_type": emotion_type}
        if emotion_intensity is not None:
            body["prompt"]["emotion_intensity"] = emotion_intensity

    resp = requests.post(f"{BASE_URL}/text-to-speech", headers=headers, json=body, timeout=120)
    if resp.status_code != 200:
        raise TypecastError(f"음성 합성 실패 ({resp.status_code}): {resp.text[:300]}")

    if out_path:
        with open(out_path, "wb") as f:
            f.write(resp.content)
        return out_path
    return resp.content


def split_text_for_tts(text, max_len=MAX_TEXT_LEN):
    """긴 스크립트를 문장 단위로 max_len 이하 청크로 분할."""
    if len(text) <= max_len:
        return [text]
    import re
    sentences = re.split(r'(?<=[.!?。！？])\s+', text)
    chunks, current = [], ""
    for s in sentences:
        if len(current) + len(s) + 1 <= max_len:
            current = f"{current} {s}".strip()
        else:
            if current:
                chunks.append(current)
            if len(s) > max_len:
                for i in range(0, len(s), max_len):
                    chunks.append(s[i:i + max_len])
                current = ""
            else:
                current = s
    if current:
        chunks.append(current)
    return chunks
