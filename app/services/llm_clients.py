"""
Claude / OpenAI(GPT) API를 SDK 없이 순수 REST 호출로 감싼 얇은 클라이언트.
(배포 환경에 어떤 패키지가 있든 requests만으로 동작하도록 — 이식성을 위해 의도적으로 SDK 미사용)
"""
import json
import requests

ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
OPENAI_URL = "https://api.openai.com/v1/chat/completions"


class LLMError(Exception):
    pass


def call_claude(api_key, model, system_prompt, user_content, max_tokens=24000, temperature=None):
    if not api_key:
        raise LLMError("Claude API 키가 설정되어 있지 않습니다. 설정 페이지에서 등록해주세요.")
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    payload = {
        "model": model,
        "max_tokens": max_tokens,
        "system": system_prompt,
        "messages": [{"role": "user", "content": user_content}],
    }
    # claude-sonnet-5 등 최신 모델은 temperature 키가 요청에 "존재하기만 해도"
    # 400 오류("temperature is deprecated for this model")를 낸다 (값이 아니라
    # 필드 존재 여부로 판단). 그래서 명시적으로 값이 주어졌을 때만 페이로드에 넣는다.
    if temperature is not None:
        payload["temperature"] = temperature
    resp = requests.post(ANTHROPIC_URL, headers=headers, json=payload, timeout=180)
    if resp.status_code != 200:
        raise LLMError(f"Claude API 오류 ({resp.status_code}): {resp.text[:500]}")
    data = resp.json()
    # stop_reason이 "max_tokens"면 응답이 중간에 잘린 것이다 (JSON이 깨져서
    # "Unterminated string"/"Expecting ',' delimiter" 같은 혼란스러운 파싱
    # 오류로 이어지기 전에 여기서 바로 명확한 원인을 알려준다).
    if data.get("stop_reason") == "max_tokens":
        raise LLMError(
            f"Claude 응답이 max_tokens({max_tokens}) 제한에 걸려 중간에 잘렸습니다. "
            "배치당 슬라이드 수를 줄이거나 max_tokens를 더 늘려야 합니다."
        )
    parts = data.get("content", [])
    text = "".join(p.get("text", "") for p in parts if p.get("type") == "text")
    return text


def call_gpt(api_key, model, system_prompt, user_content, max_tokens=4000, temperature=0.1, json_mode=True):
    if not api_key:
        raise LLMError("OpenAI(GPT) API 키가 설정되어 있지 않습니다. 설정 페이지에서 등록해주세요.")
    headers = {
        "Authorization": f"Bearer {api_key}",
        "content-type": "application/json",
    }
    payload = {
        "model": model,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content},
        ],
    }
    if json_mode:
        payload["response_format"] = {"type": "json_object"}
    resp = requests.post(OPENAI_URL, headers=headers, json=payload, timeout=180)
    if resp.status_code != 200:
        raise LLMError(f"OpenAI API 오류 ({resp.status_code}): {resp.text[:500]}")
    data = resp.json()
    return data["choices"][0]["message"]["content"]


def extract_json(text):
    """모델 응답에서 JSON 블록만 안전하게 뽑아낸다 (마크다운 코드펜스 대응)."""
    text = text.strip()
    if text.startswith("```"):
        text = text.split("```", 2)
        text = text[1] if len(text) > 1 else text[0]
        if text.startswith("json"):
            text = text[4:]
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1:
            return json.loads(text[start:end + 1])
        start = text.find("[")
        end = text.rfind("]")
        if start != -1 and end != -1:
            return json.loads(text[start:end + 1])
        raise
