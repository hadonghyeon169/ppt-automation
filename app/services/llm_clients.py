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


def call_claude(api_key, model, system_prompt, user_content, max_tokens=8000, temperature=0.2):
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
        "temperature": temperature,
        "system": system_prompt,
        "messages": [{"role": "user", "content": user_content}],
    }
    resp = requests.post(ANTHROPIC_URL, headers=headers, json=payload, timeout=180)
    if resp.status_code != 200:
        raise LLMError(f"Claude API 오류 ({resp.status_code}): {resp.text[:500]}")
    data = resp.json()
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
