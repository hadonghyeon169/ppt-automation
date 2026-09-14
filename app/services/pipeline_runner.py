"""
파이프라인 각 단계를 실제로 실행하는 오케스트레이터.
Flask 요청 스레드가 아닌 별도 백그라운드 스레드에서 실행되므로, Flask 앱 컨텍스트에
의존하지 않고 모든 값을 인자로 명시적으로 받는다 (db_path, 디렉토리, API 키 등).
"""
import os
import logging
import traceback

from .. import repo
from . import slide_extractor, translator, pptx_pipeline, render, qa_reviewer, tts_typecast

logger = logging.getLogger(__name__)


def _resolve_api_key(db_path, env_var, key_name):
    keys = repo.get_api_keys(db_path)
    return keys.get(key_name) or os.environ.get(env_var) or ""


def get_effective_keys(db_path):
    return {
        "anthropic": _resolve_api_key(db_path, "ANTHROPIC_API_KEY", "anthropic_api_key"),
        "openai": _resolve_api_key(db_path, "OPENAI_API_KEY", "openai_api_key"),
        "typecast": _resolve_api_key(db_path, "TYPECAST_API_KEY", "typecast_api_key"),
    }


def _project_dir(projects_dir, project_id):
    d = os.path.join(projects_dir, str(project_id))
    os.makedirs(d, exist_ok=True)
    return d


# ────────────────────────────────────────────────────────────────────────
# 2단계: 번역
# ────────────────────────────────────────────────────────────────────────
def run_translation_stage(db_path, projects_dir, project_id, languages, anthropic_model):
    try:
        repo.update_project(db_path, project_id, stage="translating", progress=0,
                             status_message="PPT 구조 분석 중...")
        repo.add_event(db_path, project_id, "번역 단계 시작")
        repo.clear_review_flags(db_path, project_id, "translation")  # 이번 실행 결과만 보이도록 초기화

        project = repo.get_project(db_path, project_id)
        lang_code = project["target_lang"]
        lang_meta = languages[lang_code]
        pdir = _project_dir(projects_dir, project_id)

        extracted = slide_extractor.extract_presentation(project["original_path"])
        reference_extracted = None
        if project.get("reference_path") and os.path.exists(project["reference_path"]):
            try:
                reference_extracted = slide_extractor.extract_presentation(project["reference_path"])
            except Exception:
                logger.warning("reference extraction failed", exc_info=True)

        keys = get_effective_keys(db_path)
        repo.update_project(db_path, project_id, status_message="AI 번역 계획 생성 중...")

        def progress_cb(done, total, skipped=False):
            pct = int(10 + (done / max(total, 1)) * 60)
            repo.update_project(db_path, project_id, progress=pct,
                                 status_message=f"번역 계획 생성 중... ({done}/{total} 배치)")

        plan_by_shape, plan_errors = translator.plan_translation(
            keys["anthropic"], anthropic_model, extracted, lang_code, lang_meta,
            reference_extracted=reference_extracted, progress_cb=progress_cb,
        )
        if plan_errors:
            repo.add_review_flags(db_path, project_id, "translation", [
                {"slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                 "issue": f"슬라이드 {e['slides']} 배치 번역 실패: {e['error']}", "severity": "error"}
                for e in plan_errors
            ])

        repo.update_project(db_path, project_id, progress=75, status_message="PPT 파일에 번역 반영 중...")
        out_dir = os.path.join(pdir, "translated")
        out_name = pptx_pipeline.make_output_filename(project["original_filename"], lang_meta, project.get("name"))
        out_path = os.path.join(out_dir, out_name)

        applied_records, apply_flags = pptx_pipeline.apply_translation_plan(
            project["original_path"], extracted, plan_by_shape, lang_code, lang_meta, out_path,
        )
        repo.replace_slide_texts(db_path, project_id, applied_records)
        repo.add_review_flags(db_path, project_id, "translation", apply_flags)

        repo.update_project(db_path, project_id, progress=85, status_message="AI 번역 자동 검수 중...")
        if keys["openai"]:
            qa_flags = qa_reviewer.run_translation_qa(
                keys["openai"], os.environ.get("OPENAI_MODEL", "gpt-4o"),
                [{"shape_id": r["shape_id"], "slide_index": r["slide_index"], "shape_name": r["shape_name"],
                  "source_korean": r["source_korean"], "translated_text": r["translated_text"]}
                 for r in applied_records],
            )
            repo.add_review_flags(db_path, project_id, "translation", qa_flags)
        else:
            repo.add_review_flags(db_path, project_id, "translation", [{
                "slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                "issue": "OpenAI API 키가 없어 GPT 자동 검수를 건너뛰었습니다. 설정에서 키를 등록하면 자동 검수가 실행됩니다.",
                "severity": "info",
            }])

        repo.update_project(db_path, project_id, progress=92, status_message="미리보기 렌더링 중...")
        try:
            preview_dir = os.path.join(pdir, "preview_translated")
            render.render_pptx_to_pngs(out_path, preview_dir, dpi=90)
        except Exception as e:
            repo.add_review_flags(db_path, project_id, "translation", [{
                "slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                "issue": f"미리보기 렌더링 실패 (LibreOffice): {e}", "severity": "info",
            }])

        repo.update_project(db_path, project_id, stage="review_translation", progress=100,
                             translated_path=out_path, status_message="번역 완료 — 검수 대기 중")
        repo.add_event(db_path, project_id, "번역 단계 완료, 검수 대기")
    except Exception as e:
        logger.error("translation stage failed: %s", traceback.format_exc())
        repo.update_project(db_path, project_id, stage="failed", status_message=f"번역 단계 오류: {e}")
        repo.add_event(db_path, project_id, f"번역 단계 실패: {e}", level="error")


# ────────────────────────────────────────────────────────────────────────
# 4단계: 타입캐스트 음성 생성
# ────────────────────────────────────────────────────────────────────────
def run_tts_stage(db_path, projects_dir, project_id, voice_id, typecast_model="ssfm-v30",
                   audio_format="wav", language=None, emotion_type=None):
    try:
        repo.update_project(db_path, project_id, stage="tts_running", progress=0,
                             status_message="음성 생성 준비 중...")
        project = repo.get_project(db_path, project_id)
        slide_texts = repo.list_slide_texts(db_path, project_id)

        by_slide = {}
        for r in slide_texts:
            if r["translated_text"]:
                by_slide.setdefault(r["slide_index"], []).append(r["translated_text"])
        slide_scripts = [(idx, ". ".join(texts)) for idx, texts in sorted(by_slide.items())]

        repo.replace_audio_assets_pending(db_path, project_id, slide_scripts)
        assets = repo.list_audio_assets(db_path, project_id)

        keys = get_effective_keys(db_path)
        pdir = _project_dir(projects_dir, project_id)
        audio_dir = os.path.join(pdir, "audio")
        os.makedirs(audio_dir, exist_ok=True)

        total = len(assets) or 1
        for i, asset in enumerate(assets):
            repo.update_project(db_path, project_id, progress=int(i / total * 100),
                                 status_message=f"음성 생성 중... ({i + 1}/{total})")
            if not asset["script_text"]:
                repo.update_audio_asset(db_path, asset["id"], status="done", duration_sec=0)
                continue
            try:
                chunks = tts_typecast.split_text_for_tts(asset["script_text"])
                combined = b""
                for chunk in chunks:
                    combined += tts_typecast.synthesize(
                        keys["typecast"], voice_id, chunk, model=typecast_model,
                        language=language, audio_format=audio_format, emotion_type=emotion_type,
                    )
                out_path = os.path.join(audio_dir, f"slide_{asset['slide_index']:03d}.{audio_format}")
                with open(out_path, "wb") as f:
                    f.write(combined)

                duration = None
                if audio_format == "wav":
                    duration = _wav_duration_seconds(out_path)

                repo.update_audio_asset(db_path, asset["id"], status="done", file_path=out_path,
                                         voice_id=voice_id, duration_sec=duration)
            except Exception as e:
                repo.update_audio_asset(db_path, asset["id"], status="error", error_message=str(e))
                repo.add_event(db_path, project_id, f"슬라이드 {asset['slide_index'] + 1} 음성 생성 실패: {e}",
                                level="error")

        repo.update_project(db_path, project_id, stage="tts_review", progress=100,
                             status_message="음성 생성 완료 — 검수 대기 중")
        repo.add_event(db_path, project_id, "음성 생성 단계 완료, 검수 대기")
    except Exception as e:
        logger.error("tts stage failed: %s", traceback.format_exc())
        repo.update_project(db_path, project_id, stage="failed", status_message=f"음성 생성 단계 오류: {e}")
        repo.add_event(db_path, project_id, f"음성 생성 단계 실패: {e}", level="error")


def _wav_duration_seconds(path):
    import wave
    try:
        with wave.open(path, "rb") as w:
            frames = w.getnframes()
            rate = w.getframerate()
            return round(frames / float(rate), 2) if rate else None
    except Exception:
        return None


# ────────────────────────────────────────────────────────────────────────
# 5단계: 최종(음성+PPT) 검수
# ────────────────────────────────────────────────────────────────────────
def run_final_review_stage(db_path, projects_dir, project_id, languages):
    try:
        repo.update_project(db_path, project_id, stage="final_running", progress=0,
                             status_message="최종 검수 중...")
        project = repo.get_project(db_path, project_id)
        repo.clear_review_flags(db_path, project_id, "final")

        # 1) PPT 텍스트 100% 일치 검사 (번역 승인 시점 스냅샷 대비 현재 파일)
        current_extracted = slide_extractor.extract_presentation(project["translated_path"])
        snapshot = repo.list_slide_texts(db_path, project_id)
        mismatches, match_ratio = qa_reviewer.compare_text_snapshots(current_extracted, snapshot)
        if mismatches:
            repo.add_review_flags(db_path, project_id, "final", [{
                "slide_index": m["slide_index"], "shape_name": m["shape_name"],
                "source_text": m["expected"], "translated_text": m["actual"],
                "issue": f"{m['reason']} — 승인된 번역과 현재 PPT 내용이 다릅니다.", "severity": "error",
            } for m in mismatches])
        else:
            repo.add_review_flags(db_path, project_id, "final", [{
                "slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                "issue": f"PPT 텍스트 100% 일치 확인 완료 (일치율 {match_ratio*100:.1f}%).", "severity": "info",
            }])

        # 2) 오디오 검수 (재생 길이 근사치 + 가능하면 Whisper 전사 비교)
        keys = get_effective_keys(db_path)
        assets = repo.list_audio_assets(db_path, project_id)
        for asset in assets:
            if asset["status"] != "done" or not asset["script_text"]:
                continue
            expected = qa_reviewer.estimate_speaking_seconds(asset["script_text"])
            actual = asset["duration_sec"]
            if actual and (actual < expected * 0.35 or actual > expected * 3.0):
                repo.add_review_flags(db_path, project_id, "final", [{
                    "slide_index": asset["slide_index"], "shape_name": f"오디오(슬라이드 {asset['slide_index']+1})",
                    "source_text": asset["script_text"], "translated_text": f"{actual}초",
                    "issue": f"예상 길이(~{expected:.1f}초)와 실제 길이({actual}초)가 크게 달라 음성 누락/오류 가능성이 있습니다.",
                    "severity": "warning",
                }])
            if keys["openai"] and asset["file_path"] and os.path.exists(asset["file_path"]):
                try:
                    transcript = qa_reviewer.transcribe_audio_openai(keys["openai"], asset["file_path"])
                    if not transcript.strip():
                        repo.add_review_flags(db_path, project_id, "final", [{
                            "slide_index": asset["slide_index"],
                            "shape_name": f"오디오(슬라이드 {asset['slide_index']+1})",
                            "source_text": asset["script_text"], "translated_text": "(빈 전사 결과)",
                            "issue": "음성에서 텍스트가 감지되지 않았습니다 (무음/생성 오류 의심).",
                            "severity": "warning",
                        }])
                except Exception:
                    pass  # 전사는 보조 수단이므로 실패해도 파이프라인은 계속 진행

        for a in [a for a in assets if a["status"] == "error"]:
            repo.add_review_flags(db_path, project_id, "final", [{
                "slide_index": a["slide_index"], "shape_name": f"오디오(슬라이드 {a['slide_index']+1})",
                "source_text": a["script_text"], "translated_text": None,
                "issue": f"음성 생성 실패: {a['error_message']}", "severity": "error",
            }])

        # 3) 최종 렌더링
        try:
            pdir = _project_dir(projects_dir, project_id)
            preview_dir = os.path.join(pdir, "preview_final")
            render.render_pptx_to_pngs(project["translated_path"], preview_dir, dpi=90)
        except Exception as e:
            repo.add_review_flags(db_path, project_id, "final", [{
                "slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                "issue": f"최종 렌더링 실패 (LibreOffice): {e}", "severity": "info",
            }])

        repo.update_project(db_path, project_id, stage="review_final", progress=100,
                             final_path=project["translated_path"],
                             status_message="최종 검수 완료 — 최종 승인 대기 중")
        repo.add_event(db_path, project_id, "최종 검수 완료, 최종 승인 대기")
    except Exception as e:
        logger.error("final review stage failed: %s", traceback.format_exc())
        repo.update_project(db_path, project_id, stage="failed", status_message=f"최종 검수 단계 오류: {e}")
        repo.add_event(db_path, project_id, f"최종 검수 단계 실패: {e}", level="error")
