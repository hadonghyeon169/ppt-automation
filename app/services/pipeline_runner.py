"""
파이프라인 각 단계를 실제로 실행하는 오케스트레이터.
Flask 요청 스레드가 아닌 별도 백그라운드 스레드에서 실행되므로, Flask 앱 컨텍스트에
의존하지 않고 모든 값을 인자로 명시적으로 받는다 (db_path, 디렉토리, API 키 등).
"""
import os
import json
import logging
import traceback

from .. import repo
from . import slide_extractor, translator, pptx_pipeline, render, qa_reviewer, tts_typecast, audio_embed
from . import ppt_xml_ops as ops

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
                 "issue": f"슬라이드 {e['slides']} 배치 번역 실패: {e['error']} "
                          f"(검수 화면에서 '실패한 배치만 재시도'로 다시 시도할 수 있습니다)",
                 "severity": "error"}
                for e in plan_errors
            ])

        # plan_json은 여기서 바로 저장 — failed_batches_json은 apply 단계에서 나오는
        # "번역 누락" 슬라이드까지 합쳐서 apply 이후에 저장한다 (아래).
        repo.update_project(db_path, project_id, plan_json=json.dumps(plan_by_shape, ensure_ascii=False))

        repo.update_project(db_path, project_id, progress=75, status_message="PPT 파일에 번역 반영 중...")
        out_dir = os.path.join(pdir, "translated")
        out_name = pptx_pipeline.make_output_filename(project["original_filename"], lang_meta, project.get("name"))
        out_path = os.path.join(out_dir, out_name)

        applied_records, apply_flags, untranslated_batches = pptx_pipeline.apply_translation_plan(
            project["original_path"], extracted, plan_by_shape, lang_code, lang_meta, out_path,
        )
        repo.replace_slide_texts(db_path, project_id, applied_records)
        repo.add_review_flags(db_path, project_id, "translation", apply_flags)

        # 실패한 배치(API 호출 자체가 에러)와 번역이 누락된 슬라이드(응답엔 있었지만
        # 이 도형이 빠졌거나 결과가 비었던 경우)를 합쳐 "부분 재번역" 대상으로 저장한다.
        all_retry_batches = [e["slides"] for e in plan_errors] + untranslated_batches
        repo.update_project(
            db_path, project_id,
            failed_batches_json=json.dumps(all_retry_batches, ensure_ascii=False),
        )

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


def retry_failed_translation_batches(db_path, projects_dir, project_id, languages, anthropic_model):
    """직전 번역 실행에서 실패했던 배치만 다시 호출한다. 이미 성공한 배치는 API를
    재호출하지 않고 저장해둔 plan_json을 그대로 재사용 — 토큰 낭비 없이 실패한 부분만 고친다."""
    try:
        project = repo.get_project(db_path, project_id)
        failed_batches = json.loads(project.get("failed_batches_json") or "[]")
        if not failed_batches:
            repo.add_event(db_path, project_id, "재시도할 실패 배치가 없습니다 (이전 실행이 모두 성공).")
            return

        repo.update_project(db_path, project_id, stage="translating", progress=0,
                             status_message=f"실패한 배치 {len(failed_batches)}개만 재시도 중...")
        repo.add_event(db_path, project_id, f"번역 실패 배치 재시도 시작 ({len(failed_batches)}개 배치)")

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

        existing_plan = json.loads(project.get("plan_json") or "{}")
        keys = get_effective_keys(db_path)

        def progress_cb(done, total):
            pct = int((done / max(total, 1)) * 60)
            repo.update_project(db_path, project_id, progress=pct,
                                 status_message=f"실패 배치 재시도 중... ({done}/{total})")

        plan_updates, new_errors = translator.retry_failed_batches(
            keys["anthropic"], anthropic_model, extracted, lang_code, lang_meta, failed_batches,
            reference_extracted=reference_extracted, progress_cb=progress_cb,
        )
        existing_plan.update(plan_updates)

        repo.update_project(db_path, project_id, progress=70, status_message="PPT 파일에 번역 반영 중...")
        out_dir = os.path.join(pdir, "translated")
        out_name = pptx_pipeline.make_output_filename(project["original_filename"], lang_meta, project.get("name"))
        out_path = os.path.join(out_dir, out_name)

        applied_records, apply_flags, untranslated_batches = pptx_pipeline.apply_translation_plan(
            project["original_path"], extracted, existing_plan, lang_code, lang_meta, out_path,
        )
        repo.replace_slide_texts(db_path, project_id, applied_records)
        repo.clear_review_flags(db_path, project_id, "translation")
        repo.add_review_flags(db_path, project_id, "translation", apply_flags)
        if new_errors:
            repo.add_review_flags(db_path, project_id, "translation", [
                {"slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                 "issue": f"슬라이드 {e['slides']} 배치 재시도도 실패: {e['error']}", "severity": "error"}
                for e in new_errors
            ])
        else:
            repo.add_review_flags(db_path, project_id, "translation", [{
                "slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                "issue": "실패했던 배치가 모두 재시도로 성공했습니다.", "severity": "info",
            }])

        # 이번에 새로 반영된 shape만 골라 GPT 재검수 (전체 재검수는 비용 낭비이므로 생략)
        if keys["openai"] and plan_updates:
            retried_ids = set(plan_updates.keys())
            qa_flags = qa_reviewer.run_translation_qa(
                keys["openai"], os.environ.get("OPENAI_MODEL", "gpt-4o"),
                [{"shape_id": r["shape_id"], "slide_index": r["slide_index"], "shape_name": r["shape_name"],
                  "source_korean": r["source_korean"], "translated_text": r["translated_text"]}
                 for r in applied_records if r["shape_id"] in retried_ids],
            )
            repo.add_review_flags(db_path, project_id, "translation", qa_flags)

        try:
            preview_dir = os.path.join(pdir, "preview_translated")
            render.render_pptx_to_pngs(out_path, preview_dir, dpi=90)
        except Exception as e:
            repo.add_review_flags(db_path, project_id, "translation", [{
                "slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                "issue": f"미리보기 렌더링 실패 (LibreOffice): {e}", "severity": "info",
            }])

        remaining_retry_batches = [e["slides"] for e in new_errors] + untranslated_batches
        repo.update_project(
            db_path, project_id, stage="review_translation", progress=100,
            translated_path=out_path, status_message="실패 배치 재시도 완료 — 검수 대기 중",
            plan_json=json.dumps(existing_plan, ensure_ascii=False),
            failed_batches_json=json.dumps(remaining_retry_batches, ensure_ascii=False),
        )
        repo.add_event(
            db_path, project_id,
            f"실패 배치 재시도 완료 (남은 실패: {len(new_errors)}개, 여전히 누락된 슬라이드: {len(untranslated_batches)}개)",
        )
    except Exception as e:
        logger.error("translation retry failed: %s", traceback.format_exc())
        repo.update_project(db_path, project_id, stage="failed", status_message=f"번역 재시도 오류: {e}")
        repo.add_event(db_path, project_id, f"번역 재시도 실패: {e}", level="error")


# ────────────────────────────────────────────────────────────────────────
# 4단계: 타입캐스트 음성 생성
# ────────────────────────────────────────────────────────────────────────
def run_tts_stage(db_path, projects_dir, project_id, voice_id, typecast_model="ssfm-v30",
                   audio_format="wav", language=None, emotion_type=None, only_failed=False,
                   languages=None):
    try:
        repo.update_project(db_path, project_id, stage="tts_running", progress=0,
                             status_message="음성 생성 준비 중...")
        project = repo.get_project(db_path, project_id)

        # 타입캐스트는 language를 안 주면 텍스트로 자동감지하는데, 이 앱의 번역문은
        # 한글 단어("받침" 등)가 규칙상 섞여 들어가 있어서 자동감지가 엉뚱한 언어로
        # 잘못 판단하는 경우가 있었다 (예: 러시아어인데 중국어로 읽힘). 프로젝트의
        # 목표 언어에서 명시적으로 ISO 639-3 코드를 가져와 지정한다.
        if language is None and languages:
            lang_meta_for_tts = languages.get(project["target_lang"], {})
            language = lang_meta_for_tts.get("tts_lang")

        if only_failed:
            # 이미 성공(status='done')한 슬라이드는 그대로 두고, 실패한 것만 다시 합성한다
            # (전체를 지우고 새로 만드는 기존 방식은 성공한 슬라이드까지 타입캐스트 비용을
            # 다시 써야 해서 낭비였다).
            all_assets = repo.list_audio_assets(db_path, project_id)
            assets = [a for a in all_assets if a["status"] != "done"]
            if not assets:
                repo.update_project(db_path, project_id, stage="tts_review", progress=100,
                                     status_message="재시도할 실패 슬라이드가 없습니다 (모두 완료 상태).")
                repo.add_event(db_path, project_id, "음성 재시도 대상 없음 — 모두 완료 상태")
                return
            repo.add_event(db_path, project_id, f"실패한 오디오만 재시도 시작 ({len(assets)}개 슬라이드)")
        else:
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
            if ops.has_chinese(asset["script_text"]):
                # TTS에 넘기는 스크립트 자체에 중국어가 남아있다 — 언어 코드 문제가 아니라
                # 이 슬라이드가 애초에 번역되지 않았다는 뜻이다 (번역 단계에서 skip되었거나
                # 실패한 도형). 그래도 일단 합성은 진행하되(중국어로 읽힐 것) 눈에 띄게 남긴다.
                repo.add_event(
                    db_path, project_id,
                    f"슬라이드 {asset['slide_index'] + 1} 스크립트에 중국어가 남아있습니다 "
                    f"(번역이 적용 안 된 것으로 보임 — 음성도 중국어로 나올 수 있습니다): "
                    f"{asset['script_text'][:60]}",
                    level="error",
                )
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

        # 3) 오디오를 실제 PPT 파일에 삽입 (기존에는 이 단계가 아예 없어서 TTS로
        # 생성된 음성이 디스크에만 저장되고 최종 PPT에는 전혀 반영되지 않았다).
        # 슬라이드 진입 시 자동 재생되도록 삽입하고, 원본 translated 파일은 보존한
        # 채 별도 파일로 저장한다.
        pdir = _project_dir(projects_dir, project_id)
        final_dir = os.path.join(pdir, "final")
        final_name = os.path.basename(project["translated_path"])
        final_pptx_path = os.path.join(final_dir, final_name)
        try:
            embedded_count, embed_skipped = audio_embed.embed_audio_into_pptx(
                project["translated_path"], assets, final_pptx_path,
            )
            repo.add_review_flags(db_path, project_id, "final", [{
                "slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                "issue": f"오디오 {embedded_count}개 슬라이드에 자동재생으로 삽입 완료.",
                "severity": "info",
            }])
            for s in embed_skipped:
                repo.add_review_flags(db_path, project_id, "final", [{
                    "slide_index": s["slide_index"], "shape_name": f"오디오(슬라이드 {s['slide_index']+1})",
                    "source_text": None, "translated_text": None,
                    "issue": f"오디오 삽입 건너뜀: {s['reason']}", "severity": "warning",
                }])
        except Exception as e:
            logger.error("audio embed failed: %s", traceback.format_exc())
            final_pptx_path = project["translated_path"]  # 실패 시 최소한 번역본이라도 최종본으로
            repo.add_review_flags(db_path, project_id, "final", [{
                "slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                "issue": f"오디오 삽입 실패, 번역본만 최종 파일로 사용합니다: {e}", "severity": "error",
            }])

        # 4) 최종 렌더링 (오디오가 삽입된 최종 파일 기준)
        try:
            preview_dir = os.path.join(pdir, "preview_final")
            render.render_pptx_to_pngs(final_pptx_path, preview_dir, dpi=90)
        except Exception as e:
            repo.add_review_flags(db_path, project_id, "final", [{
                "slide_index": None, "shape_name": None, "source_text": None, "translated_text": None,
                "issue": f"최종 렌더링 실패 (LibreOffice): {e}", "severity": "info",
            }])

        repo.update_project(db_path, project_id, stage="review_final", progress=100,
                             final_path=final_pptx_path,
                             status_message="최종 검수 완료 — 최종 승인 대기 중")
        repo.add_event(db_path, project_id, "최종 검수 완료, 최종 승인 대기")
    except Exception as e:
        logger.error("final review stage failed: %s", traceback.format_exc())
        repo.update_project(db_path, project_id, stage="failed", status_message=f"최종 검수 단계 오류: {e}")
        repo.add_event(db_path, project_id, f"최종 검수 단계 실패: {e}", level="error")
