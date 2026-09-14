"""
LibreOffice headless를 이용한 PPTX -> PDF -> PNG 렌더링.
스킬 문서 "최종 검수는 반드시 렌더링해서 육안 확인" 절의 자동화 버전.
"""
import os
import subprocess
import tempfile
import shutil


def render_pptx_to_pngs(pptx_path, out_dir, dpi=100, slide_limit=None):
    """pptx_path를 PDF로 변환 후 슬라이드별 PNG로 쪼갠다. PNG 경로 리스트 반환."""
    os.makedirs(out_dir, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        cmd = ["soffice", "--headless", "--norestore", "--convert-to", "pdf", "--outdir", tmp, pptx_path]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice 변환 실패: {result.stderr[:1000]}")

        pdf_name = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
        pdf_path = os.path.join(tmp, pdf_name)
        if not os.path.exists(pdf_path):
            candidates = [f for f in os.listdir(tmp) if f.endswith(".pdf")]
            if not candidates:
                raise RuntimeError("PDF 변환 결과물을 찾을 수 없습니다.")
            pdf_path = os.path.join(tmp, candidates[0])

        prefix = os.path.join(out_dir, "slide")
        ppm_cmd = ["pdftoppm", "-png", "-r", str(dpi), pdf_path, prefix]
        if slide_limit:
            ppm_cmd = ["pdftoppm", "-png", "-r", str(dpi), "-f", "1", "-l", str(slide_limit), pdf_path, prefix]
        result2 = subprocess.run(ppm_cmd, capture_output=True, text=True, timeout=300)
        if result2.returncode != 0:
            raise RuntimeError(f"PDF -> PNG 변환 실패: {result2.stderr[:1000]}")

        # 최종 산출물이 out_dir에 남아야 하므로 tmp의 pdf도 복사해둔다 (선택)
        final_pdf = os.path.join(out_dir, os.path.basename(pdf_path))
        shutil.copyfile(pdf_path, final_pdf)

    pngs = sorted(
        [os.path.join(out_dir, f) for f in os.listdir(out_dir) if f.startswith("slide") and f.endswith(".png")]
    )
    return pngs
