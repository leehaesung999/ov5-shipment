"""배포용 zip 빌드.
사용:
    1. portable Python(`입고_파레트구분기.zip` 안의 `python` 폴더)을 이 폴더로 복사
    2. python build.py  실행 (또는 build.bat)
결과:
    dist/OV5_지정출고_배포_YYYYMMDD.zip
"""
import os
import sys
import zipfile
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
ROOT = Path(__file__).parent
DIST = ROOT / "dist"
DIST.mkdir(exist_ok=True)
ts = datetime.now().strftime("%Y%m%d_%H%M")
out = DIST / f"OV5_지정출고_배포_{ts}.zip"
if out.exists():
    try:
        out.unlink()
    except PermissionError:
        # 같은 분 안에 다시 빌드한 경우 — 초 단위까지 붙여 충돌 회피
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = DIST / f"OV5_지정출고_배포_{ts}.zip"

INCLUDE_DIRS = ["core", "config", ".streamlit", "python"]
INCLUDE_FILES = ["app_streamlit.py", "실행.bat", "README.md", "requirements.txt"]

EXCLUDE_PARTS = {"__pycache__", "tmp_uploads", "output", "test_output", "dist"}
EXCLUDE_SUFFIX = (".pyc", ".bak", ".bak.xlsx")


def should_skip(path: Path) -> bool:
    if any(p in path.parts for p in EXCLUDE_PARTS):
        return True
    if path.suffix in EXCLUDE_SUFFIX:
        return True
    return False


py_exists = (ROOT / "python" / "python.exe").exists()
if not py_exists:
    print("⚠ python 폴더가 없습니다. 동봉 없이 빌드합니다.")
    print("  배포받은 PC에 시스템 Python 필요.")
    INCLUDE_DIRS = [d for d in INCLUDE_DIRS if d != "python"]

count = 0
total = 0
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
    for d in INCLUDE_DIRS:
        src_dir = ROOT / d
        if not src_dir.exists():
            print(f"  스킵 (없음): {d}/")
            continue
        for root, dirs, files in os.walk(src_dir):
            dirs[:] = [x for x in dirs if x not in EXCLUDE_PARTS]
            for f in files:
                src = Path(root) / f
                if should_skip(src):
                    continue
                arc = src.relative_to(ROOT)
                z.write(src, arc)
                count += 1
                total += src.stat().st_size

    for f in INCLUDE_FILES:
        src = ROOT / f
        if src.exists():
            z.write(src, f)
            count += 1
            total += src.stat().st_size

mb = out.stat().st_size / 1024 / 1024
print(f"\n✅ 빌드 완료: {out}")
print(f"   파일 {count}개 / 원본 {total/1024/1024:.1f}MB → 압축 {mb:.1f}MB")
print(f"\n받는 분이 할 일:")
print("  1. zip 압축 풀기 (한글 없는 경로 권장)")
print("  2. 실행.bat 더블클릭")
