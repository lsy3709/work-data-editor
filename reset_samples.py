# -*- coding: utf-8 -*-
"""samples 폴더를 처음 상태로 만든다. 처음 한 번, 그리고 연습을 다시 할 때 실행한다.
실행: python reset_samples.py
"""
import shutil
from pathlib import Path

HERE = Path(__file__).resolve().parent
SRC, DST = HERE / "samples_original", HERE / "samples"
DST.mkdir(exist_ok=True)
for f in SRC.iterdir():
    if f.is_file():
        shutil.copy2(f, DST / f.name)
        print("준비:", f.name)
print("끝. (백업 폴더는 그대로 둡니다)")
