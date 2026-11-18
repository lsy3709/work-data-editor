# -*- coding: utf-8 -*-
"""회귀 분석 연습용 가상 데이터(samples_original/주택_감정가_분석연습.xlsx)를 만든다.
실행: python make_analysis_sample.py        (난수 씨앗이 고정이라 언제 만들어도 같은 파일)

모든 값은 지어낸 것이다. 감정가는 아래 식에 약간의 흔들림을 곱해 만든다.
    감정가 = 지역 단가 × 전용면적 × 주택유형 × 노후 할인 × 층 가산 × 역세권 가산
연습 거리로 일부러 넣은 것: 빈칸 22칸, 겹친 행 4개, 잘못 입력한 값 9개(감정가에 0 하나 더 · 면적에 0 하나 더).
"""
import math
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

HERE = Path(__file__).resolve().parent
OUT = HERE / "samples_original" / "주택_감정가_분석연습.xlsx"

REGION = {"서울": 1300, "경기": 800, "부산": 620, "대전": 480, "대구": 520, "광주": 450}   # 만원/㎡
KIND = {"아파트": 1.0, "연립": 0.72, "단독": 0.80}
HEAD = ["물건번호", "지역", "주택유형", "전용면적(㎡)", "건축연수(년)", "층", "역까지거리(m)", "세대수", "감정가(만원)"]


def make_rows(n=600, seed=2026):
    rnd = random.Random(seed)
    rows = []
    for i in range(n):
        region = rnd.choices(list(REGION), weights=[22, 26, 16, 10, 14, 12])[0]
        kind = rnd.choices(list(KIND), weights=[70, 18, 12])[0]
        area = round(rnd.choice([39, 49, 59, 59, 74, 84, 84, 84, 101, 114, 135]) + rnd.uniform(-3, 3), 1)
        age = rnd.randint(0, 38)
        floor = 1 if kind == "단독" else rnd.randint(1, 5 if kind == "연립" else 28)
        dist = int(rnd.uniform(80, 2500))
        units = 1 if kind == "단독" else int(rnd.uniform(20, 2400))
        price = (REGION[region] * area * KIND[kind] * (1 - 0.012 * min(age, 30)) * (1 + 0.004 * min(floor, 20))
                 * (1 + 0.25 * math.exp(-dist / 400)) * rnd.gauss(1, 0.06))
        rows.append([f"HM-{i + 1:04d}", region, kind, area, age, floor, dist, units, int(round(price, -1))])

    pick = rnd.sample(range(n), 35)
    for r in pick[:10]:
        rows[r][4] = None                     # 건축연수 빈칸
    for r in pick[10:18]:
        rows[r][6] = None                     # 역까지거리 빈칸
    for r in pick[18:22]:
        rows[r][8] = None                     # 감정가 빈칸 (이 행은 배울 수 없다)
    for r in pick[22:28]:
        rows[r][8] *= 10                      # 감정가에 0 을 하나 더 넣은 실수
    for r in pick[28:31]:
        rows[r][3] = round(rows[r][3] * 10)   # 면적에 0 을 하나 더 넣은 실수
    for r in pick[31:35]:
        rows.append(list(rows[r]))            # 같은 행이 두 번 들어간 실수
    return rows


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "감정내역"
    ws.append(HEAD)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in make_rows():
        ws.append(row)
    for cell in ws["I"][1:]:
        cell.number_format = "#,##0"
    ws.freeze_panes = "A2"
    for col, width in zip("ABCDEFGHI", (12, 8, 10, 13, 13, 6, 15, 8, 14)):
        ws.column_dimensions[col].width = width
    wb.save(OUT)
    print("만듦:", OUT.name, ws.max_row - 1, "행")


if __name__ == "__main__":
    main()
