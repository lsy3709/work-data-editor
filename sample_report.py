# -*- coding: utf-8 -*-
"""연습용 표 네 개를 정제 → 이상 탐지 → 통계 검정 → 모델 비교 → 그래프까지 한 번에 돌려 보고서를 만든다.
실행: python sample_report.py        결과: docs/sample_analysis.md + docs/images/*.png  (1분쯤 걸림)

화면 없이 analysis.py · stat_tests.py 만 쓴다 — '매달 들어오는 파일을 자동으로 분석해 보고서로' 의 본보기.
"""
import sys
from pathlib import Path

import analysis as an
import data_editor as ed
import stat_tests as st

HERE = Path(__file__).resolve().parent
DOCS = HERE / "docs"
IMAGES = DOCS / "images"

DATASETS = [
    ("loan", "보금자리론_신청_2026_10.xlsx", "신청내역", "보금자리론 신청 (2026년 10월)",
     "수업의 기본 연습 파일. 신청금액을 다른 열로 설명할 수 있는지 본다."),
    ("merged", "S1_보금자리론_취합.xlsx", "전체내역", "보금자리론 3개월 취합 (샘플 S1 의 결과)",
     "월별 파일 3개를 합친 630행. 행이 늘면 결론이 달라지는지 본다."),
    ("pension", "주택연금_가입현황_정제전.xlsx", "가입현황", "주택연금 가입현황 — 정제 전 (샘플 S3 의 입력)",
     "금액이 '2,430,000원', 날짜가 '2026.01.03' · '20260613' 처럼 글자로 섞여 있고 겹친 행 · 빈칸이 있는 표."),
    ("house", "주택_감정가_분석연습.xlsx", "감정내역", "주택 감정가 (분석 연습용)",
     "열 사이에 실제 관계를 넣어 만든 표. 입력 실수 9건 · 빈칸 22칸 · 겹친 행 4개를 일부러 넣었다."),
]


def table(head, rows):
    out = ["| " + " | ".join(head) + " |", "| " + " | ".join("---" for _ in head) + " |"]
    out += ["| " + " | ".join(str(v) for v in row) + " |" for row in rows]
    return out


def analyze(key, file, sheet, title, note):
    source = ed.open_source(HERE / "samples_original" / file)
    df = an.to_frame(source.load(sheet))
    a = an.auto(df)
    c, o, best = a.clean, a.outliers, a.best
    tests = st.test_all(c, exclude=o.mask)
    fit = st.ols(c, exclude=o.mask)
    stats = st.describe(c, exclude=o.mask)
    curve = an.learning(c, best.key, exclude=o.mask, model=best.model)
    for name, fig in (("model", an.figure(c, o, best, a.table)), ("tests", st.figure_tests(c, tests, fit, exclude=o.mask)),
                      ("learning", an.figure_learning(best, curve))):
        fig.savefig(IMAGES / f"{key}_{name}.png", dpi=100)
    yes = [t.feature for t in tests if t.significant]
    r2 = a.table[0]["r2"]
    if r2 >= 0.75:
        verdict = f"설명하는 열로 [{c.target}] 을(를) 잘 맞힙니다 (1위 모델의 교차 검증 R² {r2:.3f})."
    elif r2 >= 0.3:
        verdict = f"어느 정도 맞히지만 설명하지 못하는 부분이 큽니다 (1위 모델의 교차 검증 R² {r2:.3f})."
    else:
        verdict = (f"**이 열들로는 [{c.target}] 을(를) 예측할 수 없습니다** (1위 모델의 교차 검증 R² {r2:.3f} — "
                   "0 이하는 '평균값만 찍는 것보다 못하다'는 뜻). 관계가 없는 데이터에서는 어떤 모델을 써도 마찬가지입니다.")
    out = [f"## {title}", "", f"`samples_original/{file}` · 시트 `{sheet}` — {note}", "",
           f"**한 줄 결론** — {verdict} 통계 검정에서 유의한 열은 {len(tests)}개 중 {len(yes)}개"
           + (f" ({', '.join(yes)})" if yes else "") + "입니다.", "",
           "### 1) 정제", ""] + [f"- {line}" for line in c.report]
    out += ["", "### 2) 이상 탐지", "", f"- {o.report[0]}"]
    shown = list(o.mask.index[o.mask])[:5]
    out += [f"  - {an.to_text(df.at[rid, df.columns[0]])}: {o.reasons[rid]}" for rid in shown]
    if int(o.mask.sum()) > len(shown):
        out.append(f"  - … 외 {int(o.mask.sum()) - len(shown)}행")
    out += ["", "### 3) 기술통계", ""]
    out += table(["열", "n", "평균", "표준편차", "최소", "중앙값", "최대", "왜도"],
                 [[s["name"], f"{s['n']:,}", f"{s['mean']:,.1f}", f"{s['std']:,.1f}", f"{s['min']:,.1f}", f"{s['median']:,.1f}",
                   f"{s['max']:,.1f}", f"{s['skew']:.2f}"] for s in stats])
    out += ["", f"### 4) 통계 검정 — 설명하는 열과 [{c.target}] 의 관계", ""]
    out += table(["설명하는 열", "검정 방법", "통계량", "자유도", "p 값", "효과 크기", "판정"],
                 [[t.feature, t.method, f"{t.stat_name} = {t.stat:,.3f}", t.df, st.p_short(t.p), st.effect_text(t.effect_name, t.effect),
                   "**유의함**" if t.significant else "유의하지 않음"] for t in tests])
    out += [""] + [f"- {t.sentence}" for t in tests]
    if len(tests) >= 5:
        out.append(f"- 열을 {len(tests)}개 검정하면 관계가 전혀 없어도 우연히 유의하게 나오는 열이 평균 {len(tests) * st.ALPHA:.1f}개쯤 생깁니다. "
                   "유의한 열이 한둘이고 효과 크기가 작다면 우연일 수 있습니다.")
    out += ["", f"![{title} — 검정 그래프](images/{key}_tests.png)", "",
            f"### 5) 다중 회귀의 계수 t 검정", "",
            f"{fit.n:,}행 · R² {fit.r2:.3f} (조정 R² {fit.adj_r2:.3f}) · 모델 전체 F = {fit.f:,.2f}, {st.p_text(fit.f_p)}"
            + (f" · {fit.note}" if fit.note else ""), ""]
    rows = sorted(fit.rows, key=lambda r: (r["name"] != "(절편)", -abs(r["t"]) if r["t"] == r["t"] else 0))[:12]
    out += table(["열", "계수", "표준오차", "t", "p 값", "95% 신뢰구간", "판정"],
                 [[r["name"], f"{r['coef']:,.2f}", f"{r['se']:,.2f}", f"{r['t']:.2f}", st.p_short(r["p"]),
                   f"{r['low']:,.1f} ~ {r['high']:,.1f}", "**유의함**" if r["significant"] else "유의하지 않음"] for r in rows])
    if len(fit.rows) > len(rows):
        out += ["", f"(계수 {len(fit.rows)}개 중 절편과 |t| 가 큰 {len(rows) - 1}개만 표시)"]
    out += ["", "### 6) 모델 9가지 비교 (5조각 교차 검증)", ""]
    out += table(["순위", "모델", "계열", "R²", "흔들림(±)", "RMSE", "시간(초)"],
                 [[i, row["name"], row["family"], f"{row['r2']:.3f}", f"{row['std']:.3f}", f"{row['rmse']:,.0f}", f"{row['seconds']:.2f}"]
                  for i, row in enumerate(a.table, 1)])
    m = best.metrics
    out += ["", f"1위 **{best.name}** 로 학습 → 따로 떼어 둔 시험 데이터 {m['n_test']}행에서 R² {m['r2_test']:.3f} · "
                f"평균 오차 {m['mae']:,.0f} · 평균 오차율 {m['mape']:.1f}% (학습 데이터에서는 R² {m['r2_train']:.3f})", "",
            f"![{title} — 모델 그래프](images/{key}_model.png)", "",
            f"### 7) 학습 과정", "",
            "배우는 데 쓴 행 수를 늘려 가며 잰 시험 점수: "
            + " → ".join(f"{n}행 {v:.3f}" for n, v in zip(curve["sizes"], curve["valid"])), "",
            f"![{title} — 학습 곡선](images/{key}_learning.png)", ""]
    summary = [title, f"{len(df):,} → {len(c.X):,}", int(o.mask.sum()), f"{len(yes)} / {len(tests)}", a.table[0]["name"], f"{r2:.3f}"]
    return out, summary


def main():
    IMAGES.mkdir(parents=True, exist_ok=True)
    body, rows = [], []
    for item in DATASETS:
        print("분석 중:", item[1], flush=True)
        text, summary = analyze(*item)
        body += text
        rows.append(summary)
    head = ["# 연습용 표 분석 결과", "",
            "> 이 문서는 `python sample_report.py` 가 자동으로 만듭니다. 표마다 **정제 → 이상 탐지 → 기술통계 → 통계 검정(t 검정 · 분산분석) → "
            "회귀 계수 t 검정 → 모델 9가지 비교 → 학습 곡선** 을 같은 순서로 돌렸습니다. 모든 값은 가상 데이터입니다.", "",
            "## 한눈에", ""]
    head += table(["표", "행 (정제 전 → 후)", "이상 행", "유의한 열", "1위 모델", "교차 검증 R²"], rows)
    head += ["",
             "- 앞의 세 표는 수업용으로 **값을 무작위로 만든** 표입니다. 그래서 '관계가 없다'가 정답이고, 분석도 그렇게 나옵니다 — "
             "유의한 열이 거의 없고 R² 가 0 근처이거나 음수입니다. 도구가 없는 관계를 지어내지 않는다는 확인이기도 합니다.",
             "- 마지막 표는 열 사이에 관계를 넣어 만든 표라서 같은 절차로 R² 0.97 이 나옵니다.",
             "- 무작위 표에서도 '유의한 열'이 나올 수 있습니다. 10월 표의 `처리상태` 가 그렇습니다(p < 0.001). 값을 무작위로 만들었으니 우연이고, "
             "3개월 치를 합친 표에서는 사라집니다. p 값이 아주 작아도 한 번의 결과로 단정하지 말고 다른 기간의 데이터로 다시 확인해야 합니다.", ""]
    (DOCS / "sample_analysis.md").write_text("\n".join(head + body) + "\n", encoding="utf-8", newline="\n")
    print("만듦: docs/sample_analysis.md, docs/images/ 그림", len(DATASETS) * 3, "장")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
