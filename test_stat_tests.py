# -*- coding: utf-8 -*-
"""stat_tests.py (기술통계 · t 검정 · 분산분석 · 회귀 계수 검정) 와 정제(tidy) · 학습 곡선 테스트.
실행: python test_stat_tests.py
"""
import math
import unittest
from pathlib import Path

import data_editor as ed

try:
    import numpy as np
    import pandas as pd
    from scipy import stats as sp

    import analysis as an
    import stat_tests as st
    MISSING = ""
except ImportError as e:
    an = st = None
    MISSING = str(e)

HERE = Path(__file__).resolve().parent
SAMPLE = HERE / "samples_original"


def load(name, sheet=None):
    src = ed.open_source(SAMPLE / name)
    return an.to_frame(src.load(sheet or src.tables()[0]))


@unittest.skipIf(an is None, f"분석 라이브러리 없음: {MISSING}")
class Base(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.house = load("주택_감정가_분석연습.xlsx")
        cls.c = an.clean(cls.house, "감정가(만원)", an.suggest(cls.house)[1])
        cls.o = an.find_outliers(cls.c)
        cls.tests = {t.feature: t for t in st.test_all(cls.c, exclude=cls.o.mask)}


class 정제_글자로_들어온_값(Base):
    def test_숫자와_날짜를_바로잡는다(self):
        raw = pd.DataFrame({
            "번호": ["A-1", "A-2", "A-3", "A-4", "A-5"],
            "이름": [" 김하나", "이두리 ", "박세나", "최네모", None],
            "금액": ["3150000", "2,430,000원", " 1,140,000 ", None, "모름"],
            "가입일": ["2026/07/15", "2026.01.03", "20260613", "2026-01-13", ""],
            "연도": ["1955", "1959", "1951", "1948", "1957"]})
        tidy, notes = an.tidy(raw)
        self.assertEqual(list(tidy["금액"][:3]), [3150000, 2430000, 1140000])
        self.assertTrue(math.isnan(tidy["금액"][3]) and math.isnan(tidy["금액"][4]))     # 못 읽는 값은 빈칸으로
        self.assertEqual(tidy["가입일"][2], pd.Timestamp(2026, 6, 13))
        self.assertTrue(pd.isna(tidy["가입일"][4]))
        self.assertEqual(tidy["이름"][0], "김하나")
        self.assertEqual(an.kind_of(tidy["연도"]), "숫자")
        self.assertEqual(tidy["번호"][0], "A-1")                                           # 번호표는 그대로
        text = " ".join(notes)
        for word in ("금액", "가입일", "앞뒤 빈칸", "읽지 못한"):
            self.assertIn(word, text)
        again, notes2 = an.tidy(tidy)                                                      # 두 번 해도 그대로
        self.assertEqual(notes2, [])
        self.assertTrue(again.equals(tidy))

    def test_수업의_정제전_표를_그대로_분석할_수_있다(self):
        df = load("주택연금_가입현황_정제전.xlsx")
        target, features = an.suggest(df)
        self.assertEqual(target, "월지급액")
        self.assertIn("지급방식", features)
        self.assertNotIn("성명", features)
        c = an.clean(df, target, features)
        self.assertEqual(c.dropped["duplicate"], 2)                                        # 공백만 달랐던 겹친 행
        self.assertGreaterEqual(len(c.X), 50)
        self.assertTrue(any("글자로 적힌 숫자" in line for line in c.report))
        self.assertTrue(any("날짜" in line for line in c.report))


class 기술통계(Base):
    def test_요약값(self):
        rows = {r["name"]: r for r in st.describe(self.c)}
        y = self.c.y
        self.assertAlmostEqual(rows["감정가(만원)"]["mean"], float(y.mean()))
        self.assertAlmostEqual(rows["감정가(만원)"]["median"], float(y.median()))
        self.assertEqual(rows["감정가(만원)"]["n"], len(y))
        self.assertIn("전용면적(㎡)", rows)
        self.assertNotIn("지역", rows)                                                     # 글자 열은 집단 표로 따로


class t검정(Base):
    def test_두_집단은_웰치_t검정이고_scipy와_같다(self):
        frame = pd.DataFrame({"구분": ["가"] * 6 + ["나"] * 7,
                              "값": [10, 12, 9, 11, 13, 10, 15, 17, 14, 18, 16, 15, 19.0]})
        t = st.compare_groups(frame["값"], frame["구분"], "구분")
        want = sp.ttest_ind(frame["값"][:6], frame["값"][6:], equal_var=False)
        self.assertIn("Welch", t.method)
        self.assertAlmostEqual(t.stat, float(want.statistic), places=10)
        self.assertAlmostEqual(t.p, float(want.pvalue), places=12)
        self.assertTrue(t.significant)
        self.assertEqual([g["name"] for g in t.groups], ["가", "나"])
        self.assertAlmostEqual(t.groups[0]["mean"], 65 / 6)
        self.assertLess(t.effect, -0.8)                                                    # 코헨 d: 큰 차이
        self.assertIn("가", t.sentence)

    def test_숫자_열은_상관_t검정이고_단순회귀의_t와_같다(self):
        x, y = self.c.X["전용면적(㎡)"][~self.o.mask], self.c.y[~self.o.mask]
        t = self.tests["전용면적(㎡)"]
        line = sp.linregress(x, y)
        self.assertIn("상관", t.method)
        self.assertAlmostEqual(t.effect, float(line.rvalue), places=10)                   # r
        self.assertAlmostEqual(t.stat, float(line.slope / line.stderr), places=6)         # t
        self.assertEqual(t.df, str(len(x) - 2))
        self.assertTrue(t.significant and t.p < 1e-10)

    def test_세_집단_이상은_분산분석과_사후_t검정(self):
        t = self.tests["지역"]
        self.assertIn("분산분석", t.method)
        self.assertEqual(t.stat_name, "F")
        self.assertTrue(t.significant)
        self.assertEqual(len(t.groups), 6)
        self.assertEqual(len(t.pairs), 15)                                                 # 6개 지역의 모든 짝
        seoul = next(p for p in t.pairs if {p["a"], p["b"]} == {"서울", "광주"})
        self.assertLess(seoul["p_adj"], 0.001)
        self.assertTrue(all(p["p_adj"] >= p["p"] for p in t.pairs))                        # 여러 번 검정한 만큼 보정
        self.assertTrue(0 < t.effect < 1)                                                  # 에타 제곱

    def test_관계가_없는_열은_유의하지_않다고_나온다(self):
        t = self.tests["세대수"]
        self.assertFalse(t.significant)
        self.assertGreater(t.p, 0.05)
        self.assertIn("유의하지 않", t.sentence)
        loan = load("보금자리론_신청_2026_10.xlsx")
        c = an.clean(loan, "신청금액(만원)", ["대출기간(년)", "상품명"])
        tests = {t.feature: t for t in st.test_all(c)}
        self.assertFalse(tests["대출기간(년)"].significant)
        self.assertFalse(tests["상품명"].significant)

    def test_모든_설명_열을_한번에(self):
        self.assertEqual(list(self.tests), self.c.features)
        for t in self.tests.values():
            self.assertTrue(0 <= t.p <= 1 and t.n > 100 and t.sentence)


class 회귀계수_t검정(Base):
    def test_한_열만_넣으면_단순회귀와_같다(self):
        c = an.clean(self.house, "감정가(만원)", ["전용면적(㎡)"])
        keep = ~self.o.mask.reindex(c.X.index, fill_value=False)
        fit = st.ols(c, exclude=~keep)
        line = sp.linregress(c.X["전용면적(㎡)"][keep], c.y[keep])
        row = next(r for r in fit.rows if r["name"] == "전용면적(㎡)")
        self.assertAlmostEqual(row["coef"], float(line.slope), places=6)
        self.assertAlmostEqual(row["se"], float(line.stderr), places=6)
        self.assertAlmostEqual(row["t"], float(line.slope / line.stderr), places=6)
        self.assertAlmostEqual(row["p"], float(line.pvalue), places=12)
        self.assertAlmostEqual(fit.r2, float(line.rvalue ** 2), places=10)

    def test_여러_열을_함께_넣은_회귀(self):
        fit = st.ols(self.c, exclude=self.o.mask)
        rows = {r["name"]: r for r in fit.rows}
        self.assertIn("(절편)", rows)
        area = rows["전용면적(㎡)"]
        self.assertGreater(area["coef"], 0)
        self.assertLess(area["p"], 1e-10)
        self.assertTrue(area["low"] < area["coef"] < area["high"])
        self.assertAlmostEqual(area["t"], area["coef"] / area["se"], places=8)
        self.assertGreater(rows["세대수"]["p"], 0.05)                                      # 관계없는 열은 계수도 유의하지 않음
        regions = [name for name in rows if name.startswith("지역=")]
        self.assertEqual(len(regions), 5)                                                  # 6개 지역 중 하나는 기준
        self.assertIn("기준", fit.note)
        self.assertTrue(0.85 < fit.adj_r2 < fit.r2 < 0.95)
        self.assertLess(fit.f_p, 1e-10)
        self.assertEqual(fit.n, int((~self.o.mask).sum()))


class 학습과정(Base):
    def test_학습_곡선(self):
        curve = an.learning(self.c, "boosting", exclude=self.o.mask)
        self.assertEqual(len(curve["sizes"]), 5)
        self.assertEqual(curve["sizes"], sorted(curve["sizes"]))
        self.assertGreater(curve["valid"][-1], curve["valid"][0])                          # 많이 배울수록 시험 점수가 오른다
        self.assertGreater(curve["valid"][-1], 0.9)
        self.assertEqual(curve["steps_name"], "나무 수")
        self.assertEqual(len(curve["steps"]), 300)
        self.assertLess(curve["steps"][-1], curve["steps"][0])                             # 나무를 보탤수록 오차가 준다
        mlp = an.learning(self.c, "mlp", exclude=self.o.mask)
        self.assertEqual(mlp["steps_name"], "반복 횟수")
        self.assertGreater(len(mlp["steps"]), 10)
        self.assertIsNone(an.learning(self.c, "linear", exclude=self.o.mask)["steps"])

    def test_진행_상황을_알려준다(self):
        seen, rows = [], []
        a = an.auto(self.house, progress=lambda text, fraction=None: seen.append((text, fraction)), on_row=rows.append)
        fractions = [f for _, f in seen if f is not None]
        self.assertEqual(fractions, sorted(fractions))
        self.assertTrue(0 < fractions[0] and fractions[-1] <= 1)
        self.assertEqual(len(rows), 9)                                                     # 모델 하나가 끝날 때마다
        self.assertEqual({r["key"] for r in rows}, {r["key"] for r in a.table})

    def test_그래프와_내보내기(self):
        import io
        from openpyxl import load_workbook
        r = an.train(self.c, "boosting", exclude=self.o.mask)
        fit = st.ols(self.c, exclude=self.o.mask)
        for fig in (st.figure_tests(self.c, list(self.tests.values()), fit, exclude=self.o.mask),
                    an.figure_learning(r, an.learning(self.c, "boosting", exclude=self.o.mask))):
            buf = io.BytesIO()
            fig.savefig(buf, format="png")
            self.assertGreater(len(buf.getvalue()), 20000)
        buf = io.BytesIO()
        an.export(buf, self.house, self.c, self.o, r, None, tests=list(self.tests.values()), ols=fit)
        wb = load_workbook(buf)
        self.assertEqual(wb.sheetnames, ["예측결과", "모델비교", "요약", "통계검정", "회귀계수"])
        self.assertEqual(wb["통계검정"].max_row, len(self.tests) + 1)
        self.assertEqual(wb["회귀계수"].max_row, len(fit.rows) + 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)
