# -*- coding: utf-8 -*-
"""analysis.py (회귀 분석 · 이상 탐지 · 모델 비교) 테스트. 화면 없이 실행.
실행: python test_analysis.py
"""
import shutil
import tempfile
import unittest
from pathlib import Path

import data_editor as ed

try:
    import analysis as an
    MISSING = ""
except ImportError as e:                       # scikit-learn · pandas 가 없으면 건너뛴다
    an, MISSING = None, str(e)

HERE = Path(__file__).resolve().parent
FILE = HERE / "samples_original" / "주택_감정가_분석연습.xlsx"
TARGET = "감정가(만원)"
FEATURES = ["지역", "주택유형", "전용면적(㎡)", "건축연수(년)", "층", "역까지거리(m)", "세대수"]


@unittest.skipIf(an is None, f"분석 라이브러리 없음: {MISSING}")
class Base(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        src = ed.open_source(FILE)
        cls.session = src.load(src.tables()[0])
        cls.df = an.to_frame(cls.session)

    def cleaned(self, **kw):
        return an.clean(self.df, TARGET, FEATURES, **kw)


class 표읽기(Base):
    def test_표를_분석용으로_바꾼다(self):
        self.assertEqual(len(self.df), 604)
        self.assertEqual(list(self.df.columns), self.session.names)
        self.assertEqual(list(self.df.index), self.session.ids())          # 행번호가 그대로 따라온다

    def test_열_살펴보기(self):
        info = {p["name"]: p for p in an.profile(self.df)}
        self.assertEqual(info[TARGET]["kind"], "숫자")
        self.assertEqual(info[TARGET]["missing"], 4)
        self.assertEqual(info["지역"]["kind"], "범주")
        self.assertEqual(info["지역"]["unique"], 6)
        self.assertEqual(info["물건번호"]["kind"], "글자")                  # 행마다 다른 값 = 번호표

    def test_변수_추천(self):
        target, features = an.suggest(self.df)
        self.assertEqual(target, TARGET)
        self.assertEqual(features, FEATURES)                               # 번호표 열은 빼고 추천


class 정제(Base):
    def test_빈칸은_중앙값으로_겹친행은_제거(self):
        c = self.cleaned(missing="median")
        self.assertEqual(c.dropped["duplicate"], 4)
        self.assertEqual(c.dropped["no_target"], 4)
        self.assertEqual(len(c.X), 604 - 4 - 4)
        self.assertEqual(int(c.X.isna().sum().sum()), 0)
        self.assertEqual(c.filled, {"건축연수(년)": 10, "역까지거리(m)": 8})
        self.assertIn("지역=서울", c.X.columns)                            # 글자 열은 0/1 열 여러 개로
        self.assertIn("주택유형=아파트", c.X.columns)
        self.assertTrue(any("겹친 행 4개" in line for line in c.report))

    def test_빈칸_있는_행을_버리기(self):
        c = self.cleaned(missing="drop")
        self.assertEqual(len(c.X), 604 - 4 - 4 - 18)
        self.assertEqual(c.filled, {})

    def test_잘못_고르면_알려준다(self):
        with self.assertRaises(ValueError):
            an.clean(self.df, "지역", ["전용면적(㎡)"])                     # 글자 열은 예측할 수 없다
        with self.assertRaises(ValueError):
            an.clean(self.df, TARGET, [])
        with self.assertRaises(ValueError):
            an.clean(self.df, TARGET, [TARGET, "층"])                      # 예측할 열을 설명 열에 넣음


class 이상탐지(Base):
    def planted(self, c):
        """심어 둔 입력 실수: (감정가에 0 하나 더 = 같은 지역 ㎡당 값의 3.5배 넘음, 면적에 0 하나 더 = 300㎡ 넘음)"""
        unit = c.y / c.frame["전용면적(㎡)"]
        usual = unit.groupby(c.frame["지역"]).transform("median")
        return set(c.y.index[unit / usual > 3.5]), set(c.frame.index[c.frame["전용면적(㎡)"] > 300])

    def test_사분위수와_표준점수(self):
        c = self.cleaned()
        _, big_area = self.planted(c)
        for method in ("iqr", "zscore"):
            o = an.find_outliers(c, method)
            found = set(o.mask.index[o.mask])
            self.assertTrue(big_area <= found, method)                     # 면적 840㎡ 같은 값은 잡는다
            self.assertTrue(all(o.reasons[i] for i in found))
            self.assertLess(len(found), 60)

    def test_모델_잔차는_0을_더_넣은_감정가를_모두_잡는다(self):
        c = self.cleaned()
        ten_times, big_area = self.planted(c)
        self.assertEqual((len(ten_times), len(big_area)), (6, 3))
        o = an.find_outliers(c, "model")
        found = set(o.mask.index[o.mask])
        self.assertTrue(ten_times <= found)
        self.assertLess(len(found), 25)
        self.assertIn("배)", o.reasons[next(iter(ten_times))])            # '몇 배' 인지 이유를 적어 준다

    def test_종합은_심어둔_값을_빠짐없이_잡는다(self):
        c = self.cleaned()
        ten_times, big_area = self.planted(c)
        o = an.find_outliers(c)                                            # 기본값 = 종합
        found = set(o.mask.index[o.mask])
        self.assertEqual(o.method, "combo")
        self.assertTrue(ten_times <= found)
        self.assertTrue(big_area <= found)
        self.assertLess(len(found), 30)

    def test_고립숲과_탐지안함(self):
        c = self.cleaned()
        o = an.find_outliers(c, "iforest")
        self.assertTrue(10 <= int(o.mask.sum()) <= 45)
        self.assertEqual(int(an.find_outliers(c, "none").mask.sum()), 0)


class 모델(Base):
    def test_소개글이_모두_있다(self):
        self.assertEqual(len(an.MODELS), 9)
        for key, m in an.MODELS.items():
            for field in (m.name, m.english, m.family, m.origin, m.history, m.what, m.when):
                self.assertTrue(field.strip(), key)
        guide = an.guide_markdown()
        for m in an.MODELS.values():
            self.assertIn(m.name, guide)

    def test_학습과_평가(self):
        c = self.cleaned()
        o = an.find_outliers(c)
        r = an.train(c, "boosting", exclude=o.mask)
        self.assertGreater(r.metrics["r2_test"], 0.93)
        self.assertLess(r.metrics["mape"], 12)
        self.assertEqual(len(r.y_test), len(r.pred_test))
        self.assertEqual(len(r.predictions), len(c.X))                     # 뺀 행에도 예측값은 준다
        self.assertAlmostEqual(float(r.importance.sum()), 1.0, places=6)
        self.assertEqual(set(r.importance.index[:2]), {"지역", "전용면적(㎡)"})   # 글자 열(지역)도 한 열로 묶어 보여 준다
        self.assertEqual(list(r.importance.index[2:3]), ["건축연수(년)"])
        again = an.train(c, "boosting", exclude=o.mask)
        self.assertEqual(r.metrics, again.metrics)                         # 같은 설정이면 같은 결과

    def test_이상값을_빼면_좋아진다(self):
        c = self.cleaned()
        dirty = an.train(c, "linear")
        clean = an.train(c, "linear", exclude=an.find_outliers(c).mask)
        self.assertGreater(clean.metrics["r2_test"], dirty.metrics["r2_test"])

    def test_모든_모델이_돈다(self):
        c = self.cleaned()
        o = an.find_outliers(c)
        for key in an.MODELS:
            r = an.train(c, key, exclude=o.mask)
            self.assertGreater(r.metrics["r2_test"], 0.6, key)

    def test_모델_비교(self):
        c = self.cleaned()
        table = an.compare(c, exclude=an.find_outliers(c).mask)
        self.assertEqual(len(table), 9)
        scores = [row["r2"] for row in table]
        self.assertEqual(scores, sorted(scores, reverse=True))
        self.assertGreater(scores[0], 0.93)


class 자동분석(Base):
    def test_한번에(self):
        a = an.auto(self.df, TARGET)
        self.assertEqual(a.clean.features, FEATURES)
        self.assertEqual(a.best.key, a.table[0]["key"])
        self.assertGreater(a.best.metrics["r2_test"], 0.93)
        self.assertGreaterEqual(len(a.log), 5)

    def test_결과_내보내기(self):
        from openpyxl import load_workbook
        a = an.auto(self.df, TARGET)
        tmp = Path(tempfile.mkdtemp(prefix="분석_"))
        try:
            out = an.export(tmp / "결과.xlsx", self.df, a.clean, a.outliers, a.best, a.table)
            wb = load_workbook(out)
            self.assertEqual(wb.sheetnames, ["예측결과", "모델비교", "요약"])
            ws = wb["예측결과"]
            head = [c.value for c in ws[1]]
            self.assertEqual(head[-3:], ["예측값", "차이(실제-예측)", "이상 여부"])
            self.assertEqual(ws.max_row, len(a.clean.X) + 1)
            self.assertEqual(wb["모델비교"].max_row, 10)
            wb.close()
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
