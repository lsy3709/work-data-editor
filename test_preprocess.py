# -*- coding: utf-8 -*-
"""전처리(빈칸 채우기 · 종류 바꾸기 · 열 고치기 …) · 복사본 · 다른 이름으로 저장 테스트. 화면 없이 실행.
실행: python test_preprocess.py
"""
import datetime as dt
import shutil
import statistics
import tempfile
import unittest
from pathlib import Path

import data_editor as ed

HERE = Path(__file__).resolve().parent
SAMPLE = HERE / "samples_original"


def has_access():
    try:
        import pyodbc
        return any("Access" in d and "accdb" in d for d in pyodbc.drivers())
    except Exception:
        return False


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="전처리_"))
        for f in SAMPLE.iterdir():
            shutil.copy2(f, self.tmp / f.name)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def open(self, name, table=None):
        src = ed.open_source(self.tmp / name)
        table = table or src.tables()[0]
        return src, table, src.load(table)

    def house(self):
        return self.open("주택_감정가_분석연습.xlsx")[2]

    def pension(self):
        return self.open("주택연금_가입현황_정제전.xlsx")[2]


class 빈칸(Base):
    def test_빈칸_세기와_열_요약(self):
        s = self.house()
        self.assertEqual(s.empty_count("건축연수(년)"), 10)
        self.assertEqual(s.empty_count("감정가(만원)"), 4)
        info = s.column_summary("건축연수(년)")
        self.assertEqual((info["kind"], info["empty"], info["count"]), ("int", 10, 604))
        self.assertEqual(info["min"], 0)
        text = s.column_summary("지역")
        self.assertEqual((text["kind"], text["unique"]), ("str", 6))

    def test_중앙값_평균_가장흔한값_직접입력으로_채우기(self):
        s = self.house()
        ages = [v for v in (s.get(r, "건축연수(년)") for r in s.ids()) if v is not None]
        self.assertEqual(s.fill_empty("건축연수(년)", "median"), 10)
        self.assertEqual(s.empty_count("건축연수(년)"), 0)
        filled = [r for r in s.ids() if s.status(r) == "changed"]
        self.assertEqual(len(filled), 10)
        self.assertEqual(s.get(filled[0], "건축연수(년)"), statistics.median(ages))
        self.assertTrue(s.undo())                                   # 한 번에 되돌림
        self.assertEqual(s.empty_count("건축연수(년)"), 10)
        self.assertEqual(s.change_count(), 0)

        dist = [v for v in (s.get(r, "역까지거리(m)") for r in s.ids()) if v is not None]
        self.assertEqual(s.fill_empty("역까지거리(m)", "mean"), 8)
        row = next(r for r in s.ids() if s.status(r) == "changed")
        self.assertAlmostEqual(s.get(row, "역까지거리(m)"), sum(dist) / len(dist), places=6)
        self.assertEqual(s.columns[s.col("역까지거리(m)")].kind, "float")     # 평균이 소수라 열 종류가 실수로
        s.undo()
        self.assertEqual(s.columns[s.col("역까지거리(m)")].kind, "int")       # 되돌리면 종류도 돌아옴

        self.assertEqual(s.fill_empty("감정가(만원)", "value", "0"), 4)
        with self.assertRaises(ValueError):
            s.fill_empty("감정가(만원)", "value", "모름")                    # 숫자 열에 글자는 안 됨
        with self.assertRaises(ValueError):
            s.fill_empty("지역", "mean")                                     # 글자 열에 평균은 안 됨

        p = self.pension()
        self.assertEqual(p.fill_empty("관할지점", "mode"), 1)
        self.assertEqual(p.empty_count("관할지점"), 0)

    def test_윗칸_값으로_채우기(self):
        s = ed.EditSession([ed.Column("본부"), ed.Column("지점")],
                           [["서울", "가"], [None, "나"], [None, "다"], ["부산", "라"], ["", "마"]])
        self.assertEqual(s.fill_empty("본부", "above"), 3)
        self.assertEqual([s.get(r, "본부") for r in s.ids()], ["서울", "서울", "서울", "부산", "부산"])

    def test_빈칸_있는_행_지우기(self):
        s = self.house()
        self.assertEqual(s.drop_empty_rows(["감정가(만원)"]), 4)
        self.assertEqual(len(s.rows), 600)
        self.assertEqual(s.drop_empty_rows(), 18)                    # 어느 열이든 빈칸이 있으면
        self.assertEqual(len(s.rows), 582)
        s.undo(); s.undo()
        self.assertEqual(len(s.rows), 604)
        blank = ed.EditSession([ed.Column("a"), ed.Column("b")], [[1, 2], [None, None], ["", None], [3, None]])
        self.assertEqual(blank.drop_empty_rows(how="all"), 2)        # 모든 칸이 빈 행만


class 행과글자(Base):
    def test_겹친_행_지우기(self):
        s = self.house()
        self.assertEqual(s.drop_duplicates(), 4)
        self.assertEqual(len(s.rows), 600)
        self.assertEqual(s.drop_duplicates(), 0)
        self.assertEqual(s.drop_duplicates(["지역", "주택유형"]), 600 - 18)    # 고른 열만 견주면: 지역 6 × 유형 3 = 18가지만 남음
        s.undo()
        self.assertEqual(len(s.rows), 600)

    def test_앞뒤_빈칸_지우면_겹친_행이_보인다(self):
        p = self.pension()
        self.assertEqual(p.strip_text(["성명"]), 12)
        self.assertEqual(p.get(p.ids()[1], "성명"), "황상철")
        self.assertEqual(p.drop_duplicates(), 2)
        self.assertEqual(len(p.rows), 60)                            # 수업 샘플 S3 의 결과와 같은 60행


class 종류바꾸기(Base):
    def test_글자로_적힌_숫자를_숫자로(self):
        p = self.pension()
        self.assertEqual(p.columns[p.col("월지급액")].kind, "str")
        changed, failed = p.convert("월지급액", "number")
        self.assertEqual((changed, failed), (61, 0))
        self.assertEqual(p.columns[p.col("월지급액")].kind, "int")
        self.assertEqual(p.get(p.ids()[1], "월지급액"), 2430000)     # '2,430,000원'
        self.assertIsNone(p.get(p.ids()[7], "월지급액"))             # 빈칸은 빈칸으로
        self.assertEqual(p.fill_empty("월지급액", "median"), 1)
        p.undo(); p.undo()
        self.assertEqual(p.columns[p.col("월지급액")].kind, "str")
        self.assertEqual(p.get(p.ids()[1], "월지급액"), "2,430,000원")
        self.assertEqual(p.change_count(), 0)

    def test_못_바꾸는_값이_있으면_멈추거나_빈칸으로(self):
        s = ed.EditSession([ed.Column("금액")], [["1,200"], ["3.5"], ["모름"], [None], ["12%"]])
        with self.assertRaises(ValueError) as e:
            s.convert("금액", "number")
        self.assertIn("모름", str(e.exception))
        self.assertEqual(s.get(0, "금액"), "1,200")                  # 하나라도 안 되면 아무것도 바꾸지 않음
        changed, failed = s.convert("금액", "number", on_error="empty")
        self.assertEqual((changed, failed), (4, 1))
        self.assertEqual([s.get(r, "금액") for r in s.ids()], [1200, 3.5, None, None, 12])
        self.assertEqual(s.columns[0].kind, "float")

    def test_날짜와_글자로(self):
        p = self.pension()
        changed, failed = p.convert("가입일", "date")
        self.assertEqual((changed, failed), (61, 0))
        self.assertEqual(p.get(p.ids()[0], "가입일"), dt.datetime(2026, 7, 15))     # '2026/07/15'
        self.assertEqual(p.get(p.ids()[4], "가입일"), dt.datetime(2026, 6, 13))     # '20260613'
        self.assertEqual(p.columns[p.col("가입일")].kind, "date")
        h = self.house()
        h.convert("층", "str")
        self.assertEqual(h.get(h.ids()[0], "층"), ed.to_text(h.original_row(h.ids()[0])[h.col("층")]))
        self.assertEqual(h.columns[h.col("층")].kind, "str")
        h.convert("층", "number")
        self.assertEqual(h.change_count(), 0)                        # 글자로 갔다가 숫자로 돌아오면 원래 값

    def test_숫자에_곱하기(self):
        h = self.house()
        first = h.get(h.ids()[0], "감정가(만원)")
        n = h.scale("감정가(만원)", 0.0001)                          # 만원 → 억원
        self.assertEqual(n, 600)
        self.assertAlmostEqual(h.get(h.ids()[0], "감정가(만원)"), first / 10000)
        with self.assertRaises(ValueError):
            h.scale("지역", 2)


class 열고치기(Base):
    def test_이름_바꾸기_삭제_추가와_되돌리기(self):
        h = self.house()
        self.assertFalse(h.structure_changed)
        h.rename_column("감정가(만원)", "감정가")
        self.assertIn("감정가", h.names)
        self.assertTrue(h.structure_changed)
        self.assertGreater(h.change_count(), 0)
        with self.assertRaises(ValueError):
            h.rename_column("층", "지역")                            # 이미 있는 이름
        h.delete_columns(["세대수", "물건번호"])
        self.assertEqual(h.names, ["지역", "주택유형", "전용면적(㎡)", "건축연수(년)", "층", "역까지거리(m)", "감정가"])
        self.assertEqual(len(h.rows[h.ids()[0]]), 7)
        h.add_column("비고", value="확인필요")
        self.assertEqual(h.get(h.ids()[3], "비고"), "확인필요")
        h.add_column("단가", kind="float", after="전용면적(㎡)")
        self.assertEqual(h.names.index("단가"), 3)
        h.set_cell(h.ids()[0], "단가", "1,250.5")
        self.assertEqual(h.get(h.ids()[0], "단가"), 1250.5)
        with self.assertRaises(ValueError):
            h.delete_columns(h.names)                                # 열을 모두 지울 수는 없음
        for _ in range(5):
            h.undo()
        self.assertEqual(h.names[0], "물건번호")
        self.assertEqual(h.names[-1], "감정가(만원)")
        self.assertEqual(len(h.names), 9)
        self.assertFalse(h.structure_changed)
        self.assertEqual(h.change_count(), 0)

    def test_열을_지워도_고친_행_표시는_맞다(self):
        h = self.house()
        rid = h.ids()[0]
        h.set_cell(rid, "층", "99")
        h.delete_columns(["물건번호"])
        self.assertEqual(h.status(rid), "changed")
        self.assertEqual(h.get(rid, "층"), 99)
        d = h.diff()
        self.assertEqual([(r, h.names[ci]) for r, ci, _, _ in d["updated"]], [(rid, "층")])
        h.undo()
        self.assertEqual(h.get(rid, "물건번호"), "HM-0001")
        self.assertEqual(h.get(rid, "층"), 99)


class 저장(Base):
    def test_열을_고친_엑셀_저장(self):
        src, table, s = self.open("2026_10_부산지사.xlsx")
        n = len(s.rows)
        s.rename_column("신청금액(만원)", "금액")
        s.delete_columns(["지역본부"])
        s.add_column("메모", value="-")
        src.save(table, s)
        again = src.load(table)
        self.assertEqual(again.names[-1], "메모")
        self.assertIn("금액", again.names)
        self.assertNotIn("지역본부", again.names)
        self.assertEqual(len(again.rows), n)
        self.assertEqual(again.get(again.ids()[0], "메모"), "-")
        self.assertEqual(again.get(again.ids()[0], "금액"), s.get(s.ids()[0], "금액"))
        self.assertEqual(len(again.names), len(s.names))

    def test_전처리한_표를_저장하고_다시_읽기(self):
        src, table, p = self.open("주택연금_가입현황_정제전.xlsx")
        p.strip_text()
        p.drop_duplicates()
        p.convert("월지급액", "number")
        p.convert("출생연도", "number")
        p.convert("가입일", "date")
        p.fill_empty("월지급액", "median")
        src.save(table, p)
        again = src.load(table)
        kinds = {c.name: c.kind for c in again.columns}
        self.assertEqual((kinds["월지급액"], kinds["출생연도"], kinds["가입일"]), ("int", "int", "date"))
        self.assertEqual(len(again.rows), 60)
        self.assertEqual(again.empty_count("월지급액"), 0)

    def test_다른_이름으로_저장하면_원본은_그대로(self):
        src, table, s = self.open("S1_보금자리론_취합.xlsx", "월별요약")
        s.set_cell(s.ids()[0], "승인건수", "999")
        target = self.tmp / "새이름.xlsx"
        new = ed.save_as(src, table, s, target)
        self.assertEqual(new.path, target)
        self.assertEqual(new.load("월별요약").get(0, "승인건수"), 999)
        self.assertEqual(len(new.load("전체내역").rows), 630)        # 다른 시트도 함께 따라옴
        self.assertNotEqual(src.load("월별요약").get(0, "승인건수"), 999)   # 원본은 그대로
        self.assertFalse((self.tmp / "백업").exists())               # 새 파일이라 백업을 만들지 않음
        with self.assertRaises(ValueError):
            ed.save_as(src, table, s, self.tmp / "새이름.csv")

    def test_복사본_만들기(self):
        src = ed.open_source(self.tmp / "지점코드표.xlsx")
        first = ed.copy_file(src.path)
        second = ed.copy_file(src.path)
        self.assertEqual(first.name, "지점코드표_복사본.xlsx")
        self.assertEqual(second.name, "지점코드표_복사본2.xlsx")
        self.assertEqual(first.read_bytes(), src.path.read_bytes())


@unittest.skipUnless(has_access(), "pyodbc 또는 Access 드라이버 없음")
class 액세스(Base):
    def test_값_전처리는_저장되고_열_구조_변경은_막는다(self):
        src, table, s = self.open("업무DB.accdb")
        n = s.fill("처리상태", "확인", s.ids()[:3])
        self.assertEqual(n, 3)
        s.rename_column("상품명", "상품")
        with self.assertRaises(ValueError) as e:
            src.save(table, s, key="신청번호")
        self.assertIn("다른 이름으로 저장", str(e.exception))
        s.undo()
        src.save(table, s, key="신청번호")
        self.assertEqual(len(src.load(table).search("확인", "처리상태", exact=True)), 3)

    def test_Access_표를_엑셀로_다른_이름으로_저장(self):
        src, table, s = self.open("업무DB.accdb")
        s.delete_columns(["고객구분"])
        target = self.tmp / "신청내역_정리.xlsx"
        new = ed.save_as(src, table, s, target, key="신청번호")
        self.assertIsInstance(new, ed.ExcelSource)
        again = new.load(new.tables()[0])
        self.assertEqual(len(again.rows), 240)
        self.assertNotIn("고객구분", again.names)
        self.assertEqual(len(src.load(table).names), len(s.names) + 1)       # 원본 Access 는 그대로


if __name__ == "__main__":
    unittest.main(verbosity=2)
