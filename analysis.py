# -*- coding: utf-8 -*-
"""
회귀 분석 도구 — 표에서 '예측할 열'을 다른 열들로 예측하는 모델을 만들고 비교한다 (화면 없이도 쓸 수 있다)

필요한 것:  pip install pandas scikit-learn matplotlib

단계 (위에서 아래로 읽으면 됩니다)
  1. 표 읽기      : to_frame · profile · suggest      편집기의 표 → 분석용 표, 열 살펴보기, 변수 추천
  2. 정제         : tidy · clean                       글자로 적힌 숫자 · 날짜 바로잡기, 겹친 행 · 빈칸 · 글자 열 처리
  3. 이상 탐지    : find_outliers                      사분위수 · 표준점수 · 고립 숲 · 모델 잔차
  4. 모델         : MODELS · train · compare · learning   9가지 모델, 학습과 평가, 교차 검증으로 비교, 학습 곡선
  5. 자동 분석    : auto                               2~4 단계를 한 번에
  6. 결과 내기    : summary_lines · figure · figure_learning · export    요약 글 · 그래프 · 엑셀
     (통계 검정 — t 검정 · 분산분석 · 회귀 계수 검정 — 은 stat_tests.py 에 따로 있다)
  7. 안내문       : guide_markdown                     모델 소개(어원 · 역사 · 기능) 문서 만들기

용어:  예측할 열 = 종속변수(결과, y)      설명하는 열 = 독립변수(원인 후보, X)
"""
import math
import os
import re
import time
import warnings
from dataclasses import dataclass, field

import numpy as np
import pandas as pd
from pandas.api import types as pdt

from data_editor import to_text  # noqa: E402,F401  (값 → 화면에 보이는 글자)

SEED = 42
# scikit-learn 이 CPU 코어 수를 알아내려고 예전 윈도우 명령(wmic)을 부르다가 경고를 내는 것을 막는다
os.environ.setdefault("LOKY_MAX_CPU_COUNT", str(os.cpu_count() or 1))
warnings.filterwarnings("ignore", message="Could not find the number of physical cores")


# =====================================================================
# 1. 표 읽기
# =====================================================================
def to_frame(session, rids=None):
    """편집기의 표(EditSession) → pandas 표. 행번호(rid)가 그대로 index 가 된다."""
    rids = session.ids() if rids is None else list(rids)
    return pd.DataFrame([session.rows[rid] for rid in rids], columns=session.names, index=rids)


def kind_of(col):
    """열의 종류: 숫자 / 날짜 / 범주(몇 가지 값이 되풀이되는 글자) / 글자(행마다 다른 값)."""
    if pdt.is_bool_dtype(col):
        return "범주"
    if pdt.is_numeric_dtype(col):
        return "숫자"
    if pdt.is_datetime64_any_dtype(col):
        return "날짜"
    values = col.dropna()
    if len(values) == 0:
        return "글자"
    if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in values):
        return "숫자"
    unique = values.astype(str).nunique()
    return "범주" if unique <= 30 and unique <= max(2, len(values) // 2) else "글자"


def profile(df):
    """열마다 종류 · 빈칸 수 · 서로 다른 값 수를 알려 준다. (글자로 적힌 숫자 · 날짜는 바로잡은 뒤에 본다)"""
    df = tidy(df)[0]
    out = []
    for name in df.columns:
        col = df[name]
        info = {"name": name, "kind": kind_of(col), "missing": int(col.isna().sum()),
                "unique": int(col.dropna().astype(str).nunique())}
        if info["kind"] == "숫자":
            nums = pd.to_numeric(col, errors="coerce")
            info.update(min=float(nums.min()), max=float(nums.max()), mean=float(nums.mean()))
        out.append(info)
    return out


def suggest(df):
    """(예측할 열, 설명하는 열들) 을 추천한다: 값이 다양한 마지막 숫자 열을 예측하고, 번호표 같은 글자 열은 뺀다."""
    info = profile(df)
    numeric = [p["name"] for p in info if p["kind"] == "숫자" and p["unique"] > 10]
    target = numeric[-1] if numeric else None
    features = [p["name"] for p in info if p["name"] != target and p["kind"] != "글자" and p["unique"] > 1]
    return target, features


# =====================================================================
# 2. 정제
# =====================================================================
_NUMBER = re.compile(r"^[-+]?\d[\d,]*(\.\d+)?\s*(원|%)?$")
_DATE = re.compile(r"^(\d{4})[-./]?(\d{1,2})[-./]?(\d{1,2})$")


def _to_date(text):
    m = _DATE.match(text)
    if not m:
        return None
    try:
        return pd.Timestamp(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


def tidy(df, rate=0.7):
    """글자로 들어온 값을 바로잡는다 — 사람이 손으로 넣은 표에서 흔한 것들.
      · 앞뒤 빈칸 지우기 (' 김하나' → '김하나'),  빈 글자는 빈칸으로
      · '2,430,000원' · ' 1,140,000 ' 같은 글자를 숫자로
      · '2026/07/15' · '2026.01.03' · '20260613' 처럼 표기가 섞인 날짜를 날짜로
    한 열의 값 중 rate(70%) 이상이 그렇게 읽힐 때만 바꾸고, 못 읽은 칸은 빈칸으로 둔다.
    (바꾼 표, 무엇을 했는지 적은 글 목록) 을 돌려준다. 원래 표는 건드리지 않는다."""
    out, notes = df.copy(), []
    for name in df.columns:
        col = df[name]
        if pdt.is_numeric_dtype(col) or pdt.is_bool_dtype(col) or pdt.is_datetime64_any_dtype(col):
            continue
        values = [None if pd.isna(v) else v for v in col.tolist()]
        texts = [v for v in values if v is not None]
        if not texts or not all(isinstance(v, str) for v in texts):
            continue
        stripped = [None if v is None else (v.strip() or None) for v in values]
        spaces = sum(1 for a, b in zip(values, stripped) if a is not None and a != b)
        filled = [v for v in stripped if v is not None]
        fixed = stripped
        if filled:
            as_number = [bool(_NUMBER.match(v)) for v in filled]
            as_date = [_to_date(v) is not None for v in filled]
            if sum(as_number) / len(filled) >= rate:
                fixed = [float(re.sub(r"[,\s원%]", "", v)) if v is not None and _NUMBER.match(v) else None for v in stripped]
                bad = len(filled) - sum(as_number)
                notes.append(f"[{name}] 글자로 적힌 숫자 {sum(as_number):,}칸을 숫자로 바꿨습니다 (쉼표 · '원' 제거)."
                             + (f" 읽지 못한 {bad:,}칸은 빈칸으로 두었습니다." if bad else ""))
                out[name] = pd.Series(fixed, index=df.index, dtype=float)
                continue
            if sum(as_date) / len(filled) >= rate:
                bad = len(filled) - sum(as_date)
                notes.append(f"[{name}] 표기가 섞인 날짜 {sum(as_date):,}칸을 날짜로 바꿨습니다."
                             + (f" 읽지 못한 {bad:,}칸은 빈칸으로 두었습니다." if bad else ""))
                out[name] = pd.to_datetime(pd.Series([None if v is None else _to_date(v) for v in stripped], index=df.index))
                continue
        if spaces:
            notes.append(f"[{name}] 앞뒤 빈칸이 있던 {spaces:,}칸을 다듬었습니다.")
            out[name] = pd.Series(fixed, index=df.index, dtype=object)
    return out, notes


@dataclass
class Clean:
    target: str
    features: list            # 실제로 쓴 설명 열 (원래 이름)
    X: pd.DataFrame           # 모델에 넣는 숫자 표 (글자 열은 '열=값' 0/1 열로 펼쳐짐)
    y: pd.Series
    frame: pd.DataFrame       # 정제한 뒤의 원래 모양 표 (X 와 같은 행)
    groups: dict              # X 의 열 이름 → 원래 열 이름
    report: list              # 무엇을 했는지 사람이 읽는 글
    dropped: dict             # 버린 행 수: duplicate / no_target / missing
    filled: dict              # 빈칸을 채운 칸 수: 열 이름 → 개수


FILL_NAME = {"median": "중앙값", "mean": "평균", "drop": "행 제거"}


def num(v):
    """숫자를 읽기 쉽게: 큰 값은 쉼표, 작은 값은 유효 숫자 3자리."""
    return f"{v:,.0f}" if abs(v) >= 100 else f"{v:,.3g}"


def clean(df, target, features, missing="median", drop_duplicates=True, max_levels=20):
    """분석할 수 있게 표를 다듬는다.
    missing: 설명 열의 빈칸을 "median"(중앙값) · "mean"(평균) 으로 채우거나 "drop"(그 행을 버림).
    글자 열의 빈칸은 가장 흔한 값으로 채운다. 예측할 열이 빈 행은 배울 수 없으므로 항상 버린다."""
    features = list(features)
    if target not in df.columns:
        raise ValueError("예측할 열을 고르세요.")
    if not features:
        raise ValueError("설명하는 열을 하나 이상 고르세요.")
    if target in features:
        raise ValueError(f"[{target}] 은 예측할 열이라서 설명하는 열에 넣을 수 없습니다.")
    if missing not in FILL_NAME:
        raise ValueError(f"빈칸 처리 방법을 알 수 없습니다: {missing}")
    report, dropped, filled = [], {"duplicate": 0, "no_target": 0, "missing": 0}, {}
    work, notes = tidy(df)
    report.append(f"처음 {len(work):,}행 · 예측할 열 [{target}] · 설명하는 열 {len(features)}개")
    report += notes

    if drop_duplicates:
        dup = work.astype(str).duplicated()
        dropped["duplicate"] = int(dup.sum())
        if dropped["duplicate"]:
            work = work[~dup.to_numpy()]
            report.append(f"겹친 행 {dropped['duplicate']:,}개를 뺐습니다. (모든 칸이 똑같은 행)")

    y = pd.to_numeric(work[target], errors="coerce")
    if y.notna().sum() < 10 or kind_of(work[target]) != "숫자":
        raise ValueError(f"[{target}] 은 숫자 열이 아니어서 예측할 수 없습니다. 숫자가 들어 있는 열을 고르세요.")
    dropped["no_target"] = int(y.isna().sum())
    if dropped["no_target"]:
        work, y = work[y.notna().to_numpy()], y[y.notna()]
        report.append(f"[{target}] 이 빈 {dropped['no_target']:,}행을 뺐습니다. (정답이 없으면 배울 수 없음)")

    used, parts, groups, keep = [], [], {}, pd.Series(True, index=work.index)
    frame = work.copy()
    for name in features:
        col, kind = work[name], kind_of(work[name])
        n_missing = int(col.isna().sum())
        if kind == "글자":
            report.append(f"[{name}] 은 행마다 값이 달라(번호 · 이름 같은 열) 쓰지 않았습니다.")
            continue
        if col.dropna().astype(str).nunique() <= 1:
            report.append(f"[{name}] 은 값이 한 가지뿐이라 쓰지 않았습니다.")
            continue
        used.append(name)
        if kind == "범주":
            text = col.astype(object).where(col.notna(), None)
            text = text.map(lambda v: None if v is None else ("예" if v is True else "아니오" if v is False else str(v)))
            if n_missing:
                if missing == "drop":
                    keep &= text.notna()
                else:
                    most = text.dropna().mode().iloc[0]
                    text = text.fillna(most)
                    filled[name] = n_missing
                    report.append(f"[{name}] 빈칸 {n_missing:,}칸을 가장 흔한 값 '{most}' 로 채웠습니다.")
            top = text.value_counts().index[:max_levels]
            if text.nunique() > max_levels:
                text = text.where(text.isin(top) | text.isna(), "기타")
                report.append(f"[{name}] 값이 많아 흔한 {max_levels}가지만 남기고 나머지는 '기타' 로 묶었습니다.")
            frame[name] = text
            dummies = pd.get_dummies(text, prefix=name, prefix_sep="=", dtype=float)
            for c in dummies.columns:
                groups[c] = name
            parts.append(dummies)
            continue
        if kind == "날짜":
            dates = pd.to_datetime(col, errors="coerce")
            nums = (dates - dates.min()).dt.total_seconds() / 86400.0        # 가장 이른 날부터 며칠째인지
            report.append(f"[{name}] 날짜를 '첫 날부터 지난 날 수' 로 바꿨습니다.")
        else:
            nums = pd.to_numeric(col, errors="coerce").astype(float)
        n_missing = int(nums.isna().sum())
        if n_missing:
            if missing == "drop":
                keep &= nums.notna()
            else:
                value = float(nums.median() if missing == "median" else nums.mean())
                nums = nums.fillna(value)
                filled[name] = n_missing
                report.append(f"[{name}] 빈칸 {n_missing:,}칸을 {FILL_NAME[missing]} {num(value)} 로 채웠습니다.")
        if kind != "날짜":
            frame[name] = nums
        groups[name] = name
        parts.append(nums.rename(name).to_frame())

    if not parts:
        raise ValueError("쓸 수 있는 설명 열이 없습니다. 숫자 열이나 값이 되풀이되는 글자 열을 고르세요.")
    X = pd.concat(parts, axis=1)
    if not keep.all():
        dropped["missing"] = int((~keep).sum())
        X, y, frame = X[keep.to_numpy()], y[keep.to_numpy()], frame[keep.to_numpy()]
        report.append(f"설명 열에 빈칸이 있는 {dropped['missing']:,}행을 뺐습니다.")
    if len(X) < 30:
        raise ValueError(f"정제하고 나니 {len(X)}행뿐입니다. 회귀 분석을 하려면 적어도 30행은 있어야 합니다.")
    report.append(f"정제 후 {len(X):,}행 · 모델에 들어가는 열 {X.shape[1]}개")
    return Clean(target, used, X.astype(float), y.astype(float), frame, groups, report, dropped, filled)


# =====================================================================
# 3. 이상 탐지
# =====================================================================
@dataclass
class Outliers:
    method: str
    mask: pd.Series           # True = 이상하다고 본 행
    reasons: pd.Series        # 왜 그렇게 봤는지 (이상이 아니면 빈 글자)
    report: list


OUTLIER_METHODS = {
    "none": ("탐지 안 함", "이상값을 찾지 않고 모든 행으로 학습합니다."),
    "iqr": ("사분위수 범위 (IQR)",
            "값을 크기순으로 세웠을 때 가운데 절반이 차지하는 폭(IQR)의 몇 배 넘게 벗어난 값을 찾습니다. "
            "터키(John Tukey)가 1977년 상자그림과 함께 널리 알린 방법으로, 계산이 단순하고 설명하기 쉽습니다."),
    "zscore": ("표준점수 (Z-score)",
               "평균에서 표준편차의 몇 배 떨어졌는지(z) 를 재서 3배 넘는 값을 찾습니다. "
               "값이 종 모양으로 퍼져 있을 때 잘 맞고, 아주 큰 이상값이 평균과 표준편차를 끌어당기는 약점이 있습니다."),
    "iforest": ("고립 숲 (Isolation Forest)",
                "무작위로 칸을 나누다 보면 남다른 행은 금방 혼자 남는다(고립된다)는 점을 씁니다. "
                "2008년 Liu · Ting · Zhou 가 발표했고, 여러 열을 함께 봐서 '각 값은 평범한데 조합이 이상한' 행도 찾습니다."),
    "model": ("모델 잔차 (예측과 크게 다른 행)",
              "랜덤 포레스트가 '그 행을 빼고' 배운 뒤 예측한 값과 실제 값을 견주어, 비율이 크게 어긋난 행을 찾습니다. "
              "0 을 하나 더 넣은 입력 실수처럼 '다른 열에 비해 말이 안 되는 값'을 가장 잘 잡습니다. 대신 조금 느립니다."),
    "combo": ("종합 (사분위수 + 모델 잔차) — 권장",
              "값 하나가 범위를 크게 벗어난 행(사분위수)과 다른 열에 비해 말이 안 되는 행(모델 잔차)을 함께 찾습니다. "
              "두 방법은 잡는 것이 서로 달라서, 합치면 입력 실수를 가장 빠짐없이 걸러 냅니다."),
}


def find_outliers(c, method="combo", level=None, seed=SEED):
    """정제한 표(Clean)에서 이상한 행을 찾는다. level: iqr=몇 배(기본 3.0), zscore=몇 배(3.0),
    iforest=이상 비율(0.05), model=허용 배수(3.5). combo 는 iqr 과 model 을 기본값으로 함께 쓴다."""
    if method not in OUTLIER_METHODS:
        raise ValueError(f"이상 탐지 방법을 알 수 없습니다: {method}")
    index = c.X.index
    mask = pd.Series(False, index=index)
    why = pd.Series("", index=index, dtype=object)
    label = OUTLIER_METHODS[method][0]
    if method == "none":
        return Outliers(method, mask, why, ["이상 탐지를 하지 않았습니다."])

    def add(rows, text):
        for rid, t in zip(rows, text):
            why[rid] = (why[rid] + " · " if why[rid] else "") + t
        mask[rows] = True

    numeric = {c.target: c.y}
    numeric.update({name: c.X[name] for name in c.features if name in c.X.columns})
    if method in ("iqr", "combo"):
        k = 3.0 if level is None or method == "combo" else float(level)
        for name, col in numeric.items():
            if col.nunique() < 5:
                continue
            q1, q3 = col.quantile(0.25), col.quantile(0.75)
            low, high = q1 - k * (q3 - q1), q3 + k * (q3 - q1)
            bad = col[(col < low) | (col > high)]
            add(bad.index, [f"[{name}] {num(v)} (정상 범위 {num(low)} ~ {num(high)})" for v in bad])
    if method == "zscore":
        k = 3.0 if level is None else float(level)
        for name, col in numeric.items():
            std = col.std()
            if not std or math.isnan(std):
                continue
            z = (col - col.mean()) / std
            bad = z[z.abs() > k]
            add(bad.index, [f"[{name}] {num(col[rid])} (평균에서 표준편차의 {abs(v):.1f}배)" for rid, v in bad.items()])
    elif method == "iforest":
        from sklearn.ensemble import IsolationForest
        rate = 0.05 if level is None else float(level)
        data = c.X.assign(**{"__y__": c.y})
        forest = IsolationForest(n_estimators=200, contamination=rate, random_state=seed).fit(data)
        score = pd.Series(-forest.score_samples(data), index=index)
        bad = score[forest.predict(data) == -1]
        add(bad.index, [f"여러 열을 함께 보면 남다름 (고립 점수 {v:.2f})" for v in bad])
    if method in ("model", "combo"):
        from sklearn.ensemble import RandomForestRegressor
        from sklearn.model_selection import KFold, cross_val_predict
        k = 3.5 if level is None or method == "combo" else float(level)
        model = RandomForestRegressor(n_estimators=200, min_samples_leaf=3, random_state=seed)
        ratio = bool((c.y > 0).all())                                    # 값이 모두 양수면 '몇 배인지' 로 견준다
        learn = np.log(c.y) if ratio else c.y                            # (그때는 로그 값으로 배운다)
        pred = pd.Series(cross_val_predict(model, c.X, learn, cv=KFold(5, shuffle=True, random_state=seed)), index=index)
        gap = learn - pred
        pred = np.exp(pred) if ratio else pred
        center = gap.median()
        spread = (gap - center).abs().median() * 1.4826 or gap.std() or 1.0   # 이상값에 덜 흔들리는 퍼짐 정도
        z = (gap - center) / spread
        bad = z[z.abs() > k]
        if ratio:
            text = [f"[{c.target}] {num(c.y[rid])} — 다른 열로 보면 {num(pred[rid])} 쯤 ({c.y[rid] / pred[rid]:.2f}배)"
                    for rid in bad.index]
        else:
            text = [f"[{c.target}] {num(c.y[rid])} — 다른 열로 보면 {num(pred[rid])} 쯤" for rid in bad.index]
        add(bad.index, text)
    n = int(mask.sum())
    return Outliers(method, mask, why, [f"이상 탐지 [{label}] : {len(index):,}행 중 {n:,}행 ({n / len(index):.1%})"])


# =====================================================================
# 4. 모델
# =====================================================================
@dataclass
class ModelInfo:
    name: str                 # 우리말 이름
    english: str
    family: str               # 머신러닝 / 딥러닝
    origin: str               # 이름의 뜻 (어원)
    history: str              # 언제 누가
    what: str                 # 어떻게 예측하나
    when: str                 # 언제 쓰면 좋나
    build: object = field(repr=False, default=None)     # seed → 모델 객체


def _scaled(estimator, scale_target=False):
    """값의 크기에 민감한 모델은 먼저 열마다 평균 0 · 표준편차 1 로 맞춘다."""
    from sklearn.compose import TransformedTargetRegressor
    from sklearn.pipeline import make_pipeline
    from sklearn.preprocessing import StandardScaler
    pipe = make_pipeline(StandardScaler(), estimator)
    return TransformedTargetRegressor(regressor=pipe, transformer=StandardScaler()) if scale_target else pipe


def _linear(seed):
    from sklearn.linear_model import LinearRegression
    return LinearRegression()


def _ridge(seed):
    from sklearn.linear_model import RidgeCV
    return _scaled(RidgeCV(alphas=np.logspace(-3, 3, 13)))


def _lasso(seed):
    from sklearn.linear_model import LassoCV
    return _scaled(LassoCV(cv=5, random_state=seed, max_iter=20000))


def _tree(seed):
    from sklearn.tree import DecisionTreeRegressor
    return DecisionTreeRegressor(max_depth=8, min_samples_leaf=5, random_state=seed)


def _forest(seed):
    from sklearn.ensemble import RandomForestRegressor
    return RandomForestRegressor(n_estimators=300, min_samples_leaf=2, random_state=seed)


def _boosting(seed):
    from sklearn.ensemble import GradientBoostingRegressor
    return GradientBoostingRegressor(n_estimators=300, max_depth=3, learning_rate=0.08, subsample=0.9, random_state=seed)


def _knn(seed):
    from sklearn.neighbors import KNeighborsRegressor
    return _scaled(KNeighborsRegressor(n_neighbors=7, weights="distance"))


def _svr(seed):
    from sklearn.svm import SVR
    return _scaled(SVR(C=10.0, epsilon=0.05), scale_target=True)


def _mlp(seed):
    from sklearn.neural_network import MLPRegressor
    return _scaled(MLPRegressor(hidden_layer_sizes=(64, 32), alpha=1e-3, max_iter=2000, random_state=seed), scale_target=True)


MODELS = {
    "linear": ModelInfo(
        "선형 회귀", "Linear Regression", "머신러닝",
        "regression 은 라틴어 regredi(되돌아가다)에서 왔습니다. 골턴이 1886년 '키가 아주 큰 부모의 자녀는 평균 쪽으로 "
        "되돌아간다'는 현상을 regression 이라 부른 것이 통계 용어가 됐습니다. linear 는 선(line)이라는 뜻입니다.",
        "계산법인 최소제곱법은 르장드르(1805)와 가우스(1809)가 천체 궤도를 구하려고 만들었고, 골턴(1886)과 "
        "피어슨 · 율(1890년대)이 지금의 회귀 분석으로 다듬었습니다. 가장 오래되고 가장 많이 쓰이는 모델입니다.",
        "설명하는 열마다 가중치(계수)를 곱해 더한 값으로 예측합니다. '면적이 1㎡ 늘면 감정가가 얼마 오른다'처럼 "
        "계수를 그대로 읽을 수 있습니다.",
        "먼저 돌려 보는 기준선. 결과를 말로 설명해야 하는 보고서. 관계가 곧은 선에 가까울 때.", _linear),
    "ridge": ModelInfo(
        "릿지 회귀", "Ridge Regression", "머신러닝",
        "ridge 는 산등성이라는 뜻입니다. 호얼(Hoerl)이 1962년 반응 표면의 능선을 따라가는 'ridge analysis' 에서 "
        "쓰던 이름을 회귀에 가져왔습니다.",
        "호얼과 케너드가 1970년에 발표했습니다. 같은 생각은 수학에서 티호노프 정규화(1943)로 먼저 나왔습니다.",
        "선형 회귀에 '계수를 너무 키우지 말라'는 벌점(계수 제곱의 합)을 더합니다. 서로 닮은 열이 여럿 있어도 "
        "계수가 널뛰지 않습니다. 벌점의 세기는 교차 검증으로 자동으로 고릅니다.",
        "설명 열이 많거나 서로 비슷한 열(면적과 방 수처럼)이 섞여 있을 때.", _ridge),
    "lasso": ModelInfo(
        "라쏘 회귀", "Lasso Regression", "머신러닝",
        "Least Absolute Shrinkage and Selection Operator(절댓값을 가장 작게 줄이고 골라내는 연산)의 머리글자입니다. "
        "올가미(lasso)라는 뜻도 겹쳐, 필요 없는 열을 올가미로 조여 없앤다는 느낌을 줍니다.",
        "팁시라니(Tibshirani)가 1996년에 발표했습니다.",
        "벌점을 계수의 절댓값 합으로 줍니다. 그러면 도움이 안 되는 열의 계수가 정확히 0 이 되어, "
        "어떤 열이 쓸모 있는지 모델이 스스로 골라 줍니다.",
        "열이 아주 많은데 그중 몇 개만 중요할 것 같을 때. 변수 선택을 겸하고 싶을 때.", _lasso),
    "tree": ModelInfo(
        "의사결정나무", "Decision Tree", "머신러닝",
        "질문이 나뭇가지처럼 갈라져 내려가는 모양에서 온 이름입니다. 맨 위 질문이 뿌리, 마지막 답이 잎입니다.",
        "모건과 손퀴스트의 AID(1963)가 시작이고, 브라이먼 등이 1984년 CART 로 회귀까지 정리했습니다. "
        "퀸란의 ID3(1986) · C4.5(1993)도 널리 쓰였습니다.",
        "'면적이 84㎡ 이상인가?' 같은 예/아니오 질문을 되풀이해 행을 무리로 나누고, 무리의 평균으로 예측합니다. "
        "굽은 관계와 열끼리 얽힌 효과를 저절로 잡습니다.",
        "규칙을 그림으로 보여 주고 싶을 때. 혼자서는 데이터에 지나치게 맞춰지기 쉬워 아래 두 모델의 재료로 더 많이 씁니다.",
        _tree),
    "forest": ModelInfo(
        "랜덤 포레스트", "Random Forest", "머신러닝",
        "나무(tree)가 모이면 숲(forest)입니다. 나무마다 행과 열을 무작위(random)로 뽑아 주므로 random forest 입니다.",
        "틴 캄 호(Tin Kam Ho)가 1995년 '무작위 결정 숲'을 제안했고, 브라이먼이 자신의 배깅(1996)과 합쳐 2001년에 "
        "지금의 랜덤 포레스트를 발표했습니다.",
        "조금씩 다른 나무 수백 그루의 예측을 평균합니다. 나무 하나의 실수가 서로 지워져 안정적이고, "
        "손볼 설정이 거의 없습니다.",
        "무엇을 쓸지 모를 때의 안전한 선택. 이상값과 굽은 관계에 강합니다.", _forest),
    "boosting": ModelInfo(
        "그레이디언트 부스팅", "Gradient Boosting", "머신러닝",
        "boost 는 밀어 올린다는 뜻입니다. 약한 모델을 차례로 보태 성능을 끌어올립니다. gradient(기울기)는 "
        "오차가 가장 빨리 줄어드는 방향을 뜻합니다.",
        "'약한 학습기를 모아 강하게 만들 수 있는가'라는 물음에 샤파이어가 1990년 답했고, 프로인트와 샤파이어의 "
        "AdaBoost(1995)를 거쳐 프리드먼이 1999~2001년 그레이디언트 부스팅으로 일반화했습니다. "
        "XGBoost(2014) · LightGBM(2017)이 그 후손입니다.",
        "작은 나무를 하나 만들고, 그 나무가 틀린 만큼을 다음 나무가 배우는 일을 수백 번 되풀이합니다. "
        "표 모양 데이터에서는 지금도 가장 정확한 축에 듭니다.",
        "정확도가 가장 중요할 때. 표 데이터 경진 대회의 단골 우승 모델입니다.", _boosting),
    "knn": ModelInfo(
        "k-최근접 이웃", "k-Nearest Neighbors", "머신러닝",
        "nearest neighbors 는 가장 가까운 이웃, k 는 몇 명의 이웃을 볼지 정하는 숫자입니다.",
        "픽스와 호지스가 1951년 미 공군 보고서에서 제안했고, 커버와 하트가 1967년 성질을 증명했습니다.",
        "새 행과 가장 닮은 k개 행을 찾아 그 값들의 (거리로 가중한) 평균으로 예측합니다. 따로 배우는 과정이 없고 "
        "'비슷한 사례를 찾아본다'는 설명이 그대로 통합니다.",
        "'비슷한 물건 몇 건의 평균'이라는 설명이 필요할 때. 열이 많아지면 '가깝다'의 뜻이 흐려져 성능이 떨어집니다.", _knn),
    "svr": ModelInfo(
        "서포트 벡터 회귀", "Support Vector Regression", "머신러닝",
        "support vector 는 경계를 떠받치는(support) 점(vector)이라는 뜻입니다. 모델을 정하는 것은 "
        "모든 행이 아니라 경계에 걸린 몇몇 행뿐입니다.",
        "바프닉과 체르보넨키스가 1960년대에 바탕을 만들었고, 1992년 커널 기법(보저 · 기용 · 바프닉), "
        "1995년 소프트 마진(코르테스 · 바프닉)을 거쳐 1996년 드러커 등이 회귀용 SVR 을 발표했습니다.",
        "예측선 둘레에 '이 정도 오차는 봐준다'는 띠를 두르고, 띠 밖으로 나간 행만 벌점을 줍니다. "
        "커널로 굽은 관계도 다룹니다.",
        "행 수가 수천 건 이하이고 관계가 매끄럽게 굽어 있을 때. 행이 많아지면 느려집니다.", _svr),
    "mlp": ModelInfo(
        "신경망 (다층 퍼셉트론)", "Neural Network · Multi-Layer Perceptron", "딥러닝",
        "neural network 는 뇌의 신경세포(neuron) 연결을 본떴다는 뜻이고, perceptron 은 perceive(알아차리다)에 "
        "기계 이름에 붙이던 -tron 을 더한 말입니다. 층을 깊게(deep) 쌓은 신경망을 배우는 일이 딥러닝입니다.",
        "매컬러와 피츠의 인공 뉴런(1943), 로젠블랫의 퍼셉트론(1958)으로 시작했습니다. 민스키와 페퍼트가 1969년 "
        "한계를 짚어 한동안 식었다가, 1986년 루멜하트 · 힌턴 · 윌리엄스의 역전파로 여러 층을 배울 수 있게 됐습니다. "
        "2006년 힌턴의 심층 신뢰망과 2012년 AlexNet 뒤로 딥러닝 시대가 열렸습니다.",
        "입력에 가중치를 곱해 더하고 굽은 함수를 통과시키는 층을 여러 겹 쌓습니다. 여기서는 64 · 32개 뉴런의 "
        "두 층짜리 작은 신경망을 씁니다. 어떤 모양의 관계든 흉내 낼 수 있습니다.",
        "행이 아주 많거나 이미지 · 글 · 소리처럼 표가 아닌 데이터. 작은 표에서는 나무 계열과 엇비슷하거나 "
        "뒤지는 일도 많고, 관계가 매끄러우면 앞서기도 합니다 — [자동 분석] 으로 직접 견주어 보세요.", _mlp),
}


@dataclass
class Result:
    key: str
    name: str
    metrics: dict             # r2_train · r2_test · mae · rmse · mape · n_train · n_test
    y_test: pd.Series
    pred_test: pd.Series
    predictions: pd.Series    # 정제한 모든 행의 예측값 (학습에서 뺀 이상 행 포함)
    importance: pd.Series     # 원래 열 이름 → 중요도 (합 1, 큰 순서)
    seconds: float
    model: object = field(repr=False, default=None)


def _usable(c, exclude):
    if exclude is None:
        return c.X, c.y
    keep = ~exclude.reindex(c.X.index, fill_value=False).to_numpy()
    if keep.sum() < 30:
        raise ValueError("이상 행을 빼고 나니 30행이 안 됩니다. 이상 탐지를 느슨하게 하거나 끄세요.")
    return c.X[keep], c.y[keep]


def train(c, key, exclude=None, test_size=0.2, seed=SEED):
    """모델 하나를 학습하고 평가한다. 전체의 20% 는 배우지 않고 남겨 두었다가 시험용으로 쓴다."""
    from sklearn.metrics import mean_absolute_error, r2_score
    from sklearn.model_selection import train_test_split
    info = MODELS[key]
    X, y = _usable(c, exclude)
    X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=test_size, random_state=seed)
    start = time.perf_counter()
    model = info.build(seed).fit(X_tr, y_tr)
    seconds = time.perf_counter() - start
    pred = pd.Series(model.predict(X_te), index=X_te.index)
    nonzero = y_te != 0
    metrics = {
        "r2_train": round(float(r2_score(y_tr, model.predict(X_tr))), 4),
        "r2_test": round(float(r2_score(y_te, pred)), 4),
        "mae": round(float(mean_absolute_error(y_te, pred)), 2),
        "rmse": round(float(np.sqrt(((y_te - pred) ** 2).mean())), 2),
        "mape": round(float(((y_te - pred).abs() / y_te.abs())[nonzero].mean() * 100), 2),
        "n_train": len(X_tr), "n_test": len(X_te),
    }
    predictions = pd.Series(model.predict(c.X), index=c.X.index)
    return Result(key, info.name, metrics, y_te, pred, predictions,
                  _importance(model, c, X_te, y_te, seed), seconds, model)


def _importance(model, c, X_te, y_te, seed, repeats=5):
    """열 중요도: 시험 데이터에서 그 열의 값을 뒤섞었을 때 정확도(R²)가 얼마나 떨어지는지로 잰다.
    어떤 모델에도 똑같이 쓸 수 있고, 글자 열에서 펼친 0/1 열들은 한 묶음으로 함께 섞는다."""
    from sklearn.metrics import r2_score
    rng = np.random.default_rng(seed)
    base = r2_score(y_te, model.predict(X_te))
    drops = {}
    for feature in c.features:
        cols = [col for col, origin in c.groups.items() if origin == feature]
        total = 0.0
        for _ in range(repeats):
            shuffled = X_te.copy()
            order = rng.permutation(len(X_te))
            shuffled[cols] = X_te[cols].to_numpy()[order]
            total += base - r2_score(y_te, model.predict(shuffled))
        drops[feature] = max(total / repeats, 0.0)
    imp = pd.Series(drops, dtype=float)
    imp = imp / imp.sum() if imp.sum() > 0 else pd.Series(1.0 / len(imp), index=imp.index)
    return imp.sort_values(ascending=False)


def compare(c, keys=None, exclude=None, cv=5, seed=SEED, progress=None, on_row=None):
    """여러 모델을 교차 검증(데이터를 cv 조각으로 나눠 돌아가며 시험)으로 견주어 점수 순으로 돌려준다.
    progress(모델 이름, 몇 번째, 전체 수) 는 모델을 시작할 때, on_row(결과 한 줄) 는 끝날 때마다 불린다."""
    from sklearn.model_selection import KFold, cross_validate
    X, y = _usable(c, exclude)
    folds = KFold(cv, shuffle=True, random_state=seed)
    table = []
    keys = list(keys or MODELS)
    for i, key in enumerate(keys):
        info = MODELS[key]
        if progress:
            progress(info.name, i, len(keys))
        start = time.perf_counter()
        scores = cross_validate(info.build(seed), X, y, cv=folds, scoring=("r2", "neg_root_mean_squared_error"))
        table.append({"key": key, "name": info.name, "family": info.family,
                      "r2": round(float(scores["test_r2"].mean()), 4), "std": round(float(scores["test_r2"].std()), 4),
                      "rmse": round(float(-scores["test_neg_root_mean_squared_error"].mean()), 2),
                      "seconds": round(time.perf_counter() - start, 2)})
        if on_row:
            on_row(table[-1])
    return sorted(table, key=lambda row: row["r2"], reverse=True)


def learning(c, key, exclude=None, seed=SEED, model=None):
    """학습 과정을 숫자로: ① 학습 곡선 — 행을 20% · 40% · … · 100% 만 써서 배웠을 때의 학습 점수와 시험 점수(R²)
    ② 단계별 오차 — 그레이디언트 부스팅은 나무를 하나 보탤 때마다, 신경망은 한 번 되풀이할 때마다 줄어드는 오차.
    model 에 이미 학습한 모델(Result.model)을 주면 ② 에 그것을 쓴다."""
    from sklearn.model_selection import KFold, learning_curve
    X, y = _usable(c, exclude)
    sizes, train_scores, valid_scores = learning_curve(
        MODELS[key].build(seed), X, y, train_sizes=np.linspace(0.2, 1.0, 5),
        cv=KFold(3, shuffle=True, random_state=seed), scoring="r2")
    if model is None and key in ("boosting", "mlp"):
        model = MODELS[key].build(seed).fit(X, y)
    steps, steps_name = None, ""
    if key == "boosting":
        steps, steps_name = [float(v) for v in model.train_score_], "나무 수"
    elif key == "mlp":
        steps, steps_name = [float(v) for v in model.regressor_[-1].loss_curve_], "반복 횟수"
    return {"key": key, "sizes": [int(s) for s in sizes],
            "train": [round(float(v), 4) for v in train_scores.mean(axis=1)],
            "valid": [round(float(v), 4) for v in valid_scores.mean(axis=1)],
            "valid_std": [round(float(v), 4) for v in valid_scores.std(axis=1)],
            "steps": steps, "steps_name": steps_name}


# =====================================================================
# 5. 자동 분석 — 정제 → 이상 탐지 → 모델 비교 → 가장 좋은 모델 학습
# =====================================================================
@dataclass
class Auto:
    clean: Clean
    outliers: Outliers
    table: list
    best: Result
    log: list


def auto(df, target=None, features=None, missing="median", outlier_method="combo", progress=None, on_row=None, seed=SEED):
    """사람이 고르던 것을 규칙으로 대신한다. target · features 를 주지 않으면 추천값을 쓴다.
    progress(글, 진행률 0~1) 로 지금 무엇을 하는지, on_row(줄) 로 모델 하나가 끝날 때마다 결과를 알려 준다."""
    say = progress or (lambda text, fraction=None: None)
    guess_target, guess_features = suggest(df)
    target = target or guess_target
    features = features or [f for f in guess_features if f != target]
    if target is None:
        raise ValueError("숫자 열이 없어 예측할 열을 정할 수 없습니다.")
    say("데이터를 정제하는 중", 0.03)
    c = clean(df, target, features, missing=missing)
    say("이상한 행을 찾는 중", 0.08)
    o = find_outliers(c, outlier_method, seed=seed)
    table = compare(c, exclude=o.mask, seed=seed, on_row=on_row,
                    progress=lambda name, i, n: say(f"모델 비교 중 ({i + 1}/{n}) · {name}", 0.15 + 0.65 * i / n))
    say(f"가장 좋은 모델({table[0]['name']})로 학습하는 중", 0.82)
    best = train(c, table[0]["key"], exclude=o.mask, seed=seed)
    log = [f"1. 변수: 예측할 열 [{target}] · 설명하는 열 {', '.join(c.features)}",
           f"2. 정제: {c.report[-1]} (빈칸은 {FILL_NAME[missing]})",
           f"3. {o.report[0]}",
           f"4. 모델 비교: {len(table)}개를 5조각 교차 검증 — 1위 {table[0]['name']} (R² {table[0]['r2']:.3f}), "
           f"꼴찌 {table[-1]['name']} (R² {table[-1]['r2']:.3f})",
           f"5. 최종 학습: {best.name} · 시험 R² {best.metrics['r2_test']:.3f} · 평균 오차율 {best.metrics['mape']:.1f}%"]
    return Auto(c, o, table, best, log)


# =====================================================================
# 6. 결과 내기 — 요약 글 · 그래프 · 엑셀
# =====================================================================
def grade(r2):
    return "매우 잘 맞음" if r2 >= 0.9 else "잘 맞음" if r2 >= 0.75 else "어느 정도 맞음" if r2 >= 0.5 else "잘 안 맞음"


def summary_lines(c, o, result, table=None):
    m = result.metrics
    lines = [f"■ 모델: {result.name}  ({MODELS[result.key].english} · {MODELS[result.key].family})",
             f"■ 예측할 열: {c.target}    설명하는 열: {', '.join(c.features)}", "",
             f"시험 데이터 {m['n_test']:,}행으로 평가 (학습 {m['n_train']:,}행, 학습 {result.seconds:.2f}초)",
             f"  · 결정계수 R²   {m['r2_test']:.3f}   → {grade(m['r2_test'])}  (1 에 가까울수록 좋음, 학습 데이터에서는 {m['r2_train']:.3f})",
             f"  · 평균 오차 MAE {m['mae']:,.0f}   · RMSE {m['rmse']:,.0f}   · 평균 오차율 {m['mape']:.1f}%"]
    if m["r2_train"] - m["r2_test"] > 0.15:
        lines.append("  ※ 학습 점수가 시험 점수보다 많이 높습니다 — 학습 데이터를 외운(과적합) 것일 수 있습니다.")
    lines += ["", "열 중요도 (그 열을 뒤섞으면 정확도가 얼마나 떨어지나)"]
    lines += [f"  {v:6.1%}  {'█' * max(1, round(v * 30))}  {name}" for name, v in result.importance.items()]
    lines += ["", "정제"] + ["  · " + line for line in c.report] + ["", "이상 탐지"] + ["  · " + line for line in o.report]
    if table:
        lines += ["", "모델 비교 (교차 검증 R²)"]
        lines += [f"  {i}. {row['name']:<14} {row['r2']:.3f} ±{row['std']:.3f}   {row['seconds']:.1f}초" for i, row in enumerate(table, 1)]
    return lines


def setup_font():
    """그래프에 한글이 나오게 이 PC 에 있는 한글 글꼴 하나를 고른다."""
    import matplotlib
    from matplotlib import font_manager
    installed = {f.name for f in font_manager.fontManager.ttflist}
    korean = [name for name in ("Malgun Gothic", "AppleGothic", "NanumGothic", "Noto Sans CJK KR") if name in installed]
    matplotlib.rcParams["font.family"] = korean[:1] or ["DejaVu Sans"]
    matplotlib.rcParams["axes.unicode_minus"] = False


def figure_learning(result, curve):
    """학습 과정 그래프: 왼쪽 = 학습 곡선(배운 행 수에 따른 점수), 오른쪽 = 단계별 오차 또는 빗나간 정도의 분포."""
    from matplotlib.figure import Figure
    setup_font()
    blue, red, gray = "#2563EB", "#DC2626", "#6B7280"
    fig = Figure(figsize=(10.5, 4.2), dpi=100, layout="constrained")
    left, right = fig.subplots(1, 2)
    sizes, valid, std = curve["sizes"], np.array(curve["valid"]), np.array(curve["valid_std"])
    left.plot(sizes, curve["train"], marker="o", color=gray, label="학습 데이터 점수")
    left.plot(sizes, valid, marker="o", color=red, label="시험 데이터 점수 (3조각 교차 검증)")
    left.fill_between(sizes, valid - std, valid + std, color=red, alpha=0.12)
    left.set_title(f"학습 곡선 — {result.name}")
    left.set_xlabel("배우는 데 쓴 행 수")
    left.set_ylabel("R²")
    left.set_ylim(min(0.0, float(valid.min()) - 0.05), 1.02)
    left.legend(loc="lower right", fontsize=8)
    if curve["steps"]:
        from matplotlib.ticker import FuncFormatter, NullFormatter
        right.plot(range(1, len(curve["steps"]) + 1), curve["steps"], color=blue)
        right.set_yscale("log")
        right.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.6g}"))
        right.yaxis.set_minor_formatter(NullFormatter())
        right.set_title(f"단계별 학습 오차 ({curve['steps_name']}가 늘수록 줄어듦)")
        right.set_xlabel(curve["steps_name"])
        right.set_ylabel("오차 (로그 눈금)")
    else:
        residual = result.y_test - result.pred_test
        right.hist(residual, bins=30, color=blue, alpha=0.8)
        right.axvline(0, color=gray, linewidth=1, linestyle="--")
        right.set_title("예측이 빗나간 정도의 분포 (시험 데이터)")
        right.set_xlabel("실제 - 예측")
        right.set_ylabel("행 수")
    for a in (left, right):
        a.grid(True, color="#E5E7EB", linewidth=0.6)
        a.set_axisbelow(True)
    return fig


def figure(c, o, result, table=None):
    """그래프 네 장을 한 그림에 그린다 (matplotlib Figure). 화면에 넣거나 fig.savefig("파일.png") 로 저장."""
    from matplotlib.figure import Figure
    setup_font()
    blue, red, gray = "#2563EB", "#DC2626", "#6B7280"
    fig = Figure(figsize=(10.5, 7.2), dpi=100, layout="constrained")
    ax = fig.subplots(2, 2)

    a = ax[0][0]                                             # ① 실제 값 vs 예측 값
    a.scatter(result.y_test, result.pred_test, s=14, color=blue, alpha=0.6)
    low, high = min(result.y_test.min(), result.pred_test.min()), max(result.y_test.max(), result.pred_test.max())
    a.plot([low, high], [low, high], color=gray, linewidth=1, linestyle="--")
    a.set_title(f"① 실제 값과 예측 값 (시험 데이터, R² {result.metrics['r2_test']:.3f})")
    a.set_xlabel(f"실제 {c.target}")
    a.set_ylabel("예측")

    a = ax[0][1]                                             # ② 잔차
    residual = result.y_test - result.pred_test
    a.scatter(result.pred_test, residual, s=14, color=blue, alpha=0.6)
    a.axhline(0, color=gray, linewidth=1, linestyle="--")
    a.set_title("② 예측이 빗나간 정도 (실제 - 예측)")
    a.set_xlabel("예측")
    a.set_ylabel("차이")

    a = ax[1][0]                                             # ③ 열 중요도
    imp = result.importance[::-1]
    a.barh(list(imp.index), imp.to_numpy() * 100, color=blue)
    a.set_title("③ 열 중요도 (%)")

    a = ax[1][1]                                             # ④ 모델 비교 또는 이상 행
    if table:
        rows = table[::-1]
        colors = [red if row["key"] == result.key else (gray if row["family"] == "머신러닝" else "#7C3AED") for row in rows]
        a.barh([row["name"] for row in rows], [max(row["r2"], 0) for row in rows], color=colors)
        a.set_xlim(0, 1)
        a.set_title("④ 모델 비교 (교차 검증 R², 빨강 = 고른 모델)")
    else:
        first = c.features[0] if c.features[0] in c.X.columns else c.X.columns[0]
        normal = ~o.mask
        a.scatter(c.X[first][normal], c.y[normal], s=12, color=blue, alpha=0.5, label="보통")
        if o.mask.any():
            a.scatter(c.X[first][o.mask], c.y[o.mask], s=28, color=red, label=f"이상 {int(o.mask.sum())}행")
            a.legend()
        a.set_title(f"④ 이상 행 ({first} 와 {c.target})")
        a.set_xlabel(first)
    for row in ax:
        for a in row:
            a.grid(True, color="#E5E7EB", linewidth=0.6)
            a.set_axisbelow(True)
            a.ticklabel_format(style="plain", axis="x") if a is not ax[1][0] and not (table and a is ax[1][1]) else None
    return fig


def export(path, df, c, o, result, table=None, tests=None, ols=None):
    """결과를 엑셀로 낸다: 예측결과(행마다 예측값 · 차이 · 이상 여부) / 모델비교 / 요약.
    tests · ols (stat_tests.py 의 결과)를 주면 통계검정 · 회귀계수 시트도 만든다."""
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook()
    ws = wb.active
    ws.title = "예측결과"
    ws.append(list(df.columns) + ["예측값", "차이(실제-예측)", "이상 여부"])
    for rid in c.X.index:
        row = [None if pd.isna(v) else (v.to_pydatetime() if isinstance(v, pd.Timestamp) else v) for v in df.loc[rid]]
        pred = float(result.predictions[rid])
        ws.append(row + [round(pred, 2), round(float(c.y[rid]) - pred, 2), o.reasons[rid] if o.mask[rid] else ""])
    ws.freeze_panes = "A2"
    ws2 = wb.create_sheet("모델비교")
    ws2.append(["순위", "모델", "계열", "교차 검증 R²", "R² 흔들림(±)", "RMSE", "걸린 시간(초)"])
    for i, row in enumerate(table or [], 1):
        ws2.append([i, row["name"], row["family"], row["r2"], row["std"], row["rmse"], row["seconds"]])
    ws3 = wb.create_sheet("요약")
    for line in summary_lines(c, o, result, table):
        ws3.append([line])
    heads = [ws, ws2]
    if tests is not None:
        ws4 = wb.create_sheet("통계검정")
        ws4.append(["설명하는 열", "종류", "검정 방법", "통계량", "값", "자유도", "p 값", "효과 크기", "값", "행 수", "유의(p<0.05)", "풀이"])
        for t in tests:
            ws4.append([t.feature, t.kind, t.method, t.stat_name, round(t.stat, 4), t.df, t.p, t.effect_name,
                        round(t.effect, 4), t.n, "예" if t.significant else "아니오", t.sentence])
        heads.append(ws4)
    if ols is not None:
        ws5 = wb.create_sheet("회귀계수")
        ws5.append(["열", "계수", "표준오차", "t", "p 값", "95% 하한", "95% 상한", "유의(p<0.05)"])
        for r in ols.rows:
            ws5.append([r["name"], r["coef"], r["se"], r["t"], r["p"], r["low"], r["high"], "예" if r["significant"] else "아니오"])
        heads.append(ws5)
    for sheet in heads:
        for cell in sheet[1]:
            cell.font = Font(bold=True)
    wb.save(path)
    return path


# =====================================================================
# 7. 안내문 — 이 파일의 설명을 문서(마크다운)로 만든다:  python analysis.py > docs/ml_guide.md
# =====================================================================
def guide_markdown():
    out = ["# 회귀 분석 · 머신러닝 안내", "",
           "> 이 문서는 `analysis.py` 안의 설명에서 자동으로 만들어집니다. (`python analysis.py guide`)", "",
           "## 1. 회귀 분석이란", "",
           "표의 한 열(**예측할 열**)을 다른 열들(**설명하는 열**)로 맞혀 보는 일입니다. "
           "예: 지역 · 면적 · 건축연수로 감정가를 예측.", "",
           "| 이 프로그램에서 | 교과서 용어 | 뜻 |", "| --- | --- | --- |",
           "| 예측할 열 | 종속변수 · 목표변수 · y | 알고 싶은 결과. 숫자여야 합니다 |",
           "| 설명하는 열 | 독립변수 · 설명변수 · 특성(feature) · X | 결과에 영향을 줄 것 같은 원인 후보 |",
           "| 결정계수 R² | coefficient of determination | 예측이 실제 값의 오르내림을 얼마나 설명하나 (1 이면 완벽, 0 이면 평균만 찍는 수준) |",
           "| 잔차 | residual | 실제 값 − 예측 값 |", "",
           "**머신러닝(machine learning)** 은 아서 새뮤얼이 1959년 체커 프로그램을 설명하며 쓴 말로, "
           "규칙을 사람이 적어 주는 대신 데이터에서 규칙을 찾게 하는 방법입니다. "
           "**딥러닝(deep learning)** 은 그중 신경망의 층을 깊게 쌓는 갈래입니다.", "",
           "## 2. 다섯 단계", "",
           "| 단계 | 하는 일 | 자동으로 해 주는 것 |", "| --- | --- | --- |",
           "| ① 변수 지정 | 예측할 열 하나, 설명하는 열 여러 개를 고름 | 숫자 열과 번호표 열을 가려 추천 |",
           "| ② 정제 | 겹친 행 제거, 빈칸 채우기, 글자 열을 숫자로 | 중앙값 · 평균으로 채우기, 글자 열을 0/1 열로 펼치기 |",
           "| ③ 이상 탐지 | 입력 실수 같은 이상한 행을 찾아 학습에서 뺌 | 다섯 가지 방법, 이유를 글로 표시 |",
           "| ④ 모델 선택 | 9가지 중 하나를 고르거나 모두 비교 | 교차 검증으로 순위 매기기 |",
           "| ⑤ 결과 · 그래프 | 정확도, 열 중요도, 그래프 4장, 엑셀로 내보내기 | 전부 |", "",
           "[자동 분석] 은 ①~⑤ 를 한 번에 돌립니다: 변수 추천 → 중앙값으로 채움 → 종합 방식으로 이상 탐지 → "
           "9개 모델 비교 → 1위 모델로 학습.", "",
           "## 3. 모델 9가지", "",
           "| 모델 | 영어 이름 | 계열 | 언제 |", "| --- | --- | --- | --- |"]
    out += [f"| {m.name} | {m.english} | {m.family} | {m.when} |" for m in MODELS.values()]
    for m in MODELS.values():
        out += ["", f"### {m.name} ({m.english})", "", f"- **이름의 뜻**: {m.origin}", f"- **역사**: {m.history}",
                f"- **하는 일**: {m.what}", f"- **언제**: {m.when}"]
    out += ["", "## 4. 이상 탐지 5가지", "", "| 방법 | 설명 |", "| --- | --- |"]
    out += [f"| {name} | {text} |" for key, (name, text) in OUTLIER_METHODS.items() if key != "none"]
    out += ["", "이상하다고 표시된 행은 지워지지 않습니다. 학습에서만 빼고, 결과 표에 이유와 함께 남깁니다. "
            "정말 틀린 값인지는 사람이 확인해야 합니다.", "",
            "## 5. 이 프로그램에 없는 딥러닝 모델 (소개만)", "",
            "| 모델 | 이름의 뜻 | 역사 | 쓰는 곳 |", "| --- | --- | --- | --- |",
            "| CNN (합성곱 신경망) | convolution 은 라틴어 convolvere(함께 말다) — 작은 창을 밀어 가며 겹쳐 계산 | "
            "후쿠시마의 네오코그니트론(1980), 르쿤의 LeNet(1989~1998), AlexNet(2012) | 이미지 · 영상 |",
            "| RNN · LSTM (순환 신경망) | recurrent 는 '되돌아오는' — 앞 계산 결과가 다음 계산으로 되돌아 들어감. "
            "LSTM 은 Long Short-Term Memory(길게 가는 단기 기억) | 엘만(1990), 호흐라이터와 슈미트후버의 LSTM(1997) | 시간 순서가 있는 값 · 글 |",
            "| 트랜스포머 | transform 은 '바꾸다' — 입력 문장을 다른 표현으로 바꿈 | "
            "바스와니 등 「Attention Is All You Need」(2017). GPT · Claude 같은 대화형 AI 의 바탕 | 글 · 번역 · 대화 |", "",
            "이 모델들은 PyTorch · TensorFlow 같은 도구와 많은 데이터 · 계산이 필요합니다. "
            "엑셀 표 수천 행은 3장의 모델 9가지로 충분한 경우가 대부분입니다.", "",
            "## 6. 알아 둘 것", "",
            "- **상관은 원인이 아닙니다.** 중요도가 높은 열이 결과를 '일으킨다'는 뜻은 아닙니다.",
            "- **배운 범위 밖은 믿지 마세요.** 39~135㎡ 로 배운 모델에 300㎡ 를 넣으면 엉뚱한 값이 나옵니다.",
            "- **시험 점수를 보세요.** 학습 데이터 점수는 외워서도 높아집니다. 이 프로그램은 20% 를 따로 떼어 시험합니다.",
            "- **행이 적으면 흔들립니다.** 100행 아래에서는 조각을 어떻게 나누느냐에 따라 순위가 바뀝니다.",
            "- 결과는 같은 데이터 · 같은 설정이면 언제나 같습니다(난수 씨앗 고정)."]
    return "\n".join(out) + "\n"


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdout.write(guide_markdown())
