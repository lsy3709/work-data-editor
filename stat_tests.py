# -*- coding: utf-8 -*-
"""
통계 검정 — '이 열이 정말 예측할 열과 관계가 있는가'를 숫자로 따진다 (화면 없이도 쓸 수 있다)

필요한 것:  pip install pandas scipy matplotlib      (scipy 는 scikit-learn 을 설치하면 함께 들어온다)

무엇을 하나 (위에서 아래로 읽으면 됩니다)
  1. 기술통계        : describe           평균 · 표준편차 · 사분위수 …
  2. 열 하나씩 검정  : test_all            열의 종류에 맞는 검정을 알아서 고른다
                         숫자 열              → 피어슨 상관 t 검정        (correlate)
                         글자 열, 두 집단     → 독립표본 t 검정 (Welch)   (compare_groups)
                         글자 열, 세 집단 이상 → 일원분산분석 + 사후 t 검정
  3. 여러 열을 함께  : ols                 다중 회귀의 계수마다 t 검정 (다른 열의 영향을 빼고도 유의한가)
  4. 결과 내기       : summary_lines · figure_tests

읽는 법
  · 귀무가설: "관계가 없다(차이가 없다)".  p 값: 정말 관계가 없는데도 우연히 이만큼 차이가 날 확률.
  · p < 0.05 이면 '통계적으로 유의하다'고 한다(우연이라 보기 어렵다). 0.05 는 관례일 뿐 절대 기준이 아니다.
  · 유의하다 ≠ 중요하다. 행이 많으면 아주 작은 차이도 유의해진다 → 효과 크기(r · d · η²)를 같이 본다.
  · t 검정의 이름은 고셋(W. S. Gosset)이 1908년 'Student' 라는 필명으로 발표한 데서 왔다(기네스 맥주 회사의 품질 관리 문제).
"""
import math
from dataclasses import dataclass, field

import numpy as np
import pandas as pd
from scipy import stats as sp

ALPHA = 0.05


@dataclass
class Test:
    feature: str              # 검정한 설명 열
    kind: str                 # 숫자 / 범주
    method: str               # 쓴 검정 방법
    stat_name: str            # t / F
    stat: float
    df: str                   # 자유도
    p: float
    effect_name: str          # r / d / η²
    effect: float
    n: int
    significant: bool
    sentence: str             # 사람이 읽는 한 줄 풀이
    groups: list = field(default_factory=list)    # 집단별 n · 평균 · 표준편차
    pairs: list = field(default_factory=list)     # 사후 검정: 집단 둘씩 견준 결과


def p_text(p):
    return "p < 0.001" if p < 0.001 else f"p = {p:.3f}"


def p_short(p):
    """표의 칸에 넣는 p 값."""
    return "-" if p != p else "< 0.001" if p < 0.001 else f"{p:.3f}"


def num(v):
    """숫자를 읽기 쉽게: 큰 값은 쉼표, 작은 값은 유효 숫자 3자리."""
    return f"{v:,.0f}" if abs(v) >= 100 else f"{v:,.3g}"


def effect_text(name, value):
    """효과 크기를 말로 (코헨의 관례)."""
    size = abs(value)
    if name == "r":
        word = "강함" if size >= 0.5 else "중간" if size >= 0.3 else "약함" if size >= 0.1 else "거의 없음"
    elif name == "d":
        word = "큼" if size >= 0.8 else "중간" if size >= 0.5 else "작음" if size >= 0.2 else "거의 없음"
    else:
        word = "큼" if size >= 0.14 else "중간" if size >= 0.06 else "작음" if size >= 0.01 else "거의 없음"
    return f"{name} = {value:.3f} ({word})"


def _usable(c, exclude):
    if exclude is None:
        return pd.Series(True, index=c.X.index)
    return ~exclude.reindex(c.X.index, fill_value=False)


# =====================================================================
# 1. 기술통계
# =====================================================================
def describe(c, exclude=None):
    """예측할 열과 숫자 설명 열의 요약값."""
    keep = _usable(c, exclude)
    columns = {c.target: c.y}
    columns.update({name: c.X[name] for name in c.features if name in c.X.columns})
    rows = []
    for name, col in columns.items():
        v = col[keep]
        rows.append({"name": name, "n": int(v.count()), "mean": float(v.mean()), "std": float(v.std()),
                     "min": float(v.min()), "q1": float(v.quantile(0.25)), "median": float(v.median()),
                     "q3": float(v.quantile(0.75)), "max": float(v.max()), "skew": float(v.skew())})
    return rows


# =====================================================================
# 2. 열 하나씩 검정
# =====================================================================
def correlate(y, x, name, alpha=ALPHA):
    """숫자 열: 피어슨 상관계수 r 이 0 과 다른지 t 검정한다.  t = r·√(n−2) / √(1−r²), 자유도 n−2
    (설명 열 하나로 단순 회귀를 했을 때 기울기의 t 검정과 똑같은 값이다.)"""
    data = pd.concat([y, x], axis=1).dropna()
    yy, xx, n = data.iloc[:, 0], data.iloc[:, 1], len(data)
    if n < 3 or xx.nunique() < 2 or yy.nunique() < 2:
        return Test(name, "숫자", "검정할 수 없음 (값이 한 가지뿐)", "t", 0.0, "-", 1.0, "r", 0.0, n, False,
                    f"[{name}] 은 값이 한 가지뿐이라 검정할 수 없습니다.")
    r, p = sp.pearsonr(xx, yy)
    r, p = float(r), float(p)
    t = r * math.sqrt((n - 2) / max(1 - r * r, 1e-300))
    ok = p < alpha
    if ok:
        text = (f"[{name}] 이(가) 클수록 [{y.name}] 이(가) {'커지는' if r > 0 else '작아지는'} 관계가 "
                f"통계적으로 유의합니다 ({effect_text('r', r)}, {p_text(p)}).")
    else:
        text = (f"[{name}] 과(와) [{y.name}] 사이의 직선 관계는 통계적으로 유의하지 않습니다 "
                f"({effect_text('r', r)}, {p_text(p)}) — 우연으로도 이 정도는 나옵니다.")
    return Test(name, "숫자", "피어슨 상관 t 검정", "t", t, str(n - 2), p, "r", r, n, ok, text)


def compare_groups(y, groups, name, alpha=ALPHA, max_pairs=8):
    """글자 열: 집단마다 예측할 열의 평균이 다른지 검정한다.
    두 집단이면 독립표본 t 검정(Welch: 두 집단의 퍼짐이 달라도 되는 방식),
    세 집단 이상이면 일원분산분석(F) 을 하고, 유의하면 어느 집단끼리 다른지 둘씩 t 검정(본페로니 보정)."""
    data = pd.DataFrame({"y": y, "g": groups}).dropna()
    data["g"] = data["g"].astype(str)
    sizes = data["g"].value_counts()
    data = data[data["g"].isin(sizes.index[sizes >= 2])]              # 한 명뿐인 집단은 퍼짐을 알 수 없어 뺀다
    names = sorted(data["g"].unique())
    parts = [data["y"][data["g"] == g] for g in names]
    info = [{"name": g, "n": len(v), "mean": float(v.mean()), "std": float(v.std())} for g, v in zip(names, parts)]
    n = len(data)
    target = y.name
    if len(names) < 2:
        return Test(name, "범주", "검정할 수 없음 (집단이 하나)", "t", 0.0, "-", 1.0, "d", 0.0, n, False,
                    f"[{name}] 은 견줄 집단이 하나뿐이라 검정할 수 없습니다.", info)

    if len(names) == 2:
        a, b = parts
        res = sp.ttest_ind(a, b, equal_var=False)
        t, p = float(res.statistic), float(res.pvalue)
        va, vb = a.var() / len(a), b.var() / len(b)
        df = (va + vb) ** 2 / (va ** 2 / (len(a) - 1) + vb ** 2 / (len(b) - 1))
        pooled = math.sqrt(((len(a) - 1) * a.var() + (len(b) - 1) * b.var()) / (len(a) + len(b) - 2))
        d = float((a.mean() - b.mean()) / pooled) if pooled else 0.0
        ok = p < alpha
        gap = f"'{names[0]}' 평균 {num(a.mean())} · '{names[1]}' 평균 {num(b.mean())}"
        if ok:
            text = f"[{name}] 의 두 집단은 [{target}] 평균이 통계적으로 유의하게 다릅니다 — {gap} ({effect_text('d', d)}, {p_text(p)})."
        else:
            text = f"[{name}] 의 두 집단은 [{target}] 평균 차이가 통계적으로 유의하지 않습니다 — {gap} ({effect_text('d', d)}, {p_text(p)})."
        return Test(name, "범주", "독립표본 t 검정 (Welch)", "t", t, f"{df:.1f}", p, "d", d, n, ok, text, info)

    res = sp.f_oneway(*parts)
    f, p = float(res.statistic), float(res.pvalue)
    grand = data["y"].mean()
    between = sum(len(v) * (v.mean() - grand) ** 2 for v in parts)
    total = float(((data["y"] - grand) ** 2).sum())
    eta = float(between / total) if total else 0.0
    ok = p < alpha
    pairs = []
    if len(names) <= max_pairs:
        count = len(names) * (len(names) - 1) // 2
        for i in range(len(names)):
            for j in range(i + 1, len(names)):
                r2 = sp.ttest_ind(parts[i], parts[j], equal_var=False)
                raw = float(r2.pvalue)
                pairs.append({"a": names[i], "b": names[j], "t": float(r2.statistic), "p": raw,
                              "p_adj": min(1.0, raw * count), "diff": float(parts[i].mean() - parts[j].mean())})
    top, low = max(info, key=lambda g: g["mean"]), min(info, key=lambda g: g["mean"])
    gap = f"가장 높은 '{top['name']}' {num(top['mean'])} · 가장 낮은 '{low['name']}' {num(low['mean'])}"
    if ok:
        differ = sum(1 for q in pairs if q["p_adj"] < alpha)
        extra = f" 둘씩 견주면 {len(pairs)}쌍 중 {differ}쌍이 다릅니다." if pairs else ""
        text = (f"[{name}] 의 {len(names)}개 집단은 [{target}] 평균이 통계적으로 유의하게 다릅니다 — {gap} "
                f"({effect_text('η²', eta)}, {p_text(p)}).{extra}")
    else:
        text = (f"[{name}] 의 {len(names)}개 집단은 [{target}] 평균 차이가 통계적으로 유의하지 않습니다 — {gap} "
                f"({effect_text('η²', eta)}, {p_text(p)}).")
    return Test(name, "범주", "일원분산분석 (ANOVA)" + (" + 사후 t 검정" if pairs else ""), "F", f,
                f"{len(names) - 1}, {n - len(names)}", p, "η²", eta, n, ok, text, info, pairs)


def test_all(c, exclude=None, alpha=ALPHA):
    """정제한 표(Clean)의 설명 열을 하나씩, 종류에 맞는 방법으로 검정한다. exclude: 이상 탐지에서 뺄 행."""
    keep = _usable(c, exclude)
    y = c.y[keep].rename(c.target)
    out = []
    for name in c.features:
        if name in c.X.columns:                                       # 숫자 · 날짜(날 수로 바뀐 것)
            out.append(correlate(y, c.X[name][keep], name, alpha))
        else:
            out.append(compare_groups(y, c.frame[name][keep], name, alpha))
    return out


test_all.__test__ = False       # (테스트 도구가 이 함수를 테스트로 착각하지 않게)


# =====================================================================
# 3. 여러 열을 함께 — 다중 회귀의 계수 t 검정
# =====================================================================
@dataclass
class Ols:
    rows: list                # 계수마다: name · coef · se · t · p · low · high (95% 신뢰구간)
    r2: float
    adj_r2: float
    f: float
    f_p: float
    n: int
    df: int                   # 잔차 자유도
    note: str                 # 기준 집단 설명


def ols(c, exclude=None, alpha=ALPHA):
    """모든 설명 열을 한꺼번에 넣은 선형 회귀를 풀고, 계수마다 't = 계수 ÷ 표준오차' 로 검정한다.
    '다른 열이 같을 때 이 열이 한 단위 늘면 예측할 열이 얼마나 달라지는가, 그것이 0 과 다른가'를 본다.
    글자 열은 가장 흔한 값을 기준 집단으로 두고, 나머지 값이 기준보다 얼마나 높은지/낮은지로 나타낸다."""
    keep = _usable(c, exclude).to_numpy()
    X, y = c.X[keep], c.y[keep]
    names, columns, refs, skipped = ["(절편)"], [np.ones(len(X))], [], []
    for feature in c.features:
        cols = [col for col, origin in c.groups.items() if origin == feature]
        if cols != [feature]:                                         # 글자 열에서 펼친 0/1 열들
            base = max(cols, key=lambda col: X[col].sum())
            refs.append(f"{feature}: {base.split('=', 1)[1]}")
            cols = [col for col in cols if col != base]
        for col in cols:
            if X[col].nunique() <= 1:
                continue
            trial = columns + [X[col].to_numpy(dtype=float)]
            if np.linalg.matrix_rank(np.column_stack(trial)) < len(trial):
                skipped.append(col)                                   # 다른 열로 완전히 설명되는 열(겹치는 정보)은 뺀다
                continue
            names.append(col)
            columns = trial
    A = np.column_stack(columns)
    n, k = A.shape
    if n - k < 1:
        raise ValueError("행 수가 설명 열 수보다 적어 회귀 계수를 검정할 수 없습니다.")
    beta, _, rank, _ = np.linalg.lstsq(A, y.to_numpy(dtype=float), rcond=None)
    resid = y.to_numpy(dtype=float) - A @ beta
    dof = n - rank
    rss = float(resid @ resid)
    tss = float(((y - y.mean()) ** 2).sum())
    cov = (rss / dof) * np.linalg.pinv(A.T @ A)
    se = np.sqrt(np.clip(np.diag(cov), 0, None))
    crit = float(sp.t.ppf(1 - alpha / 2, dof))
    rows = []
    for name, b, s in zip(names, beta, se):
        t = float(b / s) if s > 0 else float("nan")
        p = float(2 * sp.t.sf(abs(t), dof)) if s > 0 else float("nan")
        rows.append({"name": name, "coef": float(b), "se": float(s), "t": t, "p": p,
                     "low": float(b - crit * s), "high": float(b + crit * s), "significant": bool(p < alpha)})
    r2 = 1 - rss / tss if tss else 0.0
    f = ((tss - rss) / (rank - 1)) / (rss / dof) if rank > 1 and rss > 0 else float("nan")
    note = "글자 열의 기준 집단 — " + " · ".join(refs) if refs else ""
    if skipped:
        note += ("  " if note else "") + "다른 열과 정보가 겹쳐 뺀 열 — " + ", ".join(skipped)
    return Ols(rows, float(r2), float(1 - (1 - r2) * (n - 1) / dof), float(f),
               float(sp.f.sf(f, rank - 1, dof)) if rank > 1 else float("nan"), n, dof, note)


# =====================================================================
# 4. 결과 내기 — 글 · 그래프
# =====================================================================
def summary_lines(c, tests, fit=None, stats=None):
    lines = [f"■ 통계 검정 — 예측할 열 [{c.target}] 과(와) 설명하는 열의 관계 (유의수준 {ALPHA})", ""]
    yes = [t for t in tests if t.significant]
    lines.append(f"설명하는 열 {len(tests)}개 중 {len(yes)}개가 통계적으로 유의합니다"
                 + (": " + ", ".join(t.feature for t in yes) if yes else " — 이 열들로는 예측하기 어렵다는 뜻입니다."))
    if len(tests) >= 5:
        lines.append(f"(열을 {len(tests)}개 검정하면 관계가 전혀 없어도 우연히 유의하게 나오는 열이 평균 {len(tests) * ALPHA:.1f}개쯤 생깁니다.)")
    lines.append("")
    for t in tests:
        mark = "●" if t.significant else "○"
        lines.append(f"{mark} {t.feature}  —  {t.method}:  {t.stat_name}({t.df}) = {t.stat:,.3f},  {p_text(t.p)},  "
                     f"{effect_text(t.effect_name, t.effect)}")
        lines.append(f"    {t.sentence}")
        for g in t.groups[:12]:
            lines.append(f"      · {g['name']:<12} {g['n']:>5,}행   평균 {g['mean']:>12,.1f}   표준편차 {g['std']:>12,.1f}")
        shown = [q for q in t.pairs if q["p_adj"] < ALPHA][:10]
        for q in shown:
            lines.append(f"      ↔ {q['a']} · {q['b']}: 차이 {q['diff']:,.1f}, t = {q['t']:.2f}, 보정한 {p_text(q['p_adj'])}")
        lines.append("")
    if fit:
        lines += [f"■ 다중 회귀의 계수 t 검정 — {fit.n:,}행, R² {fit.r2:.3f} (조정 R² {fit.adj_r2:.3f}), "
                  f"모델 전체 F = {fit.f:,.1f}, {p_text(fit.f_p)}",
                  "  (다른 열이 같을 때, 이 열이 1 늘면 예측할 열이 '계수'만큼 달라진다. t = 계수 ÷ 표준오차)", ""]
        lines.append(f"  {'열':<22}{'계수':>14}{'표준오차':>12}{'t':>9}{'p':>11}   95% 신뢰구간")
        for r in fit.rows:
            lines.append(f"  {r['name']:<22}{r['coef']:>14,.2f}{r['se']:>12,.2f}{r['t']:>9.2f}{p_short(r['p']):>11}   "
                         f"{r['low']:,.1f} ~ {r['high']:,.1f}  {'●' if r['significant'] else ''}")
        if fit.note:
            lines += ["", "  " + fit.note]
        lines.append("")
    if stats:
        lines += ["■ 기술통계", f"  {'열':<18}{'n':>7}{'평균':>13}{'표준편차':>13}{'최소':>12}{'중앙값':>12}{'최대':>12}{'왜도':>8}"]
        lines += [f"  {s['name']:<18}{s['n']:>7,}{s['mean']:>13,.1f}{s['std']:>13,.1f}{s['min']:>12,.1f}"
                  f"{s['median']:>12,.1f}{s['max']:>12,.1f}{s['skew']:>8.2f}" for s in stats]
    return lines


def figure_tests(c, tests, fit=None, exclude=None):
    """검정 결과 그래프: p 값이 작은 순으로 열 4개(숫자 = 산점도와 추세선, 글자 = 상자그림),
    회귀 계수의 t 값, 모든 열의 p 값."""
    from matplotlib.figure import Figure

    import analysis as an
    an.setup_font()
    blue, red, gray = "#2563EB", "#DC2626", "#9CA3AF"
    keep = _usable(c, exclude).to_numpy()
    y = c.y[keep]
    fig = Figure(figsize=(12.4, 7.4), dpi=100, layout="constrained")
    ax = fig.subplots(2, 3).ravel()
    ranked = sorted(tests, key=lambda t: t.p)
    for a, t in zip(ax[:4], ranked[:4]):
        color = red if t.significant else gray
        if t.kind == "숫자":
            x = c.X[t.feature][keep]
            a.scatter(x, y, s=10, color=blue, alpha=0.45)
            if x.nunique() > 1:
                slope, inter = np.polyfit(x, y, 1)
                xs = np.array([x.min(), x.max()])
                a.plot(xs, slope * xs + inter, color=color, linewidth=2)
        else:
            groups = sorted(t.groups, key=lambda g: g["n"], reverse=True)[:8]
            order = sorted(g["name"] for g in groups)
            g_col = c.frame[t.feature][keep].astype(str)
            a.boxplot([y[(g_col == g).to_numpy()] for g in order], patch_artist=True, showmeans=True,
                      boxprops={"facecolor": "#DBEAFE", "edgecolor": blue}, medianprops={"color": color, "linewidth": 2},
                      meanprops={"marker": "D", "markerfacecolor": color, "markeredgecolor": color, "markersize": 5},
                      flierprops={"markersize": 3})
            a.set_xticks(range(1, len(order) + 1))
            a.set_xticklabels(order, rotation=20 if len(order) > 4 else 0, fontsize=8)
        a.set_title(f"{t.feature} — {t.stat_name} = {t.stat:,.2f}, {p_text(t.p)}", color=color if t.significant else "#374151", fontsize=10)
        a.set_ylabel(c.target, fontsize=8)
    for a in ax[len(ranked[:4]):4]:
        a.set_visible(False)

    a = ax[4]                                               # 회귀 계수의 t 값
    if fit:
        rows = sorted((r for r in fit.rows if r["name"] != "(절편)" and not math.isnan(r["t"])), key=lambda r: abs(r["t"]))[-12:]
        a.barh([r["name"] for r in rows], [r["t"] for r in rows], color=[red if r["significant"] else gray for r in rows])
        crit = float(sp.t.ppf(1 - ALPHA / 2, fit.df))
        for v in (-crit, crit):
            a.axvline(v, color="#111827", linewidth=1, linestyle="--")
        a.set_title(f"회귀 계수의 t 값 (점선 ±{crit:.2f} 밖 = 유의)", fontsize=10)
        a.tick_params(axis="y", labelsize=8)
    else:
        a.set_visible(False)

    a = ax[5]                                               # 모든 열의 p 값
    rows = ranked[::-1]
    a.barh([t.feature for t in rows], [-math.log10(max(t.p, 1e-300)) for t in rows],
           color=[red if t.significant else gray for t in rows])
    a.axvline(-math.log10(ALPHA), color="#111827", linewidth=1, linestyle="--")
    a.set_title("열마다의 p 값 (-log10, 점선 = 0.05)", fontsize=10)
    a.tick_params(axis="y", labelsize=8)
    for a in ax:
        a.grid(True, color="#E5E7EB", linewidth=0.6)
        a.set_axisbelow(True)
    return fig
