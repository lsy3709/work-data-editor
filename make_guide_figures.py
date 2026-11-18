# -*- coding: utf-8 -*-
"""개념을 설명하는 그림 12장을 만든다 → docs/figures/*.png
실행: python make_guide_figures.py        (난수 씨앗이 고정이라 언제 만들어도 같은 그림)

그림 속 숫자는 모두 설명하려고 지어낸 값이다(실제 조사 자료가 아니다).
"""
from pathlib import Path

import numpy as np
from matplotlib.figure import Figure
from matplotlib.patches import Circle, FancyArrowPatch, FancyBboxPatch, Rectangle
from scipy import stats as sp

import analysis as an

OUT = Path(__file__).resolve().parent / "docs" / "figures"
BLUE, RED, GRAY, INK, PALE, GREEN = "#2563EB", "#DC2626", "#6B7280", "#111827", "#DBEAFE", "#059669"
NOTE = "설명용으로 지어낸 값"


def new(rows=1, cols=1, size=(9.6, 5.4)):
    an.setup_font()
    fig = Figure(figsize=size, dpi=110, layout="constrained")
    ax = fig.subplots(rows, cols)
    return fig, ax


def clean(ax):
    ax.grid(True, color="#E5E7EB", linewidth=0.6)
    ax.set_axisbelow(True)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)


def blank(ax, w=10, h=6):
    ax.set_xlim(0, w)
    ax.set_ylim(0, h)
    ax.axis("off")


def box(ax, x, y, w, h, text, face=PALE, edge=BLUE, size=10, weight="normal", color=INK):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.12", facecolor=face, edgecolor=edge, linewidth=1.4))
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=size, fontweight=weight, color=color, linespacing=1.5)


def arrow(ax, a, b, color=GRAY):
    ax.add_patch(FancyArrowPatch(a, b, arrowstyle="-|>", mutation_scale=14, color=color, linewidth=1.4))


def save(fig, name):
    fig.savefig(OUT / name)
    print("만듦:", name)


def f01_flow():
    fig, ax = new()
    blank(ax)
    steps = [("① 변수 지정", "무엇을 예측하고\n무엇으로 설명할지"), ("② 데이터 정제", "글자 숫자 · 날짜\n겹친 행 · 빈칸"),
             ("③ 이상 탐지", "입력 실수 같은\n이상한 행 찾기"), ("④ 모델 선택", "9가지를 견주어\n가장 맞는 것"), ("⑤ 결과", "점수 · 검정\n그래프 · 엑셀")]
    for i, (title, body) in enumerate(steps):
        x = 0.25 + i * 1.95
        box(ax, x, 3.3, 1.7, 0.75, title, face=BLUE, edge=BLUE, size=11, weight="bold", color="white")
        box(ax, x, 2.0, 1.7, 1.15, body, size=9.5)
        if i < 4:
            arrow(ax, (x + 1.72, 3.67), (x + 1.93, 3.67))
    ax.text(5, 5.2, "데이터 분석의 다섯 단계", ha="center", fontsize=17, fontweight="bold", color=INK)
    ax.text(5, 4.55, "사람이 하나씩 고를 수도 있고, [자동 분석] 으로 한 번에 돌릴 수도 있다", ha="center", fontsize=10.5, color=GRAY)
    box(ax, 0.25, 0.55, 9.5, 0.95, "사람이 정할 것:  무엇을 예측할지  ·  표시된 행이 정말 틀린 값인지  ·  결과를 업무에 써도 되는지",
        face="#FEF3C7", edge="#D97706", size=10.5)
    save(fig, "f01_flow.png")


def f02_regression():
    rng = np.random.default_rng(7)
    x = rng.uniform(40, 130, 26)
    y = 600 * x + 4000 + rng.normal(0, 9000, 26)
    slope, inter = np.polyfit(x, y, 1)
    fig, ax = new()
    for xi, yi in zip(x, y):
        ax.plot([xi, xi], [yi, slope * xi + inter], color=RED, linewidth=1, alpha=0.7)
    ax.scatter(x, y, s=38, color=BLUE, zorder=3, label="실제 값 (집 한 채)")
    xs = np.array([35, 135])
    ax.plot(xs, slope * xs + inter, color=INK, linewidth=2, label="회귀선 (예측)")
    ax.plot([], [], color=RED, label="잔차 = 실제 - 예측")
    ax.set_title("회귀 분석 — 점들 사이를 가장 잘 지나가는 선 긋기", fontsize=14, fontweight="bold")
    ax.set_xlabel("전용면적 (㎡)  ← 설명하는 열 (독립변수)")
    ax.set_ylabel("감정가 (만원)  ← 예측할 열 (종속변수)")
    ax.annotate("빨간 선(잔차)의 제곱을 모두 더한 값이\n가장 작아지는 선을 고른다 = 최소제곱법", xy=(x[3], y[3]), xytext=(45, 78000),
                fontsize=10, color=INK, arrowprops={"arrowstyle": "->", "color": GRAY})
    ax.text(0.99, 0.02, NOTE, transform=ax.transAxes, ha="right", fontsize=8, color=GRAY)
    ax.legend(loc="lower right", bbox_to_anchor=(1, 0.06), fontsize=9)
    clean(ax)
    save(fig, "f02_regression.png")


def f03_galton():
    rng = np.random.default_rng(11)
    parent = rng.normal(170, 6, 220)
    child = 170 + 0.6 * (parent - 170) + rng.normal(0, 5, 220)
    slope, inter = np.polyfit(parent, child, 1)
    fig, ax = new()
    ax.scatter(parent, child, s=16, color=BLUE, alpha=0.55)
    xs = np.array([152, 188])
    ax.plot(xs, xs, color=GRAY, linestyle="--", linewidth=1.4, label="부모와 똑같이 큰다면 (기울기 1)")
    ax.plot(xs, slope * xs + inter, color=RED, linewidth=2.4, label=f"실제로 그어지는 선 (기울기 {slope:.2f})")
    ax.axhline(170, color="#9CA3AF", linewidth=0.8)
    ax.axvline(170, color="#9CA3AF", linewidth=0.8)
    ax.annotate("키가 아주 큰 부모의 자녀는\n크긴 하지만 부모만큼은 아니다\n→ 평균 쪽으로 '되돌아간다' (regress)", xy=(184, slope * 184 + inter),
                xytext=(172.5, 155.5), fontsize=10, arrowprops={"arrowstyle": "->", "color": GRAY})
    ax.set_title("'회귀'라는 이름의 유래 — 평균으로 되돌아가는 키 (골턴, 1886)", fontsize=14, fontweight="bold")
    ax.set_xlabel("부모의 키 (cm)")
    ax.set_ylabel("자녀의 키 (cm)")
    ax.text(0.99, 0.02, NOTE + " (골턴의 실제 자료가 아니라 현상을 흉내 낸 것)", transform=ax.transAxes, ha="right", fontsize=8, color=GRAY)
    ax.legend(loc="upper left", fontsize=9)
    clean(ax)
    save(fig, "f03_galton.png")


def f04_cleaning():
    fig, ax = new()
    blank(ax)
    ax.text(5, 5.55, "데이터 정제 — 계산할 수 있는 모양으로 바로잡기", ha="center", fontsize=15, fontweight="bold")
    head = ["가입번호", "성명", "가입일", "월지급액"]
    before = [["RM-0001", " 김하나", "2026/07/15", "3150000"], ["RM-0002", "이두리 ", "2026.01.03", "2,430,000원"],
              ["RM-0003", "박세나", "20260613", ""], ["RM-0002", "이두리 ", "2026.01.03", "2,430,000원"]]
    after = [["RM-0001", "김하나", "2026-07-15", "3,150,000"], ["RM-0002", "이두리", "2026-01-03", "2,430,000"],
             ["RM-0003", "박세나", "2026-06-13", "(빈칸)"]]
    marks = {(0, 1), (1, 1), (1, 2), (1, 3), (2, 2), (2, 3), (3, 0), (3, 1), (3, 2), (3, 3)}

    def table(x0, title, rows, flagged, color):
        ax.text(x0 + 2.1, 4.85, title, ha="center", fontsize=12, fontweight="bold", color=color)
        widths = [0.95, 0.85, 1.15, 1.25]
        for r, row in enumerate([head] + rows):
            x = x0
            for c, (cell, w) in enumerate(zip(row, widths)):
                face = "#E5E7EB" if r == 0 else ("#FEE2E2" if (r - 1, c) in flagged else "white")
                ax.add_patch(Rectangle((x, 4.1 - r * 0.62), w, 0.6, facecolor=face, edgecolor="#9CA3AF", linewidth=0.8))
                ax.text(x + w / 2, 4.4 - r * 0.62, cell, ha="center", va="center", fontsize=8.5, fontweight="bold" if r == 0 else "normal")
                x += w

    table(0.2, "정제 전 (사람이 손으로 넣은 표)", before, marks, RED)
    table(5.6, "정제 후", after, set(), GREEN)
    arrow(ax, (4.55, 3.2), (5.5, 3.2), color=INK)
    notes = ["앞뒤 빈칸 지우기", "섞인 날짜 표기 통일", "글자로 적힌 금액 → 숫자", "똑같은 행은 하나만", "빈칸은 채우거나 그 행을 뺌"]
    for i, text in enumerate(notes):
        ax.text(0.3 + (i % 3) * 3.3, 0.85 - (i // 3) * 0.5, "■ " + text, fontsize=10.5, color=INK)
    save(fig, "f04_cleaning.png")


def f05_outlier():
    rng = np.random.default_rng(3)
    values = np.concatenate([rng.normal(50, 9, 60), [112, 118]])
    q1, med, q3 = np.percentile(values, [25, 50, 75])
    iqr = q3 - q1
    fig, ax = new()
    ax.boxplot(values, vert=False, widths=0.45, patch_artist=True, boxprops={"facecolor": PALE, "edgecolor": BLUE},
               medianprops={"color": RED, "linewidth": 2.2}, flierprops={"marker": "o", "markerfacecolor": RED, "markeredgecolor": RED, "markersize": 8})
    ax.scatter(values, np.full_like(values, 0.55) + rng.uniform(-0.05, 0.05, len(values)), s=12, color=GRAY, alpha=0.6)
    ax.set_ylim(0.35, 1.75)
    ax.set_yticks([])
    for x, text, tx, ty in ((q1, "Q1 (아래에서 25%)", q1 - 10, 1.42), (med, "중앙값", med, 1.6), (q3, "Q3 (아래에서 75%)", q3 + 10, 1.42)):
        ax.annotate(text, xy=(x, 1.23), xytext=(tx, ty), ha="center", fontsize=9.5, arrowprops={"arrowstyle": "-", "color": GRAY})
    ax.axvline(q1 - 1.5 * iqr, color=RED, linestyle="--", linewidth=1)
    ax.annotate("", xy=(q1, 0.72), xytext=(q3, 0.72), arrowprops={"arrowstyle": "<->", "color": INK})
    ax.text((q1 + q3) / 2, 0.64, "IQR = 가운데 절반의 폭", ha="center", fontsize=9.5)
    ax.annotate("이상값 후보\n(울타리 밖)", xy=(115, 1.0), xytext=(104, 1.5), ha="center", fontsize=10.5, color=RED,
                arrowprops={"arrowstyle": "->", "color": RED})
    ax.axvline(q3 + 1.5 * iqr, color=RED, linestyle="--", linewidth=1)
    ax.text(q3 + 1.5 * iqr + 0.8, 0.42, "울타리 = Q3 + 1.5 × IQR", fontsize=9, color=RED)
    ax.set_title("이상 탐지 — 상자그림으로 '남다른 값' 찾기 (터키, 1977)", fontsize=14, fontweight="bold")
    ax.set_xlabel("값  (회색 점 = 실제 값 하나하나, " + NOTE + ")")
    clean(ax)
    save(fig, "f05_outlier.png")


def f06_ttest():
    rng = np.random.default_rng(5)
    fig, axes = new(1, 2, size=(10.4, 5.0))
    c, d = rng.normal(50, 8, 40), rng.normal(50, 8, 40)
    d = d - d.mean() + c.mean() + 1.5                     # 평균 차이가 정확히 1.5 가 되게 맞춘다
    cases = [("차이가 뚜렷하다 → 유의함", rng.normal(50, 8, 40), rng.normal(62, 8, 40)),
             ("겹쳐 있다 → 유의하지 않음", c, d)]
    for ax, (title, a, b) in zip(axes, cases):
        res = sp.ttest_ind(a, b, equal_var=False)
        assert (res.pvalue < 0.05) == ("유의함" in title), (title, res.pvalue)      # 제목과 실제 계산이 어긋나면 멈춘다
        bins = np.linspace(25, 90, 22)
        ax.hist(a, bins=bins, color=BLUE, alpha=0.55, label=f"집단 가 (평균 {a.mean():.1f})")
        ax.hist(b, bins=bins, color=RED, alpha=0.55, label=f"집단 나 (평균 {b.mean():.1f})")
        ax.axvline(a.mean(), color=BLUE, linewidth=2)
        ax.axvline(b.mean(), color=RED, linewidth=2)
        p = "p < 0.001" if res.pvalue < 0.001 else f"p = {res.pvalue:.3f}"
        ax.set_title(f"{title}\nt = {res.statistic:.2f},  {p}", fontsize=11.5, fontweight="bold")
        ax.set_xlabel("값")
        ax.set_ylabel("사람 수")
        ax.legend(fontsize=8.5, loc="upper right")
        clean(ax)
    fig.suptitle("t 검정 — 두 집단의 평균 차이가 '우연이라 보기 어려울 만큼' 큰가", fontsize=14, fontweight="bold")
    fig.text(0.99, 0.01, NOTE, ha="right", fontsize=8, color=GRAY)
    save(fig, "f06_ttest.png")


def f07_pvalue():
    df, t_obs = 30, 2.4
    xs = np.linspace(-4.5, 4.5, 400)
    ys = sp.t.pdf(xs, df)
    p = 2 * sp.t.sf(t_obs, df)
    crit = sp.t.ppf(0.975, df)
    fig, ax = new()
    ax.plot(xs, ys, color=INK, linewidth=2)
    for side in (xs >= t_obs, xs <= -t_obs):
        ax.fill_between(xs, ys, where=side, color=RED, alpha=0.6)
    for c in (-crit, crit):
        ax.axvline(c, color=GRAY, linestyle="--", linewidth=1)
    ax.text(crit + 0.05, 0.36, f"기준선 ±{crit:.2f}\n(양쪽 꼬리 합 5%)", fontsize=9, color=GRAY)
    ax.annotate(f"내 데이터의 t = {t_obs}", xy=(t_obs, sp.t.pdf(t_obs, df)), xytext=(2.9, 0.16), fontsize=11, color=RED,
                arrowprops={"arrowstyle": "->", "color": RED})
    ax.annotate(f"빨간 넓이 = p 값 = {p:.3f}\n'관계가 없어도 우연히\n이만큼 벗어날 확률'", xy=(-2.9, 0.012), xytext=(-4.3, 0.17), fontsize=10.5,
                arrowprops={"arrowstyle": "->", "color": GRAY})
    ax.text(0, 0.2, "관계가 없을 때\nt 값이 나오는 모양\n(t 분포)", ha="center", fontsize=10.5, color=INK)
    ax.set_title("p 값 — 우연히 이만큼 벗어날 확률 (작을수록 '우연이 아니다')", fontsize=14, fontweight="bold")
    ax.set_xlabel("t 값  (0 = 차이 없음, 멀어질수록 차이가 큼)")
    ax.set_yticks([])
    clean(ax)
    save(fig, "f07_pvalue.png")


def f08_trees():
    fig, axes = new(1, 3, size=(11.2, 5.0))
    a = axes[0]
    blank(a, 6, 6)
    a.set_title("의사결정나무\n질문으로 갈라 무리의 평균으로", fontsize=11.5, fontweight="bold")
    box(a, 1.6, 4.4, 2.8, 0.9, "면적 ≥ 84㎡ ?", face=BLUE, edge=BLUE, color="white", weight="bold")
    box(a, 0.2, 2.5, 2.5, 0.9, "지역 = 서울 ?", size=9.5)
    box(a, 3.3, 2.5, 2.5, 0.9, "지역 = 서울 ?", size=9.5)
    for x, text in ((0.05, "2.9억"), (1.5, "4.6억"), (3.15, "4.9억"), (4.6, "8.1억")):
        box(a, x, 0.7, 1.3, 0.8, text, face="#D1FAE5", edge=GREEN, size=9.5)
    for s, e in (((3, 4.4), (1.45, 3.4)), ((3, 4.4), (4.55, 3.4)), ((1.45, 2.5), (0.7, 1.5)), ((1.45, 2.5), (2.15, 1.5)), ((4.55, 2.5), (3.8, 1.5)), ((4.55, 2.5), (5.25, 1.5))):
        arrow(a, s, e)
    a.text(1.7, 3.85, "아니오", fontsize=8, color=GRAY)
    a.text(4.0, 3.85, "예", fontsize=8, color=GRAY)

    a = axes[1]
    blank(a, 6, 6)
    a.set_title("랜덤 포레스트\n조금씩 다른 나무 수백 그루의 평균", fontsize=11.5, fontweight="bold")
    for i, (x, v) in enumerate(((0.3, "4.4억"), (2.2, "4.9억"), (4.1, "4.6억"))):
        a.add_patch(Circle((x + 0.8, 4.6), 0.55, facecolor="#D1FAE5", edgecolor=GREEN, linewidth=1.4))
        a.plot([x + 0.8, x + 0.8], [3.5, 4.05], color="#92400E", linewidth=4)
        a.text(x + 0.8, 4.6, f"나무 {i + 1}", ha="center", va="center", fontsize=8.5)
        a.text(x + 0.8, 3.05, v, ha="center", fontsize=10)
        arrow(a, (x + 0.8, 2.85), (3.0, 1.75))
    box(a, 1.5, 0.8, 3.0, 0.9, "평균 = 4.63억", face=BLUE, edge=BLUE, color="white", weight="bold")
    a.text(3, 0.25, "한 그루의 실수가 서로 지워진다", ha="center", fontsize=9, color=GRAY)

    a = axes[2]
    blank(a, 6, 6)
    a.set_title("그레이디언트 부스팅\n앞 나무가 틀린 만큼을 다음 나무가", fontsize=11.5, fontweight="bold")
    for i, (y, text, err) in enumerate(((4.3, "나무 1 : 4.0억", "아직 0.7억 모자람"), (2.9, "나무 2 : + 0.5억", "아직 0.2억 모자람"), (1.5, "나무 3 : + 0.15억", "거의 맞음"))):
        box(a, 0.3, y, 2.7, 0.85, text, size=9.5)
        a.text(3.2, y + 0.42, "→ " + err, fontsize=9, color=RED if i < 2 else GREEN, va="center")
        if i < 2:
            arrow(a, (1.65, y), (1.65, y - 0.52))
    a.text(3, 0.6, "실제 4.7억  ←  4.0 + 0.5 + 0.15 = 4.65억", ha="center", fontsize=10, fontweight="bold")
    fig.text(0.99, 0.01, NOTE, ha="right", fontsize=8, color=GRAY)
    save(fig, "f08_trees.png")


def f09_neural():
    fig, ax = new()
    blank(ax)
    ax.text(5, 5.6, "신경망 — 뉴런 층을 겹쳐 쌓는다 (층이 깊어지면 '딥러닝')", ha="center", fontsize=14.5, fontweight="bold")
    layers = [(1.3, ["면적", "건축연수", "층", "지역"], PALE, BLUE), (4.0, [""] * 5, "#FDE68A", "#D97706"), (6.4, [""] * 4, "#FDE68A", "#D97706"), (8.8, ["감정가"], "#D1FAE5", GREEN)]
    spots = []
    for x, names, face, edge in layers:
        ys = np.linspace(4.6, 1.0, len(names)) if len(names) > 1 else [2.8]
        spots.append([(x, y) for y in ys])
    for left, right in zip(spots, spots[1:]):
        for a in left:
            for b in right:
                ax.plot([a[0], b[0]], [a[1], b[1]], color="#9CA3AF", linewidth=0.6, zorder=1)
    for (x, names, face, edge), pts in zip(layers, spots):
        for (px, py), name in zip(pts, names):
            ax.add_patch(Circle((px, py), 0.3, facecolor=face, edgecolor=edge, linewidth=1.5, zorder=3))
            if name:
                ax.text(px + (-0.42 if x < 5 else 0.42), py, name, ha="right" if x < 5 else "left", va="center", fontsize=10)
    for x, text in ((1.3, "입력층\n설명하는 열"), (5.2, "은닉층 (숨은 층)\n가중치를 곱해 더하고, 굽은 함수를 통과"), (8.8, "출력층\n예측 값")):
        ax.text(x, 0.25, text, ha="center", fontsize=9.5, color=GRAY, linespacing=1.4)
    ax.text(5.2, 5.05, "선 하나하나가 '가중치' — 학습은 이 숫자들을 조금씩 고쳐 오차를 줄이는 일 (역전파, 1986)", ha="center", fontsize=9.5, color=INK)
    save(fig, "f09_neural.png")


def f10_overfit():
    rng = np.random.default_rng(2)
    x = np.sort(rng.uniform(0, 1, 16))
    y = np.sin(2 * np.pi * x) + rng.normal(0, 0.22, 16)
    xs = np.linspace(0, 1, 300)
    fig, axes = new(1, 3, size=(11.2, 4.6))
    for ax, degree, title, color in zip(axes, (1, 3, 13), ("너무 단순 (과소적합)\n흐름을 놓친다", "알맞음\n흐름만 따라간다", "너무 복잡 (과적합)\n점 하나하나를 외운다"), (GRAY, GREEN, RED)):
        coef = np.polyfit(x, y, degree)
        ax.scatter(x, y, s=32, color=BLUE, zorder=3)
        ax.plot(xs, np.polyval(coef, xs), color=color, linewidth=2.2)
        ax.plot(xs, np.sin(2 * np.pi * xs), color="#9CA3AF", linestyle="--", linewidth=1)
        ax.set_ylim(-1.9, 1.9)
        ax.set_title(title, fontsize=11.5, fontweight="bold", color=color if color != GRAY else INK)
        ax.set_xticks([])
        ax.set_yticks([])
        clean(ax)
    fig.suptitle("과적합 — 학습 데이터는 완벽히 맞히지만 새 데이터는 틀린다 (점선 = 진짜 흐름)", fontsize=13.5, fontweight="bold")
    fig.text(0.99, 0.01, NOTE, ha="right", fontsize=8, color=GRAY)
    save(fig, "f10_overfit.png")


def f11_timeline():
    events = [(1805, "최소제곱법\n르장드르"), (1886, "regression\n골턴"), (1908, "t 검정\n고셋(Student)"), (1925, "p < 0.05 관례\n분산분석 · 피셔"),
              (1943, "인공 뉴런\n매컬러 · 피츠"), (1958, "퍼셉트론\n로젠블랫"), (1959, "'머신러닝'\n새뮤얼"), (1977, "상자그림\n터키"),
              (1984, "의사결정나무\nCART"), (1986, "역전파\n힌턴 등"), (1996, "라쏘 · SVR"), (2001, "랜덤 포레스트\n부스팅"), (2012, "AlexNet\n딥러닝 시대"), (2017, "트랜스포머\n대화형 AI 의 바탕")]
    fig, ax = new(size=(11.6, 5.0))
    ax.set_xlim(-0.6, len(events) - 0.4)
    ax.set_ylim(-2.4, 2.6)
    ax.axis("off")
    ax.plot([-0.4, len(events) - 0.6], [0, 0], color=INK, linewidth=2)
    for i, (year, text) in enumerate(events):
        up = i % 2 == 0
        color = RED if year >= 1943 and "뉴런" in text or year in (1958, 1986, 2012, 2017) else BLUE
        ax.plot([i, i], [0, 0.55 if up else -0.55], color=color, linewidth=1.4)
        ax.add_patch(Circle((i, 0), 0.09, facecolor=color, edgecolor=color, zorder=3))
        ax.text(i, 0.72 if up else -0.72, f"{year}", ha="center", va="bottom" if up else "top", fontsize=11, fontweight="bold", color=color)
        ax.text(i, 1.12 if up else -1.12, text, ha="center", va="bottom" if up else "top", fontsize=8.6, linespacing=1.35)
    ax.text(len(events) / 2 - 0.5, 2.35, "통계에서 머신러닝 · 딥러닝까지 — 200년의 연표", ha="center", fontsize=14.5, fontweight="bold")
    ax.text(len(events) / 2 - 0.5, -2.3, "파랑 = 통계 · 머신러닝   빨강 = 신경망 · 딥러닝   (간격은 실제 햇수와 비례하지 않음)", ha="center", fontsize=9, color=GRAY)
    save(fig, "f11_timeline.png")


def f12_cv():
    fig, ax = new(size=(9.6, 4.8))
    blank(ax, 10, 6)
    ax.text(5, 5.5, "교차 검증 — 데이터를 5조각으로 나눠 돌아가며 시험 본다", ha="center", fontsize=14.5, fontweight="bold")
    scores = ["0.96", "0.97", "0.95", "0.97", "0.96"]
    for r in range(5):
        y = 4.3 - r * 0.72
        ax.text(0.35, y + 0.27, f"{r + 1}회", fontsize=10, va="center")
        for c in range(5):
            test = c == r
            ax.add_patch(Rectangle((1.0 + c * 1.35, y), 1.28, 0.55, facecolor=RED if test else PALE, edgecolor="white", linewidth=2))
            ax.text(1.64 + c * 1.35, y + 0.27, "시험" if test else "학습", ha="center", va="center", fontsize=9.5, color="white" if test else INK)
        ax.text(8.1, y + 0.27, f"→ 점수 {scores[r]}", fontsize=10, va="center")
    box(ax, 2.2, 0.15, 5.6, 0.75, "다섯 점수의 평균 0.962  ·  흔들림 ±0.008", face="#FEF3C7", edge="#D97706", size=11, weight="bold")
    ax.text(9.9, 0.05, NOTE, ha="right", fontsize=8, color=GRAY)
    save(fig, "f12_cv.png")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    for make in (f01_flow, f02_regression, f03_galton, f04_cleaning, f05_outlier, f06_ttest, f07_pvalue, f08_trees, f09_neural, f10_overfit, f11_timeline, f12_cv):
        make()


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    main()
