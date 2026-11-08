"""6주차(실험 5) 암페어 법칙 레포트: 계산 + 그래프 생성."""
import json, math
from pathlib import Path
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from scipy.special import ellipk, ellipe

ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "fig"
FIG.mkdir(exist_ok=True)

for f in font_manager.findSystemFonts():
    if "wqy-zenhei" in f.lower() or "nanum" in f.lower() or "malgun" in f.lower():
        font_manager.fontManager.addfont(f)
plt.rcParams["font.family"] = ["WenQuanYi Zen Hei"]
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["mathtext.fontset"] = "dejavusans"

MU0 = 4 * math.pi * 1e-7
C_MEAS, C_THEO = "#B5333A", "#2B5C8A"

def err(meas, theo):
    return abs(meas - theo), abs(meas - theo) / theo * 100

out = {}

# ---------------- 실험 1: 직선도선 ----------------
I1 = 5.0                       # A (슬라이드 이론값에서 역산)
OFFSET = 0.0065                # m, 도선 중심 ~ 센서 감지점 초기 간격
x_meas = [0.003547, 0.013466, 0.023718, 0.033416, 0.0433]
B_meas_G = [1.045, 0.5811, 0.3662, 0.3369, 0.2588]
rows = []
for x, bg in zip(x_meas, B_meas_G):
    r = x + OFFSET
    bt = MU0 * I1 / (2 * math.pi * r)
    bm = bg * 1e-4
    e, ep = err(bm, bt)
    rows.append(dict(x=x, r=r, meas=bm, theo=bt, err=e, errp=ep))
LW = 0.40                      # m, 직선도선 길이
for d in rows:
    d["Br"] = d["meas"] * 1e4 * d["r"] * 100           # G*cm (이론: 1.000)
    d["finite"] = (LW / 2) / math.sqrt((LW / 2) ** 2 + d["r"] ** 2)
slope = np.polyfit(np.log([d["r"] for d in rows]), np.log([d["meas"] for d in rows]), 1)[0]
out["wire"] = dict(I=I1, offset=OFFSET, L=LW, slope=slope, rows=rows)

fig, ax = plt.subplots(figsize=(6.2, 3.6), dpi=200)
rr = np.linspace(0.008, 0.055, 300)
ax.plot(rr * 100, MU0 * I1 / (2 * np.pi * rr) * 1e4, color=C_THEO, lw=1.6,
        label=r"이론값  $B=\mu_0 I/2\pi r$")
ax.plot([d["r"] * 100 for d in rows], [d["meas"] * 1e4 for d in rows], "s-",
        color=C_MEAS, ms=5, lw=1.2, label="측정값")
ax.set_xlabel("도선으로부터의 거리 r (cm)")
ax.set_ylabel("자기장 B (gauss)")
ax.set_title("직선도선 주위의 거리에 따른 자기장 (I = 5 A)", fontsize=11)
ax.grid(alpha=.3); ax.legend(frameon=False); ax.set_ylim(0, 1.4)
fig.tight_layout(); fig.savefig(FIG / "fig1_wire.png"); plt.close(fig)

# ---------------- 실험 2: 원형도선 ----------------
I2 = 3.0
loops = {}
meas2 = {0.04: (4.45e-5, 1.43e-5), 0.03: (5.76e-5, 1.94e-5)}
for R, (m0, mR) in meas2.items():
    t0 = MU0 * I2 / (2 * R)
    tR = MU0 * I2 / (4 * math.sqrt(2) * R)
    loops[f"{R*100:.0f}"] = dict(R=R, m0=m0, t0=t0, e0=err(m0, t0)[1],
                                 mR=mR, tR=tR, eR=err(mR, tR)[1],
                                 ratio_meas=mR / m0, ratio_theo=tR / t0)
out["loop"] = dict(I=I2, loops=loops)

fig, ax = plt.subplots(figsize=(6.2, 3.6), dpi=200)
xx = np.linspace(0, 0.08, 300)
for (R, (m0, mR)), c, mk in zip(meas2.items(), [C_THEO, "#3E8E5E"], ["o", "^"]):
    ax.plot(xx * 100, MU0 * I2 * R**2 / (2 * (R**2 + xx**2) ** 1.5) * 1e4, color=c,
            lw=1.6, label=f"이론값 (R = {R*100:.0f} cm)")
    ax.plot([0, R * 100], [m0 * 1e4, mR * 1e4], mk, color=C_MEAS, ms=7,
            mfc=C_MEAS if R == 0.04 else "white", mew=1.5,
            label=f"측정값 (R = {R*100:.0f} cm)")
ax.set_xlabel("원형도선 중심으로부터 중심축 위의 거리 x (cm)")
ax.set_ylabel("자기장 B (gauss)")
ax.set_title("원형도선 중심축 위의 자기장 (I = 3 A)", fontsize=11)
ax.grid(alpha=.3); ax.legend(frameon=False, fontsize=8.5)
fig.tight_layout(); fig.savefig(FIG / "fig2_loop.png"); plt.close(fig)

# ---------------- 실험 3: 솔레노이드 ----------------
I3, Rs, L, N = 0.67, 0.020, 0.110, 2900
n = N / L
km = L / math.sqrt(L**2 + 4 * Rs**2)
ke = 0.5 * L / math.sqrt(L**2 + Rs**2)
Binf = MU0 * n * I3
tm, te = Binf * km, Binf * ke
mm_, me_ = -0.0198, -0.0102
out["sol"] = dict(I=I3, R=Rs, L=L, N=N, n=n, km=km, ke=ke, Binf=Binf,
                  mm=mm_, me=me_, tm=tm, te=te,
                  em=err(abs(mm_), tm)[1], ee=err(abs(me_), te)[1],
                  ratio_meas=me_ / mm_, ratio_theo=te / tm)

fig, ax = plt.subplots(figsize=(6.2, 3.6), dpi=200)
xs = np.linspace(-0.11, 0.11, 400)
a, b = L / 2 + xs, L / 2 - xs
Bx = 0.5 * MU0 * n * I3 * (b / np.sqrt(b**2 + Rs**2) + a / np.sqrt(a**2 + Rs**2))
ax.axvspan(-L / 2 * 100, L / 2 * 100, color="#999", alpha=.12, lw=0, label="코일 구간")
ax.axhline(Binf * 1e3, color="#777", ls="--", lw=1, label=r"무한 솔레노이드 $\mu_0 nI$")
ax.plot(xs * 100, Bx * 1e3, color=C_THEO, lw=1.6, label="이론값 (유한 솔레노이드)")
ax.plot([0, L / 2 * 100], [abs(mm_) * 1e3, abs(me_) * 1e3], "s", color=C_MEAS, ms=7,
        label="측정값 (크기)")
ax.set_xlabel("코일 중심으로부터 중심축 위의 거리 x (cm)")
ax.set_ylabel("자기장 B (mT)")
ax.set_title("솔레노이드 중심축 위의 자기장 (I = 0.67 A)", fontsize=11)
ax.grid(alpha=.3); ax.legend(frameon=False, fontsize=8, loc="upper center", ncol=2)
ax.set_ylim(0, 31)
fig.tight_layout(); fig.savefig(FIG / "fig3_solenoid.png"); plt.close(fig)

# ---------------- 질문 2: 자기장선 ----------------
def loop_flux(rho, z, a):
    """반지름 a 원형 전류고리의 자속함수 psi = rho*A_phi (상수배 무시)."""
    rho = np.abs(rho) + 1e-12
    m = 4 * a * rho / ((a + rho) ** 2 + z**2)
    k = np.sqrt(m)
    return rho * np.sqrt(a / rho) / k * ((1 - m / 2) * ellipk(m) - ellipe(m))

def draw_lines(ax, psi, X, Y, levels, color):
    ax.contour(X, Y, psi, levels=levels, colors=color, linewidths=1.0)

fig, axs = plt.subplots(1, 3, figsize=(10.5, 3.7), dpi=200,
                        gridspec_kw=dict(width_ratios=[1, 1, 1.45]))
# (a) 직선도선: 도선 단면(⊙)을 위에서 본 모습
ax = axs[0]
radii = 0.25 * 1.32 ** np.arange(9)
for i, r in enumerate(radii):
    ax.add_patch(plt.Circle((0, 0), r, fill=False, color=C_THEO, lw=1))
    if i % 2 == 0:
        ax.annotate("", xy=(-0.01, r), xytext=(0.01, r),
                    arrowprops=dict(arrowstyle="-|>", color=C_THEO, lw=1))
ax.add_patch(plt.Circle((0, 0), 0.12, color=C_MEAS))
ax.plot(0, 0, "o", color="white", ms=2.5)
ax.set_xlim(-2.6, 2.6); ax.set_ylim(-2.6, 2.6)
ax.set_title("(a) 직선도선 (전류: 지면에서 나옴)", fontsize=9.5)

# (b) 원형도선 단면 / (c) 솔레노이드 단면 — 같은 자속 간격의 등고선
def coil_panel(ax, centers, a, xlim, ylim, title, nlev, frac):
    X, Y = np.meshgrid(np.linspace(*xlim, 600), np.linspace(*ylim, 450))
    psi = sum(loop_flux(Y, X - c, a) for c in centers)
    # 중심 단면(x=0)에서 rho < frac*a 를 지나는 자속을 같은 간격으로 나눔
    #  -> 선 간격이 촘촘할수록 자기장이 강함
    yy = np.linspace(1e-4, ylim[1], 4000)
    psi_mid = sum(loop_flux(yy, 0 - c, a) for c in centers)
    pref = sum(loop_flux(np.array([frac * a]), 0 - c, a) for c in centers)[0]
    levels = pref * np.arange(1, nlev + 1) / nlev
    for sgn in (1, -1):
        ax.contour(X, sgn * Y, psi, levels=levels, colors=C_THEO, linewidths=0.9)
    ax.axhline(0, color=C_THEO, lw=0.9)
    ms = 4 if len(centers) < 5 else 2.2
    for c in centers:
        ax.plot(c, a, "o", color=C_MEAS, ms=ms)
        ax.plot(c, -a, "x", color=C_MEAS, ms=ms, mew=1)
    # 방향 화살표: 각 자기장선이 x=0 단면을 지나는 곳에 표시
    d = 0.03 * (xlim[1] - xlim[0])
    def arrow(y, direction):
        ax.annotate("", xy=(direction * d, y), xytext=(-direction * d, y),
                    arrowprops=dict(arrowstyle="-|>", color=C_THEO, lw=1))
    arrow(0, 1)
    inside = yy < a
    for lv in levels[::2]:
        yin = yy[inside][np.argmin(abs(psi_mid[inside] - lv))]
        out = ~inside
        yout = yy[out][np.argmin(abs(psi_mid[out] - lv))]
        for sgn in (1, -1):
            arrow(sgn * yin, 1)
            if abs(psi_mid[out][np.argmin(abs(psi_mid[out] - lv))] - lv) < 0.02 * lv:
                arrow(sgn * yout, -1)
    ax.set_xlim(*xlim); ax.set_ylim(-ylim[1], ylim[1])
    ax.set_title(title, fontsize=9.5)

coil_panel(axs[1], [0.0], 1.0, (-2.6, 2.6), (0, 2.6),
           "(b) 원형도선 단면 (●: 나옴, ×: 들어감)", 8, 0.8)
Ls, a_s, Nl = 5.5, 1.0, 40
coil_panel(axs[2], np.linspace(-Ls / 2, Ls / 2, Nl), a_s, (-5.5, 5.5), (0, 3.2),
           "(c) 솔레노이드 (단면)", 6, 0.8)
for ax in axs:
    ax.set_aspect("equal"); ax.set_xticks([]); ax.set_yticks([])
    for s in ax.spines.values(): s.set_color("#bbb")
fig.tight_layout(); fig.savefig(FIG / "fig4_fieldlines.png"); plt.close(fig)

(ROOT / "src" / "results.json").write_text(json.dumps(out, indent=1, ensure_ascii=False))
print(json.dumps(out, indent=1, ensure_ascii=False))
