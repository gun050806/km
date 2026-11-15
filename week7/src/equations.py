"""수식을 PNG 그림으로 만든다 (Word 수식을 표시하지 못하는 뷰어에서도 보이도록)."""
from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

FIG = Path(__file__).resolve().parents[1] / "fig"
plt.rcParams["mathtext.fontset"] = "cm"
EQS = {
    "eq1_faraday.png": r"$\varepsilon = -N\,\dfrac{d\Phi_B}{dt}, \qquad \Phi_B = \int \; \vec{B}\cdot d\vec{A}$",
    "eq2_transformer.png": r"$\dfrac{V_2}{V_1} = \dfrac{N_2}{N_1}, \qquad V_1 I_1 = V_2 I_2 \;\;\Rightarrow\;\; \dfrac{I_2}{I_1} = \dfrac{N_1}{N_2}$",
}
for name, tex in EQS.items():
    fig = plt.figure(figsize=(0.01, 0.01))
    fig.text(0, 0, tex, fontsize=15)
    fig.savefig(FIG / name, dpi=300, bbox_inches="tight", pad_inches=0.03, transparent=False, facecolor="white")
    plt.close(fig)
    print(name)
