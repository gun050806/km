"""7주차(실험 6) 전자기 유도와 변압기: 측정값 정리 + 계산 -> results.json"""
import json
from pathlib import Path
import numpy as np

ROOT = Path(__file__).resolve().parents[1]

# ---- 측정값 (조 데이터 사진 그대로) ----
# 표 1.1: 첫 번째 400회, 두 번째 400회
core = [("철심이 없을 때", 0.252),
        ("철심이 두 코일의 사이에 놓여졌을 때", 1.134),
        ("U자형 철심이 두 코일에 연결되었을 때", 1.716),
        ("U자형 철심에 가로대가 추가로 놓여졌을 때", 3.190)]
VIN_CORE = 3.564   # 표 1.1 입력전압: 표 1.2 400:400 측정(출력 3.190 V 동일)과 같은 조건으로 봄
# 표 1.2: (N1, N2, Vin, Vout)
turns = [(400, 400, 3.564, 3.190), (400, 200, 3.551, 1.596), (400, 800, 3.518, 6.342)]
# 표 2.1: (N1, N2, R, Vin, Iin[mA], Vout)  -- 출력전류는 Vout/R
circ = [(400, 200, 100, 3.637, 9.72, 1.750), (400, 400, 100, 3.610, 33.2, 3.190),
        (400, 800, 100, 3.585, 118.0, 6.200)]

out = {"vin_core": VIN_CORE}
out["core"] = [dict(name=n, vout=v, ratio=v / VIN_CORE) for n, v in core]

rows = []
for n1, n2, vin, vout in turns:
    ideal = n2 / n1
    rows.append(dict(n1=n1, n2=n2, vin=vin, vout=vout, vth=vin * ideal,
                     ratio=vout / vin, ideal=ideal, err=abs(vout / vin - ideal) / ideal * 100))
out["turns"] = rows
x = np.array([r["n2"] for r in rows]); y = np.array([r["vout"] for r in rows])
sl, ic = np.polyfit(x, y, 1)
out["turns_fit"] = dict(slope=sl, intercept=ic,
                        r2=1 - ((y - (sl * x + ic)) ** 2).sum() / ((y - y.mean()) ** 2).sum(),
                        ideal_slope=np.mean([r["vin"] for r in rows]) / 400,
                        mean_ratio_to_ideal=np.mean([r["ratio"] / r["ideal"] for r in rows]))

rows = []
for n1, n2, R, vin, iin_mA, vout in circ:
    iin, iout = iin_mA / 1000, vout / R
    pin, pout = vin * iin, vout * iout
    gv, gi = vout / vin, iout / iin
    rows.append(dict(n1=n1, n2=n2, R=R, vin=vin, iin=iin, pin=pin, vout=vout, iout=iout, pout=pout,
                     gv=gv, gv_ideal=n2 / n1, gv_err=abs(gv - n2 / n1) / (n2 / n1) * 100,
                     gi=gi, gi_ideal=n1 / n2, gi_err=abs(gi - n1 / n2) / (n1 / n2) * 100,
                     gp=pout / pin, loss=(1 - pout / pin) * 100,
                     r_ref=R * (n1 / n2) ** 2, iin_ref=vin / (R * (n1 / n2) ** 2)))
out["circ"] = rows

(ROOT / "src" / "results.json").write_text(json.dumps(out, indent=1, ensure_ascii=False))
for k in ("core", "turns", "turns_fit", "circ"):
    print(k, json.dumps(out[k], ensure_ascii=False, default=float)[:900])
