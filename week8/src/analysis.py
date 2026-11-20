"""8주차(실험 7) RLC 교류회로: 사진의 측정값/표를 정리하고 계산으로 검증 -> results.json"""
import json, math
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
R0, L0, C0 = 100.0, 47e-3, 4.7e-6
VT = "5.1845"
# (1) 측정값 표 (사진 그대로): f, It(mA), VR, VL, VC
MEAS = [("100", "18", "1.542", "0.495", "5.1825"), ("250", "46", "3.986", "2.968", "5.3145"),
        ("500", "42", "3.742", "5.326", "2.5485"), ("1000", "19", "1.864", "5.192", "0.6185"),
        ("1500", "11", "1.198", "5.084", "0.2615"), ("2000", "8", "0.904", "5.042", "0.1465")]
# (2) R, XL, XC 표 (사진 그대로)
T2 = [("85.7", "100", "14.3%", "27.5", "29.53", "6.9%", "287.9", "338.63", "15.0%"),
      ("86.7", "100", "13.3%", "64.5", "73.83", "12.6%", "115.5", "135.45", "14.7%"),
      ("89.1", "100", "10.9%", "126.8", "147.65", "14.1%", "60.7", "67.73", "10.4%"),
      ("98.1", "100", "1.9%", "273.3", "295.31", "7.5%", "32.6", "33.86", "3.9%"),
      ("108.9", "100", "8.9%", "462.2", "442.96", "4.3%", "23.8", "22.58", "5.3%"),
      ("113.0", "100", "13.0%", "630.3", "590.62", "6.7%", "18.3", "16.93", "8.2%")]
# (3) 임피던스, 위상차 표 (사진 그대로)
T3 = [("274.1", "324.9", "15.6%", "-71.79", "-72.07", "0.39%"),
      ("100.6", "117.5", "14.4%", "-30.48", "-31.64", "3.66%"),
      ("111.0", "128.0", "13.3%", "36.58", "38.64", "5.31%"),
      ("259.9", "279.9", "7.1%", "67.83", "69.07", "1.80%"),
      ("451.7", "432.1", "4.5%", "76.05", "76.62", "0.74%"),
      ("622.3", "582.3", "6.9%", "79.54", "80.11", "0.72%")]

def err(m, t):
    return abs(m - t) / abs(t) * 100

rows, bad = [], []
def check(label, shown, value):
    s = shown.rstrip("%"); d = len(s.split(".")[1]) if "." in s else 0
    if abs(float(s) - value) > 0.5 * 10 ** -d + 1e-9 and abs(float(s) - value) > 0.012:
        bad.append(f"{label}: 사진 {shown} / 계산 {value:.4f}")

for (f, it, vr, vl, vc), t2, t3 in zip(MEAS, T2, T3):
    f, it, vr, vl, vc = float(f), float(it) / 1000, float(vr), float(vl), float(vc)
    w = 2 * math.pi * f
    Rm, XLm, XCm = vr / it, vl / it, vc / it
    XLt, XCt = w * L0, 1 / (w * C0)
    Zm = math.hypot(vr, vl - vc) / it
    Zt = math.hypot(R0, XLt - XCt)
    pm = math.degrees(math.atan((vl - vc) / vr))
    pt = math.degrees(math.atan((XLt - XCt) / R0))
    calc2 = [Rm, R0, err(Rm, R0), XLm, XLt, err(XLm, XLt), XCm, XCt, err(XCm, XCt)]
    calc3 = [Zm, Zt, err(Zm, Zt), pm, pt, err(pm, pt)]
    for k, (s, v) in enumerate(zip(t2, calc2)): check(f"표2 f={f:g} 열{k}", s, v)
    for k, (s, v) in enumerate(zip(t3, calc3)): check(f"표3 f={f:g} 열{k}", s, v)
    rows.append(dict(f=f, it=it, vr=vr, vl=vl, vc=vc, Rm=Rm, XLm=XLm, XCm=XCm, XLt=XLt, XCt=XCt,
                     Zm=Zm, Zt=Zt, pm=pm, pt=pt, eR=err(Rm, R0), eXL=err(XLm, XLt), eXC=err(XCm, XCt),
                     eZ=err(Zm, Zt), ep=err(pm, pt), vsum=math.hypot(vr, vl - vc)))

f0 = 1 / (2 * math.pi * math.sqrt(L0 * C0))
out = dict(R=R0, L=L0, C=C0, vt=VT, f0=f0, meas=MEAS, t2=T2, t3=T3, rows=rows, mismatch=bad)
(ROOT / "src" / "results.json").write_text(json.dumps(out, indent=1, ensure_ascii=False))
print("f0 =", round(f0, 1), "Hz")
print("불일치:", bad or "없음")
for r in rows:
    print(f"{r['f']:>6g}  ep={r['ep']:.2f}  eR={r['eR']:.1f} eXL={r['eXL']:.1f} eXC={r['eXC']:.1f}  Rm-sign {'+' if r['Rm']>100 else '-'}")
