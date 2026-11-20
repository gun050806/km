// 8주차(실험 7) RLC 교류회로의 임피던스 결과 레포트 생성 스크립트
// 실행: python3 src/analysis.py && node src/build_report.js && python3 src/add_charts.py
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, BorderStyle, ShadingType, PageBreak, VerticalAlign,
  Math: OMath, MathRun, MathFraction, MathSuperScript, MathSubScript, MathRadical,
} = require("docx");

const ROOT = path.resolve(__dirname, "..");
const R = JSON.parse(fs.readFileSync(path.join(ROOT, "src", "results.json"), "utf8"));
const FONT = "맑은 고딕";
const CONTENT_W = 9026; // A4 - 좌우 여백 1440*2

// ---------- 서식 도우미 ----------
const t = (text, o = {}) => new TextRun({ text, font: FONT, ...o });
const sup = (text) => t(text, { superScript: true });
const sub = (text) => t(text, { subScript: true });

function P(children, o = {}) {
  if (typeof children === "string") children = [t(children)];
  return new Paragraph({ children, spacing: { after: 120, line: 300 }, ...o });
}
const body = (children, indent = 440) =>
  P(children, { indent: { left: indent }, alignment: AlignmentType.JUSTIFIED });
const section = (text) =>
  P([t(text, { size: 28, bold: true })], { spacing: { before: 360, after: 160 }, keepNext: true });
const sub1 = (text) =>
  P([t(text, { bold: true })], { indent: { left: 440 }, spacing: { before: 240, after: 100 }, keepNext: true });
const note = (children) =>
  P(typeof children === "string" ? [t(children, { size: 19, color: "444444" })] : children,
    { indent: { left: 440 }, spacing: { after: 100 }, keepNext: true });
const blank = () => P("");

// 과학적 표기: 1.045×10⁻⁴
function sciRuns(x, digits = 3, o = {}) {
  if (x === 0) return [t("0", o)];
  const neg = x < 0;
  const e = Math.floor(Math.log10(Math.abs(x)));
  const m = (Math.abs(x) / 10 ** e).toFixed(digits);
  return [t(`${neg ? "−" : ""}${m}×10`, o), t(String(e).replace("-", "−"), { ...o, superScript: true })];
}
const fx = (x, d) => x.toFixed(d);

// 수식
const mr = (s) => new MathRun(s);
const frac = (n, d) => new MathFraction({ numerator: n, denominator: d });
const pow = (base, e) => new MathSuperScript({ children: base, superScript: e });
const subm = (base, s) => new MathSubScript({ children: base, subScript: s });
const sqrt = (c) => new MathRadical({ children: c });
const mu0 = () => subm([mr("μ")], [mr("0")]);
function eq(children, label, prefix) {
  const runs = prefix ? [t(prefix + "   ")] : [];
  runs.push(new OMath({ children }));
  if (label) runs.push(t(`        ${label}`, { size: 19, color: "555555" }));
  return new Paragraph({ children: runs, alignment: AlignmentType.CENTER, spacing: { before: 80, after: 160 } });
}

// 표
const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };
function cell(content, w, o = {}) {
  const runs = typeof content === "string" ? [t(content, { size: 19, bold: !!o.head })] : content;
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: o.span, rowSpan: o.rowSpan,
    shading: o.head ? { fill: "EDEDED", type: ShadingType.CLEAR, color: "auto" } : undefined,
    children: [new Paragraph({ children: runs, alignment: AlignmentType.CENTER, keepNext: o.keep !== false })],
  });
}
function table(widths, rows) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    alignment: AlignmentType.CENTER, rows,
  });
}
const sci = (x, d = 3) => sciRuns(x, d, { size: 19 });
const K = (i, a) => ({ keep: i < a.length - 1 }); // 표의 마지막 행만 다음 내용과 분리 허용

// Word 차트 자리표시자 (src/add_charts.py 가 실제 차트로 교체)
function chartFig(n, caption) {
  return [
    P([t(`[[CHART:${n}]]`)], { alignment: AlignmentType.CENTER, keepNext: true }),
    P([t(caption, { size: 19, bold: true })], { alignment: AlignmentType.CENTER, spacing: { after: 240 } }),
  ];
}

// 매뉴얼에서 캡처한 그림 (png/jpg, 원본 픽셀 크기를 넘겨받음)
function figure(file, caption, widthIn, wPx, hPx) {
  const w = Math.round(widthIn * 96), h = Math.round((w * hPx) / wPx);
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 60 }, keepNext: true,
      children: [new ImageRun({ type: file.endsWith(".png") ? "png" : "jpg", data: fs.readFileSync(path.join(ROOT, "fig", file)),
        transformation: { width: w, height: h }, altText: { title: caption, description: caption, name: file } })] }),
    P([t(caption, { size: 19, bold: true })], { alignment: AlignmentType.CENTER, spacing: { after: 240 } }),
  ];
}

// ---------- 표: 조에서 정리한 표(사진)와 같은 모양 ----------
const HEAD_FILL = "E8E8E8";
const runsOf = (parts, o) => parts.map((p) => (Array.isArray(p) ? t(p[0], { ...o, subScript: true }) : t(p, o)));
function hcell(parts, w, o = {}) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 60, right: 60 },
    shading: { fill: HEAD_FILL, type: ShadingType.CLEAR, color: "auto" },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true,
      children: runsOf(typeof parts === "string" ? [parts] : parts, { size: o.size || 19, bold: true }) })],
  });
}
function bcell(text, w, o = {}) {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER, rowSpan: o.rowSpan,
    margins: { top: 100, bottom: 100, left: 60, right: 60 },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, keepNext: o.keep !== false, children: [t(text, { size: o.size || 19 })] })],
  });
}
// 표 아래 설명: 한 줄에 들어가는 짧은 줄, 가운데 정렬, 아래첨자 사용
const below = (parts, last = false) => P(runsOf(parts, { size: 19, color: "444444" }),
  { alignment: AlignmentType.CENTER, spacing: { before: last ? 0 : 80, after: last ? 200 : 0 }, keepNext: !last });
const tableTitle = (parts) => P(runsOf(parts, { bold: true, size: 22 }), { indent: { left: 440 }, spacing: { before: 240, after: 100 }, keepNext: true });

// ---------- 데이터 ----------
const pct = (x) => fx(x, 2);
const RW = R.rows, M = R.meas, T2 = R.t2, T3 = R.t3;
const at = (f) => RW.find((r) => r.f === f);
const f0 = R.f0;

// (1) 측정값 표
const w1 = [1420, 1420, 1420, 1420, 1420, 1420];
const measTable = table(w1, [
  new TableRow({ children: [hcell("주파수 f (Hz)", w1[0]), hcell(["I", ["t"], " (mA)"], w1[1]), hcell(["V", ["t"], " (V)"], w1[2]),
    hcell(["V", ["R"], " (V)"], w1[3]), hcell(["V", ["L"], " (V)"], w1[4]), hcell(["V", ["C"], " (V)"], w1[5])] }),
  ...M.map((m, i, a) => new TableRow({ children: [
    bcell(m[0], w1[0]), bcell(m[1], w1[1]),
    ...(i === 0 ? [bcell(R.vt, w1[2], { rowSpan: M.length })] : []),
    bcell(m[2], w1[3]), bcell(m[3], w1[4]), bcell(m[4], w1[5])] })),
]);

// (2) R, XL, XC 표
const w2 = [900, 850, 850, 800, 860, 860, 800, 860, 860, 880];
const H2 = ["주파수 f (Hz)", "R(측정값, Ω)", "R(이론값, Ω)", "오차율", ["X", ["L"], "(측정값, Ω)"], ["X", ["L"], "(이론값, Ω)"], "오차율",
  ["X", ["C"], "(측정값, Ω)"], ["X", ["C"], "(이론값, Ω)"], "오차율"];
const rxTable = table(w2, [
  new TableRow({ children: H2.map((h, k) => hcell(h, w2[k], { size: 17 })) }),
  ...T2.map((r, i, a) => new TableRow({ children: [M[i][0], ...r].map((v, k) => bcell(v, w2[k], { size: 18 })) })),
]);

// (3) 임피던스, 위상차 표
const w3 = [1220, 1220, 1220, 1200, 1220, 1220, 1220];
const H3 = ["주파수 f (Hz)", "Z(측정값, Ω)", "Z(이론값, Ω)", "오차율", "φ(측정값, °)", "φ(이론값, °)", "오차율"];
const zTable = table(w3, [
  new TableRow({ children: H3.map((h, k) => hcell(h, w3[k])) }),
  ...T3.map((r, i, a) => new TableRow({ children: [M[i][0], ...r].map((v, k) => bcell(v, w3[k])) })),
]);

// ---------- 수식 ----------
const Xs = (k) => subm([mr("X")], [mr(k)]), Vs = (k) => subm([mr("V")], [mr(k)]);
const tanInv = () => pow([mr("tan")], [mr("−1")]);
const eqX = eq([Xs("L"), mr("="), mr("ω"), mr("L"), mr("="), mr("2π"), mr("f"), mr("L"), mr(",    "),
  Xs("C"), mr("="), frac([mr("1")], [mr("ω"), mr("C")]), mr("="), frac([mr("1")], [mr("2π"), mr("f"), mr("C")])], "(리액턴스)");
const eqZ = eq([mr("Z"), mr("="), sqrt([pow([mr("R")], [mr("2")]), mr("+"), pow([mr("("), Xs("L"), mr("−"), Xs("C"), mr(")")], [mr("2")])]), mr(",    "),
  mr("φ"), mr("="), tanInv(), frac([Vs("L"), mr("−"), Vs("C")], [Vs("R")]), mr("="), tanInv(), frac([Xs("L"), mr("−"), Xs("C")], [mr("R")])], "(임피던스와 위상차)");
const eqF0 = eq([subm([mr("f")], [mr("0")]), mr("="), frac([mr("1")], [mr("2π"), sqrt([mr("L"), mr("C")])])], "(공진 주파수)");

// ---------- 본문 ----------
const r100 = at(100), r250 = at(250), r500 = at(500), r2000 = at(2000);
const Rms = RW.map((r) => r.Rm), eRs = RW.map((r) => r.eR), eps = RW.map((r) => r.ep);
const children = [
  P([t("8주차 레포트", { size: 20 })]),
  ...Array(6).fill(0).map(blank),
  P([t("RLC 회로의 임피던스 실험", { size: 52, bold: true })], { alignment: AlignmentType.CENTER }),
  ...Array(10).fill(0).map(blank),
  P("과목명: 일반물리실험II"), P("담당 조교: 이승윤 조교님"), P("학번: 20241501"), P("이름: 이건"),
  new Paragraph({ children: [new PageBreak()] }),

  P([t("실험 7. RLC로 이루어진 교류회로에서의 임피던스", { size: 32, bold: true })], { spacing: { after: 240 } }),

  section("1. 실험 목적"),
  body("교류회로에서 저항과 코일, 축전기의 반응저항(리액턴스)을 측정하여 임피던스의 개념을 이해하도록 한다. 이를 통해 주파수에 따른 리액턴스와 임피던스의 변화와 교류회로에서의 전압분배 및 위상차를 충분히 이해하도록 한다."),

  section("2. 데이터 정리"),
  tableTitle(["(1) 측정값 표"]),
  P([t("저항: 100 Ω, 코일: 47 mH, 축전기: 4.7 μF", { size: 19 })], { indent: { left: 440 }, spacing: { after: 80 }, keepNext: true }),
  measTable,
  below(["※ R-L-C 직렬회로에서 주파수를 바꿔가며 측정하였다."], true),
  tableTitle(["(2) R, X", ["L"], ", X", ["C"]]),
  rxTable,
  below(["※ 측정값:  R = V", ["R"], "/I", ["t"], ",   X", ["L"], " = V", ["L"], "/I", ["t"], ",   X", ["C"], " = V", ["C"], "/I", ["t"]]),
  below(["※ 이론값:  R = 100 Ω,   X", ["L"], " = 2πfL,   X", ["C"], " = 1/(2πfC)"], true),
  ...chartFig(1, "[그림 1] 주파수에 따른 저항과 리액턴스"),
  tableTitle(["(3) 임피던스, 위상차"]),
  zTable,
  below(["※ 측정값:  Z = √(V", ["R"], "² + (V", ["L"], " − V", ["C"], ")²) / I", ["t"], ",   φ = tan⁻¹((V", ["L"], " − V", ["C"], ") / V", ["R"], ")"]),
  below(["※ 이론값:  Z = √(R² + (X", ["L"], " − X", ["C"], ")²),   φ = tan⁻¹((X", ["L"], " − X", ["C"], ") / R)"], true),
  ...chartFig(2, "[그림 2] 주파수에 따른 임피던스"),
  ...chartFig(3, "[그림 3] 주파수에 따른 위상차"),

  section("3. 결과 분석"),
  body("교류회로에서는 전압과 전류의 크기와 방향이 시간에 따라 계속 변하고 코일과 축전기는 저항처럼 전류의 흐름을 방해하는 반응저항(리액턴스)을 가진다. 코일의 유도 반응저항 XL과 축전기의 용량 반응저항 XC는 다음과 같이 주파수에 따라 달라진다."),
  eqX,
  body("R-L-C 직렬회로에서는 모든 소자에 같은 전류가 흐르지만 코일에 걸리는 전압 VL은 저항에 걸리는 전압 VR보다 위상이 90° 앞서고 축전기에 걸리는 전압 VC는 90° 뒤진다. 그러므로 각 전압을 벡터로 나타내면 회로의 임피던스 Z와 위상차 φ는 다음과 같다."),
  eqZ,
  body("또한 XL과 XC가 같아지는 주파수에서는 두 리액턴스가 서로 상쇄되어 임피던스가 Z = R로 가장 작아지는데 이 주파수를 공진 주파수라고 한다."),
  eqF0,
  body(`먼저 저항 R을 측정값으로 구하면 ${fx(Math.min(...Rms), 1)}~${fx(Math.max(...Rms), 1)} Ω으로 이론값인 100 Ω과 ${fx(Math.min(...eRs), 1)}~${fx(Math.max(...eRs), 1)}%의 오차가 있었다. 이론적으로 저항은 주파수와 상관없이 일정하지만 측정값은 주파수가 커질수록 조금씩 커졌다. 다음으로 코일의 유도 반응저항 XL은 주파수가 100 Hz에서 2000 Hz로 커지는 동안 ${T2[0][3]} Ω에서 ${T2[5][3]} Ω으로 증가하였고 축전기의 용량 반응저항 XC는 ${T2[0][6]} Ω에서 ${T2[5][6]} Ω으로 감소하였다. [그림 1]을 보면 측정값이 XL = 2πfL의 직선과 XC = 1/(2πfC)의 곡선을 잘 따라가기 때문에 XL은 주파수에 비례하고 XC는 주파수에 반비례함을 확인할 수 있다. 그리고 임피던스 Z는 100 Hz에서 ${T3[0][0]} Ω이었다가 250 Hz에서 ${T3[1][0]} Ω으로 가장 작아진 뒤 주파수가 커질수록 다시 커져 2000 Hz에서 ${T3[5][0]} Ω이 되었다. 이는 낮은 주파수에서는 XC가 커서 임피던스가 크고 높은 주파수에서는 XL이 커서 임피던스가 크기 때문이며 이론적인 공진 주파수 f₀ = ${fx(f0, 1)} Hz 근처에서 XL과 XC가 서로 상쇄되어 임피던스가 가장 작아짐을 알 수 있다. 전류 It도 250 Hz에서 46 mA로 가장 컸는데 이는 임피던스가 작을수록 같은 전압에서 전류가 많이 흐르기 때문이다. 마지막으로 위상차 φ는 100 Hz와 250 Hz에서 ${T3[0][3]}°, ${T3[1][3]}°로 음수였고 500 Hz부터는 양수가 되어 2000 Hz에서 ${T3[5][3]}°까지 커졌다. 공진 주파수보다 낮은 주파수에서는 VC가 VL보다 커서 회로가 축전기처럼 동작하고 높은 주파수에서는 VL이 커서 코일처럼 동작하기 때문에 위상차의 부호가 250 Hz와 500 Hz 사이에서 바뀌는 것을 확인할 수 있다. 위상차의 오차율은 ${T3[0][5]}~${T3[2][5]}로 저항이나 리액턴스, 임피던스의 오차율보다 작았다.`),

  section("4. 토의"),
  sub1("(1) 실험 검토"),
  body("이번 실험에서는 R-L-C 직렬회로에 걸어주는 교류의 주파수를 100 Hz에서 2000 Hz까지 바꿔가며 각 소자에 걸리는 전압과 전류를 측정하였다. 이를 통해 코일의 리액턴스는 주파수에 비례하고 축전기의 리액턴스는 주파수에 반비례하며 회로의 임피던스는 공진 주파수 근처에서 가장 작아진다는 것을 알 수 있었다. 그리고 위상차의 부호가 공진 주파수를 기준으로 바뀌는 것도 확인하였다. 측정값과 이론값 사이에 오차가 있었지만 이론에서 예상한 경향을 모두 확인할 수 있었기 때문에 실험이 잘 진행되었다고 생각한다.", 800),
  sub1("(2) 오차 분석 및 개선 방안"),
  body("저항과 리액턴스, 임피던스의 오차율이 최대 15% 정도로 나타난 것을 보면 실험과정에서 오차가 발생한 것을 알 수 있다. 같은 주파수에서 R, XL, XC의 측정값이 1000 Hz 이하에서는 모두 이론값보다 작고 1500 Hz 이상에서는 모두 이론값보다 크게 나왔는데 세 값 모두 같은 전류 It로 나누어 구했기 때문에 이러한 오차의 가장 큰 원인은 전류 측정에 있었다고 생각한다. 멀티미터로 읽은 전류가 8 mA, 11 mA처럼 1 mA 단위였기 때문에 전류가 작은 높은 주파수에서는 0.5 mA만 달라져도 4~6%의 오차가 생길 수 있었다. 반면 위상차는 전류를 사용하지 않고 전압의 비로만 구하기 때문에 오차율이 작게 나왔다고 생각한다. 또한 전압을 scope에서 delta tool로 최댓값과 최솟값을 직접 찾아 읽었기 때문에 정확한 위치를 잡지 못했을 가능성도 있고 코일 도선 자체의 저항도 오차의 원인이 되었을 것이라고 생각한다. 오차를 개선하기 위해 전류를 더 작은 단위까지 측정할 수 있는 측정기를 사용하고 여러 번 측정해 평균값을 사용하는 방법이 존재한다.", 800),
  sub1("(3) 이론을 확인할 다른 실험 제안"),
  body(`이 실험에서는 100 Hz에서 2000 Hz 사이의 정해진 주파수에서만 측정하였기 때문에 공진 주파수를 정확히 확인하지 못하였다. 250 Hz와 500 Hz 사이에서 주파수를 10 Hz씩 바꿔가며 전류와 위상차를 측정하면 전류가 가장 크고 위상차가 0이 되는 공진 주파수를 직접 찾아 이론값 ${fx(f0, 1)} Hz와 비교해볼 수 있을 것이다. 또한 저항과 코일, 축전기를 병렬로 연결하여 같은 실험을 진행해볼 수 있을 것이다.`, 800),
  sub1("(4) 응용 가능성"),
  body("R-L-C 회로가 공진 주파수에서 임피던스가 가장 작아져 전류가 크게 흐르는 원리를 이용하면 라디오에서 원하는 방송의 주파수만 골라 듣는 장치를 만들 수 있다. 또한 코일은 높은 주파수의 전류를 잘 통과시키지 못하고 축전기는 낮은 주파수의 전류를 잘 통과시키지 못하는 성질을 이용하면 스피커에서 낮은 소리와 높은 소리를 나누어 보내는 필터를 만들 수 있다고 생각한다.", 800),

  section("5. 참고 문헌"),
  body("일반물리학실험 매뉴얼, 서강대학교 물리학과"),
];

const doc = new Document({
  creator: "이건",
  title: "8주차 레포트 - RLC 회로의 임피던스 실험",
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 },
      margin: { top: 1701, right: 1440, bottom: 1440, left: 1440 } } },
    children,
  }],
});

const out = path.join(ROOT, "20241501_이건_8주차_레포트.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("wrote", out); });
