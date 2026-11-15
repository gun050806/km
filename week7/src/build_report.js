// 7주차(실험 6) 전자기 유도와 변압기 결과 레포트 생성 스크립트
// 실행: python3 src/analysis.py && python3 src/equations.py && node src/build_report.js && python3 src/add_charts.py
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

// 매뉴얼에서 캡처한 그림
function figure(file, caption, widthIn) {
  const buf = fs.readFileSync(path.join(ROOT, "fig", file));
  const wPx = buf.readUInt32BE(16), hPx = buf.readUInt32BE(20); // PNG 크기
  const w = Math.round(widthIn * 96), h = Math.round((w * hPx) / wPx);
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 160, after: 60 }, keepNext: true,
      children: [new ImageRun({ type: "png", data: buf, transformation: { width: w, height: h },
        altText: { title: caption, description: caption, name: file } })] }),
    P([t(caption, { size: 19, bold: true })], { alignment: AlignmentType.CENTER, spacing: { after: 240 } }),
  ];
}

// 수식 그림 (src/equations.py 가 만든 PNG). 어느 뷰어에서나 보이도록 Word 수식 대신 사용
function eqImg(file, label) {
  const buf = fs.readFileSync(path.join(ROOT, "fig", file));
  const wPx = buf.readUInt32BE(16), hPx = buf.readUInt32BE(20);
  const w = Math.round((wPx / 300) * 0.8 * 96), h = Math.round((w * hPx) / wPx);
  const runs = [new ImageRun({ type: "png", data: buf, transformation: { width: w, height: h },
    altText: { title: label, description: label, name: file } })];
  if (label) runs.push(t(`        ${label}`, { size: 19, color: "555555" }));
  return new Paragraph({ children: runs, alignment: AlignmentType.CENTER, spacing: { before: 80, after: 160 } });
}

// ---------- 데이터 ----------
const pct = (x) => fx(x, 2);
const CORE = R.core, T = R.turns, F = R.turns_fit;
const C = [...R.circ].sort((a, b) => a.n2 - b.n2);
const Ts = [...T].sort((a, b) => a.n2 - b.n2);
const H = (s, w) => cell(s, w, { head: true });
const NN = (sym, s) => [t(sym, { size: 19, bold: true }), t(s, { size: 19, bold: true, subScript: true })];
const NNh = (pre, sym, s, post, w) => cell([t(pre, { size: 19, bold: true }), ...NN(sym, s), t(post, { size: 19, bold: true })], w, { head: true });

// 표 1.1
const w11 = [5000, 2013, 2013];
const coreTable = table(w11, [
  new TableRow({ children: [H("코일 구조", w11[0]), H("출력 전압 (V)", w11[1]), H("출력 전압 / 입력 전압", w11[2])] }),
  ...CORE.map((c, i, a) => new TableRow({ children: [cell(c.name, w11[0], K(i, a)), cell(fx(c.vout, 3), w11[1], K(i, a)), cell(fx(c.ratio, 3), w11[2], K(i, a))] })),
]);

// 표 1.2
const w12 = [1000, 1000, 1100, 1100, 1300, 1200, 1200, 1126];
const turnsTable = table(w12, [
  new TableRow({ children: [H("첫 번째 코일", w12[0]), H("두 번째 코일", w12[1]), H("입력 전압 (V)", w12[2]),
    H("출력 전압 (V)", w12[3]), H("이론 출력 전압 (V)", w12[4]), H("실제 전압비", w12[5]),
    H([t("이상 전압비 (N", { size: 19, bold: true }), t("2", { size: 19, bold: true, subScript: true }), t("/N", { size: 19, bold: true }), t("1", { size: 19, bold: true, subScript: true }), t(")", { size: 19, bold: true })], w12[6]),
    H("오차율 (%)", w12[7])] }),
  ...T.map((r, i, a) => new TableRow({ children: [cell(String(r.n1), w12[0], K(i, a)), cell(String(r.n2), w12[1], K(i, a)), cell(fx(r.vin, 3), w12[2], K(i, a)),
    cell(fx(r.vout, 3), w12[3], K(i, a)), cell(fx(r.vth, 3), w12[4], K(i, a)), cell(fx(r.ratio, 3), w12[5], K(i, a)), cell(fx(r.ideal, 3), w12[6], K(i, a)), cell(pct(r.err), w12[7], K(i, a))] })),
]);

// 표 2.1 (측정값)
const w21 = [880, 880, 880, 1000, 1000, 1100, 1000, 1000, 1286];
const circTable = table(w21, [
  new TableRow({ children: [H("첫 번째 코일", w21[0]), H("두 번째 코일", w21[1]), H("저항 (Ω)", w21[2]),
    H("입력 전압 (V)", w21[3]), H("입력 전류 (mA)", w21[4]), H("입력 전력 (W)", w21[5]),
    H("출력 전압 (V)", w21[6]), H("출력 전류 (mA)", w21[7]), H("출력 전력 (W)", w21[8])] }),
  ...C.map((r, i, a) => new TableRow({ children: [cell(String(r.n1), w21[0], K(i, a)), cell(String(r.n2), w21[1], K(i, a)), cell(String(r.R), w21[2], K(i, a)),
    cell(fx(r.vin, 3), w21[3], K(i, a)), cell(fx(r.iin * 1000, 2), w21[4], K(i, a)), cell(fx(r.pin, 4), w21[5], K(i, a)),
    cell(fx(r.vout, 3), w21[6], K(i, a)), cell(fx(r.iout * 1000, 2), w21[7], K(i, a)), cell(fx(r.pout, 4), w21[8], K(i, a))] })),
]);
// 표 2.2 (계산값)
const w22 = [1000, 1150, 1150, 1000, 1150, 1150, 1000, 1426];
const subH = (a, b, c, d) => [t(a, { size: 19, bold: true }), t(b, { size: 19, bold: true, subScript: true }), t(c, { size: 19, bold: true }), t(d, { size: 19, bold: true, subScript: true }), t(")", { size: 19, bold: true })];
const gainTable = table(w22, [
  new TableRow({ children: [H("두 번째 코일", w22[0]), H("전압 이득 (측정)", w22[1]),
    H(subH("전압 이득 (이상, N", "2", "/N", "1"), w22[2]), H("오차율 (%)", w22[3]),
    H("출력 전류 / 입력 전류", w22[4]), H(subH("전류 비 (이상, N", "1", "/N", "2"), w22[5]),
    H("오차율 (%)", w22[6]), H("전력 이득", w22[7])] }),
  ...C.map((r, i, a) => new TableRow({ children: [cell(String(r.n2), w22[0], K(i, a)), cell(fx(r.gv, 3), w22[1], K(i, a)), cell(fx(r.gv_ideal, 3), w22[2], K(i, a)),
    cell(pct(r.gv_err), w22[3], K(i, a)), cell(fx(r.gi, 3), w22[4], K(i, a)), cell(fx(r.gi_ideal, 3), w22[5], K(i, a)), cell(pct(r.gi_err), w22[6], K(i, a)),
    cell(fx(r.gp, 3), w22[7], K(i, a))] })),
]);

// ---------- 본문 ----------
const vinAvg = T.reduce((a, r) => a + r.vin, 0) / T.length;
const gpMin = Math.min(...C.map((r) => r.gp)), gpMax = Math.max(...C.map((r) => r.gp));
const lossMin = Math.min(...C.map((r) => r.loss)), lossMax = Math.max(...C.map((r) => r.loss));
const slopeShort = (1 - F.slope / F.ideal_slope) * 100;
const q = (text) => P([t(text, { bold: true })], { indent: { left: 440 }, spacing: { before: 200, after: 100 }, keepNext: true });
const Vs = (k) => subm([mr("V")], [mr(k)]), Is = (k) => subm([mr("I")], [mr(k)]), Ns = (k) => subm([mr("N")], [mr(k)]);
const Tc = (n2) => T.find((r) => r.n2 === n2), Cc = (n2) => C.find((r) => r.n2 === n2);

const children = [
  // 표지
  P([t("7주차 레포트", { size: 20 })]),
  ...Array(6).fill(0).map(blank),
  P([t("전자기 유도와 변압기 실험", { size: 52, bold: true })], { alignment: AlignmentType.CENTER }),
  ...Array(10).fill(0).map(blank),
  P("과목명: 일반물리실험II"), P("담당 조교: 이승윤 조교님"), P("학번: 20241501"), P("이름: 이건"),
  new Paragraph({ children: [new PageBreak()] }),

  P([t("실험 6. 전자기 유도와 변압기", { size: 32, bold: true })], { spacing: { after: 240 } }),

  section("1. 실험 목적"),
  body("코일의 철심 구조와 감은 수, 출력 쪽 회로를 바꿔가며 변압기의 입력과 출력 전압, 전류를 측정하도록 한다. 이를 통해 패러데이 법칙에 의한 전자기 유도와 변압기의 원리를 충분히 이해하도록 한다."),

  section("2. 데이터 정리"),
  sub1("[실험 1] 구조에 따른 전압의 변화 측정"),
  P([t("(1) 코일 구조에 따른 출력 전압", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  ...figure("manual_fig12_core.png", "[그림 1] 철심 구조 (왼쪽부터 두 코일 사이에 가로대, U자형 철심, U자형 철심 + 가로대)", 6.0),
  note(`(첫 번째 코일 = 400회, 두 번째 코일 = 400회, 입력 전압 = ${fx(R.vin_core, 3)} V)`),
  coreTable,
  ...chartFig(1, "[그림 2] 코일 구조에 따른 출력 전압"),
  P([t("(2) 두 번째 코일의 감은 수에 따른 출력 전압", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  note("(U자형 철심에 가로대를 추가로 놓은 상태에서 측정하였다. 이론 출력 전압 = 입력 전압 × N₂/N₁)"),
  turnsTable,
  ...chartFig(2, "[그림 3] 두 번째 코일의 감은 수에 따른 출력 전압"),

  sub1("[실험 2] 회로에 따른 전압 변화 측정"),
  ...figure("manual_fig13_circuit.png", "[그림 4] 출력 쪽에 부하 저항을 연결한 변압기 회로", 4.6),
  note("(두 번째 코일에 100 Ω의 부하 저항을 연결하였다. 출력 전류 = 출력 전압 / 부하 저항, 전압 이득 = Vout/Vin, 전력 이득 = Pout/Pin = IoutVout/IinVin)"),
  P([t("(1) 측정값", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  circTable,
  blank(),
  P([t("(2) 전압 이득, 전류 비, 전력 이득", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  gainTable,
  ...chartFig(3, "[그림 5] 두 번째 코일의 감은 수에 따른 전압 이득과 전류 비"),

  section("3. 결과 분석"),
  body("코일을 통과하는 자기선속이 시간에 따라 변하면 코일에 유도 기전력이 생기는데 이를 패러데이 법칙이라고 한다. 유도 기전력은 자기선속의 변화를 방해하는 방향으로 생기며(렌츠의 법칙), N번 감은 코일에서는 다음과 같이 나타낼 수 있다."),
  eqImg("eq1_faraday.png", "(패러데이 법칙)"),
  body("변압기는 첫 번째 코일에 교류 전류를 흘려 시간에 따라 변하는 자기선속을 만들고, 이 자기선속이 철심을 따라 두 번째 코일을 통과하면서 유도 기전력이 생기는 원리를 이용한다. 두 코일을 같은 자기선속이 지나간다면 한 번 감을 때마다 생기는 기전력이 같기 때문에 전압의 비는 감은 수의 비와 같다. 그리고 이상적인 변압기에서는 전력 손실이 없기 때문에 입력 전력과 출력 전력이 같고, 전류의 비는 감은 수의 비와 반대가 된다."),
  eqImg("eq2_transformer.png", "(이상적인 변압기)"),
  body(`[실험 1]에서는 400회 감은 두 코일 사이의 철심 구조를 바꿔가며 두 번째 코일의 출력 전압을 측정하였다. 철심이 없을 때 출력 전압은 ${fx(CORE[0].vout, 3)} V로 입력 전압의 약 ${fx(CORE[0].ratio * 100, 0)}%밖에 되지 않았지만, 두 코일 사이에 가로대를 놓으면 ${fx(CORE[1].vout, 3)} V, U자형 철심에 두 코일을 끼우면 ${fx(CORE[2].vout, 3)} V, 여기에 가로대를 추가로 놓으면 ${fx(CORE[3].vout, 3)} V로 입력 전압의 약 ${fx(CORE[3].ratio * 100, 0)}%까지 커졌다. 철은 공기보다 투자율이 매우 크기 때문에 첫 번째 코일이 만든 자기선속이 철심을 따라 모여 두 번째 코일을 통과하게 된다. 그러므로 철심이 두 코일을 끊김 없이 연결해 자기선속이 지나가는 닫힌 경로를 만들수록 출력 전압이 커짐을 알 수 있다. 다음으로 U자형 철심에 가로대를 놓은 상태에서 두 번째 코일의 감은 수를 200, 400, 800회로 바꿔가며 출력 전압을 측정하였다. 출력 전압은 ${fx(Tc(200).vout, 3)} V, ${fx(Tc(400).vout, 3)} V, ${fx(Tc(800).vout, 3)} V로 감은 수가 두 배가 될 때마다 거의 두 배씩 커졌고, 추세선이 y = ${fx(F.slope, 4)}x + ${fx(F.intercept, 2)}인 일차함수인 것을 보면 출력 전압과 두 번째 코일의 감은 수가 비례관계임을 확인할 수 있다. 다만 실제 전압비는 세 경우 모두 이상적인 전압비 N₂/N₁보다 약 10% 작게 나왔다. 마지막으로 [실험 2]에서는 두 번째 코일에 100 Ω의 부하 저항을 연결하고 입력과 출력의 전압, 전류를 측정하였다. 전압 이득은 ${C.map((r) => fx(r.gv, 3)).join(", ")}로 N₂/N₁(0.5, 1, 2)에 따라 커졌고, 출력 전류와 입력 전류의 비는 ${C.map((r) => fx(r.gi, 3)).join(", ")}로 N₁/N₂(2, 1, 0.5)에 따라 작아졌다. 이를 통해 전압은 감은 수에 비례하고 전류는 감은 수에 반비례한다는 변압기의 관계가 성립함을 알 수 있다. 또한 전력 이득은 ${fx(gpMin, 3)}~${fx(gpMax, 3)}로 1보다 작게 나와 입력 전력의 약 ${fx(lossMin, 0)}~${fx(lossMax, 0)}%가 변압기에서 손실되었음을 알 수 있다.`),

  section("4. 질문"),
  sub1("[실험 1]"),
  q("(1) 어떠한 코어 구성이 두 번째 코일에 전자기 효과를 최대한으로 전달해 주는가? 여러 코어 구성 사이에서 생기는 차이를 설명할 수 있는 이론을 만들어라."),
  body(`실험결과를 살펴보면 U자형 철심에 가로대를 추가로 놓았을 때 출력 전압이 ${fx(CORE[3].vout, 3)} V로 가장 컸기 때문에 이 구성이 두 번째 코일에 전자기 효과를 가장 많이 전달한다. 두 번째 코일에 유도되는 기전력은 두 번째 코일을 통과하는 자기선속의 변화에 비례하는데, 철심은 투자율이 커서 자기선속이 공기보다 철심을 따라 지나가게 된다. 철심이 없으면 첫 번째 코일이 만든 자기선속 대부분이 공기 중으로 퍼져 두 번째 코일을 거의 지나가지 않는다. [그림 1]의 왼쪽과 가운데처럼 가로대나 U자형 철심만 있을 때는 자기선속이 철심을 따라 일부 전달되지만 철심이 끊긴 부분(공기 틈)에서 다시 퍼지게 된다. 오른쪽처럼 U자형 철심에 가로대를 놓으면 철심이 하나의 닫힌 고리가 되어 자기선속이 거의 빠져나가지 않고 두 번째 코일을 통과한다. 따라서 철심이 자기선속이 지나가는 닫힌 경로를 만들수록, 즉 공기 틈이 적을수록 두 번째 코일에 전달되는 자기선속이 많아져 출력 전압이 커진다는 이론을 세울 수 있다.`, 800),
  q("(2) 테이블 1.2에서 데이터로부터 두 번째 코일에 감은 횟수의 변화에 따른 출력 전압을 그래프로 그려라. 전선의 감은 수와 결과 출력 전압 사이에 어떠한 관계가 존재하는가?"),
  body(`두 번째 코일의 감은 수에 따른 출력 전압을 그래프로 그리면 [그림 3]과 같다. 측정값이 거의 직선 위에 놓이고 추세선이 y = ${fx(F.slope, 4)}x + ${fx(F.intercept, 2)}인 일차함수이므로 전선의 감은 수와 출력 전압은 비례관계이다. 이는 V₂ = V₁N₂/N₁의 관계와 부합한다. 다만 그래프의 기울기가 이론값(V₁/N₁ ≈ ${fx(F.ideal_slope, 4)} V/회)보다 약 ${fx(slopeShort, 0)}% 작은데, 이는 첫 번째 코일이 만든 자기선속이 모두 두 번째 코일을 통과하지 못했기 때문이다.`, 800),
  q("(3) 변압기의 성능을 향상시킬 수 있는 방법을 생각해 보라. 한 코일로부터 다른 코일에 전송을 향상시키기 위해서 어떤 추가적인 변화가 있어야 하는가?"),
  body("변압기의 성능을 높이기 위해서는 첫 번째 코일이 만든 자기선속이 빠져나가지 않고 두 번째 코일로 전달되도록 해야 한다. 이를 위해 철심을 닫힌 고리 모양으로 만들고 가로대와 U자형 철심 사이의 틈을 최대한 줄여야 한다. 또한 코일과 철심 사이의 간격을 줄이거나 두 코일을 같은 철심 위에 겹쳐 감는 방법이 있고, 철보다 투자율이 큰 물질로 철심을 만들거나 얇은 철판을 절연하여 쌓은 철심을 사용해 맴돌이 전류에 의한 손실을 줄이는 방법이 존재한다. 그리고 저항이 작은 굵은 도선으로 코일을 감으면 도선에서 열로 손실되는 에너지를 줄일 수 있다.", 800),
  sub1("[실험 2]"),
  q("(1) 표 2.1에 나온 값들을 계산하여라."),
  body("입력 전력과 출력 전력은 P = IV로, 전압 이득은 Vout/Vin, 전력 이득은 Pout/Pin으로 계산하여 [표 2.1]의 (1), (2)에 정리하였다.", 800),
  q("(2) 같은 값의 저항을 가지며 감은 수가 다른 두 코일에 대해 출력 전류와 입력 전류 사이에 어떤 관계가 존재하는가?"),
  body(`실험결과를 살펴보면 같은 100 Ω 저항에서 출력 전류와 입력 전류의 비는 두 번째 코일이 200회일 때 ${fx(Cc(200).gi, 3)}, 400회일 때 ${fx(Cc(400).gi, 3)}, 800회일 때 ${fx(Cc(800).gi, 3)}였다. 이는 각각 N₁/N₂ 값인 2, 1, 0.5와 가까우므로 출력 전류와 입력 전류의 비는 감은 수의 비에 반비례한다(Iout/Iin = N₁/N₂). 이상적인 변압기에서는 VinIin = VoutIout이고 Vout/Vin = N₂/N₁이기 때문에 이러한 관계가 나온다. 또한 두 번째 코일의 감은 수가 커질수록 입력 전류가 ${C.map((r) => fx(r.iin * 1000, 2) + " mA").join(", ")}로 크게 증가하였는데, 입력 쪽에서 보면 부하 저항이 R(N₁/N₂)², 즉 ${C.map((r) => fx(r.r_ref, 0) + " Ω").join(", ")}인 것처럼 작아지기 때문이다.`, 800),
  q("(3) 이상적인 전압 이득은 두 번째 감은 횟수를 첫 번째 감은 횟수로 나눈 값과 같다. 실제 전압 이득(Vout/Vin)을 어떻게 이상적인 전압 이득과 비교할 수 있는가?"),
  body(`이상적인 전압 이득 N₂/N₁을 이론값으로, 측정한 Vout/Vin을 실험값으로 두고 오차율을 구하면 두 값을 비교할 수 있다. 오차율은 ${C.map((r) => pct(r.gv_err) + "%").join(", ")}였고 실제 전압 이득은 모두 이상적인 값보다 작았다. 부하 저항이 없을 때([실험 1])의 전압비와 비교하면 400회와 800회에서는 부하를 연결한 뒤 전압비가 ${fx(Tc(400).ratio, 3)}에서 ${fx(Cc(400).gv, 3)}로, ${fx(Tc(800).ratio, 3)}에서 ${fx(Cc(800).gv, 3)}로 더 작아졌는데, 전류가 흐르면서 코일 도선 자체의 저항에서 전압이 떨어지기 때문이다.`, 800),
  q("(4) 이상적으로, 변압기는 거의 전력 손실 없이 한 쪽으로부터 다른 쪽으로 교류 전류를 전환시킨다. 우리가 구한 전력 이득(Pout/Pin)을 보고, 저항값의 변화에 따른 전력 이득의 변화에 대해 설명해보라."),
  body(`이번 실험에서는 부하 저항을 100 Ω으로만 측정했기 때문에 저항값에 따른 전력 이득의 변화를 직접 비교하지는 못하였다. 100 Ω에서의 전력 이득은 ${fx(gpMin, 3)}~${fx(gpMax, 3)}로 이상적인 값인 1보다 작았다. 이상적인 변압기라면 저항값과 상관없이 전력 이득이 1이지만, 실제 변압기에서는 코일 도선의 저항에서 I²r만큼, 철심에서 맴돌이 전류와 자기 이력에 의해 에너지가 손실된다. 저항값이 작으면 흐르는 전류가 커져 도선에서의 손실(I²r)이 커지기 때문에 전력 이득이 작아지고, 저항값이 너무 크면 출력 전력 자체가 작아져 철심에서 생기는 손실의 비중이 커지기 때문에 역시 전력 이득이 작아질 것이다. 따라서 전력 이득은 그 사이의 적당한 저항값에서 가장 크게 나타날 것이다.`, 800),

  section("5. 토의"),
  sub1("(1) 실험 검토"),
  body("이번 실험에서는 철심의 구조, 코일의 감은 수, 출력 쪽 회로를 바꿔가며 변압기의 입력과 출력을 측정하였다. 이를 통해 철심이 닫힌 고리를 이룰수록 출력 전압이 커지고, 출력 전압은 두 번째 코일의 감은 수에 비례하며, 전류의 비는 감은 수의 비에 반비례함을 알 수 있었다. 그리고 전압비와 전력 이득이 이상적인 값보다 작게 나와 실제 변압기에서는 에너지 손실이 있다는 것을 확인할 수 있었다.", 800),
  sub1("(2) 오차 분석 및 개선 방안"),
  body("[실험 1]에서 실제 전압비가 세 경우 모두 이상적인 값보다 약 10% 작게 나온 것을 보면 실험과정에서 오차가 발생한 것을 알 수 있다. 이러한 오차의 원인으로는 U자형 철심과 가로대 사이에 작은 틈이 있어 첫 번째 코일의 자기선속 일부가 두 번째 코일을 통과하지 못하고 새어나갔기 때문이며, 철심 내부에 생기는 맴돌이 전류와 코일 도선의 저항으로 에너지가 열로 손실되었을 가능성이 있었다. 또한 [실험 2]에서는 부하 저항을 100 Ω 하나만 사용했기 때문에 저항값에 따른 변화를 확인하지 못하였다. 오차를 개선하기 위해 가로대를 U자형 철심에 단단히 고정해 틈을 줄이고, 얇은 철판을 쌓은 철심을 사용하며, 매뉴얼처럼 여러 저항값에서 측정하는 방법이 존재한다.", 800),
  sub1("(3) 이론을 확인할 다른 실험 제안"),
  body("이 실험에서는 두 코일의 위치를 고정하고 철심 구조와 감은 수를 바꿔가며 실험을 진행하였다. 철심 없이 두 코일 사이의 거리를 바꿔가며 출력 전압을 측정하면 두 코일 사이의 거리에 따라 전달되는 자기선속이 어떻게 달라지는지 확인해볼 수 있을 것이다. 또한 코일 안에서 자석을 다른 속도로 움직이며 유도 전압을 측정하면 자기선속이 빠르게 변할수록 유도 기전력이 커지는 것을 확인해볼 수 있을 것이다.", 800),
  sub1("(4) 응용 가능성"),
  body("변압기의 원리를 이용하면 발전소에서 전압을 높여 송전해 같은 전력을 작은 전류로 보내 송전선에서 열로 손실되는 전력을 줄일 수 있다. 또한 두 코일 사이의 전자기 유도는 휴대폰 무선 충전기에 이용되는데, [실험 1]에서 철심이 없을 때 출력 전압이 작았던 것처럼 두 코일이 가깝고 정확히 맞닿아야 충전이 잘 된다.", 800),

  section("6. 참고 문헌"),
  body("일반물리학실험 매뉴얼, 서강대학교 물리학과"),
];

const doc = new Document({
  creator: "이건",
  title: "7주차 레포트 - 전자기 유도와 변압기 실험",
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 },
      margin: { top: 1701, right: 1440, bottom: 1440, left: 1440 } } },
    children,
  }],
});

const out = path.join(ROOT, "20241501_이건_7주차_레포트.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("wrote", out); });
