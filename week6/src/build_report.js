// 6주차(실험 5) 암페어 법칙 결과 레포트 생성 스크립트
// 실행: python3 src/analysis.py && node src/build_report.js
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
    { indent: { left: 440 }, spacing: { after: 100 } });
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
    children: [new Paragraph({ children: runs, alignment: AlignmentType.CENTER, keepNext: !!o.head })],
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

function figure(file, caption, wPx, hPx, widthIn = 5.6) {
  const w = Math.round(widthIn * 96), h = Math.round((w * hPx) / wPx);
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 160, after: 60 }, keepNext: true,
      children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(ROOT, "fig", file)),
        transformation: { width: w, height: h },
        altText: { title: caption, description: caption, name: file } })],
    }),
    P([t(caption, { size: 19, bold: true })], { alignment: AlignmentType.CENTER, spacing: { after: 240 } }),
  ];
}

// ---------- 데이터 ----------
const W = R.wire, L = R.loop.loops, S = R.sol;
const pct = (x) => fx(x, 2);

// [실험 1] 표
const w1 = [1330, 1430, 1560, 1560, 1560, 1270];
const wireTable = table(w1, [
  new TableRow({ tableHeader: true, children: [
    cell("센서 이동 거리 (m)", w1[0], { head: true }), cell("도선으로부터 거리 r (m)", w1[1], { head: true }),
    cell("측정값 (T)", w1[2], { head: true }), cell("이론값 (T)", w1[3], { head: true }),
    cell("오차 (T)", w1[4], { head: true }), cell("오차율 (%)", w1[5], { head: true })] }),
  ...W.rows.map((d) => new TableRow({ children: [
    cell(fx(d.x, 6), w1[0]), cell(fx(d.r, 6), w1[1]), cell(sci(d.meas), w1[2]),
    cell(sci(d.theo), w1[3]), cell(sci(d.err), w1[4]), cell(pct(d.errp), w1[5])] })),
]);

// [실험 2] 표
const w2 = [1504, 1504, 1504, 1504, 1504, 1506];
function loopTable(k) {
  const d = L[k];
  return table(w2, [
    new TableRow({ children: [
      cell("r = 0 (원형도선의 중심점)", w2[0] * 3, { head: true, span: 3 }),
      cell("r = R (반지름만큼 떨어진 거리)", w2[3] * 3, { head: true, span: 3 })] }),
    new TableRow({ children: ["측정값 (T)", "이론값 (T)", "오차율 (%)", "측정값 (T)", "이론값 (T)", "오차율 (%)"]
      .map((h, i) => cell(h, w2[i], { head: true })) }),
    new TableRow({ children: [
      cell(sci(d.m0, 2), w2[0]), cell(sci(d.t0), w2[1]), cell(pct(d.e0), w2[2]),
      cell(sci(d.mR, 2), w2[3]), cell(sci(d.tR), w2[4]), cell(pct(d.eR), w2[5])] }),
  ]);
}
const w2b = [2200, 2200, 2200, 2426];
const ratioTable = table(w2b, [
  new TableRow({ children: [cell("반지름 R", w2b[0], { head: true }),
    cell("측정 B(R)/B(0)", w2b[1], { head: true }), cell("이론 B(R)/B(0)", w2b[2], { head: true }),
    cell("차이 (%)", w2b[3], { head: true })] }),
  ...["4", "3"].map((k) => new TableRow({ children: [
    cell(`${k} cm`, w2b[0]), cell(fx(L[k].ratio_meas, 3), w2b[1]), cell(fx(L[k].ratio_theo, 3), w2b[2]),
    cell(pct(Math.abs(L[k].ratio_meas - L[k].ratio_theo) / L[k].ratio_theo * 100), w2b[3])] })),
]);

// [실험 3] 표
const w3a = [3008, 3008, 3010];
const coilTable = table(w3a, [
  new TableRow({ children: [cell("반경 (R)", w3a[0], { head: true }), cell("길이 (l)", w3a[1], { head: true }),
    cell("감은 수 (n = N/l)", w3a[2], { head: true })] }),
  new TableRow({ children: [cell(`${S.R * 1000} mm`, w3a[0]), cell(`${S.L * 1000} mm`, w3a[1]),
    cell([t(`${S.N}회 / ${S.L} m = `, { size: 19 }), ...sciRuns(S.n, 3, { size: 19 }), t(" 회/m", { size: 19 })], w3a[2])] }),
]);
const w3b = [4513, 4513];
const kTable = table(w3b, [
  new TableRow({ children: [cell([t("코일 중앙 k", { size: 19, bold: true }), t("m", { size: 19, bold: true, subScript: true })], w3b[0], { head: true }),
    cell([t("코일 끝 k", { size: 19, bold: true }), t("e", { size: 19, bold: true, subScript: true })], w3b[1], { head: true })] }),
  new TableRow({ children: [cell(fx(S.km, 4), w3b[0]), cell(fx(S.ke, 4), w3b[1])] }),
]);
const w3c = [800, 1600, 1600, 1600, 1600, 1826];
const solTable = table(w3c, [
  new TableRow({ children: [
    cell("측정", w3c[0], { head: true }),
    cell([t("측정값(중앙) B", { size: 19, bold: true }), t("x=0", { size: 19, bold: true, subScript: true }), t(" (T)", { size: 19, bold: true })], w3c[1], { head: true }),
    cell([t("측정값(끝) B", { size: 19, bold: true }), t("x=l/2", { size: 19, bold: true, subScript: true }), t(" (T)", { size: 19, bold: true })], w3c[2], { head: true }),
    cell("이론값(중앙) (T)", w3c[3], { head: true }), cell("이론값(끝) (T)", w3c[4], { head: true }),
    cell("오차율 (%) (중앙)/(끝)", w3c[5], { head: true })] }),
  new TableRow({ children: [
    cell("1", w3c[0]), cell(sci(S.mm, 2), w3c[1]), cell(sci(S.me, 2), w3c[2]),
    cell(sci(S.tm), w3c[3]), cell(sci(S.te), w3c[4]), cell(`${pct(S.em)} / ${pct(S.ee)}`, w3c[5])] }),
]);

// ---------- 본문 ----------
const r = W.rows;
const children = [
  // 표지
  P([t("6주차 레포트", { size: 20 })]),
  ...Array(6).fill(0).map(blank),
  P([t("암페어 법칙 실험", { size: 52, bold: true })], { alignment: AlignmentType.CENTER }),
  ...Array(10).fill(0).map(blank),
  P("과목명: 일반물리실험II"), P("담당 조교: 이승윤 조교님"), P("학번: 20241501"), P("이름: 이건"),
  new Paragraph({ children: [new PageBreak()] }),

  P([t("실험 5. 암페어 법칙 (도선 주위의 자기장 측정)", { size: 32, bold: true })], { spacing: { after: 240 } }),

  // 1. 실험 목적
  section("1. 실험 목적"),
  body("직선도선, 원형도선, 솔레노이드 코일에 전류를 흘려 도선 주위에 생기는 자기장의 세기를 측정하고 이론적인 자기장 분포와 비교해보도록 한다. 이를 통해 암페어 법칙과 비오-사바르 법칙이 성립하는지 확인하고 전류가 만드는 자기장에 대해 충분히 이해하도록 한다."),

  // 2. 데이터 정리
  section("2. 데이터 정리"),
  sub1(`[실험 1] 직선도선 주위의 자기장 측정 (전류 I = ${W.I} A)`),
  note(`(센서의 흰 점이 도선 중심에서 ${W.offset * 1000} mm 떨어진 곳에서부터 측정을 시작했기 때문에 도선으로부터의 거리 r은 측정 거리에 ${W.offset * 1000} mm를 더해 구하였다. 1 G = 10⁻⁴ T)`),
  wireTable,
  ...figure("fig1_wire.png", "[그림 1] 직선도선으로부터의 거리에 따른 자기장 (측정값과 이론값)", 1240, 720),

  sub1(`[실험 2] 원형도선 주위의 자기장 측정 (전류 I = ${R.loop.I} A)`),
  P([t("(1) R = 4 cm", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  loopTable("4"),
  blank(),
  P([t("(2) R = 3 cm", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  loopTable("3"),
  blank(),
  P([t("(3) 중심 대비 반지름 거리에서의 자기장 비율", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  ratioTable,
  ...figure("fig2_loop.png", "[그림 2] 원형도선 중심축 위의 자기장 (이론 곡선과 측정 지점)", 1240, 720),

  sub1(`[실험 3] 솔레노이드 중심축에서의 자기장 측정 (전류 I ≈ ${S.I} A)`),
  note("(측정값 일부가 누락되어 코일 상수와 k는 일부 역산하여 구하였다.)"),
  P([t("(1) 코일 상수 측정", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  coilTable,
  note("*코일의 길이는 전체 길이가 아니라 코일이 감겨진 부분만의 길이이다."),
  P([t("(2) 보정계수의 계산, k = ½(cos θ", { bold: true }), t("1", { bold: true, subScript: true }), t(" + cos θ", { bold: true }), t("2", { bold: true, subScript: true }), t(")", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  kTable,
  blank(),
  P([t("(3) 측정값과 이론값 비교", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  solTable,
  note("(측정값의 − 부호는 자기장의 방향이 센서의 기준 방향과 반대라는 의미이기 때문에 이론값과는 크기만 비교하였다.)"),
  ...figure("fig3_solenoid.png", "[그림 3] 솔레노이드 중심축 위의 자기장 (이론 곡선과 측정 지점)", 1240, 720),

  // 3. 결과 분석
  section("3. 결과 분석"),
  body("전류가 흐르는 도선 주위에는 자기장이 만들어진다. 전류요소 i ds가 거리 r만큼 떨어진 점에 만드는 자기장은 비오-사바르 법칙으로 나타낼 수 있고, 이를 도선 전체에 대해 적분하면 전체 자기장을 구할 수 있다."),
  eq([mr("d"), mr("B"), mr("="), frac([mu0()], [mr("4π")]), frac([mr("i d"), mr("s"), mr("×"), mr("r")], [pow([mr("r")], [mr("3")])])], "(비오-사바르 법칙)"),
  body("그리고 전류 분포가 대칭성을 가지고 있을 때에는 암페어 법칙을 이용해 자기장을 더 간단하게 구할 수 있다. 여기서 μ₀ = 4π × 10⁻⁷ T·m/A는 진공의 투자율이다."),
  eq([mr("∮"), mr("B"), mr("·"), mr("d"), mr("s"), mr("="), mu0(), subm([mr("i")], [mr("enc")])], "(암페어 법칙)"),
  body("이번 실험에서는 두 법칙으로 구한 아래의 식을 이용해 이론값을 구하고 측정값과 비교한다."),
  eq([mr("B"), mr("("), mr("r"), mr(")"), mr("="), frac([mu0(), mr("I")], [mr("2πr")])], null, "직선도선:"),
  eq([mr("B"), mr("("), mr("0"), mr(")"), mr("="), frac([mu0(), mr("I")], [mr("2R")]), mr(",    B"), mr("("), mr("R"), mr(")"), mr("="), frac([mu0(), mr("I")], [mr("4"), sqrt([mr("2")]), mr("R")])], null, "원형도선:"),
  eq([subm([mr("B")], [mr("x=0")]), mr("="), mu0(), mr("n"), mr("I"), frac([mr("l")], [sqrt([pow([mr("l")], [mr("2")]), mr("+4"), pow([mr("R")], [mr("2")])])]), mr(",    "), subm([mr("B")], [mr("x=l/2")]), mr("="), frac([mr("1")], [mr("2")]), mu0(), mr("n"), mr("I"), frac([mr("l")], [sqrt([pow([mr("l")], [mr("2")]), mr("+"), pow([mr("R")], [mr("2")])])])], null, "솔레노이드:"),

  body(`[실험 1]에서는 직선도선 주위의 자기장을 측정하였다. 전류를 ${W.I} A로 고정하면 이론적으로 B(r) = μ₀I/2πr에 따라 자기장과 거리 r은 반비례관계이다. 실험 결과 거리가 약 1 cm에서 5 cm로 멀어지는 동안 자기장이 ${fx(r[0].meas * 1e4, 3)} G에서 ${fx(r[4].meas * 1e4, 3)} G로 줄어든 것을 보면 거리가 멀어질수록 자기장이 약해지는 것을 확인할 수 있다. 그래프 또한 이론 곡선과 비슷한 분수함수 형태인 것을 보면 자기장과 거리가 반비례관계임을 알 수 있다. 다만 측정값이 모든 지점에서 이론값보다 조금씩 컸고 거리가 먼 4 cm, 5 cm에서는 오차율이 각각 ${pct(r[3].errp)}%, ${pct(r[4].errp)}%로 커졌다. 다음으로 [실험 2]에서는 반지름이 4 cm, 3 cm인 원형도선의 중심(r = 0)과 반지름 거리(r = R)에서 자기장을 측정하였다. 중심에서는 오차율이 ${pct(L["4"].e0)}%, ${pct(L["3"].e0)}%로 이론값과 잘 맞았고, 반지름이 작은 원형도선의 자기장이 더 크게 측정된 것을 보면 B(0) = μ₀I/2R에 따라 자기장과 반지름이 반비례관계임을 알 수 있다. 또한 이론적으로 B(R)/B(0)은 약 0.354인데 측정값은 ${fx(L["4"].ratio_meas, 3)}, ${fx(L["3"].ratio_meas, 3)}으로 비슷하게 나와 중심에서 멀어질 때 자기장이 줄어드는 비율도 이론과 부합함을 확인할 수 있다. 마지막으로 [실험 3]에서는 솔레노이드의 중앙(x = 0)과 끝(x = l/2)에서 자기장을 측정하였다. 무한히 긴 솔레노이드라면 내부 자기장이 μ₀nI로 일정하지만 실제 솔레노이드는 길이가 유한하기 때문에 보정계수 k를 곱한 값을 이론값으로 사용하였다. 측정 결과 중앙에서 ${fx(Math.abs(S.mm) * 1e3, 1)} mT, 끝에서 ${fx(Math.abs(S.me) * 1e3, 1)} mT로 오차율이 각각 ${pct(S.em)}%, ${pct(S.ee)}%였고, 끝의 자기장이 중앙의 약 절반인 것을 보면 이론과 부합함을 알 수 있다. [실험 3]을 통해 길이가 유한한 솔레노이드에서는 보정을 해주어야 한다는 것을 알 수 있다.`),

  // 4. 질문
  section("4. 질문"),
  P([t("(1) 세가지 도체(직선도선, 원형도선, 솔레노이드)에서 자기장 센서를 움직이면 자기장의 변화는 어떠한가? 암페르 법칙과 비오-사바르 법칙으로부터 나온 공식을 유도해보고 측정값과 비교해보시오.", { bold: true })], { indent: { left: 440, hanging: 0 }, keepNext: true }),
  body("실험결과를 살펴보면 직선도선은 도선에서 멀어질수록 자기장이 거리에 반비례하여 감소하였다. 원형도선은 중심에서 자기장이 가장 크고 중심축을 따라 멀어질수록 자기장이 감소하였다. 솔레노이드는 코일 중앙에서 자기장이 가장 크고 거의 일정하며 끝으로 갈수록 감소해 끝에서는 중앙의 약 절반이 되었다. 각 도체의 공식은 다음과 같이 유도할 수 있다.", 800),
  body([t("① 직선도선 (암페어 법칙): ", { bold: true }), t("도선을 중심으로 반지름이 r인 원을 암페어 고리로 잡으면 고리 위의 자기장은 크기가 모두 같고 방향은 고리의 접선 방향이다. 따라서")], 800),
  eq([mr("∮"), mr("B"), mr("·"), mr("d"), mr("s"), mr("="), mr("B"), mr("("), mr("2πr"), mr(")"), mr("="), mu0(), mr("I"), mr("   ⇒   "), mr("B"), mr("="), frac([mu0(), mr("I")], [mr("2πr")])]),
  body(`이 된다. 비오-사바르 법칙으로 길이가 40 cm인 도선을 계산해도 측정한 거리(1~5 cm)에서는 무한 도선과의 차이가 3% 이내이기 때문에 무한 도선의 식을 사용해도 된다. 실험결과 측정값은 이론처럼 거리에 반비례하여 감소하였지만 이론값보다 5~34% 정도 크게 측정되었다.`, 800),
  body([t("② 원형도선 (비오-사바르 법칙): ", { bold: true }), t("원형도선의 중심축 위에서 x만큼 떨어진 점에서 도선요소 dl이 만드는 자기장은 dB = μ₀I dl / 4π(R² + x²)이다. 대칭성에 의해 중심축에 수직한 성분은 서로 상쇄되고 중심축 방향 성분만 남기 때문에")], 800),
  eq([mr("B"), mr("("), mr("x"), mr(")"), mr("="), mr("∮"), frac([mu0(), mr("I d"), mr("l")], [mr("4π"), mr("("), pow([mr("R")], [mr("2")]), mr("+"), pow([mr("x")], [mr("2")]), mr(")")]), frac([mr("R")], [sqrt([pow([mr("R")], [mr("2")]), mr("+"), pow([mr("x")], [mr("2")])])]), mr("="), frac([mu0(), mr("I"), pow([mr("R")], [mr("2")])], [mr("2"), pow([mr("("), pow([mr("R")], [mr("2")]), mr("+"), pow([mr("x")], [mr("2")]), mr(")")], [mr("3/2")])])]),
  body(`이 된다. 그러므로 B(0) = μ₀I/2R, B(R) = μ₀I/4√2R이다. 실험결과 중심에서는 오차율이 ${pct(L["4"].e0)}%, ${pct(L["3"].e0)}%로 이론값과 잘 맞았고 반지름 거리에서는 12~14% 정도 작게 측정되었다.`, 800),
  body([t("③ 솔레노이드 (암페어 법칙, 비오-사바르 법칙): ", { bold: true }), t("무한히 긴 솔레노이드는 내부 자기장이 균일하고 외부 자기장은 0이다. 길이가 h인 직사각형 암페어 고리를 잡으면 고리를 지나는 전류가 nhI이기 때문에")], 800),
  eq([mr("∮"), mr("B"), mr("·"), mr("d"), mr("s"), mr("="), mr("Bh"), mr("="), mu0(), mr("nhI"), mr("   ⇒   "), mr("B"), mr("="), mu0(), mr("n"), mr("I")]),
  body("이 된다. 길이가 유한한 솔레노이드는 원형도선을 여러 개 쌓은 것으로 생각하고 원형도선의 식을 적분해 구할 수 있다. 위치 x′에서 두께 dx′인 부분에 흐르는 전류는 nI dx′이므로", 800),
  eq([subm([mr("B")], [mr("x")]), mr("="), mr("∫"), frac([mu0(), mr("n"), mr("I"), pow([mr("R")], [mr("2")]), mr("d"), mr("x′")], [mr("2"), pow([mr("("), pow([mr("R")], [mr("2")]), mr("+"), pow([mr("("), mr("x′−x"), mr(")")], [mr("2")]), mr(")")], [mr("3/2")])]), mr("="), frac([mr("1")], [mr("2")]), mu0(), mr("n"), mr("I"), mr("("), frac([mr("b")], [sqrt([pow([mr("b")], [mr("2")]), mr("+"), pow([mr("R")], [mr("2")])])]), mr("+"), frac([mr("a")], [sqrt([pow([mr("a")], [mr("2")]), mr("+"), pow([mr("R")], [mr("2")])])]), mr(")"), mr("="), frac([mr("1")], [mr("2")]), mu0(), mr("n"), mr("I"), mr("("), mr("cos"), subm([mr("θ")], [mr("1")]), mr("+"), mr("cos"), subm([mr("θ")], [mr("2")]), mr(")")]),
  body(`가 된다. (a, b는 점 x에서 코일 양 끝까지의 거리) 이 식에 따라 중앙의 k는 ${fx(S.km, 3)}, 끝의 k는 ${fx(S.ke, 3)}이고, 실험결과 측정값은 이론값과 ${pct(S.em)}%, ${pct(S.ee)}% 차이로 잘 맞았다.`, 800),

  P([t("(2) 자기장의 세기로부터 도체 주위의 자기장 분포를 자기장선으로 그려보시오. (자기장선의 특징은 다음과 같다. 1. 자기장선 위 한 점에서의 접선방향은 그 점에서의 자기장의 방향이다. 2. 자기장선이 촘촘한 지역에서 자기장이 세고, 성근 지역에서 자기장이 약하다.)", { bold: true })], { indent: { left: 440 }, spacing: { before: 240, after: 120 }, keepNext: true }),
  body("자기장선 사이의 간격이 자기장의 세기를 나타내도록 세 도체 주위의 자기장선을 그리면 [그림 4]와 같다. (●는 전류가 지면에서 나오는 방향, ×는 들어가는 방향이다.)", 800),
  ...figure("fig4_fieldlines.png", "[그림 4] 직선도선(a), 원형도선(b), 솔레노이드(c) 주위의 자기장선", 2100, 740, 6.2),
  body("직선도선의 자기장선은 도선을 중심으로 하는 동심원 모양이고 도선에서 멀어질수록 자기장이 약해지기 때문에 원 사이의 간격이 넓어진다. 원형도선은 도선 가까이에서 도선을 감싸는 고리 모양이고 도선 안쪽에서는 모든 자기장선이 같은 방향으로 통과하게 된다. 솔레노이드는 내부에서 자기장선이 축과 평행하고 간격이 일정하기 때문에 균일한 자기장을 만들고, 바깥에서는 자기장선이 성기기 때문에 자기장이 약하다. 이는 막대자석의 자기장선과 비슷한 모양이다.", 800),

  // 5. 토의
  section("5. 토의"),
  sub1("(1) 실험 검토"),
  body("이번 실험에서는 직선도선, 원형도선, 솔레노이드에 전류를 흘려 위치에 따른 자기장의 세기를 측정하고 이론값과 비교하였다. 이를 통해 직선도선의 자기장은 거리에 반비례하고, 원형도선은 반지름이 작을수록 중심의 자기장이 커지며, 솔레노이드는 끝의 자기장이 중앙의 약 절반이 된다는 것을 알 수 있었다. 그리고 원형도선의 중심과 솔레노이드에서는 오차율이 10% 이내로 이론과 잘 맞았지만 직선도선의 먼 거리와 원형도선의 반지름 거리에서는 오차가 비교적 크게 나타났다.", 800),
  sub1("(2) 오차 분석 및 개선 방안"),
  body("측정값과 이론값 사이에 최대 30% 이상의 오차율이 나타난 것을 보면 실험과정에서 오차가 발생한 것을 알 수 있다. 이러한 오차의 원인으로는 센서의 위치를 손으로 직접 맞췄기 때문에 측정 지점과 정확히 일치하지 않았을 가능성이 있었으며, 특히 자기장이 급하게 변하는 원형도선의 반지름 거리에서는 위치가 조금만 달라져도 측정값이 크게 달라질 수 있었다. 또한 처음에 한 번만 영점을 맞췄기 때문에 지구 자기장이나 주변 전자기기의 영향을 받았을 수 있고 전원장치의 전류가 설정값과 달랐을 가능성도 있다. 그리고 솔레노이드의 코일 상수 일부를 역산했기 때문에 이론값에도 오차가 포함되었다. 오차를 개선하기 위해 센서를 고정 장치로 고정해 정확한 위치에서 측정하고 전류계로 실제 전류를 함께 측정하는 방법이 존재한다.", 800),
  sub1("(3) 이론을 확인할 다른 실험 제안"),
  body("이 실험에서는 전류를 고정하고 위치에 따라 자기장의 세기를 측정하는 실험을 진행하였다. 반대로 위치를 고정하고 전류의 세기를 바꿔가며 자기장을 측정하거나, 원형도선 두 개를 반지름만큼 떨어뜨려 놓은 헬름홀츠 코일을 이용해 코일 사이의 자기장을 측정하는 실험을 진행해볼 수 있을 것이다.", 800),
  sub1("(4) 응용 가능성"),
  body("솔레노이드 내부에 균일한 자기장이 생기는 원리를 이용하면 전자석이나 MRI 장비를 만들 수 있고, 도선 주위의 자기장이 전류에 비례하는 원리를 이용하면 도선을 끊지 않고 전류를 측정하는 전류 센서를 만들 수 있다.", 800),

  // 6. 참고 문헌
  section("6. 참고 문헌"),
  body("일반물리학실험 매뉴얼, 서강대학교 물리학과"),
];

const doc = new Document({
  creator: "이건",
  title: "6주차 레포트 - 암페어 법칙 실험",
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 },
      margin: { top: 1701, right: 1440, bottom: 1440, left: 1440 } } },
    children,
  }],
});

const out = path.join(ROOT, "20241501_이건_6주차_레포트.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("wrote", out); });
