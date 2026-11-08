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
  body("직선도선, 원형도선, 솔레노이드 코일에 전류를 흘려 주위에 생기는 자기장의 세기를 자기장 센서로 측정하고, 비오-사바르 법칙과 암페어 법칙으로 구한 이론값과 비교한다. 이를 통해 전류가 만드는 자기장이 도체의 모양과 측정 위치에 따라 어떻게 분포하는지 이해하고, 두 법칙이 실제로 성립하는지 확인하도록 한다."),

  // 2. 데이터 정리
  section("2. 데이터 정리"),
  sub1(`[실험 1] 직선도선 주위의 자기장 측정 (전류 I = ${W.I} A)`),
  note(`※ 센서의 감지점(막대 끝의 흰 점)이 도선 중심에서 ${W.offset * 1000} mm 떨어진 곳에서 측정을 시작하였으므로, 도선으로부터의 거리 r은 센서 이동 거리에 ${W.offset * 1000} mm를 더하여 구하였다. 측정값은 gauss 단위로 얻었으며 1 G = 10⁻⁴ T로 환산하였다.`),
  wireTable,
  ...figure("fig1_wire.png", "[그림 1] 직선도선으로부터의 거리에 따른 자기장 (측정값과 이론값)", 1240, 720),

  sub1(`[실험 2] 원형도선 주위의 자기장 측정 (전류 I = ${R.loop.I} A)`),
  note("※ 실험 2부터 PASCO Capstone 프로그램에서 위치 데이터와 자기장 데이터의 연동이 지연되는 현상이 계속 발생하였다. 센서를 움직여도 거리 변화가 그래프에 반영되지 않고 자기장 값이 일정한 수평선(y = c) 형태로만 기록되어, 연속적인 위치-자기장 그래프를 얻을 수 없었다. 따라서 담당 조교님의 지도에 따라 원형도선은 중심(r = 0)과 반지름 거리(r = R), 솔레노이드는 코일 중앙(x = 0)과 코일 끝(x = l/2)에서 각각 따로 자기장을 측정하여 이론값과 비교하였다."),
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
  note("※ 측정값 일부가 누락되어 코일 상수와 보정계수 k는 일부 역산하여 구하였다."),
  P([t("(1) 코일 상수 측정", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  coilTable,
  note("*코일의 길이는 전체 길이가 아니라 코일이 감겨진 부분만의 길이이다."),
  P([t("(2) 보정계수의 계산, k = ½(cos θ", { bold: true }), t("1", { bold: true, subScript: true }), t(" + cos θ", { bold: true }), t("2", { bold: true, subScript: true }), t(")", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  kTable,
  blank(),
  P([t("(3) 측정값과 이론값 비교", { bold: true })], { indent: { left: 440 }, keepNext: true }),
  solTable,
  note("※ 측정값의 음(−) 부호는 자기장의 방향이 센서의 기준 방향과 반대였음을 뜻하므로, 이론값과는 크기를 비교하였다."),
  ...figure("fig3_solenoid.png", "[그림 3] 솔레노이드 중심축 위의 자기장 (이론 곡선과 측정 지점)", 1240, 720),

  // 3. 결과 분석
  section("3. 결과 분석"),
  body("전하가 움직이면, 즉 전류가 흐르면 주위 공간에 자기장이 만들어진다. 전류요소 i ds가 거리 r만큼 떨어진 점에 만드는 자기장은 비오-사바르 법칙으로 나타낼 수 있고, 이를 도체 전체에 대해 적분하면 전체 자기장을 얻는다."),
  eq([mr("d"), mr("B"), mr("="), frac([mu0()], [mr("4π")]), frac([mr("i d"), mr("s"), mr("×"), mr("r")], [pow([mr("r")], [mr("3")])])], "(비오-사바르 법칙)"),
  body("한편 전류 분포가 원통 대칭이나 평면 대칭처럼 대칭성을 가지면, 닫힌 경로(암페어 고리)를 따라 자기장을 적분한 값이 고리 안을 지나는 알짜 전류에 비례한다는 암페어 법칙으로 자기장을 훨씬 간단히 구할 수 있다. 여기서 μ₀ = 4π × 10⁻⁷ T·m/A는 진공의 투자율이다."),
  eq([mr("∮"), mr("B"), mr("·"), mr("d"), mr("s"), mr("="), mu0(), subm([mr("i")], [mr("enc")])], "(암페어 법칙)"),
  body("이번 실험에서는 두 법칙에서 얻은 다음 식들을 측정값과 비교하였다. (각 식의 유도는 4. 질문 (1)에 정리하였다.)"),
  eq([mr("B"), mr("("), mr("r"), mr(")"), mr("="), frac([mu0(), mr("I")], [mr("2πr")])], null, "직선도선:"),
  eq([mr("B"), mr("("), mr("0"), mr(")"), mr("="), frac([mu0(), mr("I")], [mr("2R")]), mr(",    B"), mr("("), mr("R"), mr(")"), mr("="), frac([mu0(), mr("I")], [mr("4"), sqrt([mr("2")]), mr("R")])], null, "원형도선:"),
  eq([subm([mr("B")], [mr("x=0")]), mr("="), mu0(), mr("n"), mr("I"), frac([mr("l")], [sqrt([pow([mr("l")], [mr("2")]), mr("+4"), pow([mr("R")], [mr("2")])])]), mr(",    "), subm([mr("B")], [mr("x=l/2")]), mr("="), frac([mr("1")], [mr("2")]), mu0(), mr("n"), mr("I"), frac([mr("l")], [sqrt([pow([mr("l")], [mr("2")]), mr("+"), pow([mr("R")], [mr("2")])])])], null, "솔레노이드:"),

  body(`[실험 1]에서는 직선도선에 ${W.I} A의 전류를 흘리고 도선에서 멀어지는 방향으로 센서를 움직이며 자기장을 측정하였다. 전류가 일정하면 이론적으로 B는 거리 r에 반비례하므로 B·r이 일정해야 한다. 측정 결과 r이 약 1 cm에서 5 cm로 커지는 동안 자기장은 ${fx(r[0].meas * 1e4, 3)} G에서 ${fx(r[4].meas * 1e4, 3)} G로 줄어들어, 거리가 멀어질수록 자기장이 약해지는 것을 확인하였다. 측정값으로 구한 B·r은 ${fx(Math.min(...r.map((d) => d.Br)), 2)}~${fx(Math.max(...r.map((d) => d.Br)), 2)} G·cm로 이론값 ${fx(2e-7 * W.I * 1e4 * 100, 2)} G·cm와 비슷하였고, log B와 log r의 관계에서 구한 기울기는 ${fx(W.slope, 2)}로 이론의 −1에 가까웠다. 따라서 직선도선의 자기장이 거리에 반비례한다는 암페어 법칙의 결과가 대체로 성립함을 알 수 있다. 다만 측정값이 모든 지점에서 이론값보다 조금씩 컸고, 오차율은 1 cm에서 ${pct(r[0].errp)}%로 가장 작고 4 cm, 5 cm에서 각각 ${pct(r[3].errp)}%, ${pct(r[4].errp)}%로 커졌다. 멀리 떨어질수록 도선이 만드는 자기장이 0.2~0.3 G 정도로 작아져, 지구 자기장(약 0.5 G)이나 주변 기기가 만드는 자기장의 변화가 상대적으로 크게 작용했기 때문으로 보인다.`),
  body(`[실험 2]에서는 반지름이 4 cm, 3 cm인 원형도선에 ${R.loop.I} A의 전류를 흘리고 중심(r = 0)과 중심축 위로 반지름만큼 떨어진 곳(r = R)에서 자기장을 측정하였다. 중심에서의 오차율은 R = 4 cm일 때 ${pct(L["4"].e0)}%, R = 3 cm일 때 ${pct(L["3"].e0)}%로 작았으며, 반지름이 작은 도선일수록 중심의 자기장이 더 크다는 것(측정 ${fx(L["3"].m0 / L["4"].m0, 3)}배, 이론 ${fx(L["3"].t0 / L["4"].t0, 3)}배)도 확인할 수 있었다. 이론적으로 B(R)/B(0) = 1/(2√2) ≈ 0.354로 반지름과 무관한데, 측정값의 비율은 ${fx(L["4"].ratio_meas, 3)}, ${fx(L["3"].ratio_meas, 3)}으로 이론과 비슷하여 중심축 위에서 거리에 따라 자기장이 줄어드는 비율이 비오-사바르 법칙을 따름을 알 수 있다. 반면 r = R에서의 오차율은 ${pct(L["4"].eR)}%, ${pct(L["3"].eR)}%로 중심보다 컸는데, [그림 2]처럼 r = R 부근은 거리에 따라 자기장이 가장 급하게 변하는 구간이어서 센서 위치가 조금만 어긋나도 측정값이 크게 달라지기 때문이다.`),
  body(`[실험 3]에서는 솔레노이드의 중앙(x = 0)과 끝(x = l/2)에서 자기장을 측정하였다. 무한히 긴 솔레노이드라면 내부 자기장은 μ₀nI = ${fx(S.Binf * 1e3, 2)} mT로 위치와 관계없이 일정해야 하지만, 실제 솔레노이드는 길이가 유한하므로 보정계수 k를 곱한 값을 이론값으로 사용하였다. 측정 결과 중앙에서 ${fx(Math.abs(S.mm) * 1e3, 1)} mT, 끝에서 ${fx(Math.abs(S.me) * 1e3, 1)} mT로 이론값과의 오차율은 각각 ${pct(S.em)}%, ${pct(S.ee)}%였다. 특히 끝과 중앙의 자기장 비율은 측정 ${fx(S.ratio_meas, 3)}, 이론 ${fx(S.ratio_theo, 3)}으로 거의 같았는데, 이는 솔레노이드 끝에서는 한쪽에만 코일이 있어 자기장이 중앙의 약 절반이 된다는 이론을 잘 보여준다. 또한 측정값이 무한 솔레노이드의 값보다 보정된 이론값에 더 가까우므로, 유한한 솔레노이드에는 보정이 필요함을 확인할 수 있었다.`),

  // 4. 질문
  section("4. 질문"),
  P([t("(1) 세가지 도체(직선도선, 원형도선, 솔레노이드)에서 자기장 센서를 움직이면 자기장의 변화는 어떠한가? 암페르 법칙과 비오-사바르 법칙으로부터 나온 공식을 유도해보고 측정값과 비교해보시오.", { bold: true })], { indent: { left: 440, hanging: 0 }, keepNext: true }),
  body("직선도선에서는 센서를 도선에서 멀리 움직일수록 자기장이 거리에 반비례하여 줄어든다. 원형도선에서는 중심에서 자기장이 가장 크고, 중심축을 따라 멀어지면 처음에는 완만하다가 반지름 정도의 거리에서 빠르게 감소한 뒤 멀리서는 거리의 세제곱에 반비례하여 약해진다. 솔레노이드는 코일 내부의 중앙 부근에서 자기장이 거의 일정하게 가장 크고, 끝 부분에 가까워지면 줄어들어 끝에서는 중앙의 약 절반이 되며 코일 밖으로 나가면 급격히 약해진다.", 800),
  body([t("① 직선도선 (암페어 법칙): ", { bold: true }), t("도선을 중심으로 반지름 r인 원을 암페어 고리로 잡으면 대칭성에 의해 고리 위의 자기장은 크기가 같고 방향은 고리의 접선 방향이다. 따라서")], 800),
  eq([mr("∮"), mr("B"), mr("·"), mr("d"), mr("s"), mr("="), mr("B"), mr("("), mr("2πr"), mr(")"), mr("="), mu0(), mr("I"), mr("   ⇒   "), mr("B"), mr("="), frac([mu0(), mr("I")], [mr("2πr")])]),
  body(`같은 결과를 비오-사바르 법칙으로 길이가 L인 도선에 대해 구하면 B = (μ₀I/2πr)·(L/2)/√((L/2)² + r²)이고, L → ∞이면 위 식과 같아진다. 실험에 사용한 도선(L = 40 cm)에서 이 보정값은 r = 1 cm에서 ${fx(r[0].finite, 3)}, r = 5 cm에서 ${fx(r[4].finite, 3)}로 1에 가까워, 측정 범위에서는 무한 도선 식을 써도 3% 이내로 맞다. 측정값은 이론처럼 거리에 반비례하여 감소하였으나(log-log 기울기 ${fx(W.slope, 2)}) 이론값보다 5~34% 크게 측정되었다.`, 800),
  body([t("② 원형도선 (비오-사바르 법칙): ", { bold: true }), t("반지름 R인 원형도선의 중심축 위 거리 x인 점에서, 도선요소 dl은 그 점을 향하는 벡터 r과 항상 수직이고 r² = R² + x²이다. 따라서 dB = μ₀I dl / 4π(R² + x²)이다. 대칭성에 의해 중심축에 수직한 성분은 모두 상쇄되고 중심축 방향 성분(dB sin θ, sin θ = R/√(R² + x²))만 남으므로")], 800),
  eq([mr("B"), mr("("), mr("x"), mr(")"), mr("="), mr("∮"), frac([mu0(), mr("I d"), mr("l")], [mr("4π"), mr("("), pow([mr("R")], [mr("2")]), mr("+"), pow([mr("x")], [mr("2")]), mr(")")]), frac([mr("R")], [sqrt([pow([mr("R")], [mr("2")]), mr("+"), pow([mr("x")], [mr("2")])])]), mr("="), frac([mu0(), mr("I"), pow([mr("R")], [mr("2")])], [mr("2"), pow([mr("("), pow([mr("R")], [mr("2")]), mr("+"), pow([mr("x")], [mr("2")]), mr(")")], [mr("3/2")])])]),
  body(`이고, x = 0이면 B(0) = μ₀I/2R, x = R이면 B(R) = μ₀I/4√2R이다. 측정값은 B(0)에서 이론값과 ${pct(L["4"].e0)}%, ${pct(L["3"].e0)}% 차이로 잘 맞았고, B(R)에서는 12~14% 작게 측정되었지만 B(R)/B(0) 비율은 이론값 0.354에 가까웠다.`, 800),
  body([t("③ 솔레노이드 (암페어 법칙, 비오-사바르 법칙): ", { bold: true }), t("무한히 긴 솔레노이드의 내부 자기장은 축에 평행하고 균일하며 외부 자기장은 0이다. 길이 h의 한 변은 내부에, 다른 변은 외부에 있는 직사각형 암페어 고리를 잡으면 고리를 지나는 전류는 nhI이므로")], 800),
  eq([mr("∮"), mr("B"), mr("·"), mr("d"), mr("s"), mr("="), mr("Bh"), mr("="), mu0(), mr("nhI"), mr("   ⇒   "), mr("B"), mr("="), mu0(), mr("n"), mr("I")]),
  body("길이가 유한한 솔레노이드는 원형도선을 축을 따라 쌓은 것으로 보고 원형도선의 식을 적분한다. 위치 x′에서 두께 dx′인 부분에는 n dx′번 감긴 코일, 즉 전류 nI dx′가 흐르므로, 중심축 위의 점 x에서", 800),
  eq([subm([mr("B")], [mr("x")]), mr("="), mr("∫"), frac([mu0(), mr("n"), mr("I"), pow([mr("R")], [mr("2")]), mr("d"), mr("x′")], [mr("2"), pow([mr("("), pow([mr("R")], [mr("2")]), mr("+"), pow([mr("("), mr("x′−x"), mr(")")], [mr("2")]), mr(")")], [mr("3/2")])]), mr("="), frac([mr("1")], [mr("2")]), mu0(), mr("n"), mr("I"), mr("("), frac([mr("b")], [sqrt([pow([mr("b")], [mr("2")]), mr("+"), pow([mr("R")], [mr("2")])])]), mr("+"), frac([mr("a")], [sqrt([pow([mr("a")], [mr("2")]), mr("+"), pow([mr("R")], [mr("2")])])]), mr(")"), mr("="), frac([mr("1")], [mr("2")]), mu0(), mr("n"), mr("I"), mr("("), mr("cos"), subm([mr("θ")], [mr("1")]), mr("+"), mr("cos"), subm([mr("θ")], [mr("2")]), mr(")")]),
  body(`가 된다. 여기서 a, b는 점 x에서 코일 양 끝까지의 거리이다. 중앙(a = b = l/2)에서 k = l/√(l² + 4R²) = ${fx(S.km, 3)}, 끝(a = l, b = 0)에서 k = l/2√(l² + R²) = ${fx(S.ke, 3)}이며, l ≫ R이면 중앙은 μ₀nI, 끝은 그 절반이 된다. 측정값은 보정된 이론값과 ${pct(S.em)}%, ${pct(S.ee)}% 차이로 잘 맞았다.`, 800),

  P([t("(2) 자기장의 세기로부터 도체 주위의 자기장 분포를 자기장선으로 그려보시오. (자기장선의 특징은 다음과 같다. 1. 자기장선 위 한 점에서의 접선방향은 그 점에서의 자기장의 방향이다. 2. 자기장선이 촘촘한 지역에서 자기장이 세고, 성근 지역에서 자기장이 약하다.)", { bold: true })], { indent: { left: 440 }, spacing: { before: 240, after: 120 }, keepNext: true }),
  body("앞에서 구한 식으로 세 도체 주위의 자기장을 계산하고, 이웃한 자기장선 사이를 지나는 자기 선속이 같도록 자기장선을 그리면 [그림 4]와 같다. 이렇게 그리면 자기장이 센 곳일수록 선이 촘촘해진다. 전류의 방향은 지면에서 나오는 방향(●)과 들어가는 방향(×)으로 표시하였고, 자기장의 방향은 오른손 법칙을 따른다.", 800),
  ...figure("fig4_fieldlines.png", "[그림 4] 직선도선(a), 원형도선(b), 솔레노이드(c) 주위의 자기장선", 2100, 740, 6.2),
  body("(a) 직선도선의 자기장선은 도선을 중심으로 하는 동심원이며, 자기장이 거리에 반비례하므로 도선에 가까울수록 원 사이 간격이 좁고 멀어질수록 넓어진다. (b) 원형도선에서는 도선 가까이에서는 도선을 감싸는 작은 고리 모양이고, 도선 안쪽에서는 자기장선이 모두 같은 방향으로 고리를 통과한 뒤 바깥으로 돌아 나간다. 중심에서는 선이 비교적 촘촘하고 중심축을 따라 멀어지면 퍼지면서 성겨진다. (c) 솔레노이드 내부에서는 자기장선이 축에 거의 평행하고 간격이 일정하여 균일한 자기장을 이루며, 양 끝에서는 선이 퍼져 나가고 바깥에서는 매우 성기게 되돌아오므로 외부 자기장은 약하다. 이는 막대자석의 자기장선 모양과 비슷하다.", 800),

  // 5. 토의
  section("5. 토의"),
  sub1("(1) 실험 검토"),
  body("이번 실험에서는 직선도선, 원형도선, 솔레노이드에 전류를 흘리고 위치에 따른 자기장의 세기를 측정하여 비오-사바르 법칙과 암페어 법칙에서 얻은 이론값과 비교하였다. 직선도선에서는 자기장이 거리에 반비례하여 줄어드는 경향을, 원형도선에서는 반지름이 작을수록 중심의 자기장이 커지고 중심축을 따라 멀어지면 이론과 비슷한 비율로 줄어드는 것을, 솔레노이드에서는 끝의 자기장이 중앙의 약 절반이 되는 것을 확인하였다. 원형도선 중심과 솔레노이드의 측정값은 이론값과 10% 이내로 잘 맞았으며, 직선도선의 먼 거리와 원형도선의 r = R 지점에서 상대적으로 오차가 컸다. 다만 실험 2와 3은 프로그램 문제로 연속적인 그래프 대신 두 지점만 측정하였기 때문에 전체 분포는 이론 곡선과 측정 지점을 비교하는 방식으로 확인하였다.", 800),
  sub1("(2) 오차 분석 및 개선 방안"),
  body("첫째, 실험 2부터 Capstone에서 위치 데이터와 자기장 데이터가 제대로 연동되지 않아 연속 측정을 하지 못하였다. 이 때문에 측정 지점이 두 곳뿐이라 거리에 따른 분포 전체를 실험적으로 확인하지 못했다. 로터리 모션 센서 연결과 샘플링 설정을 미리 점검하고, 문제가 생기면 자를 이용해 여러 고정 위치에서 하나씩 측정하는 방법으로 보완할 수 있다.", 800),
  body("둘째, 센서 위치를 손으로 맞추었기 때문에 감지점(흰 점)이 원형도선의 중심축이나 도선의 수직 방향과 정확히 일치하지 않았을 수 있다. 자기장이 위치에 따라 빠르게 변하는 원형도선의 r = R 지점에서 오차가 컸던 것이 그 예이다. 센서를 고정하는 지그나 눈금이 있는 이동 장치를 사용하면 위치 오차를 줄일 수 있다.", 800),
  body("셋째, 전원장치의 전류가 설정값과 정확히 같지 않거나 측정 중 도선이 가열되어 저항이 커지면서 전류가 변했을 수 있다. 특히 솔레노이드는 전류를 약 0.67 A로 읽어 이론값을 계산하였으므로 전류값의 작은 차이가 그대로 오차가 된다. 직렬로 전류계를 연결해 측정하는 동안 전류를 함께 기록하고, 이론값을 실제 전류로 계산하면 개선할 수 있다.", 800),
  body("넷째, 측정 시작 전에 한 번만 영점을 맞추었기 때문에 지구 자기장과 실험실 내 전자기기, 철제 책상 등이 만드는 자기장의 영향이 센서 방향이나 위치가 바뀔 때마다 달라졌을 수 있다. 직선도선의 먼 거리처럼 측정하는 자기장이 작을수록 이 영향이 상대적으로 커진다. 측정 지점마다 전류를 끈 상태의 값을 따로 측정해 빼 주고, 주변의 자성 물체를 치우면 오차를 줄일 수 있다. 마지막으로 솔레노이드의 코일 상수와 보정계수를 일부 역산하여 구했으므로, 버니어 캘리퍼스로 반경과 감긴 부분의 길이를 직접 측정해야 이론값을 더 정확히 구할 수 있다.", 800),
  sub1("(3) 이론을 확인할 다른 실험 제안"),
  body("이번 실험에서는 전류를 고정하고 위치에 따른 자기장을 측정하였으므로, 반대로 측정 위치를 고정하고 전류를 바꾸어 가며 자기장이 전류에 비례하는지 확인하는 실험을 할 수 있다. 또한 반지름이 같은 원형도선 두 개를 반지름만큼 떨어뜨려 놓은 헬름홀츠 코일을 만들어 두 코일 사이의 자기장이 균일해지는지 측정하면 비오-사바르 법칙과 자기장의 중첩을 함께 확인할 수 있다. 원형도선이나 솔레노이드의 중심축이 아닌 곳에서 자기장을 측정하여 [그림 4]의 자기장선 분포와 비교하는 실험도 가능하다.", 800),
  sub1("(4) 응용 가능성"),
  body("솔레노이드 내부에 균일한 자기장이 생기는 원리는 전자석, 릴레이, MRI 장비의 강하고 균일한 자기장을 만드는 데 이용된다. 또 직선도선 주위의 자기장이 전류에 비례한다는 점을 이용하면 도선을 끊지 않고 주위의 자기장만 측정하여 전류를 알아내는 클램프 미터나 홀 센서 방식의 전류 센서를 만들 수 있다.", 800),

  // 6. 참고 문헌
  section("6. 참고 문헌"),
  body("일반물리학실험 매뉴얼, 서강대학교 물리학과"),
  body("D. Halliday, R. Resnick, J. Walker, 일반물리학 (Fundamentals of Physics), 29장 전류에 의한 자기장"),
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
