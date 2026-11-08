"""build_report.js 가 만든 docx 의 [[CHART:n]] 자리에 Word 차트(편집 가능)를 넣는다.

실행: python3 src/add_charts.py   (node src/build_report.js 다음에)
각 차트에는 엑셀 데이터가 포함되어 Word 에서 '데이터 편집'으로 수정할 수 있다.
"""
import io, json, math, re, shutil, zipfile
from pathlib import Path
from xml.sax.saxutils import escape
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
DOCX = ROOT / "20241501_이건_6주차_레포트.docx"
STYLE = ROOT / "src" / "chart_style"          # Office 기본 차트 스타일/색 파일
R = json.loads((ROOT / "src" / "results.json").read_text(encoding="utf8"))
MU0 = 4 * math.pi * 1e-7
FONT = "맑은 고딕"

# ---------------------------------------------------------------- 데이터
def frange(a, b, step):
    n = int(round((b - a) / step))
    return [round(a + i * step, 6) for i in range(n + 1)]

def wire_chart():
    W = R["wire"]
    xs = [round(d["r"] * 100, 4) for d in W["rows"]]
    ys = [round(d["meas"] * 1e4, 4) for d in W["rows"]]
    tx = frange(0.8, 5.5, 0.1)
    ty = [round(MU0 * W["I"] / (2 * math.pi * x / 100) * 1e4, 4) for x in tx]
    return dict(
        title="직선도선으로부터의 거리에 따른 자기장",
        xtitle="도선으로부터의 거리 r (cm)", ytitle="자기장 B (gauss)",
        xmin=0, xmax=6, ymin=0, ymax=1.4,
        series=[
            dict(name="이론값", x=tx, y=ty, color="accent1", line=True, marker=None),
            dict(name="측정값", x=xs, y=ys, color="accent2", line=True, marker="square",
                 trend="power"),
        ])

def loop_chart():
    I, L = R["loop"]["I"], R["loop"]["loops"]
    tx = frange(0, 8, 0.2)
    ser = []
    for k, col, mk in (("4", "accent1", "circle"), ("3", "accent6", "triangle")):
        Rr = L[k]["R"]
        ty = [round(MU0 * I * Rr**2 / (2 * (Rr**2 + (x / 100) ** 2) ** 1.5) * 1e4, 4) for x in tx]
        ser.append(dict(name=f"이론값 (R={k}cm)", x=tx, y=ty, color=col, line=True, marker=None))
    for k, col, mk in (("4", "accent1", "circle"), ("3", "accent6", "triangle")):
        ser.append(dict(name=f"측정값 (R={k}cm)", x=[0, int(k)],
                        y=[round(L[k]["m0"] * 1e4, 4), round(L[k]["mR"] * 1e4, 4)],
                        color=col, line=False, marker=mk))
    return dict(title="원형도선 중심축 위의 거리에 따른 자기장",
                xtitle="원형도선 중심으로부터의 거리 x (cm)", ytitle="자기장 B (gauss)",
                xmin=0, xmax=8, ymin=0, ymax=0.7, series=ser)

def sol_chart():
    S = R["sol"]
    n, I, Rs, Lc = S["n"], S["I"], S["R"], S["L"]
    tx = frange(-12, 12, 0.25)
    ty = []
    for x in tx:
        a, b = Lc / 2 + x / 100, Lc / 2 - x / 100
        ty.append(round(0.5 * MU0 * n * I * (b / math.hypot(b, Rs) + a / math.hypot(a, Rs)) * 1e3, 4))
    return dict(title="솔레노이드 중심축 위의 거리에 따른 자기장",
                xtitle="코일 중심으로부터의 거리 x (cm)", ytitle="자기장 B (mT)",
                xmin=-12, xmax=12, xunit=4, ymin=0, ymax=25, ycross="min",
                series=[
                    dict(name="이론값", x=tx, y=ty, color="accent1", line=True, marker=None),
                    dict(name="측정값", x=[0, Lc / 2 * 100],
                         y=[round(abs(S["mm"]) * 1e3, 3), round(abs(S["me"]) * 1e3, 3)],
                         color="accent2", line=False, marker="square"),
                ])

# ---------------------------------------------------------------- XML 조각
GRAY = '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>'
def font(sz, b=0):
    return (f'<a:defRPr sz="{sz}" b="{b}" i="0" u="none" strike="noStrike" kern="1200" baseline="0">{GRAY}'
            f'<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/><a:cs typeface="+mn-cs"/></a:defRPr>')
BODY = '<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/>'
NOFILL = '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>'

def rich_title(text, sz, rot=None):
    body = BODY if rot is None else BODY.replace('rot="0"', f'rot="{rot}"').replace('vert="horz"', 'vert="horz"')
    return (f'<c:title><c:tx><c:rich>{body}<a:p><a:pPr>{font(sz)}</a:pPr>'
            f'<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="{sz}"/><a:t>{escape(text)}</a:t></a:r></a:p></c:rich></c:tx>'
            f'<c:overlay val="0"/>{NOFILL}</c:title>')

def txpr(sz, rot="-60000000"):
    body = BODY.replace('rot="0"', f'rot="{rot}"')
    return f'<c:txPr>{body}<a:p><a:pPr>{font(sz)}</a:pPr><a:endParaRPr lang="ko-KR"/></a:p></c:txPr>'

def numref(ref, vals, fmt="General"):
    pts = "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(vals))
    return (f'<c:numRef><c:f>{ref}</c:f><c:numCache><c:formatCode>{fmt}</c:formatCode>'
            f'<c:ptCount val="{len(vals)}"/>{pts}</c:numCache></c:numRef>')

def series_xml(i, s, col_x, col_y):
    n = len(s["x"])
    if s["line"]:
        ln = f'<a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="{s["color"]}"/></a:solidFill><a:round/></a:ln>'
    else:
        ln = '<a:ln w="25400" cap="rnd"><a:noFill/><a:round/></a:ln>'
    if s["marker"]:
        marker = (f'<c:marker><c:symbol val="{s["marker"]}"/><c:size val="7"/><c:spPr>'
                  f'<a:solidFill><a:schemeClr val="{s["color"]}"/></a:solidFill>'
                  f'<a:ln w="9525"><a:solidFill><a:schemeClr val="{s["color"]}"/></a:solidFill></a:ln>'
                  f'<a:effectLst/></c:spPr></c:marker>')
    else:
        marker = '<c:marker><c:symbol val="none"/></c:marker>'
    trend = ""
    if s.get("trend"):
        trend = (f'<c:trendline><c:name>{escape(s["name"])} 추세선</c:name><c:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:schemeClr val="{s["color"]}"/></a:solidFill>'
                 f'<a:prstDash val="sysDot"/></a:ln><a:effectLst/></c:spPr><c:trendlineType val="{s["trend"]}"/>'
                 f'<c:dispRSqr val="0"/><c:dispEq val="1"/><c:trendlineLbl><c:layout><c:manualLayout>'
                 f'<c:x val="0.05"/><c:y val="-0.12"/></c:manualLayout></c:layout>'
                 f'<c:numFmt formatCode="#,##0.000" sourceLinked="0"/>{NOFILL}{txpr(900, "0")}</c:trendlineLbl></c:trendline>')
    xr = f"Sheet1!${col_x}$2:${col_x}${n + 1}"
    yr = f"Sheet1!${col_y}$2:${col_y}${n + 1}"
    return (f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>'
            f'<c:tx><c:strRef><c:f>Sheet1!${col_y}$1</c:f><c:strCache><c:ptCount val="1"/>'
            f'<c:pt idx="0"><c:v>{escape(s["name"])}</c:v></c:pt></c:strCache></c:strRef></c:tx>'
            f'<c:spPr>{ln}<a:effectLst/></c:spPr>{marker}{trend}'
            f'<c:xVal>{numref(xr, s["x"])}</c:xVal><c:yVal>{numref(yr, s["y"])}</c:yVal>'
            f'<c:smooth val="{1 if s["line"] and not s["marker"] else 0}"/></c:ser>')

GRID = ('<c:majorGridlines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1">'
        '<a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr></c:majorGridlines>')
AXLINE = ('<c:spPr><a:noFill/><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1">'
          '<a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>')

def val_ax(ax_id, cross_id, pos, title, vmin, vmax, grid, rot=None, unit=None, crosses="autoZero"):
    return (f'<c:valAx><c:axId val="{ax_id}"/><c:scaling><c:orientation val="minMax"/>'
            f'<c:max val="{vmax}"/><c:min val="{vmin}"/></c:scaling><c:delete val="0"/><c:axPos val="{pos}"/>'
            f'{GRID if grid else ""}{rich_title(title, 1000, rot)}'
            f'<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
            f'<c:tickLblPos val="low"/>{AXLINE}{txpr(900)}<c:crossAx val="{cross_id}"/><c:crosses val="{crosses}"/>'
            f'<c:crossBetween val="midCat"/>{f"<c:majorUnit val={chr(34)}{unit}{chr(34)}/>" if unit else ""}</c:valAx>')

def chart_xml(spec, cols):
    sers = "".join(series_xml(i, s, cx, cy) for i, (s, (cx, cy)) in enumerate(zip(spec["series"], cols)))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<c:date1904 val="0"/><c:lang val="ko-KR"/><c:roundedCorners val="0"/>'
        f'<c:chart>{rich_title(spec["title"], 1400)}<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>'
        f'<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>{sers}'
        '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/>'
        '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
        '<c:axId val="500100"/><c:axId val="500200"/></c:scatterChart>'
        f'{val_ax(500100, 500200, "b", spec["xtitle"], spec["xmin"], spec["xmax"], False, unit=spec.get("xunit"))}'
        f'{val_ax(500200, 500100, "l", spec["ytitle"], spec["ymin"], spec["ymax"], True, "-5400000", crosses=spec.get("ycross", "autoZero"))}'
        f'{NOFILL.replace("c:spPr", "c:spPr")}</c:plotArea>'
        f'<c:legend><c:legendPos val="b"/><c:overlay val="0"/>{NOFILL}{txpr(900, "0")}</c:legend>'
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
        '<c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
        '<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill>'
        '<a:round/></a:ln><a:effectLst/></c:spPr>'
        f'<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:defRPr></a:pPr>'
        '<a:endParaRPr lang="ko-KR"/></a:p></c:txPr>'
        '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>')

# 문서에 테마가 없어도 같은 색이 나오도록 Office 테마 색을 RGB 로 직접 지정
# (5주차 레포트와 같은 Office 2023 테마: accent1 156082, accent2 E97132, accent6 4EA72E)
RGB = {
    '<a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr>': "595959",
    '<a:schemeClr val="tx1"><a:lumMod val="25000"/><a:lumOff val="75000"/></a:schemeClr>': "BFBFBF",
    '<a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr>': "D9D9D9",
    '<a:schemeClr val="bg1"/>': "FFFFFF",
    '<a:schemeClr val="accent1"/>': "156082",
    '<a:schemeClr val="accent2"/>': "E97132",
    '<a:schemeClr val="accent6"/>': "4EA72E",
}
def to_rgb(xml):
    for k, v in RGB.items():
        xml = xml.replace(k, f'<a:srgbClr val="{v}"/>')
    assert "schemeClr" not in xml
    return xml

def workbook(spec):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    cols, c = [], 1
    for s in spec["series"]:
        cx, cy = openpyxl.utils.get_column_letter(c), openpyxl.utils.get_column_letter(c + 1)
        ws.cell(1, c, f"{s['name']} x"); ws.cell(1, c + 1, s["name"])
        for i, (x, y) in enumerate(zip(s["x"], s["y"]), start=2):
            ws.cell(i, c, x); ws.cell(i, c + 1, y)
        cols.append((cx, cy)); c += 3
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue(), cols

def drawing(n, rid, cx=5120640, cy=2926080):
    return (
        f'<w:p><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="60"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{9000 + n}" name="차트 {n}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
        '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"/>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')

# ---------------------------------------------------------------- docx 수정
def main():
    specs = {1: wire_chart(), 2: loop_chart(), 3: sol_chart()}
    zin = zipfile.ZipFile(DOCX)
    files = {i.filename: zin.read(i.filename) for i in zin.infolist() if not i.filename.endswith("/")}
    zin.close()
    doc = files["word/document.xml"].decode("utf8")
    rels = files["word/_rels/document.xml.rels"].decode("utf8")
    ct = files["[Content_Types].xml"].decode("utf8")

    for n, spec in specs.items():
        xlsx, cols = workbook(spec)
        files[f"word/charts/chart{n}.xml"] = to_rgb(chart_xml(spec, cols)).encode("utf8")
        files[f"word/charts/style{n}.xml"] = (STYLE / "style.xml").read_bytes()
        files[f"word/charts/colors{n}.xml"] = (STYLE / "colors.xml").read_bytes()
        files[f"word/embeddings/Microsoft_Excel_Worksheet{n}.xlsx"] = xlsx
        files[f"word/charts/_rels/chart{n}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet{n}.xlsx"/>'
            f'<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style{n}.xml"/>'
            f'<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors{n}.xml"/>'
            '</Relationships>').encode("utf8")
        rid = f"rIdChart{n}"
        rels = rels.replace("</Relationships>",
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart{n}.xml"/></Relationships>')
        ct = ct.replace("</Types>",
            f'<Override PartName="/word/charts/chart{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            f'<Override PartName="/word/charts/style{n}.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>'
            f'<Override PartName="/word/charts/colors{n}.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/></Types>')
        pat = re.compile(r"<w:p>(?:(?!</w:p>).)*?\[\[CHART:%d\]\](?:(?!</w:p>).)*?</w:p>" % n, re.S)
        doc, k = pat.subn(drawing(n, rid), doc)
        assert k == 1, f"placeholder {n} not found ({k})"
    if 'Extension="xlsx"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Default ', 1)

    files["word/document.xml"] = doc.encode("utf8")
    files["word/_rels/document.xml.rels"] = rels.encode("utf8")
    files["[Content_Types].xml"] = ct.encode("utf8")
    tmp = DOCX.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", files.pop("[Content_Types].xml"))
        for name, data in files.items():
            z.writestr(name, data)
    shutil.move(tmp, DOCX)
    print("charts added:", list(specs))

if __name__ == "__main__":
    main()
