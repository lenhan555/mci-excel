"""Thêm PivotTable (và Slicer, Timeline, PivotChart) vào file .xlsx do openpyxl tạo, bằng cách ghi thẳng OOXML.

openpyxl không tạo được PivotTable mới; Excel cho Mac cũng không tạo được qua AppleScript. Module này ghi phần
khai báo tối thiểu (cache rỗng, refreshOnLoad="1"), giống cách EPPlus làm: khi mở file, Excel tự Refresh và dựng
lại toàn bộ Pivot. Vì vậy sau khi dựng PHẢI mở file bằng Excel và lưu lại một lần (excel_resave), để Pivot có số
liệu thật khi người khác xem trước file, và để kiểm tra Excel đọc được file (không hiện hộp thoại sửa lỗi).

    from pivot_xml import Cache, Pivot, inject, excel_resave
"""
import os, re, subprocess, zipfile
from dataclasses import dataclass, field
from xml.sax.saxutils import escape

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT = "application/vnd.openxmlformats-officedocument.spreadsheetml."

# Show Values As (thuộc tính showDataAs của dataField)
SHOW_AS = {"pct_col": "percentOfCol", "pct_row": "percentOfRow", "pct_total": "percentOfTotal",
           "pct_diff": "percentDiff", "diff": "difference", "run_total": "runTotal", "pct": "percent", "index": "index"}
PREV_ITEM = 1048828          # baseItem = (previous)


@dataclass
class Value:
    field: str                        # tên cột nguồn, hoặc tên Calculated Field
    caption: str                      # tên hiện trên Pivot (không trùng tên cột)
    func: str = "sum"                 # sum | count | average | max | min
    fmt: str = "#,##0"                # định dạng số (phải có trong styles.xml; xem inject)
    show_as: str = None               # khóa trong SHOW_AS
    base_field: str = None            # trường gốc cho pct_diff / run_total / diff
    base_prev: bool = True            # pct_diff: so với mục liền trước


@dataclass
class Pivot:
    sheet: str                        # sheet đặt Pivot (phải có sẵn trong file)
    anchor: str                       # ô góc trên trái, ví dụ "A3"
    name: str
    cache: str                        # tên Cache
    rows: list = field(default_factory=list)
    cols: list = field(default_factory=list)
    filters: list = field(default_factory=list)
    values: list = field(default_factory=list)
    layout: str = "outline"           # compact | outline | tabular
    style: str = "PivotStyleMedium2"
    row_caption: str = None           # chữ thay cho "Row Labels" (compact)
    col_caption: str = None
    total_row: bool = True            # dòng Grand Total ở cuối (Excel: rowGrandTotals)
    total_col: bool = True            # cột Grand Total bên phải (Excel: colGrandTotals)
    sort_desc_by: str = None          # sắp xếp mục của trường hàng đầu tiên giảm dần theo giá trị (caption)
    top_n: tuple = None               # (trường, n, caption giá trị): Value Filter Top n
    hidden_items: dict = field(default_factory=dict)   # {trường: [giá trị bị ẩn]}


@dataclass
class Cache:
    name: str
    table: str                        # tên Excel Table nguồn (tblSales)
    columns: list                     # tên cột theo đúng thứ tự trong Table
    calc_fields: dict = field(default_factory=dict)    # {tên: công thức Pivot, ví dụ "=Gross_Margin/Revenue"}
    date_fields: tuple = ()
    items: dict = field(default_factory=dict)          # {cột: [giá trị theo thứ tự]} cho trường đặt vào hàng/cột/lọc
    sheet: str = None                 # nguồn là vùng thường (không phải Table): tên sheet …
    ref: str = None                   # … và vùng, ví dụ "A1:I6188"


def _col_to_num(c):
    n = 0
    for ch in c:
        n = n * 26 + ord(ch) - 64
    return n


def _num_to_col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _cell(a):
    m = re.match(r"([A-Z]+)(\d+)", a)
    return _col_to_num(m.group(1)), int(m.group(2))


def _xml_val(v):
    if isinstance(v, str):
        return f'<s v="{escape(v, {chr(34): "&quot;"})}"/>'
    return f'<n v="{v}"/>'


def _shared(values):
    if not values:
        return "<sharedItems/>"
    if all(isinstance(v, str) for v in values):
        return f'<sharedItems count="{len(values)}">' + "".join(_xml_val(v) for v in values) + "</sharedItems>"
    ints = ' containsInteger="1"' if all(float(v).is_integer() for v in values) else ""
    return (f'<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"{ints} '
            f'minValue="{min(values)}" maxValue="{max(values)}" count="{len(values)}">' + "".join(_xml_val(v) for v in values) + "</sharedItems>")


def _cache_xml(c, rid):
    fields = []
    for col in c.columns:
        fmt = ' numFmtId="14"' if col in c.date_fields else ' numFmtId="0"'
        fields.append(f'<cacheField name="{escape(col)}"{fmt}>{_shared(c.items.get(col))}</cacheField>')
    for nm, f in c.calc_fields.items():
        fields.append(f'<cacheField name="{escape(nm)}" numFmtId="0" formula="{escape(f.lstrip("="))}" databaseField="0"/>')
    src = (f'<worksheetSource name="{c.table}"/>' if c.table else f'<worksheetSource ref="{c.ref}" sheet="{escape(c.sheet)}"/>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="{NS}" xmlns:r="{NS_R}" r:id="{rid}" '
            f'refreshOnLoad="1" refreshedBy="MCI" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="0">'
            f'<cacheSource type="worksheet">{src}</cacheSource>'
            f'<cacheFields count="{len(fields)}">{"".join(fields)}</cacheFields></pivotCacheDefinition>')


def _records_xml():
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="{NS}" xmlns:r="{NS_R}" count="0"/>'


def _pivot_xml(p, c, cache_id, numfmt_ids):
    allf = list(c.columns) + list(c.calc_fields)
    idx = {f: i for i, f in enumerate(allf)}
    data_idx = {idx[v.field] for v in p.values}
    compact = p.layout == "compact"
    pfs = []
    for i, f in enumerate(allf):
        attrs = []
        axis = ("axisRow" if f in p.rows else "axisCol" if f in p.cols else "axisPage" if f in p.filters else None)
        if axis: attrs.append(f'axis="{axis}"')
        if i in data_idx: attrs.append('dataField="1"')
        if not compact: attrs += ['compact="0"', 'outline="0"' if p.layout == "tabular" else 'outline="1"']
        if f in c.calc_fields: attrs += ['dragToRow="0"', 'dragToCol="0"', 'dragToPage="0"']
        attrs.append('showAll="0"')
        if axis and p.sort_desc_by and p.rows and f == p.rows[0]:
            attrs.append('sortType="descending"')
        inner = ""
        if axis:
            its = c.items.get(f) or []
            hid = set(p.hidden_items.get(f, []))
            H = ' h="1"'
            inner = f'<items count="{len(its) + 1}">' + "".join(
                f'<item{H if v in hid else ""} x="{k}"/>' for k, v in enumerate(its)) + '<item t="default"/></items>'
            if axis and p.sort_desc_by and p.rows and f == p.rows[0]:
                k = [v.caption for v in p.values].index(p.sort_desc_by)
                inner += (f'<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="1">'
                          f'<reference field="4294967294" count="1" selected="0"><x v="{k}"/></reference></references></pivotArea></autoSortScope>')
        elif f in c.calc_fields:
            attrs.append('defaultSubtotal="0"')
        pfs.append(f'<pivotField {" ".join(attrs)}>{inner}</pivotField>' if inner else f'<pivotField {" ".join(attrs)}/>')
    x = lambda names: "".join(f'<field x="{idx[n]}"/>' for n in names)
    rows = f'<rowFields count="{len(p.rows)}">{x(p.rows)}</rowFields>' if p.rows else ""
    colnames = [idx[n] for n in p.cols] + ([-2] if len(p.values) > 1 else [])
    cols = (f'<colFields count="{len(colnames)}">' + "".join(f'<field x="{k}"/>' for k in colnames) + "</colFields>") if colnames else ""
    pages = (f'<pageFields count="{len(p.filters)}">' + "".join(f'<pageField fld="{idx[n]}" hier="-1"/>' for n in p.filters) + "</pageFields>") if p.filters else ""
    dfs = []
    for v in p.values:
        a = [f'name="{escape(v.caption)}"', f'fld="{idx[v.field]}"']
        if v.func != "sum": a.append(f'subtotal="{v.func}"')
        if v.show_as:
            a.append(f'showDataAs="{SHOW_AS[v.show_as]}"')
            bf = idx[v.base_field] if v.base_field else 0
            a += [f'baseField="{bf}"', f'baseItem="{PREV_ITEM if v.show_as in ("pct_diff", "diff") and v.base_prev else 0}"']
        else:
            a += ['baseField="0"', 'baseItem="0"']
        a.append(f'numFmtId="{numfmt_ids[v.fmt]}"')
        dfs.append(f'<dataField {" ".join(a)}/>')
    datas = f'<dataFields count="{len(dfs)}">{"".join(dfs)}</dataFields>'
    filt = ""
    if p.top_n:
        fld, n, cap = p.top_n
        k = [v.caption for v in p.values].index(cap)
        filt = (f'<filters count="1"><filter fld="{idx[fld]}" type="count" evalOrder="-1" id="1" iMeasureFld="{k}">'
                f'<autoFilter ref="A1"><filterColumn colId="0"><top10 val="{n}" filterVal="{n}"/></filterColumn></autoFilter></filter></filters>')
    c0, r0 = _cell(p.anchor)
    ncol = 1 + max(1, len(p.values)) * 2
    ref = f"{p.anchor}:{_num_to_col(c0 + ncol)}{r0 + 4}"
    top = 2 if p.filters else 0
    head = (f'<pivotTableDefinition xmlns="{NS}" name="{escape(p.name)}" cacheId="{cache_id}" applyNumberFormats="0" applyBorderFormats="0" '
            f'applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" '
            f'updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" '
            + ('' if compact else 'compact="0" compactData="0" ') + ('outline="1" outlineData="1" ' if p.layout != "tabular" else '')
            + ('' if p.total_row else 'rowGrandTotals="0" ') + ('' if p.total_col else 'colGrandTotals="0" ')
            + (f'rowHeaderCaption="{escape(p.row_caption)}" ' if p.row_caption else '') + (f'colHeaderCaption="{escape(p.col_caption)}" ' if p.col_caption else '')
            + 'multipleFieldFilters="0">')
    loc = f'<location ref="{ref}" firstHeaderRow="1" firstDataRow="{2 if p.cols or len(p.values) > 1 else 1}" firstDataCol="1"' + (f' rowPageCount="{len(p.filters)}" colPageCount="1"' if p.filters else "") + "/>"
    style = (f'<pivotTableStyleInfo name="{p.style}" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + head + loc
            + f'<pivotFields count="{len(pfs)}">{"".join(pfs)}</pivotFields>' + rows + cols + pages + datas + style + filt + "</pivotTableDefinition>")


def _rels(items):
    body = "".join(f'<Relationship Id="{i}" Type="{t}" Target="{g}"/>' for i, t, g in items)
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>'


def _add_rel(xml, rtype, target):
    ids = [int(i) for i in re.findall(r'Id="rId(\d+)"', xml)]
    rid = f"rId{max(ids, default=0) + 1}"
    return xml.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/></Relationships>'), rid


def _numfmt_ids(styles, fmts):
    """Tìm numFmtId của các định dạng cần dùng; định dạng tùy chỉnh chưa có thì thêm vào styles.xml."""
    builtin = {"General": 0, "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10}
    found = {m.group(2).replace("&quot;", '"'): int(m.group(1)) for m in re.finditer(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', styles)}
    ids = {}
    for f in fmts:
        if f in builtin: ids[f] = builtin[f]; continue
        if f in found: ids[f] = found[f]; continue
        nid = max([163] + list(found.values())) + 1
        found[f] = nid; ids[f] = nid
        code = escape(f, {'"': "&quot;"})
        if "<numFmts" in styles:
            styles = re.sub(r'<numFmts count="(\d+)">', lambda m: f'<numFmts count="{int(m.group(1)) + 1}">', styles)
            styles = styles.replace("</numFmts>", f'<numFmt numFmtId="{nid}" formatCode="{code}"/></numFmts>')
        else:
            styles = re.sub(r"(<styleSheet[^>]*>)", rf'\1<numFmts count="1"><numFmt numFmtId="{nid}" formatCode="{code}"/></numFmts>', styles, 1)
    return styles, ids


@dataclass
class Slicer:
    sheet: str
    field: str
    pivots: list                      # tên các Pivot được nối (Report Connections); cùng một Cache
    cell_from: str                    # ô góc trên trái, ví dụ "H3"
    cell_to: str                      # ô góc dưới phải (không tính)
    caption: str = None
    columns: int = 1
    style: str = "SlicerStyleLight1"


@dataclass
class Timeline:
    sheet: str
    field: str
    pivots: list
    cell_from: str
    cell_to: str
    caption: str = None
    level: int = 2                    # 0 năm · 1 quý · 2 tháng · 3 ngày
    start: str = "2025-01-01"
    end: str = "2025-04-01"


@dataclass
class Chart:
    sheet: str
    pivot: str                        # tên Pivot nguồn (PivotChart)
    kind: str                         # line | col | bar | pie | col_stack | col_pct
    title: str
    cell_from: str
    cell_to: str
    cats: str                         # vùng nhãn trục, ví dụ "PV!$A$5:$A$94"
    series: list                      # [(ô tên chuỗi, vùng giá trị)]
    legend: bool = False
    val_fmt: str = '#,##0,,"tr"'
    colors: list = field(default_factory=lambda: ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "4A3AA7"])
    labels: bool = False


XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
MSREL07, MSREL11 = "http://schemas.microsoft.com/office/2007/relationships/", "http://schemas.microsoft.com/office/2011/relationships/"


def _anchor(a, b, body, one_cell=False):
    (c0, r0), (c1, r1) = _cell(a), _cell(b)
    pos = lambda c, r: f"<xdr:col>{c - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff>"
    edit = ' editAs="oneCell"' if one_cell else ""
    return (f"<xdr:twoCellAnchor{edit}><xdr:from>{pos(c0, r0)}</xdr:from><xdr:to>{pos(c1, r1)}</xdr:to>"
            f"{body}<xdr:clientData/></xdr:twoCellAnchor>")


def _frame(nid, name, uri, inner):
    return (f'<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{nid}" name="{escape(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
            f'<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="{uri}">{inner}</a:graphicData></a:graphic></xdr:graphicFrame>')


def _fallback(nid, text):
    return (f'<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{nid + 500}" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>'
            f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="2524125"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
            f'<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>{text}</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>')


def _run(text, size, bold=False, color="1B2230"):
    return (f'<a:r><a:rPr lang="vi-VN" sz="{size}" b="{1 if bold else 0}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>{escape(text)}</a:t></a:r>')


def _chart_xml(ch, wbname):
    ser = []
    for k, (name_ref, val_ref) in enumerate(ch.series):
        col = ch.colors[k % len(ch.colors)]
        if ch.kind == "line":
            sp = f'<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="{col}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>'
        elif ch.kind == "pie":
            sp = ""
        else:
            sp = f'<c:spPr><a:solidFill><a:srgbClr val="{col}"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>'
        dl = ('<c:dLbls><c:numFmt formatCode="' + escape(ch.val_fmt, {'"': "&quot;"}) + '" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
              '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>') if ch.labels else ""
        ser.append(f'<c:ser><c:idx val="{k}"/><c:order val="{k}"/><c:tx><c:strRef><c:f>{name_ref}</c:f></c:strRef></c:tx>{sp}{dl}'
                   f'<c:cat><c:strRef><c:f>{ch.cats}</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>{val_ref}</c:f></c:numRef></c:val>'
                   + ('<c:smooth val="0"/>' if ch.kind == "line" else "") + "</c:ser>")
    ser = "".join(ser)
    ax = '<c:axId val="50010001"/><c:axId val="50010002"/>'
    if ch.kind == "line":
        plot = f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{ser}<c:marker val="1"/>{ax}</c:lineChart>'
    elif ch.kind == "pie":
        plot = f'<c:pieChart><c:varyColors val="1"/>{ser}<c:firstSliceAng val="0"/></c:pieChart>'
    else:
        grouping = {"col_stack": "stacked", "col_pct": "percentStacked"}.get(ch.kind, "clustered")
        bar_dir = "bar" if ch.kind == "bar" else "col"
        overlap = '<c:overlap val="100"/>' if grouping != "clustered" else ""
        plot = f'<c:barChart><c:barDir val="{bar_dir}"/><c:grouping val="{grouping}"/><c:varyColors val="0"/>{ser}<c:gapWidth val="60"/>{overlap}{ax}</c:barChart>'
    grid = '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E5E7EC"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>'
    axis_ln = '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="C3C7CF"/></a:solidFill></a:ln></c:spPr>'
    vf = escape(ch.val_fmt, {'"': "&quot;"})
    axes = "" if ch.kind == "pie" else (
        f'<c:catAx><c:axId val="50010001"/><c:scaling><c:orientation val="{"maxMin" if ch.kind == "bar" else "minMax"}"/></c:scaling><c:delete val="0"/>'
        f'<c:axPos val="{"l" if ch.kind == "bar" else "b"}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>'
        f'<c:tickLblPos val="low"/>{axis_ln}<c:crossAx val="50010002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
        f'<c:valAx><c:axId val="50010002"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="{"b" if ch.kind == "bar" else "l"}"/>{grid}'
        f'<c:numFmt formatCode="{vf}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        f'<c:spPr><a:ln><a:noFill/></a:ln></c:spPr><c:crossAx val="50010001"/><c:crosses val="{"max" if ch.kind == "bar" else "autoZero"}"/><c:crossBetween val="between"/></c:valAx>')
    legend = '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' if ch.legend else ""
    title = (f'<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr>{_run(ch.title, 1200, True, "203864")}</a:p></c:rich></c:tx>'
             f'<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>')
    txpr = ('<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="585858"/></a:solidFill><a:latin typeface="Arial"/>'
            '<a:cs typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="vi-VN"/></a:p></c:txPr>')
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="{C}" xmlns:a="{A}" xmlns:r="{NS_R}">'
            f'<c:date1904 val="0"/><c:lang val="vi-VN"/><c:roundedCorners val="0"/>'
            f'<c:pivotSource><c:name>[{escape(wbname)}]{escape(ch.sheet)}!{escape(ch.pivot)}</c:name><c:fmtId val="0"/></c:pivotSource>'
            f'<c:chart>{title}<c:plotArea><c:layout/>{plot}{axes}<c:spPr><a:noFill/></c:spPr></c:plotArea>{legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
            f'<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9DEE8"/></a:solidFill></a:ln></c:spPr>{txpr}'
            f'<c:extLst><c:ext uri="{{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}}" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">'
            f'<c14:pivotOptions><c14:dropZoneFilter val="1"/><c14:dropZoneCategories val="1"/><c14:dropZoneData val="1"/><c14:dropZonesVisible val="1"/></c14:pivotOptions></c:ext></c:extLst>'
            f'</c:chartSpace>')


def inject(path, caches, pivots, slicers=(), timelines=(), charts=()):
    """Ghi thêm Pivot, Slicer, Timeline, PivotChart vào file path (ghi đè)."""
    z = zipfile.ZipFile(path)
    files = {n: z.read(n) for n in z.namelist()}
    z.close()
    get = lambda n: files[n].decode("utf-8")
    wb, wbrels, ct = get("xl/workbook.xml"), get("xl/_rels/workbook.xml.rels"), get("[Content_Types].xml")
    add_ct = lambda part, typ: ct.replace("</Types>", f'<Override PartName="/{part}" ContentType="{typ}"/></Types>')
    sheets = {m.group(1): m.group(2) for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="(rId\d+)"', wb)}
    sheet_ids = {m.group(1): m.group(2) for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?sheetId="(\d+)"', wb)}
    targets = {m.group(1): m.group(2) for m in re.finditer(r'<Relationship Id="(rId\d+)"[^>]*?Target="([^"]+)"', wbrels)}
    targets.update({m.group(2): m.group(1) for m in re.finditer(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="(rId\d+)"', wbrels)})
    sheet_part = lambda nm: "xl/" + targets[sheets[nm]].lstrip("/").replace("xl/", "")
    rels_of = lambda part: part.replace("worksheets/", "worksheets/_rels/") + ".rels"
    styles, numfmt = _numfmt_ids(get("xl/styles.xml"), {v.fmt for p in pivots for v in p.values})
    files["xl/styles.xml"] = styles.encode("utf-8")
    cache_ids, pc_elems = {}, []
    for k, c in enumerate(caches, start=1):
        d, r = f"xl/pivotCache/pivotCacheDefinition{k}.xml", f"xl/pivotCache/pivotCacheRecords{k}.xml"
        xml = _cache_xml(c, "rId1").replace("</pivotCacheDefinition>",
              f'<extLst><ext uri="{{725AE2AE-9491-48be-B2B4-4EB974FC3084}}" xmlns:x14="{X14}"><x14:pivotCacheDefinition pivotCacheId="{k}"/></ext></extLst></pivotCacheDefinition>')
        files[d] = xml.encode("utf-8")
        files[r] = _records_xml().encode("utf-8")
        files[f"xl/pivotCache/_rels/pivotCacheDefinition{k}.xml.rels"] = _rels([("rId1", REL + "pivotCacheRecords", f"pivotCacheRecords{k}.xml")]).encode()
        wbrels, rid = _add_rel(wbrels, REL + "pivotCacheDefinition", f"pivotCache/pivotCacheDefinition{k}.xml")
        cid = 10 + k
        cache_ids[c.name] = (cid, c, k)
        pc_elems.append(f'<pivotCache xmlns:r="{NS_R}" cacheId="{cid}" r:id="{rid}"/>')
        ct = add_ct(d, CT + "pivotCacheDefinition+xml"); ct = add_ct(r, CT + "pivotCacheRecords+xml")
    block = f"<pivotCaches>{''.join(pc_elems)}</pivotCaches>"
    if "<calcPr" in wb:
        wb = re.sub(r"(<calcPr[^>]*/>)", r"\1" + block, wb, 1)
    elif "<extLst>" in wb:
        wb = wb.replace("<extLst>", block + "<extLst>", 1)
    else:
        wb = wb.replace("</workbook>", block + "</workbook>")
    piv_by_name = {}
    for j, p in enumerate(pivots, start=1):
        cid, c, k = cache_ids[p.cache]
        piv_by_name[p.name] = p
        part = f"xl/pivotTables/pivotTable{j}.xml"
        files[part] = _pivot_xml(p, c, cid, numfmt).encode("utf-8")
        files[f"xl/pivotTables/_rels/pivotTable{j}.xml.rels"] = _rels([("rId1", REL + "pivotCacheDefinition", f"../pivotCache/pivotCacheDefinition{k}.xml")]).encode()
        ct = add_ct(part, CT + "pivotTable+xml")
        srels = rels_of(sheet_part(p.sheet))
        cur = files[srels].decode() if srels in files else _rels([])
        cur, _ = _add_rel(cur, REL + "pivotTable", f"../pivotTables/pivotTable{j}.xml")
        files[srels] = cur.encode()

    # ---- Slicer, Timeline, PivotChart: mỗi sheet một drawing
    drawings, sheet_ext = {}, {}          # sheet -> list anchor xml ; sheet -> {uri: [r:id]}
    wb_ext, defined = {}, []
    nid = [1]
    def next_id():
        nid[0] += 1; return nid[0]
    def sheet_rel(sheet, rtype, target):
        srels = rels_of(sheet_part(sheet))
        cur = files[srels].decode() if srels in files else _rels([])
        cur, rid = _add_rel(cur, rtype, target)
        files[srels] = cur.encode()
        return rid
    for s_i, sl in enumerate(slicers, start=1):
        p0 = piv_by_name[sl.pivots[0]]
        _, c, k = cache_ids[p0.cache]
        cache_name = "Slicer_" + re.sub(r"\W", "_", sl.field) + (f"{s_i}" if s_i > 1 else "")
        its = c.items.get(sl.field) or []
        tabs = "".join(f'<pivotTable tabId="{sheet_ids[piv_by_name[n].sheet]}" name="{escape(n)}"/>' for n in sl.pivots)
        items = f'<items count="{len(its)}">' + "".join(f'<i x="{q}" s="1"/>' for q in range(len(its))) + "</items>"
        scpart = f"xl/slicerCaches/slicerCache{s_i}.xml"
        files[scpart] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<slicerCacheDefinition xmlns="{X14}" xmlns:mc="{MC}" mc:Ignorable="x" xmlns:x="{NS}" '
                         f'name="{cache_name}" sourceName="{escape(sl.field)}"><pivotTables>{tabs}</pivotTables><data><tabular pivotCacheId="{k}">{items}</tabular></data>'
                         f'</slicerCacheDefinition>').encode()
        ct = add_ct(scpart, "application/vnd.ms-excel.slicerCache+xml")
        wbrels, rid = _add_rel(wbrels, MSREL07 + "slicerCache", f"slicerCaches/slicerCache{s_i}.xml")
        wb_ext.setdefault("slicer", []).append(rid)
        defined.append(f'<definedName name="{cache_name}">#N/A</definedName>')
        sname = sl.caption or sl.field
        spart = f"xl/slicers/slicer{s_i}.xml"
        files[spart] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<slicers xmlns="{X14}" xmlns:mc="{MC}" mc:Ignorable="x" xmlns:x="{NS}">'
                        f'<slicer name="{escape(sname)}" cache="{cache_name}" caption="{escape(sname)}"'
                        + (f' columnCount="{sl.columns}"' if sl.columns > 1 else "") + f' style="{sl.style}" rowHeight="241300"/></slicers>').encode()
        ct = add_ct(spart, "application/vnd.ms-excel.slicer+xml")
        rid = sheet_rel(sl.sheet, MSREL07 + "slicer", f"../slicers/slicer{s_i}.xml")
        sheet_ext.setdefault(sl.sheet, {}).setdefault("slicer", []).append(rid)
        n_ = next_id()
        body = (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">'
                + _frame(n_, sname, "http://schemas.microsoft.com/office/drawing/2010/slicer",
                         f'<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="{escape(sname)}"/>')
                + "</mc:Choice>" + _fallback(n_, "Slicer: cần Excel 2010 trở lên.") + "</mc:AlternateContent>")
        drawings.setdefault(sl.sheet, []).append(_anchor(sl.cell_from, sl.cell_to, body, one_cell=True))
    for t_i, tl in enumerate(timelines, start=1):
        p0 = piv_by_name[tl.pivots[0]]
        _, c, k = cache_ids[p0.cache]
        cache_name = "NativeTimeline_" + re.sub(r"\W", "_", tl.field)
        tabs = "".join(f'<pivotTable tabId="{sheet_ids[piv_by_name[n].sheet]}" name="{escape(n)}"/>' for n in tl.pivots)
        tcpart = f"xl/timelineCaches/timelineCache{t_i}.xml"
        files[tcpart] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<timelineCacheDefinition xmlns="{X15}" xmlns:mc="{MC}" mc:Ignorable="x" xmlns:x="{NS}" '
                         f'name="{cache_name}" sourceName="{escape(tl.field)}"><pivotTables>{tabs}</pivotTables>'
                         f'<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="{k}" filterType="unknown">'
                         f'<bounds startDate="{tl.start}T00:00:00" endDate="{tl.end}T00:00:00"/></state></timelineCacheDefinition>').encode()
        ct = add_ct(tcpart, "application/vnd.ms-excel.timelineCache+xml")
        wbrels, rid = _add_rel(wbrels, MSREL11 + "timelineCache", f"timelineCaches/timelineCache{t_i}.xml")
        wb_ext.setdefault("timeline", []).append(rid)
        defined.append(f'<definedName name="{cache_name}">#N/A</definedName>')
        tname = tl.caption or tl.field
        tpart = f"xl/timelines/timeline{t_i}.xml"
        files[tpart] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<timelines xmlns="{X15}" xmlns:mc="{MC}" mc:Ignorable="x" xmlns:x="{NS}">'
                        f'<timeline name="{escape(tname)}" cache="{cache_name}" caption="{escape(tname)}" level="{tl.level}" selectionLevel="{tl.level}" '
                        f'scrollPosition="{tl.start}T00:00:00"/></timelines>').encode()
        ct = add_ct(tpart, "application/vnd.ms-excel.timeline+xml")
        rid = sheet_rel(tl.sheet, MSREL11 + "timeline", f"../timelines/timeline{t_i}.xml")
        sheet_ext.setdefault(tl.sheet, {}).setdefault("timeline", []).append(rid)
        n_ = next_id()
        body = (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice xmlns:tsle="http://schemas.microsoft.com/office/drawing/2012/timeslicer" Requires="tsle">'
                + _frame(n_, tname, "http://schemas.microsoft.com/office/drawing/2012/timeslicer",
                         f'<tsle:timeslicer xmlns:tsle="http://schemas.microsoft.com/office/drawing/2012/timeslicer" name="{escape(tname)}"/>')
                + "</mc:Choice>" + _fallback(n_, "Timeline: cần Excel 2013 trở lên.") + "</mc:AlternateContent>")
        drawings.setdefault(tl.sheet, []).append(_anchor(tl.cell_from, tl.cell_to, body, one_cell=True))
    wbname = os.path.basename(path)
    chart_rels = {}
    for c_i, ch in enumerate(charts, start=1):
        cpart = f"xl/charts/chart{c_i}.xml"
        files[cpart] = _chart_xml(ch, wbname).encode("utf-8")
        ct = add_ct(cpart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
        chart_rels.setdefault(ch.sheet, []).append(c_i)
        n_ = next_id()
        body = _frame(n_, f"Chart {c_i}", C, f'<c:chart xmlns:c="{C}" xmlns:r="{NS_R}" r:id="rIdC{c_i}"/>')
        drawings.setdefault(ch.sheet, []).append(_anchor(ch.cell_from, ch.cell_to, body))
    for d_i, (sheet, anchors) in enumerate(drawings.items(), start=1):
        dpart = f"xl/drawings/drawing{d_i}.xml"
        files[dpart] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}">' + "".join(anchors) + "</xdr:wsDr>").encode()
        ct = add_ct(dpart, "application/vnd.openxmlformats-officedocument.drawing+xml")
        files[f"xl/drawings/_rels/drawing{d_i}.xml.rels"] = _rels(
            [(f"rIdC{c_i}", REL + "chart", f"../charts/chart{c_i}.xml") for c_i in chart_rels.get(sheet, [])]).encode()
        rid = sheet_rel(sheet, REL + "drawing", f"../drawings/drawing{d_i}.xml")
        sp = sheet_part(sheet)
        sx = get(sp)
        tag = f'<drawing xmlns:r="{NS_R}" r:id="{rid}"/>'
        m = re.search(r"<(legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b", sx)
        sx = sx[:m.start()] + tag + sx[m.start():] if m else sx.replace("</worksheet>", tag + "</worksheet>")
        files[sp] = sx.encode()
    for sheet, exts in sheet_ext.items():
        sp = sheet_part(sheet)
        sx = get(sp)
        e = ""
        if "slicer" in exts:
            e += (f'<ext uri="{{A8765BA9-456A-4dab-B4F3-ACF838C121DE}}" xmlns:x14="{X14}"><x14:slicerList>'
                  + "".join(f'<x14:slicer xmlns:r="{NS_R}" r:id="{r}"/>' for r in exts["slicer"]) + "</x14:slicerList></ext>")
        if "timeline" in exts:
            e += (f'<ext uri="{{7E03D99C-DC04-49d9-9315-930204A7B6E9}}" xmlns:x15="{X15}"><x15:timelineRefs>'
                  + "".join(f'<x15:timelineRef xmlns:r="{NS_R}" r:id="{r}"/>' for r in exts["timeline"]) + "</x15:timelineRefs></ext>")
        sx = sx.replace("</extLst>", e + "</extLst>") if "<extLst>" in sx else sx.replace("</worksheet>", f"<extLst>{e}</extLst></worksheet>")
        files[sp] = sx.encode()
    if wb_ext:
        e = ""
        if "slicer" in wb_ext:
            e += (f'<ext uri="{{BBE1A952-AA13-448e-AADC-164F8A28A991}}" xmlns:x14="{X14}"><x14:slicerCaches>'
                  + "".join(f'<x14:slicerCache xmlns:r="{NS_R}" r:id="{r}"/>' for r in wb_ext["slicer"]) + "</x14:slicerCaches></ext>")
        if "timeline" in wb_ext:
            e += (f'<ext uri="{{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}}" xmlns:x15="{X15}"><x15:timelineCacheRefs>'
                  + "".join(f'<x15:timelineCacheRef xmlns:r="{NS_R}" r:id="{r}"/>' for r in wb_ext["timeline"]) + "</x15:timelineCacheRefs></ext>")
        wb = wb.replace("</extLst>", e + "</extLst>") if "<extLst>" in wb else wb.replace("</workbook>", f"<extLst>{e}</extLst></workbook>")
    if defined:
        if "<definedNames/>" in wb:
            wb = wb.replace("<definedNames/>", "<definedNames>" + "".join(defined) + "</definedNames>")
        elif "<definedNames>" in wb:
            wb = wb.replace("</definedNames>", "".join(defined) + "</definedNames>")
        else:
            wb = wb.replace("<calcPr", "<definedNames>" + "".join(defined) + "</definedNames><calcPr", 1)
    files["xl/workbook.xml"], files["xl/_rels/workbook.xml.rels"], files["[Content_Types].xml"] = wb.encode(), wbrels.encode(), ct.encode()
    files.pop("xl/calcChain.xml", None)
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for n in (["[Content_Types].xml"] + [n for n in files if n != "[Content_Types].xml"]):
            out.writestr(n, files[n])
    os.replace(tmp, path)


EXCEL_DOCS = os.path.expanduser("~/Library/Containers/com.microsoft.Excel/Data/Documents")


def excel_resave(path, check=None):
    """Mở file bằng Microsoft Excel (Mac, AppleScript), Refresh All, lưu, đóng; trả về giá trị các ô check ("Sheet!A1").

    File được chép vào thư mục Documents trong sandbox của Excel rồi mới mở: mở thẳng trong thư mục OneDrive thì Excel
    có thể lấy bản cũ trên đám mây (vừa ghi đè nhưng chưa đồng bộ) và lưu đè lên bản mới dựng. Lưu xong chép ngược về."""
    import shutil
    work = os.path.join(EXCEL_DOCS, os.path.basename(path))
    shutil.copyfile(path, work)
    before = os.path.getmtime(work)
    cells = check or []
    reads = "".join(f'set end of out to (value of range "{a.split("!")[1]}" of worksheet "{a.split("!")[0]}" of wb) as text\n' for a in cells)
    name = os.path.basename(path)
    script = f'''
set p to POSIX file "{work}"
with timeout of 300 seconds
tell application "Microsoft Excel"
    set display alerts to false
    repeat with w in (every workbook whose name is "{name}")
        close w saving no
    end repeat
    if (count of (every workbook whose name is "{name}")) > 0 then error "Đang mở một workbook cùng tên, không đóng được"
    open p
    set wb to workbook "{name}"
    if (full name of wb) does not contain "com.microsoft.Excel" then error "Excel mở nhầm bản khác: " & (full name of wb)
    refresh all wb
    calculate
    set out to {{}}
    {reads}
    save wb
    set t0 to 0
    repeat while (saved of wb) is false and t0 < 120
        delay 0.5
        set t0 to t0 + 1
    end repeat
    delay 2
    close wb saving no
end tell
end timeout
set AppleScript's text item delimiters to "|"
return out as text
'''
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=600)
    if r.returncode:
        raise RuntimeError(r.stderr)
    import time
    for _ in range(60):                       # chờ Excel ghi xong
        if os.path.getmtime(work) > before:
            size = os.path.getsize(work); time.sleep(1)
            if os.path.getsize(work) == size: break
        time.sleep(1)
    else:
        raise RuntimeError(f"Excel không lưu {path}")
    shutil.copyfile(work, path)
    os.remove(work)
    return r.stdout.strip().split("|") if cells else []
