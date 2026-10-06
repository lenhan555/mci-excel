r"""Thêm Power Query (queries M), load ra Table / Connection Only + Data Model, và dựng PivotTable trên bảng do
Power Query nạp ra — bằng cách điều khiển Excel qua COM (PowerShell), vì openpyxl không tạo được Power Query.

Chỉ chạy trên Windows (Excel qua COM). Cách dùng:

    from powerquery_com import Query, TableLoad, ModelLoad, Pivot, KiemTra, run

    run(path,
        queries=[Query("fnCleanSales", m_fn, is_function=True), Query("Sales", m_sales)],
        table_loads=[TableLoad(query="Sales", sheet="Sales_PQ", table_name="Sales")],
        model_loads=[ModelLoad("Product"), ModelLoad("Store")],
        pivots=[Pivot(sheet="PV_Refresh", name="PV_Refresh", source_table="Sales",
                       rows=["Store_Code"], cols=["Thang"],
                       values=[("Revenue", "Doanh_thu", "sum")])],
        kiemtra=KiemTra(sheet="Kiem_tra_B4", title="KIỂM TRA / ĐỐI SOÁT — Brick output B4", rows=[...]),
        checks=["Sales_PQ!A1"])

Mỗi lần gọi `run` mở đúng MỘT phiên Excel, làm hết mọi việc (thêm query, load, dựng Pivot, RefreshAll, ghi
Kiem_tra), lưu, đóng — rồi trả về giá trị các ô `checks` để script gọi kiểm tra kết quả.
"""
import json, os, subprocess, sys, tempfile, shutil
from dataclasses import dataclass, field

XL_SUM, XL_COUNT, XL_AVERAGE, XL_MAX, XL_MIN = -4157, -4112, -4106, -4136, -4139
FUNC_CODE = {"sum": XL_SUM, "count": XL_COUNT, "average": XL_AVERAGE, "max": XL_MAX, "min": XL_MIN}


@dataclass
class Query:
    name: str
    formula: str                      # mã M
    description: str = ""
    is_function: bool = False         # True: query hàm (Content as binary) => ..., không Load


@dataclass
class TableLoad:
    query: str                        # tên Query đã Add
    sheet: str                        # sheet mới để đặt bảng (tạo nếu chưa có)
    table_name: str                   # tên ListObject sau khi đổi tên (dùng trong công thức =Ten[Cot])
    anchor: str = "A1"
    also_model: bool = True           # đồng thời Add to Data Model (dùng connection thứ hai)
    background: bool = False          # Enable background refresh: tắt để Refresh All chạy xong query rồi mới cập nhật Pivot


@dataclass
class ModelLoad:
    query: str                        # Connection Only + Add to Data Model


@dataclass
class Pivot:
    sheet: str                        # sheet mới đặt Pivot
    name: str
    source_table: str                 # tên ListObject nguồn (phải nằm trong table_loads)
    anchor: str = "A3"
    rows: list = field(default_factory=list)
    cols: list = field(default_factory=list)
    values: list = field(default_factory=list)   # [(field, caption, func)]
    filters: list = field(default_factory=list)  # trường đặt ở vùng Filters (page field)


@dataclass
class KiemTra:
    sheet: str
    title: str
    note: str
    rows: list                        # [(label, formula_or_None, value, expected)] — formula=None: value là số, hoặc "model:<Bảng>"
                                      # để ghi RecordCount của bảng đó trong Data Model (đọc qua COM sau Refresh All)


def _esc_ps_dq(s: str) -> str:
    """Escape một chuỗi để nhúng an toàn trong PowerShell here-string double-quote (@\"...\"@): nhân đôi backtick và $."""
    return s.replace("`", "``").replace("$", "`$")


def _ps_str(s: str) -> str:
    """Chuỗi PowerShell single-quote (an toàn cho mọi ký tự, kể cả $), nhân đôi dấu nháy đơn."""
    return "'" + str(s).replace("'", "''") + "'"


def run(path, queries, table_loads=(), model_loads=(), pivots=(), kiemtra=None, checks=()):
    if sys.platform != "win32":
        raise RuntimeError("powerquery_com chỉ chạy trên Windows (Excel qua COM).")
    path = os.path.abspath(path)
    tmpdir = os.path.join(tempfile.gettempdir(), "mci_excel_pq")
    os.makedirs(tmpdir, exist_ok=True)
    work = os.path.join(tmpdir, os.path.basename(path))
    res, ps1 = work + ".json", work + ".ps1"
    for f in (work, res):
        if os.path.exists(f):
            os.remove(f)
    shutil.copyfile(path, work)
    before_mtime = os.path.getmtime(work)

    lines = [_PS_HEADER.replace("__WORK__", _ps_str(work))]

    for q in queries:
        var = f"m_{_safe(q.name)}"
        lines.append(f'${var} = @"\n{_esc_ps_dq(q.formula)}\n"@')
        lines.append(f'Com {{ [void]$wb.Queries.Add({_ps_str(q.name)}, ${var}, {_ps_str(q.description)}) }}')

    for tl in table_loads:
        conn = f"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={tl.query};Extended Properties=\"\""
        lines.append(f'''
Com {{
    $sh = $null
    try {{ $sh = $wb.Worksheets.Item({_ps_str(tl.sheet)}) }} catch {{ }}
    if ($null -eq $sh) {{ $sh = $wb.Worksheets.Add(); $sh.Name = {_ps_str(tl.sheet)} }}
    $srcArr = ,({_ps_str(conn)})
    $lo = $sh.ListObjects.Add(0, $srcArr, [Type]::Missing, 1, $sh.Range({_ps_str(tl.anchor)}))
    $qt = $lo.QueryTable
    $qt.CommandType = 2
    $qt.CommandText = "SELECT * FROM [{tl.query}]"
    $qt.BackgroundQuery = {"$true" if tl.background else "$false"}
    $qt.Refresh($false)
    $lo.Name = {_ps_str(tl.table_name)}
}}''')
        if tl.also_model:
            model_loads = list(model_loads) + [ModelLoad(tl.query)]

    for ml in model_loads:
        conn = f"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={ml.query};Extended Properties=\"\""
        lines.append(f'Com {{ [void]$wb.Connections.Add2({_ps_str("Query - " + ml.query)}, "", {_ps_str(conn)}, {_ps_str(ml.query)}, 6, $true, $false) }}')

    lines.append('Com { $wb.RefreshAll() }')
    lines.append('Com { $x.CalculateUntilAsyncQueriesDone() }')
    lines.append('Start-Sleep -Seconds 2')

    for pv in pivots:
        rows_ps = ", ".join(_ps_str(r) for r in pv.rows)
        cols_ps = ", ".join(_ps_str(c) for c in pv.cols)
        filt_ps = ", ".join(_ps_str(c) for c in pv.filters)
        vals = []
        for fld, cap, func in pv.values:
            vals.append(f'Com {{ [void]$pt.AddDataField($pt.PivotFields({_ps_str(fld)}), {_ps_str(cap)}, {FUNC_CODE[func]}) }}')
        vals_ps = "\n    ".join(vals)
        lines.append(f'''
Com {{
    $sh = $null
    try {{ $sh = $wb.Worksheets.Item({_ps_str(pv.sheet)}) }} catch {{ }}
    if ($null -eq $sh) {{ $sh = $wb.Worksheets.Add(); $sh.Name = {_ps_str(pv.sheet)} }}
    $caches = $wb.PivotCaches()
    $cache = $caches.Create(1, {_ps_str(pv.source_table)})
    $pt = $cache.CreatePivotTable($sh.Range({_ps_str(pv.anchor)}), {_ps_str(pv.name)})
    foreach ($f in @({rows_ps})) {{ $pt.PivotFields($f).Orientation = 1 }}
    foreach ($f in @({cols_ps})) {{ $pt.PivotFields($f).Orientation = 2 }}
    foreach ($f in @({filt_ps})) {{ $pt.PivotFields($f).Orientation = 3 }}
    {vals_ps}
}}''')

    if kiemtra:
        lines.append(f'''
Com {{
    $sh = $null
    try {{ $sh = $wb.Worksheets.Item({_ps_str(kiemtra.sheet)}) }} catch {{ }}
    if ($null -eq $sh) {{ $sh = $wb.Worksheets.Add(); $sh.Name = {_ps_str(kiemtra.sheet)} }}
    $sh.Range("A1").Value2 = {_ps_str(kiemtra.title)}
    $sh.Range("A2").Value2 = {_ps_str(kiemtra.note)}
    $sh.Range("A4").Value2 = "Chỉ tiêu"
    $sh.Range("C4").Value2 = "Kết quả"
    $sh.Range("D4").Value2 = "Đáp án"
    $sh.Range("E4").Value2 = "Khớp?"
}}''')
        for i, (label, formula, value, expected) in enumerate(kiemtra.rows, start=5):
            lines.append(f'Com {{ $sh = $wb.Worksheets.Item({_ps_str(kiemtra.sheet)}); $sh.Range("A{i}").Value2 = {_ps_str(label)} }}')
            if formula:
                lines.append(f'Com {{ $sh = $wb.Worksheets.Item({_ps_str(kiemtra.sheet)}); $sh.Range("C{i}").Formula = {_ps_str(formula)} }}')
            elif isinstance(value, str) and value.startswith("model:"):
                mt = value.split(":", 1)[1]
                lines.append(f'Com {{ $sh = $wb.Worksheets.Item({_ps_str(kiemtra.sheet)}); $sh.Range("C{i}").Value2 = [double]($wb.Model.ModelTables.Item({_ps_str(mt)}).RecordCount) }}')
            else:
                lines.append(f'Com {{ $sh = $wb.Worksheets.Item({_ps_str(kiemtra.sheet)}); $sh.Range("C{i}").Value2 = {value} }}')
            lines.append(f'Com {{ $sh = $wb.Worksheets.Item({_ps_str(kiemtra.sheet)}); $sh.Range("D{i}").Value2 = {expected} }}')
            lines.append(f'Com {{ $sh = $wb.Worksheets.Item({_ps_str(kiemtra.sheet)}); $sh.Range("E{i}").Formula = "=IF(ROUND(C{i},2)=ROUND(D{i},2),""✔"",""✘"")" }}')

    lines.append('Com { $x.CalculateFull() }')
    cells_ps = ", ".join(_ps_str(c) for c in checks)
    lines.append(f'$cellsToRead = @({cells_ps})')
    lines.append(_PS_FOOTER.replace("__RES__", _ps_str(res)))

    code = "\n".join(lines)
    with open(ps1, "w", encoding="utf-8-sig") as f:
        f.write(code)
    try:
        r = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", ps1],
                            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900)
        if r.returncode or not os.path.exists(res):
            raise RuntimeError(f"Power Query build thất bại cho {path}:\n{r.stdout}\n{r.stderr}")
        if os.path.getmtime(work) <= before_mtime:
            raise RuntimeError(f"Excel không lưu {path}")
        out = json.load(open(res, encoding="utf-8"))
        if out.get("error"):
            raise RuntimeError(f"Lỗi khi dựng {path}:\n{out['error']}")
        shutil.copyfile(work, path)
        return out.get("cells", [])
    finally:
        for f in (work, res, ps1):
            if os.path.exists(f):
                try:
                    os.remove(f)
                except OSError:
                    pass


def _safe(name: str) -> str:
    return "".join(c if c.isalnum() else "_" for c in name)


_PS_HEADER = r"""
$ErrorActionPreference = 'Stop'
$work = __WORK__
function Com([scriptblock]$b) {
  for ($i = 0; ; $i++) {
    try { return (& $b) }
    catch {
      $busy = $_.Exception.ToString() -match '0x80010001|0x8001010A|RPC_E_CALL_REJECTED|RPC_E_SERVERCALL_RETRYLATER'
      if ($busy -and $i -lt 600) { Start-Sleep -Milliseconds 500 } else { throw }
    }
  }
}
$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
$x = New-Object -ComObject Excel.Application
$mine = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
$errOut = $null
$cellsOut = @()
try {
  Com { $x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false }
  Com { $x.AskToUpdateLinks = $false; $x.EnableEvents = $false }
  $wb = Com { $x.Workbooks.Open($work, 0, $false) }
  if (Com { $wb.ReadOnly }) { throw "Excel mo file o che do chi doc: $work" }
"""

_PS_FOOTER = r"""
  foreach ($c in $cellsToRead) {
    $sh, $addr = $c.Split('!', 2)
    $cellsOut += [string](Com { $wb.Worksheets.Item($sh).Range($addr).Value2 })
  }
  Com { $wb.Save() }
  Com { $wb.Close($false) }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
} catch {
  $errOut = $_.Exception.ToString()
} finally {
  if ($mine.Count -gt 0) { try { Com { $x.Quit() } } catch { } }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  $x = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Start-Sleep -Milliseconds 800
  foreach ($procId in $mine) { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue }
}
$result = @{ error = $errOut; cells = $cellsOut }
[IO.File]::WriteAllText(__RES__, (ConvertTo-Json -InputObject $result -Compress), (New-Object Text.UTF8Encoding $false))
"""
