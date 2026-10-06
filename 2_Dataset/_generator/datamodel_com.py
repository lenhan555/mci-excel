r"""Dựng Data Model (relationship, measure DAX), PivotTable From Data Model và ghi ô (giá trị / công thức CUBEVALUE)
bằng cách điều khiển Excel qua COM (PowerShell) — openpyxl không đọc / ghi được Data Model, và ghi lại bằng openpyxl
sẽ làm mất nó.

Chỉ chạy trên Windows. Cùng khuôn với powerquery_com.py: mỗi lần gọi `run` mở đúng MỘT phiên Excel ẩn trên bản sao
tạm của file, làm hết mọi việc, tính lại (kể cả truy vấn CUBEVALUE bất đồng bộ), lưu, đóng, rồi chép đè file đích
và trả về giá trị các ô `checks`.

KHÔNG gọi RefreshAll: file khởi đầu B5 có query Power Query trỏ tới thư mục của máy dựng B4; refresh có thể làm đổi
dữ liệu. Thêm relationship / measure / Pivot từ Data Model không cần refresh query.

Giới hạn của mô hình đối tượng Excel (không có API, script không làm được — ghi thành bước làm tay nếu cần):
calculated column, Mark as Date Table, Hide from Client Tools.

    from datamodel_com import Rel, Measure, CubePivot, Cells, run
    run(src, dst, rels=[Rel("Sales", "Product_Code", "Product", "Product_Code")],
        measures=[Measure("Total Revenue", "Sales", "SUM(Sales[Revenue])", "int")],
        pivots=[CubePivot("PV_X", "PV_X", rows=["[Store].[Region]"], values=["Total Revenue"],
                          filters={"[Calendar].[Nam]": ["2025"]})],
        cells=[Cells("Kiem_tra", [("A1", "Tiêu đề"), ("C5", "=CUBEVALUE(...)")])],
        checks=["Kiem_tra!E5"])
"""
import json, os, shutil, subprocess, sys, tempfile
from dataclasses import dataclass, field

FMT = {  # định dạng measure → biểu thức PowerShell trên $m (Workbook.Model)
    "int": "$m.ModelFormatWholeNumber($true)",
    "pct1": "$m.ModelFormatPercentageNumber($true, 1)",
    "pct2": "$m.ModelFormatPercentageNumber($true, 2)",
    "dec0": "$m.ModelFormatDecimalNumber($true, 0)",
}


@dataclass
class Rel:
    fk_table: str          # bảng fact (phía nhiều)
    fk_col: str
    pk_table: str          # bảng lookup (phía một)
    pk_col: str


@dataclass
class Measure:
    name: str
    table: str             # bảng chứa measure
    formula: str           # DAX, không có dấu "=" đầu
    fmt: str = "int"       # khóa của FMT
    description: str = ""


@dataclass
class CubePivot:
    sheet: str
    name: str
    anchor: str = "A8"                               # chừa dòng 1–2 cho tiêu đề, 4–7 cho vùng Filters
    rows: list = field(default_factory=list)        # cube field "[Bảng].[Cột]"
    cols: list = field(default_factory=list)
    values: list = field(default_factory=list)      # tên measure
    filters: dict = field(default_factory=dict)     # "[Bảng].[Cột]" → [khóa thành viên] (1 khóa = chọn một; nhiều = multi)
    row_items: dict = field(default_factory=dict)   # "[Bảng].[Cột]" → [khóa] chỉ hiện các item này ở Rows/Columns
    title: str = ""                                  # ghi vào A1 của sheet
    note: str = ""                                   # ghi vào A2


@dataclass
class Cells:
    sheet: str
    items: list                                      # [(địa chỉ, giá trị | "=công thức")]
    widths: dict = field(default_factory=dict)       # cột → độ rộng
    numfmt: dict = field(default_factory=dict)       # vùng → number format
    bold: list = field(default_factory=list)         # vùng in đậm
    fill: dict = field(default_factory=dict)         # vùng → màu nền hex RRGGBB
    font_color: dict = field(default_factory=dict)   # vùng → màu chữ hex
    wrap: list = field(default_factory=list)         # vùng xuống dòng
    title: str = ""                                  # vùng A1 kiểu tiêu đề (đậm, 14, navy) nếu khác rỗng


def _ps(s) -> str:
    return "'" + str(s).replace("'", "''") + "'"


def _bgr(hex_rgb: str) -> int:
    r, g, b = int(hex_rgb[0:2], 16), int(hex_rgb[2:4], 16), int(hex_rgb[4:6], 16)
    return r + (g << 8) + (b << 16)


def _member(field_: str, key) -> str:
    return f"{field_}.&[{key}]"


def build_ps(work, res, delete_sheets=(), rels=(), measures=(), pivots=(), cells=(), checks=(), active_sheet=None, first_sheets=()):
    L = [_HEADER.replace("__WORK__", _ps(work))]
    L.append("$m = Com { $wb.Model }")
    L.append("for ($i=0; $i -lt 40; $i++) { $n = 0; try { $m = $wb.Model; $n = $m.ModelTables.Count } catch { }; if ($n -gt 0) { break }; try { $x.CalculateUntilAsyncQueriesDone() } catch { }; Start-Sleep -Seconds 3 }")
    L.append('$log = New-Object System.Collections.ArrayList')
    L.append('[void]$log.Add("tables=" + $m.ModelTables.Count); [void]$log.Add("start: relationships=" + (Com { $m.ModelRelationships.Count }) + ", measures=" + (Com { $m.ModelMeasures.Count }))')

    for sh in delete_sheets:
        L.append(f"try {{ Com {{ $wb.Worksheets.Item({_ps(sh)}).Delete() }} }} catch {{ [void]$log.Add('khong xoa duoc sheet ' + {_ps(sh)}) }}")

    for r in rels:
        L.append(f"""Com {{
  $fk = $m.ModelTables.Item({_ps(r.fk_table)}).ModelTableColumns.Item({_ps(r.fk_col)})
  $pk = $m.ModelTables.Item({_ps(r.pk_table)}).ModelTableColumns.Item({_ps(r.pk_col)})
  [void]$m.ModelRelationships.Add($fk, $pk)
}}""")

    for ms in measures:
        L.append(f"""$t = Com {{ $m.ModelTables.Item({_ps(ms.table)}) }}
try {{ $f = Com {{ {FMT[ms.fmt]} }} }} catch {{ $f = [Type]::Missing; [void]$log.Add('khong tao duoc dinh dang cho ' + {_ps(ms.name)}) }}
try {{ Com {{ [void]$m.ModelMeasures.Add({_ps(ms.name)}, $t, {_ps(ms.formula)}, $f, {_ps(ms.description)}) }} }}
catch {{ Com {{ [void]$m.ModelMeasures.Add({_ps(ms.name)}, $t, {_ps('=' + ms.formula)}, $f, {_ps(ms.description)}) }}; [void]$log.Add('measure can dau = : ' + {_ps(ms.name)}) }}""")

    L.append('[void]$log.Add("sau khi dung: relationships=" + (Com { $m.ModelRelationships.Count }) + ", measures=" + (Com { $m.ModelMeasures.Count }))')

    for pv in pivots:
        L.append(_sheet_get_or_add(pv.sheet))
        if pv.title:
            L.append(f"Com {{ $sh.Range('A1').Value2 = {_ps(pv.title)}; $sh.Range('A1').Font.Bold = $true; $sh.Range('A1').Font.Size = 13; $sh.Range('A1').Font.Color = {_bgr('203864')} }}")
        if pv.note:
            L.append(f"Com {{ $sh.Range('A2').Value2 = {_ps(pv.note)}; $sh.Range('A2').Font.Italic = $true; $sh.Range('A2').Font.Color = {_bgr('585858')} }}")
        # biến gán trong scriptblock của Com chỉ sống trong scope con → lấy $pt qua giá trị trả về
        L.append(f"""$pt = Com {{
  $pc = $wb.PivotCaches().Create(2, $wb.Connections.Item('ThisWorkbookDataModel'), 6)
  $pc.CreatePivotTable($sh.Range({_ps(pv.anchor)}), {_ps(pv.name)})
}}""")
        for fld, keys in pv.filters.items():
            L.append(f"Com {{ $pt.CubeFields.Item({_ps(fld)}).Orientation = 3 }}")
            pf = f"{fld}.[{fld.split('[')[-1].rstrip(']')}]"
            if len(keys) == 1:
                L.append(f"Com {{ $pt.PivotFields({_ps(pf)}).CurrentPageName = {_ps(_member(fld, keys[0]))} }}")
            else:
                arr = ", ".join(_ps(_member(fld, k)) for k in keys)
                L.append(f"Com {{ $pt.CubeFields.Item({_ps(fld)}).EnableMultiplePageItems = $true }}")
                L.append(f"Com {{ $pt.PivotFields({_ps(pf)}).VisibleItemsList = [object[]]@({arr}) }}")
        for fld in pv.rows:
            L.append(f"Com {{ $pt.CubeFields.Item({_ps(fld)}).Orientation = 1 }}")
        for fld in pv.cols:
            L.append(f"Com {{ $pt.CubeFields.Item({_ps(fld)}).Orientation = 2 }}")
        for meas in pv.values:
            L.append(f"Com {{ $pt.CubeFields.Item({_ps('[Measures].[' + meas + ']')}).Orientation = 4 }}")
        for fld, keys in pv.row_items.items():
            pf = f"{fld}.[{fld.split('[')[-1].rstrip(']')}]"
            arr = ", ".join(_ps(_member(fld, k)) for k in keys)
            L.append(f"Com {{ $pt.PivotFields({_ps(pf)}).VisibleItemsList = [object[]]@({arr}) }}")
        L.append("try { Com { $pt.RowAxisLayout(1) } } catch { }")

    for c in cells:
        L.append(_sheet_get_or_add(c.sheet))
        for addr, v in c.items:
            if isinstance(v, str) and v.startswith("="):
                L.append(f"Com {{ $sh.Range({_ps(addr)}).Formula = {_ps(v)} }}")
            elif isinstance(v, (int, float)):
                L.append(f"Com {{ $sh.Range({_ps(addr)}).Value2 = [double]{repr(float(v))} }}")
            elif v is not None:
                L.append(f"Com {{ $sh.Range({_ps(addr)}).Value2 = {_ps(v)} }}")
        for col, w in c.widths.items():
            L.append(f"Com {{ $sh.Columns({_ps(col)}).ColumnWidth = {w} }}")
        for rng, f in c.numfmt.items():
            L.append(f"Com {{ $sh.Range({_ps(rng)}).NumberFormat = {_ps(f)} }}")
        for rng in c.bold:
            L.append(f"Com {{ $sh.Range({_ps(rng)}).Font.Bold = $true }}")
        for rng, col in c.fill.items():
            L.append(f"Com {{ $sh.Range({_ps(rng)}).Interior.Color = {_bgr(col)} }}")
        for rng, col in c.font_color.items():
            L.append(f"Com {{ $sh.Range({_ps(rng)}).Font.Color = {_bgr(col)} }}")
        for rng in c.wrap:
            L.append(f"Com {{ $sh.Range({_ps(rng)}).WrapText = $true; $sh.Range({_ps(rng)}).VerticalAlignment = -4160 }}")
        if c.title:
            L.append(f"Com {{ $sh.Range('A1').Font.Bold = $true; $sh.Range('A1').Font.Size = 14; $sh.Range('A1').Font.Color = {_bgr('203864')} }}")

    # thứ tự sheet: các sheet trong first_sheets đặt lên đầu, theo đúng thứ tự
    for name in reversed(list(first_sheets)):
        L.append(f"Com {{ $wb.Worksheets.Item({_ps(name)}).Move($wb.Worksheets.Item(1)) }}")
    if active_sheet:
        L.append(f"Com {{ $x.ScreenUpdating = $true; $wb.Worksheets.Item({_ps(active_sheet)}).Activate(); $wb.Worksheets.Item({_ps(active_sheet)}).Range('A1').Select() }}")

    L.append("Com { $x.CalculateUntilAsyncQueriesDone() }")
    L.append("Com { $x.CalculateFull() }")
    L.append("Com { $x.CalculateUntilAsyncQueriesDone() }")
    L.append("Start-Sleep -Seconds 2")
    L.append(f"$cellsToRead = @({', '.join(_ps(c) for c in checks)})")
    L.append(_FOOTER.replace("__RES__", _ps(res)))
    return "\n".join(L)


def run(src, dst, **kw):
    if sys.platform != "win32":
        raise RuntimeError("datamodel_com chỉ chạy trên Windows (Excel qua COM).")
    tmpdir = os.path.join(tempfile.gettempdir(), "mci_excel_dm")
    os.makedirs(tmpdir, exist_ok=True)
    work = os.path.join(tmpdir, os.path.basename(dst))
    res, ps1 = work + ".json", work + ".ps1"
    for f in (work, res):
        if os.path.exists(f):
            os.remove(f)
    shutil.copyfile(src, work)
    before = os.path.getmtime(work)
    with open(ps1, "w", encoding="utf-8-sig") as f:
        f.write(build_ps(work, res, **kw))
    try:
        r = subprocess.run(["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", ps1],
                           capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=1800)
        if r.returncode or not os.path.exists(res):
            raise RuntimeError(f"Dựng Data Model thất bại cho {dst}:\n{r.stdout}\n{r.stderr}")
        out = json.load(open(res, encoding="utf-8"))
        if out.get("error"):
            raise RuntimeError(f"Lỗi khi dựng {dst}:\n{out['error']}\nNhật ký: {out.get('log')}")
        if os.path.getmtime(work) <= before:
            raise RuntimeError(f"Excel không lưu {dst}")
        shutil.copyfile(work, dst)
        return out.get("cells", []), out.get("log", [])
    finally:
        for f in (work, res, ps1):
            if os.path.exists(f):
                try:
                    os.remove(f)
                except OSError:
                    pass


def _sheet_get_or_add(name):
    return f"""Com {{
  $sh = $null
  try {{ $sh = $wb.Worksheets.Item({_ps(name)}) }} catch {{ }}
  if ($null -eq $sh) {{ $sh = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)); $sh.Name = {_ps(name)} }}
}}
$sh = Com {{ $wb.Worksheets.Item({_ps(name)}) }}"""


_HEADER = r"""
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
$log = $null
try {
  Com { $x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false }
  Com { $x.AskToUpdateLinks = $false; $x.EnableEvents = $false }
  $wb = Com { $x.Workbooks.Open($work, 0, $false) }
  if (Com { $wb.ReadOnly }) { throw "Excel mo file o che do chi doc: $work" }
"""

_FOOTER = r"""
  foreach ($c in $cellsToRead) {
    $shn, $addr = $c.Split('!', 2)
    $cellsOut += [string](Com { $wb.Worksheets.Item($shn).Range($addr).Text })
  }
  Com { $wb.Save() }
  Com { $wb.Close($false) }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
} catch {
  $errOut = $_.Exception.ToString() + "`n" + $_.InvocationInfo.PositionMessage
} finally {
  if ($mine.Count -gt 0) { try { Com { $x.Quit() } } catch { } }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  $x = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Start-Sleep -Milliseconds 800
  foreach ($procId in $mine) { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue }
}
$result = @{ error = $errOut; cells = $cellsOut; log = @($log) }
[IO.File]::WriteAllText(__RES__, (ConvertTo-Json -InputObject $result -Compress -Depth 4), (New-Object Text.UTF8Encoding $false))
"""
