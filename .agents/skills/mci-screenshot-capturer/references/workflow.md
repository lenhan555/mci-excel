# Quy trình nguồn đã chuyển từ Claude

Đọc phần thích nghi trong ../SKILL.md trước; phần đó thay thế mọi
chỉ dẫn Mac/Claude không phù hợp bên dưới.

Bạn chụp ảnh minh họa cho khóa **Excel for Business Intelligence Analyst** của MCI. Thư mục gốc trên Mac: `/Users/leopham/Library/CloudStorage/OneDrive-Personal/MCI Excel`; trên Windows: `D:\Onedrive\MCI Excel`. Hai máy dùng chung thư mục này qua OneDrive, mỗi lần ghi mất vài giây đến khoảng một phút mới sang máy kia. Trả lời bằng tiếng Việt.

Việc của bạn kết thúc ở chỗ **PNG nằm đúng tên trong `Materials/B<n>/Images/`**. Ảnh vào deck/handout do `session-builder` (chế độ `chot`) dựng lại; bạn không chạy script dựng và không ghi `.pptx` / `.docx`.

## Đọc trước khi chụp

1. `AGENTS.md`, mục Ảnh minh họa: Excel 365 Windows, zoom 100%, dữ liệu MCI Mart; ảnh handout chụp trên bản `_HocVien` để tên sheet khớp file học viên nhận.
2. `Materials/B<n>/Images/_DANH_SACH_ANH.md`: đoạn đầu (điều kiện máy chụp, file nào dùng bản sao, trạng thái thư mục dữ liệu) và bảng Deck / Handout (Nội dung cần chụp · File · sheet · vùng).
3. Với ảnh slide `S<nn>`: phần `### Slide <nn> · …` trong `1_Sessions/B<n>_…/Slide_Script_B<n>.md`: lời giảng nói thao tác nào trước khi tới trạng thái cần chụp, và con số nào phải hiện trên màn hình. Với ảnh handout `H<nn>`: lời gọi `shot=(…)` / `figure(…)` trong `2_Dataset/_generator/build_docs_b<n>.py`.

## Bộ công cụ trong `1_Sessions/B<n>_…/_capture/`

- `watcher.ps1` chạy trên Windows: lấy file `cmd/*.ps1` có tên nhỏ nhất, xóa file đó, chạy, ghi kết quả vào `out/<tên>.log` (UTF-8 có BOM, dòng cuối `DONE`). Lệnh chứa `#STOP` làm watcher dừng; watcher tự dừng sau 4 giờ.
- `lib.ps1` (dot-source ở đầu mỗi lệnh): `Snap name.png`, `SnapN 'S11a' 3` (chụp 3 ảnh toàn màn hình `out/S11a_1..3.png`), `Place 'phần tiêu đề cửa sổ' x y w h`, `Wins` (liệt kê cửa sổ), `Get-Xl`, `Get-XlFor 'đường dẫn'` (Excel qua COM), `Retry { … }` (thử lại khi Excel bận). Không đặt biến `$log`, `$f`, `$code`, `$res`, `$stop` (biến của watcher).
- `crop.py <ảnh gốc> <ảnh ra> x0 y0 x1 y1`: cắt theo pixel thật. `cropf.py <tên> x0 y0 x1 y1 [ảnh ra]`: chọn ảnh ít nhiễu nhất trong `out/<tên>_*.png` rồi cắt, nhưng tọa độ tính theo khung hiển thị 1389 px của màn hình 1920 px; màn hình khác cỡ thì dùng `crop.py`.
- Buổi chưa có `_capture/`: chép `watcher.ps1`, `lib.ps1`, `crop.py`, `cropf.py` từ `1_Sessions/B4_Power_Query/_capture/`, sửa `$CapDir`, `$PracDir` và dòng chú thích đầu trong `lib.ps1` cho đúng buổi. Không sửa logic watcher.

## Giao thức gửi lệnh

1. Ghi lệnh vào `_capture/cmd/<NNN>_<việc>.ps1`. `NNN` là số 3 chữ số lớn hơn mọi số đã có trong `out/`, để chạy đúng thứ tự.
2. Dòng đầu mỗi lệnh: `. 'D:\Onedrive\MCI Excel\1_Sessions\<thư mục buổi>\_capture\lib.ps1'`. In ra mọi thứ cần biết (`"opened $($wb.Name)"`, kết quả `Place`, `Wins`): log là cách duy nhất bạn thấy được máy Windows.
3. Gửi **từng lệnh một** và chờ `out/<NNN>_<việc>.log` có dòng `DONE` bằng vòng lặp Python (đọc `utf-8-sig`, `time.sleep(2)`, tối đa 3 phút). Hết giờ thì kiểm tra file lệnh còn trong `cmd/` hay không (còn = watcher không chạy hoặc OneDrive chưa đồng bộ).
4. Ảnh `out/*.png` về Mac sau khi đồng bộ: chờ file tồn tại và kích thước đứng yên rồi mới mở.

## Quy trình

**0. Kiểm tra kết nối và máy chụp.** Gửi lệnh `info`: culture/region, độ phân giải màn hình, tiến trình Excel, `Wins`, trạng thái quyền admin của watcher.
- Không có phản hồi: dừng, báo người dùng bật watcher trên Windows: `powershell -ExecutionPolicy Bypass -File 'D:\Onedrive\MCI Excel\1_Sessions\<thư mục buổi>\_capture\watcher.ps1'`.
- `Get-XlFor` trả về rỗng ("no COM") khi watcher và Excel chạy khác mức quyền: nhờ người dùng chạy lại watcher cùng mức quyền với Excel (B4 cần "Run as Administrator").
- So điều kiện ở đầu `_DANH_SACH_ANH.md` (ví dụ B4: Region English (United States), `Practice/Data_Sales/` đúng 4 file). Không đạt thì không chụp những ảnh phụ thuộc điều kiện đó, ghi vào báo cáo.

**1. Lập kế hoạch.** Với mỗi ảnh cần chụp: file nguồn, sheet/query, vùng, trạng thái cần dựng (theo lời giảng). Gom theo file nguồn để mỗi file mở một lần. Chia hai loại:
- **Tự làm được qua COM:** mở workbook, chọn sheet, chọn/cuộn tới vùng, zoom 100%, ẩn tạm cột, lọc, mở và xếp cửa sổ CSV, đọc mã M của query, Refresh.
- **Cần người dùng thao tác:** hộp thoại (Merge, Group By, Change Type with Locale, PivotTable Options, Manage Parameters, Insert Calculated Item…), các bước trong Power Query Editor, menu đang mở, ScreenTip, F2/Range Finder. Viết hướng dẫn từng bước ngắn, lấy từ lời giảng của slide, để người dùng dựng trạng thái trên Windows. Khi được gọi tiếp với "đã dựng xong <tên ảnh>", bạn chỉ cần `Place` + `SnapN` + cắt.

**2. Làm trên bản sao.** Chép file nguồn sang `_capture/<tên>_Chup.xlsx` bằng `Copy-Item` ngay trên Windows rồi mở bản sao. Ảnh handout dùng bản `_HocVien` nếu file có bản đó (Demo); file Practice không có bản riêng thì dùng bản Practice. **Không bao giờ lưu file gốc**; đóng workbook bằng `$wb.Close($false)`. Thay đổi thư mục dữ liệu (ví dụ chép tạm file tháng mới) thì trả lại trạng thái cũ ngay sau khi chụp.

**3. Chụp.** Đặt cửa sổ ở vị trí và kích thước cố định cho cả buổi (ví dụ `Place '<tên cửa sổ>' 0 0 1366 900`) để các ảnh đồng đều; zoom 100%; `SnapN '<S11a>' 3`.

**4. Xem, cắt, tự kiểm.**
- Mở ảnh toàn màn hình bằng công cụ xem ảnh/đọc file hiện có. Lấy kích thước thật bằng Pillow; nếu ảnh hiển thị nhỏ hơn thì quy đổi tọa độ trước khi cắt.
- Chỉ giữ phần Excel cần cho nội dung ảnh: tiêu đề cột, số dòng, Applied Steps… như mô tả ghi. Không lấy cửa sổ khác, thanh taskbar, cửa sổ Codex, đường dẫn cá nhân. Không phóng to/thu nhỏ, không vẽ thêm, không sửa nội dung ảnh; chỉ cắt.
- Lưu vào `_capture/shots/<tên file đích>.png`, mở lại bằng công cụ xem ảnh/đọc file hiện có và so **từng chi tiết** với cột "Nội dung cần chụp" và lời giảng: con số, dòng/ô được chọn, hộp thoại, lỗi hiện trong ô, vùng. Sai do thao tác thì chụp lại.
- Nếu màn hình thật khác mô tả (số khác, Excel không hiển thị như lời giảng nói): **không** cố tạo ra ảnh cho khớp. Ghi là "lệch mô tả", kèm ảnh làm bằng chứng; nguồn cần sửa là kịch bản, phiên chính sẽ quyết định.

**5. Lưu.** Chép `_capture/shots/<tên>.png` sang `Materials/B<n>/Images/<tên>.png`, đúng từng ký tự tên trong `_DANH_SACH_ANH.md`. Đã có ảnh cùng tên (chụp lại) thì chép ảnh cũ vào `_capture/backup_anh_cu/<YYYY-MM-DD>/` trước.

**6. Dọn.** Đóng mọi workbook bạn đã mở, không lưu. Kiểm tra mtime file gốc không đổi và thư mục dữ liệu đúng trạng thái ban đầu. Không đóng ứng dụng hay cửa sổ khác của người dùng. Xong hết và không còn ảnh chờ thao tác tay thì gửi lệnh `#STOP`; còn ảnh chờ thì để watcher chạy.

## Không được làm

- Chạy script trong `_generator/`, ghi `.pptx` / `.docx`, sửa kịch bản, `_DANH_SACH_ANH.md` hay script.
- Lưu hay sửa file trong `Practice/`, `Demo/`, `Homework/`.
- Chụp bằng Excel trên Mac, hoặc dựng ảnh giả lập giao diện Excel.
- Xóa file trong `out/`, `shots/` hay các thư mục backup.

## Báo cáo

1. **Kết nối:** watcher có chạy không, mức quyền, region, độ phân giải; điều kiện máy chụp nào không đạt.
2. **Bảng ảnh:** Tên file · Trạng thái (Đã lưu / Chờ thao tác tay / Lệch mô tả / Lỗi) · Ảnh gốc đã dùng (`out/…`) · Ghi chú kiểm tra.
3. **Hướng dẫn thao tác tay** cho các ảnh "Chờ thao tác tay", nhóm theo file, đánh số từng bước.
4. **Lệch mô tả:** ảnh nào, mô tả nói gì, màn hình thật hiện gì, đề xuất sửa ở kịch bản/`_DANH_SACH_ANH.md`.
5. **Dọn dẹp:** file đã đóng, thư mục dữ liệu đã trả lại, watcher đã dừng hay còn chạy.
6. **Bước tiếp theo** cho phiên chính: nhờ người dùng thao tác tay rồi gọi lại bạn, hoặc chạy `session-builder` chế độ `chot`.
