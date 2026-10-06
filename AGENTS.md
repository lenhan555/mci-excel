# Quy tắc cho Codex — dự án MCI Excel

Cấu trúc thư mục và quy ước chung: xem `README.md`. Tài liệu gốc: `Course_Plan.md`.

## Mỗi demo trong kịch bản slide phải có file demo

- Mọi thao tác giảng viên làm mẫu trong `Slide_Script_B<n>.md` — dòng *(→ demo …)*, *(→ kéo xuống…)*, *(→ cuộn tới…)*, hoặc lời giảng "các bạn làm cùng mình" trên một bảng số liệu cụ thể — phải có file Excel để mở ra làm ngay.
- Được dùng lại file đã có (`Practice/…Dirty`, `Practice/…Answer`) **nếu demo chạy được nguyên trạng trên file đó**. Demo cần dữ liệu riêng (bảng nhỏ tự gõ, dữ liệu còn giữ lỗi, dòng để dán thêm…) thì tạo file trong `1_Sessions/B<n>_…/Demo/`.
- Tên file: `B<n>_Demo_<Chủ_đề>.xlsx`. Mỗi file có sheet `Huong_dan`: slide · thao tác · kết quả mong đợi.
- File demo dựng bằng `2_Dataset/_generator/build_demos.py`, số liệu tính từ dataset, không gõ tay. Sau khi dựng, mở bằng Excel để kiểm tra kết quả mong đợi (ví dụ đúng số ô ra lỗi, đúng tổng).
- **Không** gộp sheet demo vào file `Practice/`: file Answer của buổi trước là file khởi đầu phát cho học viên buổi sau.
- Khi viết hoặc sửa kịch bản: liệt kê file demo trong bảng "File dùng trong buổi" (ghi rõ slide nào), và cập nhật `2_Dataset/README.md`.
- **Ghi rõ file trong lời giảng và speaker notes.** Mỗi chỉ dẫn demo phải nêu đường dẫn file tính từ thư mục buổi, kèm sheet (và ô, nếu cần), ví dụ: *(→ demo: mở `Demo/B1_Demo_PctShare.xlsx`, sheet Demo_PctShare)*. Không viết chung chung như "file đáp án", "file case", "(→ demo)". Kịch bản và notes trong `.pptx` phải khớp nhau.
- Khi review kịch bản hoặc deck: đối chiếu từng demo với file demo; con số nói trong lời giảng phải ra đúng khi làm trên file đó.

## Bản giảng viên và bản học viên

- File nào phát cho học viên thì phải có bản riêng hậu tố `_HocVien`, dựng lại bằng script mỗi khi bản giảng viên đổi:
  - **Slide:** `Slide_B<n>_HocVien.pptx` không có speaker notes và **không có đáp án quiz** (`_generator/build_student_slides.py` xóa các shape tên `Answer …`, `Dim …` và animation hiện đáp án). Khi làm slide quiz mới, đặt tên shape đáp án theo quy ước này.
  - **Demo:** `Demo/…_HocVien.xlsx` không có câu chỉ dẫn dành cho giảng viên ("làm lại trước lớp", "demo cho lớp xem"…), vì học viên đọc thấy sẽ nghĩ giảng viên đang được người khác hướng dẫn. Viết thành bước làm cho người học. Tiêu đề "DEMO" đổi thành "VÍ DỤ".
  - **Practice:** không cần bản riêng; file đề (Dirty) và đáp án (Answer) đã tách sẵn.
- Bản giảng viên của file demo mở ra ở sheet làm việc, sheet `Huong_dan` để cuối, để khi chiếu lên màn hình học viên không thấy hướng dẫn.
- Bài tập về nhà là **bắt buộc** (không dùng chữ "tự luyện"); mỗi bài có file đề trong `Homework/` và file `_Answer` có sheet `Kiem_tra` để chấm.
- Bài tập về nhà là **bước tập dượt cho capstone B7** (`Course_Plan.md`, Nguyên tắc 11): đặt trong một tình huống công việc, có lỗi cài sẵn chỉ lộ ra khi đối soát, gắn với tiêu chí rubric B7 (`1_Sessions/B7_Capstone/Rubric_B7_Capstone.md`). Khi soạn bài mới, ghi rõ bài tập dượt tiêu chí nào.
- Trước khi sửa `.pptx`, kiểm tra PowerPoint không đang mở file đó; nếu đang mở, nhờ người dùng lưu và đóng trước, tránh bị ghi đè.

## Ảnh minh họa: ưu tiên ảnh chụp thật

- Minh họa thao tác, hàm, kết quả trong Excel (slide, handout) **ưu tiên ảnh chụp thật** Excel 365 Windows trên chính file dataset / demo / practice của khóa (theo Design System: zoom 100%, dữ liệu MCI Mart). Không tự dựng lưới mô phỏng Excel bằng shape thay cho ảnh chụp.
- **[VẼ]** chỉ dùng cho sơ đồ khái niệm không có trên màn hình Excel: quy trình, hình xương sống Chuỗi C, viên gạch, trục số thống kê, thẻ số KPI.
- Chưa chụp được thì để khung chờ ảnh ghi rõ **[CHỤP]**: nội dung, file, sheet, ô; liệt kê trong bảng "Danh sách ảnh cần chụp" của kịch bản. Không lặng lẽ thay bằng hình vẽ.

**Buổi nào cũng phải có ảnh minh họa** (Course Plan, nguyên tắc 2 và 3):
- Mỗi slide **CÔNG CỤ** có ít nhất một ảnh chụp (hoặc khung [CHỤP] khi chưa có ảnh). Slide KHÁI NIỆM có demo trên Excel nên có ảnh; slide BÌA, CHIA PHẦN, QUIZ, THỰC HÀNH không bắt buộc.
- Mỗi mục công cụ trong handout có ít nhất một ảnh ở những chỗ giao diện hoặc kết quả trên màn hình giúp hiểu hơn bảng chữ (hộp thoại, vùng tràn, định dạng có điều kiện, kết quả trên dữ liệu case…).
- Khi soạn một buổi mới, lập danh sách ảnh ngay từ đầu, cùng lúc với kịch bản, không để đến cuối.

**Quy ước chèn ảnh:**
- Thư mục: `Materials/B<n>/Images/`. Tên file: `S<số slide 2 chữ số>_<chủ_đề>.png` cho slide (nhiều ảnh cùng slide thì thêm a, b: `S07a_…`), `H<số thứ tự 2 chữ số>_<chủ_đề>.png` cho handout. PNG, chụp Excel 365 Windows, zoom 100%.
- `_DANH_SACH_ANH.md` trong thư mục đó có hai bảng Deck và Handout, cột: Tên file · Nội dung cần chụp · File · sheet · vùng · Slide/Mục. Bảng Deck trùng với bảng "Danh sách ảnh cần chụp" trong kịch bản. Ảnh cho handout chụp trên bản `_HocVien` để tên sheet khớp file học viên nhận.
- Script dựng slide (`build_slides_b<n>.py`) và handout (`build_docs…py`) **tự chèn** ảnh có trong thư mục vào đúng khung: giữ tỉ lệ, canh giữa khung, viền hairline, không bo góc. Thiếu ảnh thì giữ khung [CHỤP] ghi tên file cần lưu, và script in ra danh sách khung còn thiếu.
- **Chân slide:** ghi file demo chứa ảnh, dạng "Ảnh: Demo/B2_Demo_Lookup.xlsx, sheet Demo_XLOOKUP (cột B)", kèm ghi chú nội dung nếu cần. **Không** ghi trích dẫn tài liệu PDF nguồn ("Formulas PDF p…", "Statistics PDF p…") trên slide; số trang PDF chỉ để trong kịch bản (`[PDF p…]`) cho người soạn. Không để nhãn sản xuất như [VẼ], [CHỤP] trên slide đã có ảnh; vùng chụp chi tiết (file · sheet · vùng → tên file ảnh) nằm trong khung [CHỤP], biến mất khi ảnh đã chèn.
- Không chèn hay sửa ảnh bằng tay trong `.pptx` / `.docx`: lần dựng lại sau sẽ ghi đè.
- Thêm ảnh xong: chạy lại script dựng slide → `build_student_slides.py B<n>` → script handout → xuất PDF bằng Word.
- Trước khi phát tài liệu cho học viên: deck `_HocVien` và handout PDF **không còn khung [CHỤP]**. Còn thì báo người dùng danh sách ảnh thiếu thay vì phát.
- B1 dựng deck trước khi có quy ước này (ảnh đã chèn sẵn, tên cũ `H01…` và `_capture/shots/`); giữ nguyên, không cần đổi tên.

## Infographic tóm tắt buổi

- Buổi nào cũng xuất một infographic tóm tắt: `Handout/B<n>_Infographic_TomTat.png` (A4 dọc, in được). Phát **riêng**, không chèn vào handout.
- Nguồn: `2_Dataset/_generator/infographic_b<n>.html`, dựng bằng `python3 _generator/build_infographic.py B<n>`. Sửa HTML rồi dựng lại, không sửa ảnh PNG. Lấy `infographic_b1.html` làm khuôn (Design System, cùng bố cục).
- Nội dung: câu hỏi mở buổi, vị trí buổi trong quy trình BI và 7 viên gạch, các mục công cụ chính (tên công cụ, phím tắt, lỗi hay gặp), khối số liệu để học viên tự đối chiếu file, câu insight cuối buổi, việc trước buổi sau (bài tập về nhà bắt buộc) và câu hỏi buổi sau.
- Số liệu trên infographic lấy từ file Answer / meta của buổi, không gõ tay; phải khớp kịch bản và handout. Buổi có `_generator/build_b<n>_meta.json` thì HTML viết `{{khóa|n}}` (số) hoặc `{{khóa|pct}}` (phần trăm), script tự điền (xem `infographic_b2.html`).
- Làm sau khi kịch bản và handout đã chốt; mỗi lần sửa kịch bản hay handout thì rà lại infographic và dựng lại. Liệt kê trong bảng "File dùng trong buổi" của kịch bản.

## Sửa file Excel đã lưu bằng Excel

File đã được mở và lưu bằng Excel (ví dụ file Answer) có Data Validation trỏ sang sheet khác, lưu trong `extLst`. openpyxl sẽ **xóa** phần này khi ghi lại. Muốn sửa thì dùng chính Excel (AppleScript), hoặc dựng lại file bằng script trong `_generator/`.

## Dựng tài liệu một buổi (Mac ↔ Windows)

Mac chạy mọi script dựng; Windows chỉ chụp ảnh (và chạy lệnh ghi "CHỈ WINDOWS" trong `2_Dataset/README.md`, ví dụ `build_b4.py`). Ảnh vào khung bằng cách chạy lại script dựng, không dán tay. Phiên chính điều phối các agent theo thứ tự:

1. `session-builder` chế độ `khung` (`.codex/agents/session-builder.toml`): dựng dữ liệu, deck, handout; trả về danh sách khung [CHỤP] còn thiếu, nhóm theo file nguồn.
2. `screenshot-capturer` (`.codex/agents/screenshot-capturer.toml`): người dùng bật `_capture/watcher.ps1` trên Windows trước; agent chụp, cắt, tự kiểm, lưu PNG vào `Materials/B<n>/Images/`. Ảnh cần mở hộp thoại hay Query Editor thì agent trả hướng dẫn để người dùng dựng trạng thái, rồi gọi lại agent để chụp.
3. `session-builder` chế độ `chot`: sao lưu, dựng lại deck → `_HocVien`, handout → PDF, infographic; kiểm tra không còn [CHỤP].
4. `consistency-checker` (mục dưới).

Bước 2 và 3 làm được theo từng đợt ảnh. Trước khi agent ghi `.pptx` / `.docx`, hỏi người dùng đã đóng các file đó trên cả Mac và Windows chưa.

## Chấm bài tập về nhà

Bài học viên nộp tải về `1_Sessions/B<n>_…/Homework/Bai_nop/`, rồi gọi agent `homework-grader` (`.codex/agents/homework-grader.toml`). Agent chấm theo De_bai và file `_Answer` bằng script `2_Dataset/_generator/grade_b<n>.py`, ghi nhận xét từng học viên, `Tong_hop_B<n>.csv` và `Bao_cao_lop_B<n>.md` vào `Homework/Cham_bai/`. Không sửa file học viên nộp.
- **Điểm số làm tròn thành số nguyên** (integer) ở mọi nơi: nhận xét, `Tong_hop_B<n>.csv`, báo cáo lớp.

- Đọc cả ô và XML drawing/text box để kiểm tra insight; openpyxl chỉ đọc ô có thể bỏ sót phần chữ. Không giả định có Cowork. Phần không đọc được ghi rõ Không kiểm được, cần GV xem.

## Văn phong tiếng Việt

Skill dùng chung: `.agents/skills/viet-van-phong-mci/SKILL.md` (kèm `references/`). Ba agent đọc nó trước khi viết hoặc kiểm chữ:
- `homework-grader` viết nhận xét học viên (sắc thái đào tạo, xưng "bạn") và báo cáo lớp (nội bộ, khách quan), hai lượt.
- `consistency-checker` kiểm văn phong ở mục K, chỉ báo cáo.
- `session-builder` viết báo cáo cuối theo skill và quét tell trên deck `_HocVien` và PDF vừa dựng, chỉ báo.

Quy ước đã chốt: giữ ký hiệu `→` trong `(→ demo …)` và `·` trong chân slide, không coi là tell; không đụng công thức, tên hàm, địa chỉ ô, tên file/sheet, mã dữ liệu; không dùng "tay nghề" (dùng "cách làm", "tư duy làm"). Sửa văn phong trong kịch bản thì dựng lại deck → `_HocVien` → handout theo quy trình ở mục trên.

## Rà soát tính nhất quán

Trước khi phát tài liệu của một buổi, hoặc sau khi sửa nhiều file, gọi agent `consistency-checker` (`.codex/agents/consistency-checker.toml`) để đối chiếu Course_Plan, kịch bản, deck, handout, demo, practice và homework. Agent chỉ báo cáo, không sửa file.


---

Các hướng dẫn Codex/Windows dưới đây thay thế phần trên nếu có mâu thuẫn về môi trường, đường dẫn công cụ hoặc việc xác nhận file đang mở.

# MCI Excel — hướng dẫn cho Codex

Workspace xây dựng và vận hành khóa Excel for Business Intelligence Analyst
của MCI: 7 buổi × 150 phút, case MCI Mart, capstone B7.
Trả lời và viết tài liệu bằng tiếng Việt.

## Nguồn và quy tắc

- Đọc README.md và các phần liên quan của CLAUDE.md trước khi sửa tài liệu.
  CLAUDE.md là quy ước nội dung dùng chung; phần chỉ dành cho Claude/Mac được
  thay bằng hướng dẫn Codex/Windows trong các skill tương ứng.
- Course_Plan.md là nguồn tổng thể. Thay đổi chương trình thì sửa plan trước,
  sau đó kịch bản và script sinh tài liệu. Khi đối soát con số/hành vi Excel,
  dùng file và ảnh thật làm bằng chứng; báo sai lệch của tài liệu nguồn.
- Tái sử dụng 2_Dataset/_generator/; tra lệnh trong 2_Dataset/README.md.
  Không sửa tay sản phẩm sinh từ script rồi bỏ qua nguồn.
- Không ghi lại bằng openpyxl file đã được Excel lưu nếu có validation mở
  rộng, Pivot, Power Query hoặc Data Model. Dùng Excel COM hoặc dựng lại bằng
  backend phù hợp, có sao lưu trước khi ghi đè sản phẩm.
- Bài học viên nộp chỉ đọc. Không dùng _capture/ hoặc bản backup làm nguồn
  phát học viên. Tách bản giảng viên/học viên theo quy ước hiện có.
- Dùng ảnh chụp Excel thật. Tài liệu phát phải hết [CHỤP], bản _HocVien không
  lộ notes/đáp án, PDF và infographic được dựng từ nguồn hiện hành.

## Skill và agent

Skill dự án nằm trong .agents/skills/. Chọn đúng phạm vi và đọc SKILL.md:

| Yêu cầu | Skill | Agent |
|---|---|---|
| Dựng khung/chốt B<n> | mci-session-builder | session-builder |
| Chụp ảnh Excel B<n> | mci-screenshot-capturer | screenshot-capturer |
| Chấm bài B<n> | mci-homework-grader | homework-grader |
| Rà tài liệu | mci-consistency-checker | consistency-checker |
| Viết/kiểm văn phong | viet-van-phong-mci | Dùng trong vai trò liên quan |

Khi dựng trọn một buổi, phiên chính điều phối lần lượt: khung → chụp ảnh thiếu
→ chốt → rà nhất quán. Có thể giao agent đúng vai trò khi công cụ delegation
có sẵn; không spawn để chỉ trả lời câu hỏi đơn giản. Agent chỉ ghi trong phạm
vi vai trò được giao, không tự gọi thêm agent. Không chạy song song tác vụ
cùng ghi một file hay điều khiển Office. Phiên chính xử lý lỗi nguồn nằm ngoài
phạm vi agent, theo yêu cầu người dùng.

## Windows và quyền thực thi

Phát hiện Python/dependency và khả năng COM trước khi chạy. Không dùng đường
dẫn Mac cố định, không giả định AppleScript chạy trên Windows. Không nhập
permissions.allow của Claude thành quyền Codex; tuân thủ sandbox hiện tại.
Chỉ yêu cầu lưu/đóng tài liệu khi phát hiện đúng file sắp ghi đang mở hoặc có
xung đột chưa giải quyết. Không hỏi lại quyền đã được cấp trong phiên.
Hook dự án chỉ bổ sung ngữ cảnh, không dựng tài liệu/chấm bài/tự cấp quyền.
