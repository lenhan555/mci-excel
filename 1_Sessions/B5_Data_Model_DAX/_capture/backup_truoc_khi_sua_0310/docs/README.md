# Excel for Business Intelligence Analyst — MCI

Khóa học 7 buổi × 150 phút. Tài liệu gốc: `Course_Plan.md`. Khi chỉnh nội dung, sửa plan trước, các file khác theo sau.

## Cấu trúc thư mục

```
MCI Excel/
├── README.md                          (file này)
├── CLAUDE.md                          quy tắc cho Claude (file demo, sửa file Excel)
├── .claude/agents/consistency-checker.md   agent rà tính nhất quán giữa các tài liệu
├── .claude/agents/homework-grader.md       agent chấm bài tập về nhà (đọc Homework/Bai_nop/, ghi Homework/Cham_bai/)
├── .claude/agents/session-builder.md       agent chạy script dựng một buổi trên Mac: chế độ khung (tạo khung [CHỤP]) và chot (chèn ảnh, _HocVien, PDF, infographic)
├── .claude/agents/screenshot-capturer.md   agent chụp ảnh Excel 365 trên Windows qua _capture/watcher.ps1, lưu PNG vào Materials/B<n>/Images/
├── Course_Plan.md                     plan tổng thể — tài liệu gốc
├── 1_Sessions/
│   ├── B1_Nen_tang_Excel_Du_lieu/     Slide_Script_B1.md · Slide_B1.pptx (+ _HocVien) · Handout/ · Practice/ · Demo/ · Homework/
│   │                                   (_capture/: thư mục làm việc để chụp ảnh, chứa bản sao cũ — không phát, không dùng làm nguồn)
│   ├── B2_Ham_va_KPI/                 Slide_Script_B2.md · Slide_B2.pptx (+ _HocVien) · Handout/ · Practice/ · Demo/ · Homework/
│   ├── B3_PivotTable/                 Slide_Script_B3.md · Slide_B3.pptx (+ _HocVien) · Handout/ · Practice/ · Demo/ · Homework/
│   ├── B4_Power_Query/                Slide_Script_B4.md · Slide_B4.pptx (+ _HocVien) · Handout/ · Practice/ (Data_Sales/, New_Month/, lookup, target) · Homework/ (không có Demo/)
│   ├── B5_Data_Model_DAX/             Slide_Script_B5.md · Slide_B5.pptx (+ _HocVien, còn 19 khung [CHỤP]) · Practice/ (đáp án) · Demo/ (2 file + _HocVien) · Homework/ (đề + đáp án) · Handout/ (chỉ infographic; chưa có handout) — dựng 02/10
│   ├── B6_Dashboard/                  (chưa có tài liệu)
│   └── B7_Capstone/                   Rubric_B7_Capstone.md
├── 2_Dataset/                         MCI_Mart_Master.xlsx · README.md (từ điển dữ liệu, câu chuyện cài sẵn) · _generator/
├── 3_Design_Reviews/                  Review_01, Review_02, Review_04 + Coverage_Audit.csv
├── 4_Reference/                       PDF nguồn, bảng tra Excel/Power BI, Curriculum_Guide.html (bản nháp)
└── Materials/                         Design/ (design system + theme Office) · B1/Images/ … B5/Images/ (ảnh chụp cho deck và handout; danh sách trong _DANH_SACH_ANH.md)
```

## Quy ước trong mỗi thư mục buổi

| Vị trí | Nội dung |
|---|---|
| Gốc thư mục buổi | Kịch bản slide (`Slide_Script_B<n>.md`) và deck dựng từ kịch bản (`Slide_B<n>.pptx`) |
| `Handout/` | Tài liệu phát cho học viên: handout, checklist (`.docx` + `.pdf`), infographic tóm tắt buổi (`B<n>_Infographic_TomTat.png`) |
| `Practice/` | File thực hành và đáp án. File đáp án của buổi trước là **file khởi đầu** của buổi sau |
| `Demo/` | File demo cho giảng viên (`B<n>_Demo_<Chủ_đề>.xlsx`), mỗi file có sheet `Huong_dan`; bản phát học viên là `…_HocVien.xlsx`. Không gộp vào file Practice |
| `Homework/` | Bài tập về nhà **bắt buộc**: file đề (phát học viên) và file `_Answer` (chỉ giảng viên / trợ giảng chấm). Bài học viên nộp tải về `Homework/Bai_nop/` (thư mục con theo lớp nếu có); agent `homework-grader` ghi nhận xét và bảng tổng hợp vào `Homework/Cham_bai/` |

**Bản giảng viên và bản học viên.** File nào phát cho học viên thì có bản riêng hậu tố `_HocVien`: slide bỏ speaker notes và đáp án quiz (`Slide_B<n>_HocVien.pptx`), file demo bỏ các câu chỉ dẫn dành cho giảng viên. Practice không cần bản riêng vì đã tách sẵn file đề (Dirty) và đáp án (Answer).

Handout, checklist và file thực hành được dựng bằng script trong `2_Dataset/_generator/`. Muốn sửa nội dung thì sửa script rồi dựng lại, đừng sửa tay file `.docx` (xem `2_Dataset/README.md`).
