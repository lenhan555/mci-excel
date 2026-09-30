# Ảnh minh họa Buổi 2 (deck + handout)

Chụp trên Excel 365 **Windows**, zoom 100%, trên đúng file và vùng ghi dưới đây (đường dẫn tính từ `1_Sessions/B2_Ham_va_KPI/`). Lưu thành **PNG** vào thư mục này, **đúng tên file**. Ảnh được script tự chèn vào khung [CHỤP]; không chèn tay vào `.pptx` / `.docx` (bị ghi đè khi dựng lại).

Sau khi thêm ảnh, chạy trong `2_Dataset/`:

```bash
python _generator/build_slides_b2.py           # in ra các khung [CHỤP] còn thiếu ảnh
python _generator/build_student_slides.py B2
python _generator/build_docs_b2.py             # rồi xuất lại B2_Handout.pdf bằng Word
```

## Deck `Slide_B2.pptx` (S = số slide)

Chụp trên bản giảng viên của file Demo.

| Tên file | Nội dung cần chụp | File · sheet · vùng | Slide |
|---|---|---|---|
| S04_grain.png | Đơn HD2501-00002 có 4 dòng (Line_No 1–4), chọn dòng 3–6 | `../B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx` · Sales · A1:K7 (ẩn tạm cột D–F, I–J) | 4 |
| S07_if_ifs.png | Cột Margin_%, IF, IFS sau khi kéo; P041 ra #DIV/0! và "Thấp" | `Demo/B2_Demo_Logic.xlsx` · Demo_IF · A1:H6 + dòng 42–43 (ẩn tạm 7–41) | 7 |
| S08_and_or_not.png | Cột I, J, K (AND, OR, NOT) ra TRUE/FALSE trên các dòng Sữa và Bánh kẹo | `Demo/B2_Demo_Logic.xlsx` · Demo_IF · A1:K1 + dòng 9–13, 22–28 | 8 |
| S09_iferror.png | P041, P042: #DIV/0!, "Chưa có giá"; L42 = IFERROR(…,0) ra 0% | `Demo/B2_Demo_Logic.xlsx` · Demo_IF · A41:L43 | 9 |
| S11_average_median.png | Khối I2:I6 (COUNT, AVERAGE, MEDIAN, MIN, MAX) và Status Bar khi chọn cột E | `Demo/B2_Demo_ThongKe.xlsx` · Demo_ThongKe · E1:I13 + Status Bar | 11 |
| S12_stdev_outlier.png | Khối I7:I13: STDEV.S, SKEW, Q1, Q3, IQR, ngưỡng 896.500, 138 đơn | `Demo/B2_Demo_ThongKe.xlsx` · Demo_ThongKe · G7:I13 | 12 |
| S13_rank.png | RANK.EQ 12 cửa hàng, F2 hiện `$D$2:$D$13` | `Demo/B2_Demo_ThongKe.xlsx` · Demo_RANK · A1:E13 | 13 |
| S16_sumifs_screentip.png | ScreenTip khi gõ `=SUMIFS(` (sum_range, criteria_range1, criteria1…) | `Demo/B2_Demo_SUMIFS.xlsx` · Demo_SUMIFS · ô trống | 16 |
| S17_sumifs_grid.png | Bảng Miền × Tháng sau khi kéo; F2 ở C13 hiện `$A13`, `C$11` | `Demo/B2_Demo_SUMIFS.xlsx` · Demo_SUMIFS · A11:E15 | 17 |
| S20_vlookup.png | VLOOKUP FALSE vs bỏ trống (P041, P21 → 71.500) | `Demo/B2_Demo_Lookup.xlsx` · Demo_VLOOKUP · A1:D7 | 20 |
| S21_index_match.png | INDEX/MATCH, F2 ở B2: chỉ ô A2 được tô màu (vùng trên sheet Product không tô) | `Demo/B2_Demo_Lookup.xlsx` · Demo_XLOOKUP · A1:B7 | 21 |
| S22_xlookup.png | XLOOKUP "Không có mã" và tra ngược tên → mã | `Demo/B2_Demo_Lookup.xlsx` · Demo_XLOOKUP · A1:F7 | 22 |
| S23_p021_duplicate.png | Product dòng 22–23: P021 hai lần | `Demo/B2_Demo_Lookup.xlsx` · Product · A1:E23 (ẩn tạm 2–20) | 23 |
| S24_trim_value.png | Khóa khác kiểu: #N/A, sửa bằng VALUE / TRIM, LEN = 5 | `Demo/B2_Demo_Lookup.xlsx` · Demo_KieuKhoa · A3:E14 | 24 |
| S26_text.png | LEFT, MID, RIGHT, TRIM, &, TEXT cho 6 cửa hàng | `Demo/B2_Demo_Text_Date.xlsx` · Demo_Text · A1:J7 | 26 |
| S27_date.png | EOMONTH, TEXT, WEEKDAY và NETWORKDAYS (23 → 17) | `Demo/B2_Demo_Text_Date.xlsx` · Demo_Date · A1:H15 | 27 |

## Handout `B2_Handout.docx` (H = thứ tự ảnh)

File Demo dùng bản `_HocVien`, để tên sheet trong ảnh khớp với file học viên nhận.

| Tên file | Nội dung cần chụp | File · sheet · vùng | Mục trong handout |
|---|---|---|---|
| H01_kpi_card.png | Sheet KPI_Card với 3 card đã điền và dòng giá trị Q1/2025 | `Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx` · KPI_Card · A4:D16 | 1. KPI Definition Card |
| H02_index_match_2d.png | Gõ INDEX/MATCH hai chiều ở G6, nhấn F2: vùng B6:D8, A6:A8, B5:D5 được tô màu; kết quả 129.438.000 | `Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx` · KPI_SUMIFS · A4:G9 | 4. INDEX / MATCH hai chiều |
| H03_search_find.png | Thêm tạm 2 cột SEARCH / FIND vào tblProduct; dòng "Hộp quà Tết cao cấp" khác nhau giữa hai cột | `Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx` · Product · A1:G41 | 5. SEARCH / FIND |
| H04_workday.png | Khối NETWORKDAYS (tháng 1: 23 → 17), ngày lễ H10:H15, một ô `=WORKDAY(DATE(2025,1,24),3,H10:H15)` ra 05/02/2025 | `Demo/B2_Demo_Text_Date_HocVien.xlsx` · Vi_du_Date · A8:H15 | 5. TODAY / NOW · WEEKDAY · WORKDAY |
| H05_cf_formula.png | Hộp New Formatting Rule với `=$T2="Cao"` và các dòng phân khúc Cao được tô màu | `Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx` · Sales · A1:T21 | 7. Conditional Formatting bằng công thức |
| H06_dynamic_array.png | Vùng tràn của `=SORT(UNIQUE(…))` (viền xanh), kết quả FILTER, ô A21 dùng `A4#` | `Demo/B2_Demo_SUMIFS_HocVien.xlsx` · Vi_du_DynamicArray · A3:I22 | 8. PRO TIP Dynamic Arrays |
| H07_duplicate_values.png | Duplicate Values tô màu hai dòng P021 (dòng 22–23) | `Practice/B2_Product_DuplicateKey.xlsx` · Product · A18:E26 | 8. Tìm giá trị trùng |

Ảnh chưa có thì deck và handout giữ khung "[CHỤP] …" (viền nét đứt). Trước khi phát tài liệu cho học viên, không còn khung nào.
