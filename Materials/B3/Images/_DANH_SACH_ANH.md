# Ảnh minh họa Buổi 3 (deck + handout)

Chụp trên Excel 365 **Windows**, zoom 100%, trên đúng file và vùng ghi dưới đây (đường dẫn tính từ `1_Sessions/B3_PivotTable/`). Lưu thành **PNG** vào thư mục này, **đúng tên file**. Ảnh được script tự chèn vào khung [CHỤP]; không chèn tay vào `.pptx` / `.docx` (bị ghi đè khi dựng lại).

Nhiều ảnh cần làm thao tác trước khi chụp (tạo Pivot, Group, dán dòng, double-click): làm đúng các bước ở sheet `Huong_dan` của file demo (hoặc lời giảng của slide tương ứng trong `Slide_Script_B3.md`) rồi chụp; chụp xong đóng file **không lưu**.

Sau khi thêm ảnh, chạy trong `2_Dataset/`:

```bash
python3 _generator/build_slides_b3.py           # in ra các khung [CHỤP] còn thiếu ảnh
python3 _generator/build_student_slides.py B3
python3 _generator/build_docs_b3.py             # rồi xuất lại B3_Handout.pdf bằng Word
```

## Deck `Slide_B3.pptx` (S = số slide)

Chụp trên bản giảng viên của file Demo. Bảng này trùng với bảng "Danh sách ảnh cần chụp" trong `Slide_Script_B3.md`.

| Tên file | Nội dung cần chụp | File · sheet · vùng | Slide |
|---|---|---|---|
| S09_create_pivot.png | Hộp Create PivotTable: Table/Range = tblSales, New Worksheet | `Demo/B3_Demo_Pivot.xlsx` · Sales · chọn A2 → Insert → PivotTable → From Table/Range | 9 |
| S10_field_list.png | Pivot Region × Thang (3 × 3 + Grand Total) và khung PivotTable Fields: Region ở Rows, Thang ở Columns, Sum of Revenue ở Values | `Demo/B3_Demo_Pivot.xlsx` · PV_MienThang (sheet vừa tạo) · A3:E8 + khung Fields | 10 |
| S11_pivot_vs_sumifs.png | Pivot Miền × Tháng và khối đối soát; chọn B13, thanh công thức hiện `=GETPIVOTDATA("Doanh thu",$A$4,"Region","Bắc","Thang",3)`; chênh lệch 0 | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_MienThang · A4:D15 | 11 |
| S12a_refresh_table_vs_range.png | Sau khi dán 10 dòng tháng 4 và Refresh All: PV_Table tổng 942.937.000, PV_Vung vẫn 941.454.000 (hai cửa sổ cạnh nhau) | `Demo/B3_Demo_Refresh.xlsx` · PV_Table, PV_Vung · A3:B7 | 12 |
| S12b_change_source.png | Hộp Change PivotTable Data Source: `Sales_Vung!$A$1:$I$6198` | `Demo/B3_Demo_Refresh.xlsx` · PV_Vung | 12 |
| S14_value_field_settings.png | Hộp Value Field Settings (Custom Name: Doanh thu) và hộp Format Cells → Number, Use 1000 Separator | `Demo/B3_Demo_Pivot.xlsx` · PV_MienThang | 14 |
| S15a_compact.png | Region + Store_Code trên hàng, Compact Form | `Demo/B3_Demo_Pivot.xlsx` · PV_MienThang · A3:F20 | 15 |
| S15b_outline.png | Cùng Pivot, Outline Form | `Demo/B3_Demo_Pivot.xlsx` · PV_MienThang · A3:G20 | 15 |
| S15c_tabular.png | Cùng Pivot, Tabular Form + Repeat All Item Labels | `Demo/B3_Demo_Pivot.xlsx` · PV_MienThang · A3:G20 | 15 |
| S18a_sort_thu.png | Pivot Thu × Doanh thu sắp xếp giảm dần: Sun 159.845.000 … Thu 111.408.000 | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_Thu · A4:B12 | 18 |
| S18b_top5.png | Hộp Top 10 Filter (Top 5 Items by Sum of Revenue) và kết quả 5 sản phẩm, Hộp quà Tết cao cấp 135.850.000 đứng đầu | `Demo/B3_Demo_Pivot.xlsx` · PV_Top5 (sheet vừa tạo) · A3:B9 | 18 |
| S19a_group_7days.png | Hộp Grouping (Days, Number of days: 7) và Pivot 13 tuần (08–14/01 118.778.000; 29/01–04/02 31.567.000) | `Demo/B3_Demo_Pivot.xlsx` · PV_Tuan (sheet vừa tạo) · A3:B17 | 19 |
| S19b_cannot_group.png | Thông báo "Cannot group that selection." và 3 mục ngày dạng chữ ở đầu danh sách | `Demo/B3_Demo_Loi_Du_lieu.xlsx` · Pivot mới từ tblSalesLoi · Order_Date ở Rows | 19 |
| S20a_store_month.png | Pivot Store_Code × Thang, ô ST07 × 3 = 8.539.000 được chọn | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_CuaHang · A4:E18 | 20 |
| S20b_show_details.png | Sheet chi tiết sau khi double-click ô D12 (ST07 × 3), sắp xếp Order_Date: 55 dòng, từ 10/03 thưa hẳn | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · sheet mới sau double-click PV_CuaHang!D12 · A1:J30 | 20 |
| S21a_slicer_connections.png | Slicer Store_Code (2 cột) cạnh Pivot Miền × Tháng và hộp Report Connections tick 3 Pivot | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_MienThang · A4:I20 | 21 |
| S21b_timeline.png | Timeline Order_Date (mức Months) phía trên PivotChart | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_Ngay · D4:N11 | 21 |
| S23_count_default.png | Pivot từ tblSalesLoi, Region × **Count of Revenue** (tổng 6.186) | `Demo/B3_Demo_Loi_Du_lieu.xlsx` · Pivot mới · A3:B7 + khung Values | 23 |
| S24_pct_column.png | Category × Thang, % of Column Total (Bánh kẹo 45,6% → 30,9% → 25,5%) | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_NganhHang · A4:E12 | 24 |
| S25_mom_running.png | Thang × Doanh thu, MoM %, Lũy kế, Số đơn, AOV, Gross Margin % | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_Thang · A4:G9 | 25 |
| S26_calc_field.png | Hộp Insert Calculated Field: Name AOV, Formula `=Revenue/So_don` | `Demo/B3_Demo_Pivot.xlsx` · PV_Thang (sau khi thêm cột So_don) | 26 |
| S28a_pie_wrong.png | PivotChart Pie doanh thu 3 tháng | `Demo/B3_Demo_Pivot.xlsx` · PV_Thang | 28 |
| S28b_line_right.png | Cùng dữ liệu, PivotChart Line | `Demo/B3_Demo_Pivot.xlsx` · PV_Thang | 28 |
| S29_pivotchart_ngay.png | PivotChart đường doanh thu theo ngày Quý 1 và Timeline | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_Ngay · D4:P33 | 29 |

## Handout `B3_Handout.docx` (H = thứ tự ảnh)

File Demo dùng bản `_HocVien`, để tên sheet trong ảnh khớp với file học viên nhận.

| Tên file | Nội dung cần chụp | File · sheet · vùng | Mục trong handout |
|---|---|---|---|
| H01_analyze_tab.png | Tab PivotTable Analyze trên Ribbon khi chọn một ô trong Pivot (nhóm Active Field, Group, Filter, Data, Calculations) | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_MienThang · Ribbon | 1. Tóm tắt |
| H02_show_values_index.png | Pivot Category × Region, Show Values As → Index (Chăm sóc cá nhân × Nam = 1,17) | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · Pivot mới từ tblSales · A3:E10 | 2. Show Values As mở rộng |
| H03_calculated_item.png | Hộp Insert Calculated Item (Hàng FMCG = 'Đồ uống' + 'Sữa' + 'Bánh kẹo') và Pivot có Grand Total 1.406.955.000 | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · Pivot mới Category × Revenue | 3. Calculated Item |
| H04_report_filter_pages.png | 12 sheet ST01 … ST12 sau Show Report Filter Pages | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · Pivot Category × Thang, Store_Code ở Filters · thanh tab sheet | 4. Report Filter Pages |
| H05_field_buttons.png | PivotChart doanh thu theo thứ, menu Field Buttons đang mở | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_Thu · D4:L22 | 5. PivotChart |
| H06_subtotal.png | Hộp Subtotal (At each change in Region, Sum, Revenue) và vùng thu gọn mức 2: 3 dòng tổng miền | `Demo/B3_Demo_Refresh_HocVien.xlsx` · Sales_Vung · A1:I6192 | 6. Subtotal |
| H07_cf_pivot.png | Pivot PV_CuaHang tô Color Scale, lựa chọn “All cells showing ‘Doanh thu’ values for ‘Store_Code’ and ‘Thang’” | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_CuaHang · A4:E18 | 7. PRO TIP Conditional Formatting |
| H08_pivot_options.png | Hộp PivotTable Options, tab Data | `Practice/B3_MCI_Mart_Q1_2025_Answer.xlsx` · PV_MienThang → PivotTable Options | 7. PRO TIP PivotTable Options |

Ảnh chưa có thì deck và handout giữ khung "[CHỤP] …" (viền nét đứt). Trước khi phát tài liệu cho học viên, không còn khung nào.

**Trạng thái (28/09/2026):** đã chụp đủ 24 ảnh deck + 8 ảnh handout trên Excel 365 Windows (1920×1200, zoom 100%, Office Theme White); deck, bản học viên, handout và PDF đã dựng lại, không còn khung [CHỤP].
