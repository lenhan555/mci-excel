# Ảnh minh họa Buổi 6 (deck + handout)

Chụp trên Excel 365 **Windows**, zoom 100%, trên đúng file, sheet và cửa sổ ghi dưới đây (đường dẫn tính từ `1_Sessions/B6_Dashboard/`). Lưu thành **PNG** vào thư mục này, **đúng tên file**. Ảnh được script tự chèn vào khung [CHỤP]; không chèn tay vào `.pptx` / `.docx` (bị ghi đè khi dựng lại).

Ảnh S08–S14, S23–S24 chụp trên các file `Demo/` (bản `_HocVien`, để tên sheet khớp file học viên); ảnh dashboard S18–S21, S30, S33 chụp trên `B6_<HoTen>.xlsx` (từ file khởi đầu `../B5_Data_Model_DAX/Practice/B5_MCI_Mart_Answer.xlsx`) hoặc `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` khi đã dựng. Ba file `Demo/` đã dựng 03/10/2026 (`build_b6_demos.py`); `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` đã dựng nên ảnh S19–S21, S30, S33 chụp trên file này. **Đã chụp đủ 23 ảnh ngày 04/10/2026** (chụp riêng cửa sổ Excel bằng `_generator/snapwatch.ps1` + `_generator/snap.sh`; ảnh H06 dựng bằng `_generator/vba/B6Shots2.bas`). Chụp xong đóng file **không lưu**. Không bấm Refresh All.

Sau khi thêm ảnh, chạy trong `2_Dataset/`:

```bash
python3 _generator/build_slides_b6.py           # in ra các khung còn thiếu ảnh (hiện: đủ ảnh)
python3 _generator/build_student_slides.py B6
```

## Deck `Slide_B6.pptx` (S = số slide)

Bảng này trùng với bảng "Danh sách ảnh cần chụp" trong `Slide_Script_B6.md`.

| Tên file | Nội dung cần chụp | File · sheet · vùng | Slide |
|---|---|---|---|
| S08_dashboard_xau.png | Dashboard xấu: 12 visual, pie 3D, trục cắt khúc, không target (toàn trang) | `Demo/B6_Demo_DashboardXau.xlsx` · Dashboard_Xau | 8 |
| S09_truc_y.png | Hai biểu đồ cột T3/2024 và T4/2024 cạnh nhau: trục từ 236.000.000 và từ 0 | `Demo/B6_Demo_DashboardXau.xlsx` · Demo_TrucY | 9 |
| S12_combo_nganh.png | Combo: cột Doanh thu + đường GM% (trục phụ) theo 6 ngành | `Demo/B6_Demo_Chart.xlsx` · Demo_Combo · A3:C9 | 12 |
| S13_waterfall.png | Waterfall từ Target 285.000.000 qua 12 cửa hàng đến Actual 269.608.000, hai cột Total | `Demo/B6_Demo_Chart.xlsx` · Demo_Waterfall · A3:B17 | 13 |
| S14a_histogram.png | Histogram giá trị đơn Q1, bin 100.000, đuôi phải | `Demo/B6_Demo_Chart.xlsx` · Demo_Phan_phoi | 14 |
| S14b_scatter.png | Scatter 12 cửa hàng (tuổi × doanh thu) + trendline + R², ô CORREL = 0,78 | `Demo/B6_Demo_Chart.xlsx` · Demo_Scatter | 14 |
| S18a_cubevalue.png | Ô CUBEVALUE trên thanh công thức, gợi ý autocomplete đang hiện | `B6_<HoTen>.xlsx` · Data_Dash · B3 | 18 |
| S18b_data_dash.png | Vùng `Data_Dash` 6 KPI: giá trị · target · mức đạt (tháng 4) | Như trên · A3:E9 | 18 |
| S19_rag.png | 6 thẻ KPI tháng 4/2025 với biểu tượng ✔ ▲ ✘ và nhãn chữ | `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` · Dashboard | 19 |
| S20_report_connections.png | Hộp Report Connections của Slicer Tháng: tick các Pivot, bỏ Pivot combo | `Practice/…` · chuột phải Slicer | 20 |
| S21_sparkline.png | Thẻ Doanh thu có sparkline T1–T4 | `Practice/…` · Dashboard | 21 |
| S23_trendline.png | Line chart doanh thu ngày + Linear Trendline + Forward 30 + R² 0,045 | `Demo/B6_Demo_DuBao.xlsx` · Demo_Trendline | 23 |
| S24a_forecast_dialog.png | Hộp Create Forecast Worksheet với Options mở rộng (Seasonality = 7, 95%) | `Demo/B6_Demo_DuBao.xlsx` · Data → Forecast Sheet | 24 |
| S24b_forecast_chart.png | Biểu đồ dự báo 30 ngày, đường Forecast và hai đường Confidence Bound | `Demo/B6_Demo_DuBao.xlsx` · Demo_ETS | 24 |
| S30_dashboard_t4.png | Dashboard hoàn chỉnh, tháng 4/2025, toàn trang | `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` · Dashboard | 30 |
| S33_protect_sheet.png | Hộp Protect Sheet, tick Use PivotTable & PivotChart | `Practice/…` · Review → Protect Sheet | 33 |

## Handout `B6_Handout.docx` (H = số thứ tự mục)

Đã soạn (`2_Dataset/_generator/build_docs_b6.py`, 03/10/2026); chạy `python3 _generator/build_docs_b6.py` để in ra khung còn thiếu. Ảnh handout chụp trên bản `_HocVien` (Demo) hoặc trên bản sao file khởi đầu / `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` khi đã dựng.

| Tên file | Nội dung cần chụp | File · sheet · vùng | Mục |
|---|---|---|---|
| H01_chart_it_dung.png | Insert → All Charts liệt kê Histogram, Box & Whisker, Treemap, Sunburst, Waterfall, Funnel; cạnh đó histogram giá trị đơn Q1 với bin 100.000 | `Demo/B6_Demo_Chart_HocVien.xlsx` · Demo_Phan_phoi · Insert → All Charts | 1 |
| H02_convert_to_formulas.png | Menu PivotTable Analyze → OLAP Tools đang mở, dòng Convert to Formulas được chọn | Bản sao `../B5_Data_Model_DAX/Practice/B5_MCI_Mart_Answer.xlsx` · một sheet Pivot từ Data Model | 3 |
| H03_cube_na.png | Ô CUBEVALUE báo #N/A vì gõ sai tên measure; thanh công thức và gợi ý autocomplete | Bản sao file khởi đầu · sheet mới, ô B3 | 3 |
| H04_rules_manager.png | Hộp Rules Manager (This Worksheet) liệt kê quy tắc Icon Set của 6 thẻ KPI; hộp Edit Rule mở với Type = Formula | `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` · Dashboard · Home → Conditional Formatting → Manage Rules | 4 |
| H05_theme.png | Page Layout → Themes mở; hộp Customize Theme Colors với sáu màu nhấn | `Practice/B6_MCI_Mart_Dashboard_Answer.xlsx` · Page Layout → Themes | 4 |
| H06_form_control.png | Hộp Format Control của Combo Box (Input range, Cell link, Drop down lines) cạnh chart đổi theo cửa hàng | Bản sao `Demo/B6_Demo_Chart_HocVien.xlsx`, tự dựng theo handout mục 5 | 5 |
| H07_forecast_options.png | Hộp Create Forecast Worksheet, Options mở rộng: Seasonality = Set Manually 7, Fill Missing Points, Aggregate Duplicates | `Demo/B6_Demo_DuBao_HocVien.xlsx` · tblNgay · Data → Forecast Sheet → Options | 6 |
