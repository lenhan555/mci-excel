# Danh sách ảnh cần chụp — nhận xét bài tập về nhà B2

Mỗi ảnh lưu **PNG**, đúng tên, vào `Cham_bai/<Học viên>/Images/`. Chụp Excel 365 Windows, zoom 100%.
Ảnh **"bài của bạn"** chụp từ file trong `Homework/Bai_nop/`; ảnh **"đáp án"** chụp từ file `_Answer` (hoặc `Practice/`, `Demo/` khi ghi rõ).
Ô cần nhấn mạnh: tô viền đỏ trên **bản sao** của file, không sửa file gốc của học viên.
Chạy lại `python 2_Dataset/_generator/build_nhanxet_pdf.py B2` để ảnh tự vào PDF.

Còn thiếu: **13** ảnh.

| Học viên | Tên file | Chụp gì (file · sheet · vùng · cần thấy) |
|---|---|---|
| ToThiThu | `NX01_kpi1_ban.png` | File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `KPI` · Vùng: A1:H5 · Cần thấy: khối Miền × Ngành 3 × 6 + dòng/cột Tổng, H5 = 296.965.000; chọn B2 để thanh công thức hiện `SUMIFS(... $A2 ... B$1)`. |
| ToThiThu | `NX02a_dong_tong_ban.png` | File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `KPI` · Vùng: A1:L5 (hoặc H1:L5 nếu quá rộng) · Cần thấy: dòng 5 "Tổng" với I5 và J5 trống; viền đỏ I5 và J5; cột Số đơn (317 / 202 / 447) và AOV từng miền ở trên. |
| ToThiThu | `NX02b_dong_tong_dap_an.png` | File: `Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx` · Sheet: `KPI` · Vùng: A11:E16 · Cần thấy: dòng Tổng có C16 = 966, D16 = 307.417, E16 = 33,8%; chọn D16 để thanh công thức hiện `=B16/C16`. |
| ToThiThu | `NX03_demo_tong_chia_tong.png` | File: `Demo/B2_Demo_SUMIFS.xlsx` · Sheet: `Demo_SUMIFS` · Vùng: khối Gross Margin % theo ngành (dòng 20–28) · Cần thấy: ô AVERAGE = 31,2% và ô tổng ÷ tổng = 34,2% cạnh nhau. |
| ToThiThu | `NX04a_phan_khuc_ban.png` | File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `Sales` · Vùng: I1:P8 · Cần thấy: cột P có nhãn "trung bình" / "thấp"; chọn một ô P để thanh công thức hiện công thức IFS của bạn. |
| ToThiThu | `NX04b_loc_phan_khuc_ban.png` | Cùng file và sheet · Mở: bấm mũi tên Filter ở cột P · Cần thấy: các giá trị lọc "Cao", "thấp", "trung bình". |
| ToThiThu | `NX04c_phan_khuc_dap_an.png` | File: `Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx` · Sheet: `Sales` · Vùng: I1:P8 · Cần thấy: cột P có nhãn "Cao" / "Trung" / "Thấp"; thanh công thức `=IFS(...,TRUE,"Thấp")`. |
| ToThiThu | `NX05a_card_cong_thuc_ban.png` | File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `KPI_Card` · Vùng: A3:B8 · Cần thấy: B6 (công thức có `KPI!A2`) và B7 (có "Order_ID_line"); viền đỏ hai ô này. Bật Wrap Text, kéo rộng cột B. |
| ToThiThu | `NX05b_card_cong_thuc_dap_an.png` | File: `Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx` · Sheet: `KPI_Card` · Vùng: A3:B8 · Cần thấy: B6 và B7 theo đáp án. |
| ToThiThu | `NX06_don_nhieu_dong.png` | File: `../B1_Nen_tang_Excel_Du_lieu/Practice/B1_MCI_Mart_Q1_2025_Answer.xlsx` · Sheet: `Sales` · Vùng: A1:K7 (chọn dòng 3–6, ẩn tạm cột D–F, I–J) · Cần thấy: đơn HD2501-00002 chiếm 4 dòng Line_No 1–4. |
| ToThiThu | `NX06b_card_dimensions_ban.png` | File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `KPI_Card` · Vùng: A9:B13 · Cần thấy: B9 (Region, Store_Code, Thang, Category, Phan_khuc) và B11 (304,000 VND). Bật Wrap Text. |
| ToThiThu | `NX07_kpi_theo_mien.png` | File: `Homework/Bai_nop/b2_btvn_to-thuXRBUKA.xlsx` · Sheet: `KPI` · Vùng: A1:L5 (bản rộng, hoặc ghép H1:L5 với A1:A5) · Cần thấy: Tổng theo miền (95.038.000 / 59.866.000 / 142.061.000), Số đơn (317 / 202 / 447), AOV (≈299.804 / 296.366 / 317.810). |
| ToThiThu | `NX07b_insight_mau.png` | File: `Homework/B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx` · Sheet: `Insight_mau` · Vùng: A1:A5 · Cần thấy: câu mẫu (bật Wrap Text, kéo rộng cột A). |
