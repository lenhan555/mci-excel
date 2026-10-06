---
name: viet-van-phong-mci
description: Chuẩn văn phong tiếng Việt cho tài liệu và báo cáo khóa MCI Excel. Dùng khi viết nhận xét bài tập về nhà, báo cáo lớp, báo cáo rà soát, báo cáo dựng tài liệu, hoặc khi kiểm văn phong của kịch bản, notes, handout, hướng dẫn demo. Rút từ hai skill mã nguồn mở viet-dieuchinh-vanphong và vietnamizer.
---

# Văn phong tiếng Việt, khóa MCI Excel

Nguồn: chắt lọc từ [viet-dieuchinh-vanphong](https://github.com/namdangntd-sys/viet-dieuchinh-vanphong) (MIT, SEONGON) và [vietnamizer](https://github.com/fioenix/vietnamizer) v0.9.8 (MIT, fioenix). Đã lược bớt và đổi ví dụ cho khóa Excel/BI.

Mục tiêu: chữ tiếng Việt sắc, gọn, tự nhiên, sạch dấu hiệu "giọng AI", mà **không làm sai nội dung**. Với tài liệu khóa học, đúng bằng chứng quan trọng hơn hay.

## Chọn chế độ

| Chế độ | Ai dùng | Việc làm |
|---|---|---|
| **Viết** | `homework-grader`, báo cáo cuối của `session-builder` và `consistency-checker` | Viết hai lượt, tự quét trước khi ghi |
| **Kiểm** | `consistency-checker` (mục K), bước quét của `session-builder` | Chỉ báo lỗi, không sửa file |

## Sắc thái (không hỏi, chọn theo đầu ra)

- Nhận xét gửi học viên, hướng dẫn trong demo/homework: **đào tạo**, xưng "bạn", giọng người cùng làm với nhau. Chi tiết ở `references/sac-thai-mci.md`.
- Báo cáo cho giảng viên (báo cáo lớp, báo cáo rà soát, báo cáo dựng): **nội bộ, khách quan**, bỏ đại từ, nêu thẳng việc, số liệu, vị trí.

## Vùng không đụng

Không sửa, không "làm mượt", không đổi dấu trong:
- tên hàm và cú pháp Excel/Power Query/DAX, công thức, địa chỉ ô (`C5`, `$C$5`, `Sheet!A1:K4`);
- tên file, tên sheet, tên bảng, đường dẫn, mã dữ liệu (`ST07`, `P021`, `tblSales_T4`);
- chữ trích nguyên văn (lời học viên, thông báo lỗi `#N/A`, nhãn UI của Excel);
- nhãn sản xuất `[CHỤP]`, `[VẼ]`, `[PDF p…]` và ký hiệu quy ước của dự án: `→` trong `(→ demo …)`, `·` trong chân slide và bảng `Slide · Thao tác · Kết quả`.

Văn phong chỉ áp lên chữ thường của tác giả. Kịch bản là tài liệu nguồn: một câu "sạch" mà sai số liệu là lỗi nặng hơn một câu "máy" mà đúng.

## Quy trình viết (hai lượt)

1. **Lượt 1.** Viết đủ ý, đúng bằng chứng (sheet, ô, giá trị), đúng thứ tự ưu tiên. Chưa lo giọng.
2. **Lượt 2.** Làm cho ra người theo `references/luat-ma-sat.md`: cắt lặp ý, bỏ câu tự dán nhãn và câu chốt khôn khôn, bỏ khung "không phải X mà là Y", đổi từ nối dịch, xen câu dài ngắn, cắt mùi dịch (thừa "là/một/việc/sự").
3. **Quét hình thức** theo `references/quy-tac-chung.md`: gạch dài, ba chấm, ngoặc cong, emoji, số kiểu Việt.
4. **Quét dùng từ** theo `references/loi-dung-tu.md`: thiếu bổ ngữ kết quả, cặp liên từ thiếu vế, thừa "của", câu trợ lý còn sót.
5. **Soát từ vựng khóa** theo `references/tu-vung-khoa.md`.
6. **Đọc to.** Chỗ nào nghe như đọc biểu mẫu thì sửa. Chỉ ghi file khi nghe ra một người thật đang nói.

Liều lượng: hư từ tình thái, nhiệt độ, chi tiết lẻ là gia vị. Mỗi bài một hai chỗ. Rắc nhiều thành giả giọng.

## Quy trình kiểm

1. Xác định thể loại và vùng không đụng của đoạn đang đọc.
2. Quét hình thức bằng regex ở `references/quy-tac-chung.md` (đọc-only).
3. Đọc lượt hai: luật ma sát, lỗi dùng từ, từ vựng khóa.
4. Chỉ báo chỗ **có lỗi thật**; câu đã tự nhiên thì để yên. Không đề xuất viết lại cả đoạn chỉ vì "chưa hay".
5. Mỗi lỗi: vị trí, trích nguyên văn ngắn, luật vi phạm (tên hoặc mã `V…`/`T…`), câu thay đề xuất. Gộp theo luật, tối đa 3 ví dụ mỗi luật.
6. Mức lỗi: **Nhỏ** (hình thức, một từ), **Cần sửa** (lỗi lặp nhiều chỗ, văn dịch làm khó hiểu hướng dẫn). Không đặt Nghiêm trọng trừ khi câu làm người đọc hiểu sai thao tác.

## Tài nguyên

- `references/quy-tac-chung.md`: cụm cấm, ký tự, số kiểu Việt, regex quét.
- `references/luat-ma-sat.md`: 16 luật ma sát.
- `references/loi-dung-tu.md`: lỗi dùng từ và cấu trúc (từ V1-V25, T1-T6).
- `references/sac-thai-mci.md`: hai sắc thái và mẫu.
- `references/tu-vung-khoa.md`: thuật ngữ giữ nguyên, từ nên tránh.
