# Sắc thái dùng cho khóa

Chọn theo đầu ra, không hỏi người dùng. Dựa trên hai sắc thái "Đào tạo & hướng dẫn" và "Nội bộ, khách quan" của `viet-dieuchinh-vanphong` (MIT), điều chỉnh cho khóa.

## A. Đào tạo (nhận xét học viên, hướng dẫn demo/homework, lời giảng)

- Xưng "bạn"; giảng viên nói trong lời giảng xưng "mình" ("mình làm cùng các bạn"). Không lẫn ngôi trong một văn bản.
- Câu ngắn, cụ thể, nói với một người. Mỗi ý chỉ ra **ở đâu** (sheet, ô, cột), **đang thế nào**, **sửa ra sao**.
- Khen đúng chỗ, một hai điểm, nêu cụ thể. Không khen chung ("làm tốt lắm").
- Không nhắc tới agent, script, file đáp án, ghi chú dành cho giảng viên.
- Không dùng "tay nghề"; dùng "cách làm", "tư duy làm", "thao tác".
- Lời giảng: viết để đọc thành tiếng; câu một hơi, không ngoặc lồng ngoặc.

Mẫu (nhận xét):

> **Đạt - cần sửa nhỏ.** Bảng Sales đúng 1.248 dòng và tổng doanh thu khớp.
>
> **Làm tốt:** Cột Revenue là công thức của Table, kéo thêm dòng vẫn tự tính.
>
> **Cần sửa:**
> 1. Sheet Kiem_tra, ô C9: Pct_Share chia cho 4.250.000 gõ tay. Đổi thành `=D2/$C$5` để đổi dữ liệu thì tỷ trọng đổi theo (xem slide 18).
> 2. Store_Code chưa có thông báo lỗi tiếng Việt. Vào Data > Data Validation > Error Alert.

Mẫu (hướng dẫn demo cho học viên):

> Bấm Data > Remove Duplicates, chỉ tick cột Order_ID và Line_No. Excel báo xóa 3 dòng trùng, còn 1.245 dòng.

## B. Nội bộ, khách quan (báo cáo lớp, báo cáo rà soát, báo cáo dựng)

- Bỏ đại từ, gọi thẳng đối tượng ("Slide 14", "file `Demo/B2_Demo_Lookup.xlsx`", "học viên Nguyễn A").
- Mỗi nhận định có bằng chứng: vị trí, giá trị đọc được, cách kiểm. Chưa kiểm chứng thì ghi "chưa kiểm chứng" và lý do.
- Xếp theo mức nghiêm trọng, lỗi nặng trước. Số liệu đứng trước nhận xét.
- Đề xuất sửa ở đâu: kịch bản, script `_generator/…`, deck. Một câu một việc.
- Không đánh giá người; chỉ nêu sự việc.

Mẫu (báo cáo rà soát):

> **Cần sửa.** Slide 11, notes: "tổng doanh thu là 4,25 tỷ", file `Practice/B1_MCI_Mart_T4_Answer.xlsx` sheet Kiem_tra ô C5 ra 4.253.100.000. Làm tròn 4,25 tỷ đúng, nhưng slide ghi 4,2 tỷ. Sửa ở `Slide_Script_B1.md` mục Slide 11, rồi dựng lại deck.

Mẫu (báo cáo lớp):

> 14/20 bài **Đạt**, 4 **Đạt - cần sửa nhỏ**, 2 **Chưa đạt**. Lỗi phổ biến nhất: 9 bài chưa đổi 6 ô Quantity dạng chữ (dòng 3 của `Dap_an_loi`), kéo lệch cả tổng Quantity lẫn Pivot đối soát.

Các mẫu trên chỉ minh họa giọng; số liệu, tên file và ô trong mẫu không phải số thật của khóa. Không chép vào nhận xét hay báo cáo.

## Giữ nguyên trong cả hai sắc thái

- Con số, vị trí, mã, tên hàm.
- Thứ tự ưu tiên tài liệu: Course_Plan, kịch bản, deck/handout/demo.
- Các câu hỏi chưa chắc: ghi "Cần GV xem" kèm lý do.
