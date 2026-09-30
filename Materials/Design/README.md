# Design — MCI Excel BI

Nguồn chuẩn của hệ thống thiết kế: artifact **MCI Excel BI**, tại https://claude.ai/artifact/Es8iAxiYzhSV6q8D7isomX

Thư mục này chứa hai file theme Office, để slide, Excel và Word dùng đúng màu và font của hệ thống.

| File | Nội dung |
|---|---|
| `MCI_Excel_BI_Colors.xml` | Theme màu. Accent 1–6 là 6 màu dữ liệu, Dark 2 là navy, Light 2 là surface-alt, Hyperlink là link. |
| `MCI_Excel_BI_Fonts.xml` | Theme font: Arial cho cả tiêu đề và thân bài. |

## Cài đặt

Chép hai file vào thư mục theme của Office, rồi khởi động lại Office.

- **Windows:**
  - file màu → `%APPDATA%\Microsoft\Templates\Document Themes\Theme Colors\`
  - file font → `%APPDATA%\Microsoft\Templates\Document Themes\Theme Fonts\`
- **Mac:**
  - file màu → `~/Library/Group Containers/UBF8T346G9.Office/User Content/Themes/Theme Colors/`
  - file font → `~/Library/Group Containers/UBF8T346G9.Office/User Content/Themes/Theme Fonts/`

## Áp dụng

- **Excel:** Page Layout → Colors → chọn **MCI Excel BI**; Page Layout → Fonts → chọn **MCI Excel BI**.
- **PowerPoint:** Design → Variants → Colors / Fonts.

Sau khi áp theme, biểu đồ mới sẽ tự dùng thứ tự màu Bắc → Trung → Nam, và TableStyleMedium2 có màu xanh `series-1`.
