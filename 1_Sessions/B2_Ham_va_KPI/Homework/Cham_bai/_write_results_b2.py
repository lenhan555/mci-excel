import json, pathlib, importlib.util, csv, collections, hashlib, difflib, warnings
warnings.filterwarnings('ignore')
root=pathlib.Path(__file__).resolve().parents[4]
h=pathlib.Path(__file__).resolve().parent.parent
out=h/'Cham_bai'
sp=importlib.util.spec_from_file_location('grade_b2',root/'2_Dataset/_generator/grade_b2.py')
g=importlib.util.module_from_spec(sp);sp.loader.exec_module(g)
ans=h/'B2_BTVN_MCI_Mart_T4_2025_Answer.xlsx'
submissions=sorted((h/'Bai_nop').glob('*.xlsx'))
hashes={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in submissions}
results={}
for p in [ans,h/'B2_BTVN_MCI_Mart_T4_2025.xlsx']+submissions:
 results[p.name]=g.grade(str(p),str(ans));print(p.name,'đã đọc',flush=True)
manual={
 'b2_btvn_to-thu_lan2.xlsx':{'name':'Tô Thu','v5':'Một phần','insight':3,'D4':'Competent','e5':'KPI_Card!B4:B13 đủ 10 trường; B6 trộn tổng doanh thu toàn bộ với mẫu số theo miền và tham chiếu cũ; B7 ghi tblSales/Order_ID_line không có; B9 cho phép Category/Phan_khuc chưa xử lý đơn nhiều ngành'},
 'b2_btvn_to-thuXRBUKA.xlsx':{'name':'Tô Thu','v5':'Một phần','insight':3,'D4':'Competent','e5':'KPI_Card!B4:B13 đủ 10 trường; B6 công thức khác phạm vi tử/mẫu; B7 tên bảng/cột sai; B9 chiều Category/Phan_khuc không thống nhất grain đơn'},
 'b2_btvn_vu-le-thai-ngoc.xlsx':{'name':'Vũ Lê Thái Ngọc','v5':'Một phần','insight':1,'D4':'Developing','e5':'KPI_Card!B13 trống; B9 ghi 1 line_no=1 đơn; B8 nguồn tblSales khác tblSales_T4; B12 chưa định lượng target AOV','e6':'Insight!A3: nguyên nhân đạt; số dẫn 49,9% và 137,2% sai (đúng 49,5% và 137,3%); thiếu hàm ý'},
 'b2_btvn_xlsx_Pham Nhu Tam.xlsx':{'name':'Phạm Như Tâm','v5':'Một phần','insight':3,'D4':'Competent','e5':' KPI_Card!B2:B11 đủ 10 trường; B6 dữ liệu grain đúng nhưng mức báo cáo ghi cấp đơn, chưa nêu kỳ tháng; tên sheet có khoảng trắng đầu'},
 'b2-le-lam-tuyen.xlsx':{'name':'Lê Lâm Tuyên','v5':'Chưa đạt','insight':0,'D4':'Emerging','e5':'KPI CARD!A2:D3 có ba KPI Q1, không có card AOV đủ 10 trường; chữ Q1 tại A6:A17','e6':'Không có sheet Insight; KPI CARD!A6:A17 có chữ nhận xét Q1, không có nguyên nhân cho so sánh T4'},
}
for n,m in manual.items():
 r=results[n];g.ap_cham_tay(r,m['v5'],m['insight']);r['viec']['5'][1]=m['e5'];r['viec']['6'][1]=m.get('e6','Câu insight có số so sánh khớp dữ liệu, nguyên nhân khả dĩ và hàm ý');r['muc_D4']=m['D4'];r['ho_ten']=m['name']
 if n=='b2-le-lam-tuyen.xlsx':r['phan_chu_khac']=[{'sheet':'KPI CARD','range':'A6:A17','note':'Đã đọc nhận xét Q1, không coi là insight của đề T4'}]
 r['ghi_chu'].append('Tên file không đúng B2_BTVN_<HoTen>.xlsx')
 if 'to-thu' in n:r['ghi_chu'].append('Hai lần nộp cùng học viên; insight và KPI card giống nhau. Tổng hợp ưu tiên lan2. Không tính là dấu hiệu giữa hai học viên.')
 (out/(pathlib.Path(n).stem+'_Bang_chung.json')).write_text(json.dumps(r,ensure_ascii=False,indent=1,default=str),encoding='utf8')
(out/'_evidence_b2.json').write_text(json.dumps(results,ensure_ascii=False,indent=1,default=str),encoding='utf8')
(out/'cham_tay_B2.json').write_text(json.dumps(manual,ensure_ascii=False,indent=1),encoding='utf8')
tests={n:results[n] for n in (ans.name,'B2_BTVN_MCI_Mart_T4_2025.xlsx')}
tests[ans.name]['viec']['5']=['Đạt','Đã đọc KPI_Card!B4:B13: 10 trường đúng AOV'];tests[ans.name]['ghi_chu'].append('Answer thiếu Insight, có Insight_mau!A3/A5. Chấm học viên theo De_bai!C11.')
(out/'Thu_script_B2.json').write_text(json.dumps(tests,ensure_ascii=False,indent=1,default=str),encoding='utf8')
selected=[n for n in manual if n!='b2_btvn_to-thuXRBUKA.xlsx']
kt=list(results[ans.name]['kiem_tra']);vn=['Tra cứu (XLOOKUP)','Cột tính toán','Bảng KPI 1 - Doanh thu Miền x Ngành hàng','Bảng KPI 2 - theo Miền','KPI Definition Card','Viết insight']
cols=['Ho_ten','File','Ket_qua']+kt+vn+['Insight','Muc_D4','Ghi_chu_GV']
with (out/'Tong_hop_B2.csv').open('w',encoding='utf-8-sig',newline='') as f:
 w=csv.DictWriter(f,fieldnames=cols);w.writeheader()
 for n in selected:
  r=results[n];d={'Ho_ten':r['ho_ten'],'File':n,'Ket_qua':r['ket_qua'],'Insight':f'{r["insight_thanh_phan_dat"]}/3','Muc_D4':r['muc_D4'],'Ghi_chu_GV':' | '.join(r['ghi_chu'])}
  if n=='b2-le-lam-tuyen.xlsx':d['Ghi_chu_GV']+=' | Cần GV xem: nộp dữ liệu Q1, kiểm tra có gửi nhầm file. Số KPI2 trong JSON là cột T1 B13/B20/B27, không phải chỉ tiêu T4; Q1 ở E13/E20/E27.'
  if n=='b2_btvn_to-thu_lan2.xlsx':d['Ghi_chu_GV']+=' | Bản XRBUKA: 6/8 chỉ tiêu; bản lan2: 7/8. Đã lưu riêng bằng chứng và nhận xét bản trước.'
  for k,x in r['kiem_tra'].items():d[k]=f'{x["hoc_vien"]} | '+('Khớp' if x['khop'] else 'Lệch/thiếu')
  for i,v in enumerate(vn,1):d[v]=r['viec'][str(i)][0]+' | '+r['viec'][str(i)][1]
  w.writerow(d)
texts={
'ToThu_NhanXet.md':'''Chưa đạt - nộp lại

File: b2_btvn_to-thu_lan2.xlsx

**Làm tốt:**

- Bạn đã giữ đủ 1.919 dòng ở Sales. Bốn mã Product_Code tại G53, G230, G1311 và G1520 đúng, không còn dấu cách thừa hay lỗi tra cứu.
- Bảng KPI!B4:G6 dùng SUMIFS và tham chiếu hỗn hợp đúng. Tổng tại H7 là 296.965.000 VNĐ; số đơn tại C14 là 966.

**Cần sửa:**

1. KPI!D14 đang dùng `=SUM(D11:D13)`, ra 913.980,596 VNĐ. AOV toàn hệ thống phải là 296.965.000 chia 966, tức 307.417,184 VNĐ. Đổi thành `=B14/C14`. Đây là cộng ba AOV miền, nên số tổng bị sai. Xem slide 17, mục tổng chia tổng.
2. KPI!C11:C13 dùng SUMIFS trên cột LINE1. Cách này đang cho số đúng, nhưng đề yêu cầu COUNTIFS với Line_No = 1. Chẳng hạn C11 dùng `=COUNTIFS(Sales!$L:$L,A11,Sales!$B:$B,1)`, rồi kéo xuống. Xem slide 16.
3. KPI_Card!B6 ghi `=SUM(tblSales[Revenue]) /COUNTIFS(Sales!$K:$K,KPI!A2,Sales!$B:$B,1)`. Bảng thực tế tên tblSales_T4, K là Revenue và A2 trên KPI trống; tử số toàn hệ thống cũng chưa cùng phạm vi với mẫu số theo miền. Bạn sửa về tổng Revenue chia COUNTIFS(Line_No,1), rồi cập nhật B7 bằng đúng tên bảng và cột. Xem slide 5 và 16.
4. KPI_Card!B9 cho phép xem AOV theo Category và Phan_khuc. Một đơn có thể có nhiều ngành hàng, nên bạn cần ghi rõ cách đếm và phân bổ trước khi dùng các chiều này. Giữ các chiều miền, cửa hàng và tháng cho card hiện tại; B8 bổ sung mức báo cáo theo tháng. Xem slide 4-5.
5. Sales!Q10 và 909 ô cùng phân khúc đang ghi "trung bình". Ngưỡng IFS đúng, nhưng nhãn đề yêu cầu "Trung". Bạn đổi nhãn trong công thức để phân khúc thống nhất. Xem slide 7.

**Câu insight:**

insight!A1 đạt 3/3 thành phần, mức D4 Competent. Tỷ trọng miền Nam gần 48% và AOV 317.810 VNĐ/đơn khớp dữ liệu. Câu đã có giả thuyết về Bánh kẹo, Gia dụng và đề xuất combo. Bạn nên viết nguyên nhân dưới dạng giả thuyết cần kiểm tra, vì cơ cấu ngành hàng chưa chứng minh tác động nhân quả đến AOV. Giữ ý này khi sửa AOV tổng.
''',
'ToThu_XRBUKA_NhanXet.md':'''Chưa đạt - nộp lại

File: b2_btvn_to-thuXRBUKA.xlsx (bản trước lần 2)

**Làm tốt:**

- Sales có đủ 1.919 dòng. Product_Code ở F53, F230, F1311 và F1520 đã sạch; ba cột tra cứu không còn ô lỗi.
- Khối KPI!B2:G4 kéo SUMIFS đúng, H5 ra 296.965.000 VNĐ. COGS toàn bảng là 196.595.000 VNĐ.

**Cần sửa:**

1. KPI!I5 và J5 còn trống. Bạn bổ sung tổng số đơn `=SUM(I2:I4)` để ra 966, rồi AOV tổng `=H5/I5` để ra 307.417,184 VNĐ. Xem slide 16-17.
2. KPI_Card!B6 dùng tổng doanh thu toàn hệ thống chia số đơn của một miền. Tên bảng tblSales trong B6:B7 cũng khác tblSales_T4, cột Order_ID_line tại B7 không có trong Sales. Bạn viết lại công thức và nguồn cho cùng phạm vi AOV. Xem slide 5 và 16.
3. KPI_Card!B9 nêu Category và Phan_khuc là chiều phân tích AOV nhưng chưa xử lý đơn nhiều ngành hàng. Bạn ghi rõ giới hạn này, giữ miền, cửa hàng và tháng cho card hiện tại. B8 bổ sung kỳ báo cáo theo tháng. Xem slide 4-5.
4. Sales!P10 và 909 ô cùng phân khúc ghi "trung bình". Ngưỡng đúng; bạn đổi nhãn thành "Trung" trong IFS. Xem slide 7.

**Câu insight:**

insight!A1 đạt 3/3, mức D4 Competent. Con số, so sánh miền Nam và đề xuất combo có đủ. Bạn diễn đạt sức mua hai ngành hàng như nguyên nhân khả dĩ cần kiểm tra thêm, tránh khẳng định đã xác định được nguyên nhân.
''',
'VuLeThaiNgoc_NhanXet.md':'''Đạt - cần sửa nhỏ

File: b2_btvn_vu-le-thai-ngoc.xlsx

**Làm tốt:**

- Bạn giữ đủ 1.919 dòng ở Sales; các cột tra cứu, COGS, Gross_Margin và Phan_khuc đều là công thức, không còn lỗi dữ liệu trong các cột này.
- KPI_SUMIFS!H9 ra 296.965.000 VNĐ, B16 ra 966 đơn. B23 và C30 dùng tổng chia tổng, lần lượt ra AOV 307.417,184 VNĐ và Gross Margin 33,8%.

**Cần sửa:**

1. Insight!A3 ghi miền Nam cao hơn Bắc 49,9% và Trung 137,2%. Từ 142.061.000, 95.038.000 và 59.866.000 VNĐ, hai tỷ lệ đúng khi làm tròn một chữ số là 49,5% và 137,3%. Bạn tính `(doanh thu Nam/doanh thu miền so sánh-1)` trong ô trước khi đưa vào câu. Xem slide 17 và mục Viết insight của bài tập về nhà.
2. KPI_Card!B9 ghi "1 line_no = 1 đơn hàng". Một dòng là một sản phẩm trong một đơn; chỉ dòng Line_No = 1 được dùng để đếm đơn. Bạn sửa grain này và cập nhật nguồn ở B8 từ tblSales thành tblSales_T4, nêu Revenue, Order_ID, Line_No. Xem slide 4-5.
3. KPI_Card!B13 còn trống; B12 chưa nêu target AOV bằng số hay một nguồn target AOV cụ thể. Bạn điền target và ngưỡng Warning/Critical tương ứng. Xem slide 5 và mẫu B2_KPI_Definition_Card.
4. Sheet đang tên KPI_SUMIFS, trong đề yêu cầu KPI. Bạn đổi tên sheet để thống nhất tên bảng nộp. Các công thức SUMIFS hiện tại kéo đúng. Xem mục Bảng KPI 1 và Bảng KPI 2 trong De_bai.

**Câu insight:**

Insight!A3 đạt 1/3 thành phần, mức D4 Developing. Câu đã nêu nguyên nhân khả dĩ bằng số đơn và AOV của miền Nam. Hai tỷ lệ so sánh cần sửa; câu còn thiếu hành động tiếp theo. Bạn có thể nêu việc phân tích cửa hàng hoặc cơ cấu giỏ hàng để kiểm tra vì sao số đơn và AOV cao, rồi xác định cách cải thiện hai miền còn lại. Xem mục Viết insight trong bài tập về nhà.
''',
'PhamNhuTam_NhanXet.md':'''Đạt - cần sửa nhỏ

File: b2_btvn_xlsx_Pham Nhu Tam.xlsx

**Làm tốt:**

- Sales có đủ 1.919 dòng. Ba cột tra cứu và cột tính toán đều là công thức; bốn mã Product_Code ở G53, G230, G1311 và G1520 đã sạch.
- KPI!B3:G5 kéo SUMIFS đúng. H6 và B13 đều ra 296.965.000 VNĐ, C13 ra 966 đơn. D13 dùng tổng chia tổng, AOV là 307.417,184 VNĐ.

**Cần sửa:**

1. Sheet ` KPI_Card` có khoảng trắng đầu tên. Bạn đổi thành `KPI_Card` theo yêu cầu bài nộp. Xem mục KPI Definition Card trong De_bai.
2. ` KPI_Card`!B6 đã mô tả đúng một dòng dữ liệu, nhưng ghi KPI tính "ở cấp đơn hàng" chưa nêu kỳ báo cáo. Bạn bổ sung AOV được tổng hợp theo tháng, có thể xem theo miền và cửa hàng như B7 đã ghi. Xem slide 4-5 và mẫu B2_KPI_Definition_Card.

**Câu insight:**

Insight!A1 đạt 3/3 thành phần, mức D4 Competent. AOV miền Nam 317.810 đồng/đơn, 447 đơn và doanh thu Bánh kẹo 45.695.000 đồng, Gia dụng 34.927.000 đồng đều khớp. Câu có giả thuyết và hướng kiểm tra để cải thiện hai miền còn lại. Bạn giữ cách viết nguyên nhân ở dạng khả dĩ này.
''',
'LeLamTuyen_NhanXet.md':'''Chưa đạt - nộp lại

File: b2-le-lam-tuyen.xlsx

**Làm tốt:**

- Bạn đã có các cột Category, Unit_Cost, COGS và Gross_Margin bằng công thức trong Sales.
- KPI_SUMIFS!E20 tính AOV Quý 1 bằng `=E6/E13`, đúng cách lấy tổng chia tổng.

**Cần sửa:**

1. Sales có 6.187 dòng dữ liệu, tổng Revenue 941.454.000 VNĐ. KPI_SUMIFS!B2:D2 ghi tháng 1, 2, 3 và E2 ghi Quý 1. Bài về nhà yêu cầu tháng 4/2025, 1.919 dòng và 296.965.000 VNĐ. Bạn kiểm tra lại file đã gửi, rồi làm trên B2_BTVN_MCI_Mart_T4_2025.xlsx. Các chênh lệch số liệu hiện tại cùng xuất phát từ khác kỳ dữ liệu. Xem slide 39 và bảng tự kiểm tra trong De_bai.
2. KPI_SUMIFS!A2:E6 đang là Miền theo tháng của Quý 1. Trên dữ liệu tháng 4, bạn dựng lại ba miền theo sáu ngành hàng, viết một SUMIFS với tham chiếu hỗn hợp rồi kéo. Xem slide 17.
3. KPI CARD!A2:D3 đang có ba KPI Quý 1, chưa có card AOV đủ 10 trường. Bạn tạo KPI_Card và điền theo mẫu B2_KPI_Definition_Card, ghi rõ công thức, nguồn, grain và kỳ tháng. Xem slide 4-5.
4. File chưa có sheet Insight. KPI CARD!A6:A17 đã có nhận xét và đề xuất cho Quý 1, nhưng bạn cần viết lại một câu dựa trên bảng tháng 4, có so sánh, nguyên nhân khả dĩ và hàm ý. Xem slide 39, mục Viết insight của bài tập về nhà.

**Câu insight:**

Phần chữ hiện tại ở KPI CARD!A6:A17 chưa đáp ứng insight tháng 4, ghi 0/3 theo yêu cầu bài này, mức D4 Emerging. Bạn giữ cách dẫn số theo miền, đổi sang đúng dữ liệu tháng 4, rồi thêm một giả thuyết cần kiểm tra và hành động cụ thể gắn với giả thuyết đó.
'''
}
# Lượt hai: văn phong đã biên tập trong bản dưới, kiểm hình thức trước khi ghi.
for name,text in texts.items():
 assert not any(x in text for x in ('—','–','“','”','‘','’','Hơn nữa','Bên cạnh đó','tay nghề'))
 (out/name).write_text(text,encoding='utf8')
report='''Báo cáo lớp B2 - chấm ngày 07/10/2026

Chưa có thư mục con theo lớp. Đã chấm 5 file của 4 học viên. Tô Thu có hai bản, tổng hợp theo học viên dùng b2_btvn_to-thu_lan2.xlsx; bản XRBUKA có nhận xét và bằng chứng riêng.

| Phạm vi | Đạt | Đạt - cần sửa nhỏ | Chưa đạt - nộp lại |
|---|---:|---:|---:|
| 4 học viên, lấy bản mới của Tô Thu | 0 | 2 | 2 |
| 5 file, tính cả hai bản Tô Thu | 0 | 2 | 3 |

Không quy điểm. Thang /10 trong kết quả cũ là thang tạm, chưa có thang và trọng số do người dùng đưa ở lần chấm này. Kết quả cũ được lưu trong _backup_20261007, gồm ToThiThu/, báo cáo, CSV và grade_b2.py cũ. Các file cũ tại ToThiThu/ và _DANH_SACH_ANH_NX.md vẫn được giữ nguyên; kết quả hiện tại ở các file nhận xét tại thư mục Cham_bai gốc.

| Học viên / lần nộp | Chỉ tiêu khớp | Insight | D4 | Kết quả |
|---|---:|---:|---|---|
| Tô Thu, lan2 | 7/8 | 3/3 | Competent | Chưa đạt - nộp lại |
| Tô Thu, XRBUKA | 6/8 | 3/3 | Competent | Chưa đạt - nộp lại |
| Vũ Lê Thái Ngọc | 8/8 | 1/3 | Developing | Đạt - cần sửa nhỏ |
| Phạm Như Tâm | 8/8 | 3/3 | Competent | Đạt - cần sửa nhỏ |
| Lê Lâm Tuyên | 1/8 | 0/3 theo đề T4 | Emerging | Chưa đạt - nộp lại |

Lỗi phổ biến, tính trên 4 học viên theo bản được chọn:

- 1/4 sai AOV tổng: Tô Thu, KPI!D14 = SUM(D11:D13), ra 913.980,5964378964 thay vì 307.417,1842650103. Đây là cộng AOV ba miền, cùng họ lỗi tổng tỷ lệ ở Dap_an_loi dòng 7; đề xuất bổ sung trường hợp SUM tỷ lệ bên cạnh AVERAGE trong script dựng homework.
- 1/4 chưa dùng đúng COUNTIFS theo yêu cầu: Tô Thu, KPI!C11:C13 dùng SUMIFS cột Sales!C:C (LINE1). Số miền và tổng 966 đúng. Dap_an_loi dòng 6 nói lỗi đếm dòng thành đơn; bài này đã xử lý grain bằng cờ nhưng khác hàm yêu cầu, chỉ đánh giá Một phần cho phần phương pháp, không tính thêm lỗi số đơn.
- 3/4 có vấn đề trong card AOV: Tô Thu đủ 10 trường nhưng KPI_Card!B6 trộn phạm vi tử/mẫu, B7 sai bảng/cột, B9 dùng Category và Phan_khuc không định nghĩa đơn nhiều ngành; Ngọc KPI_Card!B13 trống, B9 grain sai, B8 nguồn sai và B12 chưa định lượng target AOV; Tâm ` KPI_Card`!B6 chưa nêu kỳ báo cáo tháng, tên sheet có khoảng trắng đầu. Lê Lâm Tuyên là một trường hợp khác: KPI CARD!A2:D3 là ba KPI Q1, chưa có card AOV đủ 10 trường. Chưa có trong Dap_an_loi; đề xuất bổ sung lỗi card ở script dựng homework, giữ phần đọc hiểu chấm tay.
- 1/4 dẫn sai phần trăm và thiếu hàm ý: Ngọc, Insight!A3 ghi 49,9% và 137,2%. Tính từ KPI_SUMIFS!H6:H8 ra 49,4781034954% và 137,2982995356%, làm tròn 49,5% và 137,3%. Có nguyên nhân bằng số đơn/AOV; chưa có hành động tiếp theo. Đề xuất bổ sung vào Dap_an_loi của script dựng homework.
- 1/4 nộp khác kỳ: Lê Lâm Tuyên, Sales 6.187 dòng, Revenue 941.454.000 VNĐ, KPI_SUMIFS!B2:E2 là tháng 1-3/Quý 1. Không có khóa Order_ID + Line_No chung với dữ liệu T4; thiếu 1.919 khóa T4 và có 6.187 khóa Q1. Mọi lệch số liệu ở bài này gộp về lỗi khác kỳ, không tính nhiều lỗi gốc. Đề xuất thêm bước xác nhận kỳ vào danh sách nộp bài.
- 1/4 khác nhãn Phan_khuc nhưng đúng ngưỡng: Tô Thu lan2, Sales!Q10 và 909 dòng khác ghi "trung bình" thay vì "Trung". Không đánh giá sai ngưỡng. Bản XRBUKA cùng vấn đề ở Sales!P10.

Ba điểm nên chữa đầu buổi sau:

1. AOV tổng và Gross Margin % phải tính tổng chia tổng. Dùng KPI!D14 của bản lan2 để chỉ ra lỗi SUM tỷ lệ, sau đó đối chiếu KPI!B14/C14.
2. Card AOV phải cùng phạm vi ở tử/mẫu, đúng grain một dòng sản phẩm trong đơn và có kỳ báo cáo. Dùng card Tô Thu, Ngọc để phân biệt dữ liệu gốc với mức tổng hợp KPI.
3. Tính lại số so sánh trước khi viết insight, rồi nêu nguyên nhân khả dĩ và hành động. Dùng hai tỷ lệ sai tại Insight!A3 của Ngọc.

Bài cần GV xem:

- Lê Lâm Tuyên: cần xác nhận gửi nhầm file Q1 hay chưa làm bài T4. COGS/Gross_Margin/Phan_khuc có công thức nhưng không có khóa chung để đối chiếu ngưỡng với bài T4, ghi Không kiểm được ở việc 2. Các chỉ tiêu KPI2 trong JSON đọc cột T1 (B13/B20/B27), vì bài chỉ có bảng Q1; Q1 tổng ở E13/E20/E27. Không trình bày các số T1 này như số T4.
- Tô Thu: ưu tiên bản lan2 đã có tổng số đơn nhưng AOV tổng sai. Không suy đoán lịch sử nộp ngoài hậu tố tên file.
- Tài liệu gốc: Answer thiếu Insight; đề hiện đã sửa lỗi cài sẵn và có nhiều phần lời giải, cần GV xem trước khi giao lại bài.

Dấu hiệu bất thường:

Không phát hiện sheet Dap_an_loi hoặc Insight_mau trong 5 bài nộp. docProps/core.xml được đọc ở cả 5 bài; không thấy người tạo/người sửa cuối trùng tên một học viên khác trong 4 học viên được chấm. Lam Ngoc Tran ở hai bản Tô Thu chưa khớp một học viên trong danh sách này, nên không kết luận danh tính từ metadata. Hai bản cùng Tô Thu có insight và card giống nhau, là lịch sử nộp của cùng người. Không dùng dữ liệu sạch giống nhau làm dấu hiệu. Không kết luận chép bài.

Tên file / định dạng:

Cả 5 file đều mở được dạng .xlsx; cả 5 không theo đúng chuỗi B2_BTVN_<HoTen>.xlsx: b2_btvn_to-thu_lan2.xlsx, b2_btvn_to-thuXRBUKA.xlsx, b2_btvn_vu-le-thai-ngoc.xlsx, b2_btvn_xlsx_Pham Nhu Tam.xlsx, b2-le-lam-tuyen.xlsx. Tên ghi theo tên file, chưa xác minh danh sách lớp để bổ sung họ tên. Không trừ thêm ở các chỉ tiêu số vì lỗi tên file.

Thử script và lệch tài liệu:

- grade_b2.py đã sửa: bỏ điểm tạm, kiểm toàn cột, tìm tiêu đề từ dòng 1-10, đọc ArrayFormula, bảng theo nhãn thay vì tên sheet, hỗ trợ KPI2 tách nhiều khối, đọc drawing XML và VML. Kiểm toàn khối 18 ô, dòng/cột Tổng, dữ liệu theo khóa, COUNTIFS Line_No=1, công thức tổng tỷ lệ và từng miền. Sau thay đổi cuối, chạy lại Answer, đề và cả 5 bài.
- Answer: 8/8 Kiem_tra khớp. Việc 1-4 Đạt; KPI_Card!B4:B13 đủ 10 trường đúng AOV nên việc 5 Đạt sau đọc tay. Việc 6 Chưa đạt vì thiếu sheet Insight, trái De_bai!C11; Insight_mau!A3/A5 chỉ là tham chiếu chấm. Đây là lệch của Answer, chấm học viên theo De_bai.
- Đề hiện tại: 6/8 chỉ tiêu có giá trị khớp; tổng KPI1 thiếu, GM% thiếu. Sales đã có cột Region/Category/UnitCost/COGS/Gross Margin/Phan_khuc và Sheet1 có SUMIFS, COUNTIFS, AOV. Bốn mã Product_Code tại Sales!F53/F230/F1311/F1520 đều đã sạch. Không còn lỗi 4 mã có dấu cách như Dap_an_loi dòng 4 mô tả. Vì đề đã làm một phần, phép thử đề lỗi cài sẵn không đạt điều kiện dữ liệu ban đầu; những việc còn thiếu vẫn trả Chưa đạt. Ghi rõ ngoại lệ này thay vì sửa tiêu chí để làm đề trượt giả. Không sửa đề hoặc Answer.
- Phần chữ được đọc trong ô và XML drawing/VML của cả 5 file. Không có text box chứa chữ trong 5 file; chữ học viên ở các ô Insight và KPI CARD. Không mở Excel và không lưu lại bài nộp.
- Kết quả dùng giá trị cache đã lưu kết hợp đọc công thức và tính lại tổng từ dữ liệu, COGS từ Quantity x Unit_Cost, so sánh từng dòng và từng ô KPI. Không chạy lại engine Excel; không chứng minh trạng thái giao diện hay lịch sử thao tác.

Đầu ra: 5 file *_NhanXet.md, Tong_hop_B2.csv (UTF-8 BOM, 4 dòng học viên), 5 file *_Bang_chung.json, cham_tay_B2.json, Thu_script_B2.json và _evidence_b2.json. Báo cáo cũ và điểm cũ nằm trong backup, không dùng thay cho kết quả hiện tại.
'''
assert not any(x in report for x in ('—','–','“','”','‘','’','Hơn nữa','Bên cạnh đó','tay nghề'))
(out/'Bao_cao_lop_B2.md').write_text(report,encoding='utf8')
after={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in submissions}
assert hashes==after,'Bài nộp bị thay đổi'
(out/'Kiem_tra_file_goc_B2.json').write_text(json.dumps({'sha256_before':hashes,'sha256_after':after,'unchanged':hashes==after},indent=1),encoding='utf8')
print('Kết quả',collections.Counter(results[n]['ket_qua'] for n in selected),flush=True)
