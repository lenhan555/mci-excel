import pathlib,json,csv,shutil,hashlib,re
from decimal import Decimal,ROUND_HALF_UP
out=pathlib.Path(__file__).resolve().parent; homework=out.parent
data=json.loads((out/'_evidence_b2.json').read_text(encoding='utf8'))
backup=out/'_backup_truoc_ap_mau_20261007';backup.mkdir(exist_ok=True)
for p in out.iterdir():
 if p.name.startswith('_') or p.name=='Images':continue
 if p.is_file():shutil.copy2(p,backup/p.name)
 elif p.is_dir() and p.name in ('ToThiThu','VuLeThaiNgoc','PhamNhuTam','LeLamTuyen'):
  shutil.copytree(p,backup/p.name,dirs_exist_ok=True)
configs=[('ToThiThu','Tô Thị Thu','b2_btvn_to-thu_lan2.xlsx',[100,50,100,50,50],3,'Competent'),('VuLeThaiNgoc','Vũ Lê Thái Ngọc','b2_btvn_vu-le-thai-ngoc.xlsx',[100,100,100,100,50],1,'Developing'),('PhamNhuTam','Phạm Như Tâm','b2_btvn_xlsx_Pham Nhu Tam.xlsx',[100,100,100,100,50],3,'Competent'),('LeLamTuyen','Lê Lâm Tuyên','b2-le-lam-tuyen.xlsx',[0,0,0,0,0],0,'Emerging')]
weight=[2,2,1,1,1]
names=['Tra cứu XLOOKUP','Cột tính toán','Bảng KPI 1: Miền x Ngành hàng','Bảng KPI 2: theo Miền','KPI Definition Card (AOV)']
state={100:'Đạt',50:'Một phần',0:'Chưa có phần việc đáp ứng bài T4'}
def number(x,name=''):
 if x is None:return 'Trống / chưa có bảng T4'
 if 'Margin %' in name:return f'{x*100:.1f}%'.replace('.',',')
 if isinstance(x,(int,float)):return f'{x:,.0f}'.replace(',','.')
 return str(x)
ref='''## Tài liệu để xem lại

Slide nằm trong `Slide_B2_HocVien.pptx`. Đường dẫn Excel dưới đây tính từ thư mục Buổi 2.

| Slide | Nội dung | File để làm lại |
|---|---|---|
| 4-5 | Grain và KPI Definition Card | `Handout/B2_KPI_Definition_Card.docx` |
| 7 | IFS và nhãn phân khúc | `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF |
| 16-17 | COUNTIFS Line_No = 1 và tổng chia tổng | `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS |
| 30-31 | Thực hành bảng KPI và card | `Practice/B2_MCI_Mart_Q1_2025_Answer.xlsx`, sheet KPI_Card |
| 40 | Câu insight có số, nguyên nhân, hàm ý | Sheet Insight trong bài của bạn |
'''
errors={
'ToThiThu':[
('AOV dòng Tổng đang cộng ba miền','KPI!D14 dùng `=SUM(D11:D13)`, ra 913.980,596 VNĐ. Số đơn ở C14 đã đúng 966.','AOV toàn hệ thống = 296.965.000 / 966 = 307.417,184 VNĐ.','Cộng ba AOV miền làm con số tổng lớn gần ba lần, dù AOV từng miền đúng.','Đổi D14 thành `=B14/C14`. C11:C13 hiện SUMIFS cột LINE1 cho số đúng; bạn đổi về COUNTIFS với Line_No = 1 theo yêu cầu, ví dụ `=COUNTIFS(Sales!$L:$L,A11,Sales!$B:$B,1)`, rồi kéo xuống.','Slide 16-17, `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS.','KPI','A10:G14','NX01_aov_tong.png','Chọn D14 để thấy SUM(D11:D13); cả số tổng doanh thu, số đơn và AOV phải hiện.'),
('Card AOV chưa thống nhất công thức và nguồn','KPI_Card!B6 ghi `=SUM(tblSales[Revenue]) /COUNTIFS(Sales!$K:$K,KPI!A2,Sales!$B:$B,1)`; B7 ghi tblSales và Order_ID_line.','Bảng đang tên tblSales_T4, các cột đúng là Revenue, Order_ID, Line_No. Tử số và mẫu số cần cùng kỳ và cùng miền.','B6 tham chiếu tên bảng khác và trộn tổng toàn bảng với điều kiện miền; trên bản lan2, Sales!K:K còn là Revenue và KPI!A2 trống. Người dùng card sẽ không tính đúng AOV.','B6 ghi `=SUM(tblSales_T4[Revenue])/COUNTIFS(tblSales_T4[Line_No],1)`. B7 sửa tên bảng/cột. B8 ghi dữ liệu một sản phẩm trong một đơn, KPI báo cáo theo tháng. B9 giữ miền, cửa hàng, tháng; muốn xem theo Category/Phan_khuc cần định nghĩa cách phân bổ đơn nhiều ngành. Target tại B11 cần ghi nguồn hoặc đánh dấu giả định.','Slide 4-5 và 30, mẫu `Handout/B2_KPI_Definition_Card.docx`.','KPI_Card','A3:B13','NX02_card_aov.png','Cần đọc rõ B6, B7, B8 và B9; bật Wrap Text trên bản sao nếu cần.'),
('Nhãn Phan_khuc khác nhãn yêu cầu','Sales!Q10 và 909 ô cùng phân khúc ghi "trung bình". Ngưỡng của IFS vẫn đúng.','Đề dùng nhãn Cao / Trung / Thấp.','Nhóm "trung bình" có nghĩa đúng nhưng khác chuỗi "Trung", dễ làm bảng lọc hoặc đối chiếu dùng nhãn chuẩn bỏ sót nhóm. Chữ hoa/thường không bị trừ riêng.','Đổi nhánh IFS thành "Trung". Có thể dùng `=IFS([@Unit_Price]>=100000,"Cao",[@Unit_Price]>=30000,"Trung",TRUE,"Thấp")`.','Slide 7, `Demo/B2_Demo_Logic.xlsx`, sheet Demo_IF.',None,None,None,None)],
'VuLeThaiNgoc':[
('Card còn thiếu ngưỡng và sai grain','KPI_Card!B13 trống; B9 ghi "1 line_no = 1 đơn hàng". B8 ghi tblSales; B12 chưa có target AOV bằng số hoặc nguồn cụ thể.','Một dòng dữ liệu là một sản phẩm trong một đơn. Chỉ Line_No = 1 được dùng đếm đơn. Card cần đủ 10 trường, ghi target AOV và hai ngưỡng.','Định nghĩa một dòng thành một đơn khiến người đọc có thể đếm số dòng thay số đơn. Thiếu ngưỡng nên chưa phân loại được KPI đạt hay cảnh báo.','B9 sửa grain và giữ kỳ tháng; B8 ghi tblSales_T4 với Revenue, Order_ID, Line_No. Điền B12 bằng target hoặc nguồn target AOV, rồi B13 ghi Warning/Critical tương ứng.','Slide 4-5, `Handout/B2_KPI_Definition_Card.docx`.','KPI_Card','A5:B14','NX01_card_grain.png','Cần đọc B8/B9/B12 và thấy B13 trống.'),
('Insight cần sửa tỷ lệ và thêm hành động','Insight!A3 ghi miền Nam hơn Bắc 49,9% và Trung 137,2%, câu kết chỉ nhận xét biên lợi nhuận.','Từ KPI_SUMIFS!H6:H8, Nam hơn Bắc 49,4781% và Trung 137,2983%; làm tròn một chữ số là 49,5% và 137,3%. Insight còn cần hàm ý.','Người đọc nhận tỷ lệ chưa đúng và chưa biết cần kiểm tra hoặc làm gì tiếp theo.','Tính `(doanh thu Nam/doanh thu miền so sánh-1)` trong ô trước khi viết. Thêm hành động kiểm tra cửa hàng hoặc cơ cấu giỏ hàng để giải thích số đơn/AOV miền Nam và tìm cách áp dụng cho miền còn lại.','Slide 17 và 40, `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS.','Insight','A2:A3','NX02_insight.png','Cần đọc nguyên văn hai tỷ lệ và phần kết chưa có hành động.')],
'PhamNhuTam':[
('Card cần nêu kỳ báo cáo','` KPI_Card`!B6 ghi dữ liệu một sản phẩm trong một đơn, KPI tính "ở cấp đơn hàng"; B7 có miền, cửa hàng, tháng.','AOV được tổng hợp cho một kỳ báo cáo, ở bài này là tháng 4/2025; đơn hàng là đối tượng được đếm ở mẫu số.','Người đọc chưa phân biệt giá trị mỗi đơn với AOV trung bình của các đơn trong tháng.','B6 bổ sung: dữ liệu một sản phẩm trong một đơn; KPI tổng hợp theo tháng, có thể xem theo miền và cửa hàng. Tên sheet có khoảng trắng đầu chỉ nên sửa để gọn, không trừ thêm điểm. Target B9 cần nêu nguồn hoặc giả định.','Slide 4-5, `Handout/B2_KPI_Definition_Card.docx`.',' KPI_Card','A2:B11','NX01_card_ky.png','Cần đọc B6 với mức KPI cấp đơn và B7 có chiều tháng.')],
'LeLamTuyen':[
('File đang là Quý 1, khác dữ liệu bài tháng 4','Sales có 6.187 dòng, tổng Revenue 941.454.000 VNĐ. KPI_SUMIFS!B2:D2 ghi 1, 2, 3 và E2 ghi Quý 1.','Bài yêu cầu tháng 4/2025, 1.919 dòng, doanh thu 296.965.000 VNĐ và 966 đơn.','Không có khóa Order_ID + Line_No nào chung với dữ liệu tháng 4 nên chưa có bằng chứng để chấm các phần số liệu của bài T4. Đây là một lỗi gốc khác kỳ dữ liệu; không tách các chênh lệch thành nhiều lỗi độc lập.','Bạn kiểm tra có gửi nhầm file không. Nếu đây là file thực hành trên lớp, gửi đúng bài đã làm từ `Homework/B2_BTVN_MCI_Mart_T4_2025.xlsx`. Với dữ liệu T4, dựng Miền x sáu ngành hàng và bảng KPI miền bằng COUNTIFS/tổng chia tổng.','Mục bài tập về nhà trong De_bai và slide 16-17, `Demo/B2_Demo_SUMIFS.xlsx`, sheet Demo_SUMIFS.','KPI_SUMIFS','A1:E6','NX01_sai_ky.png','Cần thấy tiêu đề Quý 1, tháng 1-3 và tổng 941.454.000.'),
('Chưa có card AOV và insight của bài tháng 4','KPI CARD!A2:D3 có ba KPI Q1; A6:A17 có nhận xét và đề xuất cho Q1. File chưa có sheet Insight.','Bài T4 cần card AOV đủ 10 trường và insight gồm con số có so sánh, nguyên nhân khả dĩ, hàm ý.','Phần chữ hiện tại chưa trả lời yêu cầu của dữ liệu T4, nên chưa ghi nhận điểm ở việc 5 và 6.','Sau khi xác nhận đúng file, tạo KPI_Card và Insight. Điền card theo mẫu, viết câu dựa trên bảng T4; giữ số so sánh, thêm nguyên nhân khả dĩ và một hành động cụ thể.','Slide 4-5 và 40, `Handout/B2_KPI_Definition_Card.docx`.','KPI CARD','A1:D17','NX02_card_q1.png','Cần thấy ba KPI Q1, phần chữ A6:A17 và chưa có card AOV đủ 10 trường.')]
}
good={
'ToThiThu':['Sales giữ đủ 1.919 dòng; Product_Code tại G53, G230, G1311 và G1520 đã sạch, các cột tra cứu không còn lỗi.','KPI!B4:G6 kéo SUMIFS đúng với $ hỗn hợp, H7 ra 296.965.000 VNĐ; C14 ra đúng 966 đơn.'],
'VuLeThaiNgoc':['Sales đủ 1.919 dòng; ba cột tra cứu và cột tính toán đều là công thức, khớp theo khóa dòng.','KPI_SUMIFS!H9 = 296.965.000, B16 = 966; B23 = 307.417,184 VNĐ, C30 = 33,8%, đều dùng tổng chia tổng.'],
'PhamNhuTam':['Sales đủ 1.919 dòng; bốn mã Product_Code tại G53, G230, G1311, G1520 sạch và cột tính toán đúng.','KPI!B3:G5 kéo SUMIFS đúng. H6/B13 = 296.965.000, C13 = 966, D13 = 307.417,184 VNĐ.'],
'LeLamTuyen':['Sales có công thức Category, Unit_Cost, COGS và Gross_Margin; đã thực hành cách làm giàu bảng.','KPI_SUMIFS!E20 dùng `=E6/E13`, đúng cách tổng chia tổng cho AOV của Quý 1. Những phần này được ghi nhận là thao tác đã làm, chưa thay cho bằng chứng của bài T4.']}
ins_comments={
'ToThiThu':['Đạt: doanh thu 296.965.000, Nam 142.061.000 (47,8%) và AOV Nam 317.810 đúng.','Đạt ở mức giả thuyết: Bánh kẹo/Gia dụng được nêu như lý do khả dĩ. Hai ngành này cũng lớn ở các miền khác, nên cần kiểm tra cơ cấu giỏ hàng và số đơn trước khi kết luận.','Đạt: đã nêu Marketing và combo. "Giữ vững đà tăng trưởng" chưa có căn cứ thời gian trong một tháng; đổi thành việc cần kiểm chứng tiếp.'],
'VuLeThaiNgoc':['Chưa đạt: 49,9% và 137,2% cần sửa thành 49,5% và 137,3%. Các số doanh thu, số đơn và AOV còn lại đúng.','Đạt: đã giải thích bằng số đơn 447 và AOV 317.810 của miền Nam.','Chưa đạt: chưa có hành động hoặc hướng kiểm tra cụ thể sau nhận định biên lợi nhuận.'],
'PhamNhuTam':['Đạt: AOV Nam 317.810, 447 đơn, Bánh kẹo 45.695.000 và Gia dụng 34.927.000 đều đúng; có so sánh cao nhất.','Đạt: viết nhu cầu và giá trị sản phẩm dưới dạng yếu tố khả dĩ, giữ đúng mức bằng chứng.','Đạt: đã đề xuất tìm hiểu nguyên nhân AOV miền Nam để tìm cách cải thiện hai miền còn lại.'],
'LeLamTuyen':['Chưa có con số cho bài T4. KPI CARD!A6 chỉ dẫn doanh thu Q1 và số theo miền.','Chưa có nguyên nhân khả dĩ cho so sánh ở bài T4.','KPI CARD!A7/A11:A17 có đề xuất cho Q1, nhưng chưa gắn với số liệu T4 nên chưa ghi nhận trong phần insight của bài này.']}
all_images=[];csv_rows=[]
for folder,name,file,pcts,tp,d4 in configs:
 d=out/folder;d.mkdir(exist_ok=True);(d/'Images').mkdir(exist_ok=True)
 shutil.copy2(homework/'Bai_nop'/file,d/file)
 r=data[file];raw=sum(Decimal(w)*Decimal(pc)/100 for w,pc in zip(weight,pcts))+tp;score=int(raw.quantize(Decimal('1'),rounding=ROUND_HALF_UP))
 s=f'**Điểm: {score}/10**\n\n# Nhận xét bài tập về nhà Buổi 2 - {name}\n\nFile: `{file}`\n\n'
 if folder=='LeLamTuyen':s+='Điểm hiện tại phản ánh phần việc có bằng chứng trên đúng dữ liệu T4. Bạn kiểm tra lại file đã gửi để xác nhận có nhầm với bài thực hành Q1 không.\n\n'
 s+='| # | Việc | Kết quả | Trọng số | % đạt / Thành phần |\n|---|---|---|---:|---|\n'
 for i,(v,pc,w) in enumerate(zip(names,pcts,weight),1):s+=f'| {i} | {v} | {state[pc]} | {w} | {pc}% |\n'
 s+=f'| 6 | Insight | {tp}/3 thành phần | 3 | {tp}/3 |\n| | **Tổng đã làm tròn** | | **10** | **{score}/10** |\n\n'
 s+='Thang điểm: việc 1-2 mỗi việc 2 điểm; việc 3-5 mỗi việc 1 điểm. Đạt được 100%, Một phần được 50%, chưa có phần việc đáp ứng bài được 0%. Insight có ba thành phần, mỗi thành phần đạt được 1 điểm. Cộng các phần theo trọng số trước, rồi làm tròn tổng về số nguyên theo quy tắc từ nửa điểm trở lên làm tròn lên. Không cộng các điểm từng mục đã làm tròn.\n\n'
 s+=ref+'\n## Đối soát nhanh\n\n| Chỉ tiêu | Số cần đạt | File của bạn | Khớp? |\n|---|---:|---:|---|\n'
 for k,v in r['kiem_tra'].items():
  nv='Chưa có chỉ tiêu T4' if folder=='LeLamTuyen' else number(v['hoc_vien'],k)
  s+=f'| {k} | {number(v["dap_an"],k)} | {nv} | '+('Có' if v['khop'] and folder!='LeLamTuyen' else 'Chưa')+' |\n'
 s+='\nSố tiền trong bảng làm tròn đến VNĐ để dễ đọc; phần đối soát dùng giá trị đầy đủ đến sáu chữ số thập phân.\n\n## Làm tốt\n\n'+''.join('- '+x+'\n' for x in good[folder])+'\n## Điểm cần lưu ý\n\n'
 checks=[]
 for i,e in enumerate(errors[folder],1):
  title,done,correct,effect,fix,learn,sheet,region,img,need=e
  s+=f'### {i}. {title}\n\n**Bạn đã làm:** {done}\n\n**Đúng cần là:** {correct}\n\n**Hậu quả:** {effect}\n\n**Cách sửa:** {fix}\n\n**Học lại:** {learn}\n\n'
  checks.append(title)
  if img:
   note=f'File: `Homework/Cham_bai/{folder}/{file}` · Sheet: `{sheet}` · Vùng: {region} · Cần thấy: {need} · Chụp Excel 365 Windows, zoom 100%, chỉ thao tác trên bản sao.'
   s+=f'![Hình {i} · {title}](Images/{img})\n> **[CHỤP] {img}** · {note}\n\n'
   all_images.append((folder,img,note))
 s+='## Câu insight\n\n'
 if r['insight']:quote='\n\n'.join(x['text'] for x in r['insight']);s+='Câu của bạn tại '+r['insight'][0]['sheet']+'!'+r['insight'][0]['o']+':\n\n'+''.join('> '+line+'\n' for line in quote.splitlines())+'\n'
 else:
  inspect=json.loads((out/'_inspection.json').read_text(encoding='utf8'))[file]['cells']['KPI CARD']
  quote='\n'.join(str(v) for coord,v in inspect if coord in ('A6','A7'))
  s+='File chưa có sheet Insight. Phần chữ hiện tại ở KPI CARD!A6:A7:\n\n'+''.join('> '+line+'\n' for line in quote.splitlines())+'\n'
 s+='| Thành phần | Nhận xét |\n|---|---|\n'+''.join('| '+part+' | '+comment+' |\n' for part,comment in zip(['Con số có so sánh','Nguyên nhân khả dĩ','Hàm ý'],ins_comments[folder]))
 s+=f'\n**Mức D4:** {d4}. Phần insight đạt {tp}/3 thành phần theo bài T4.\n\n## Việc cần làm trước Buổi 3\n\n'+''.join('- [ ] '+x+'.\n' for x in checks)+'\nẢnh minh họa hiện là khung chờ chụp, đã ghi rõ file, sheet và vùng.\n'
 # Lượt hai đã cắt lặp và giữ bằng chứng; nguyên văn học viên được bảo toàn.
 (d/(folder+'_NhanXet.md')).write_text(s,encoding='utf8')
 r['diem_10']=score;r['ty_le_dat_v1_v5']=pcts;r['insight_thanh_phan_dat']=tp;r['muc_D4']=d4;r['ket_qua']='Đã nhận xét';r['ho_ten']=name
 for i,pc in enumerate(pcts,1):r['viec'][str(i)][0]='Đạt' if pc==100 else 'Một phần' if pc==50 else 'Chưa đạt'
 if folder=='ToThiThu':r['viec']['4'][1]='Một phần: số đơn các miền/tổng và GM% đúng; AOV tổng cộng ba miền; COUNTIFS cần dùng theo đề'
 if folder=='VuLeThaiNgoc':r['viec']['3'][1]+='; không trừ riêng tên KPI_SUMIFS';r['viec']['4'][1]='Đạt: các miền, dòng Tổng và công thức đúng; không trừ tên sheet'
 if folder=='LeLamTuyen':r['viec']['2'][1]='Chưa có bằng chứng T4 do khác kỳ dữ liệu; không chấm cột tính Q1 thành hoàn thành bài T4'
 (d/(folder+'_Bang_chung.json')).write_text(json.dumps(r,ensure_ascii=False,indent=1),encoding='utf8')
 csv_rows.append({'Ho_ten':name,'File':file,'Diem_10':score,'Ket_qua':'Đã nhận xét',**{k:str(v['hoc_vien'])+' | '+('Khớp' if v['khop'] else 'Lệch/thiếu') for k,v in r['kiem_tra'].items()},**{names[i]:state[pc]+f' ({pc}%)' for i,pc in enumerate(pcts)},'Insight':f'{tp}/3','Muc_D4':d4,'Ghi_chu_GV':('Cần GV xem: có thể gửi nhầm file Q1; các số KPI2 lưu trong bằng chứng là cột T1, không phải T4.' if folder=='LeLamTuyen' else 'Thu dùng lan2, bản trước giữ Lan1.' if folder=='ToThiThu' else '')})
lan1=out/'ToThiThu'/'Lan1';lan1.mkdir(exist_ok=True)
oldfile='b2_btvn_to-thuXRBUKA.xlsx';shutil.copy2(homework/'Bai_nop'/oldfile,lan1/oldfile)
if (backup/'ToThiThu'/'ToThiThu_NhanXet.md').exists():
 shutil.copy2(backup/'ToThiThu'/'ToThiThu_NhanXet.md',lan1/'ToThiThu_Lan1_NhanXet.md')
if (backup/'ToThiThu'/'ToThiThu_NhanXet.pdf').exists():shutil.copy2(backup/'ToThiThu'/'ToThiThu_NhanXet.pdf',lan1/'ToThiThu_Lan1_NhanXet.pdf')
shutil.copy2(out/'b2_btvn_to-thuXRBUKA_Bang_chung.json',lan1/'ToThiThu_Lan1_Bang_chung.json')
for row in csv_rows:
 if row['Ho_ten']=='Lê Lâm Tuyên':
  for k in data[oldfile]['kiem_tra']:row[k]='Chưa có chỉ tiêu T4 | File đang là Q1'
with (out/'Tong_hop_B2.csv').open('w',encoding='utf-8-sig',newline='') as f:
 w=csv.DictWriter(f,fieldnames=list(csv_rows[0]));w.writeheader();w.writerows(csv_rows)
(out/'_evidence_b2.json').write_text(json.dumps(data,ensure_ascii=False,indent=1),encoding='utf8')
images='Danh sách ảnh nhận xét B2 - bản theo mẫu, 07/10/2026\n\n'+f'Còn {len(all_images)} khung chờ ảnh cho 4 học viên. Lưu PNG vào Homework/Cham_bai/<học viên>/Images/. Chụp Excel 365 Windows, zoom 100%, chỉ mở bản sao trong thư mục chấm. Danh sách ảnh mẫu cũ nằm trong _backup_truoc_ap_mau_20261007.\n\n| Học viên | Tên ảnh | Nguồn và trạng thái cần chụp |\n|---|---|---|\n'+''.join(f'| {fo} | {im} | {note} |\n' for fo,im,note in all_images)
(out/'_DANH_SACH_ANH_NX.md').write_text(images,encoding='utf8')
report=(backup/'Bao_cao_lop_B2.md').read_text(encoding='utf8')
start=report.index('Lỗi phổ biến,');end=report.index('Tên file / định dạng:')
kept=report[start:end].replace('Tô Thu, KPI!D14','Tô Thị Thu, KPI!D14').replace('Tô Thu:','Tô Thị Thu:')
report='''Báo cáo chấm bài B2 - áp dụng mẫu ToThiThu, 07/10/2026

Đã đọc 5 file của 4 học viên; Tô Thị Thu dùng bản lan2. Không có thư mục lớp riêng. Mỗi học viên có thư mục chứa bản sao .xlsx, nhận xét .md và .pdf. Bản XRBUKA của Thu nằm trong ToThiThu/Lan1; mẫu và đầu ra trước được sao lưu trước khi ghi mới.

| Học viên | Điểm /10 (số nguyên) | Insight | D4 | Ghi chú |
|---|---:|---|---|---|
| Tô Thị Thu | 8 | 3/3 | Competent | AOV tổng sai; V2, V4, V5 đạt 50% |
| Vũ Lê Thái Ngọc | 8 | 1/3 | Developing | Card đạt 50%; tỷ lệ insight cần sửa và thiếu hàm ý |
| Phạm Như Tâm | 10 | 3/3 | Competent | Card đạt 50%, tổng được làm tròn lên |
| Lê Lâm Tuyên | 0 | 0/3 theo đề T4 | Emerging | Cần xác nhận file Q1 có gửi nhầm |

Thang điểm áp dụng theo mẫu được người dùng yêu cầu: V1=2, V2=2, V3=1, V4=1, V5=1, insight=3 (mỗi thành phần 1). V1-V5 Đạt 100%, Một phần 50%, chưa có phần việc đáp ứng T4 0%. Cộng theo trọng số trước rồi làm tròn half-up về số nguyên. Bảng điểm thành phần chỉ hiển thị trọng số nguyên và mức hoàn thành, tránh làm tròn riêng từng mục. Điểm 10 của Tâm vẫn kèm phần card cần sửa; kết quả làm tròn không xóa nhận xét đó. Không dùng kết luận yêu cầu nộp lại.

KPI2 Thu: số đơn các miền và tổng, AOV từng miền, GM% đều đúng; sai AOV dòng Tổng nên chấm Một phần 50% theo cùng nguyên tắc của mẫu thiếu ô tổng. Tên sheet, khoảng trắng và hoa/thường không bị trừ riêng khi phần việc đúng. V2 Thu khác chuỗi "trung bình" với nhãn "Trung", đúng ngưỡng nên 50% theo mẫu. Với Tuyên, chưa có bằng chứng T4 nên không ghi nhận cột tính Q1 là hoàn thành T4; các chênh lệch được gộp về một lỗi khác kỳ.

'''+kept+'''Thử script và tài liệu gốc:

grade_b2.py đã kiểm toàn cột và từng ô bảng, drawing/VML. Bằng chứng chấm trước được dùng lại; không mở hay tính lại 7 workbook trong bước chuyển mẫu. Answer 8/8 chỉ tiêu khớp, việc 1-5 Đạt sau đọc card; thiếu sheet Insight so với De_bai!C11. Đề 6/8 chỉ tiêu khớp, đã có cột tính và bảng Sheet1, bốn mã Sales!F53/F230/F1311/F1520 đã sạch nên không còn đúng trạng thái lỗi cài sẵn Dap_an_loi dòng 4. Các ngoại lệ vẫn lưu ở Thu_script_B2.json; không sửa đề/Answer.

5 bài nộp đều mở được .xlsx; tên file không theo chuỗi B2_BTVN_<HoTen>.xlsx. Tên nhận xét Thu theo mẫu đã có; cần đối chiếu danh sách lớp nếu muốn xác nhận họ tên đầy đủ cho mọi học viên. Không phát hiện dấu hiệu bất thường đủ căn cứ giữa các học viên. Hai bản cùng Thu giống card/insight được coi là lịch sử cùng một người, không phải dấu hiệu giữa hai học viên.

Ảnh minh họa: 7 khung chờ, chưa có ảnh chụp. MD/PDF ghi rõ nguồn bản sao, sheet, vùng và trạng thái. Không dựng ảnh Excel giả. Danh sách tại _DANH_SACH_ANH_NX.md; Images/ có sẵn ở mỗi thư mục.

Đầu ra chính: ToThiThu/, VuLeThaiNgoc/, PhamNhuTam/, LeLamTuyen/, Tong_hop_B2.csv (UTF-8 BOM, 4 dòng, Diem_10 nguyên). PDF dùng build_nhanxet_pdf.py hiện có theo từng file nhận xét; không gọi main để tránh lấy thư mục backup. Kiểm PDF và trạng thái tạo được ghi trong Kiem_tra_PDF_B2.json.
'''
(out/'Bao_cao_lop_B2.md').write_text(report,encoding='utf8')
print('Đã áp mẫu',[(r['Ho_ten'],r['Diem_10']) for r in csv_rows],len(all_images),'khung ảnh',flush=True)
