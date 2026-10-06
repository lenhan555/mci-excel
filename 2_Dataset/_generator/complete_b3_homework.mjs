// Author only supplemental sheets; Python preserves the original native Excel parts.
import fs from 'node:fs/promises';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';
try {
const [input,output,preview]=process.argv.slice(2);
const d=JSON.parse(await fs.readFile(input,'utf8'));
const w=Workbook.create();
function sheet(name,values,widths){
 const s=w.worksheets.add(name); const count=Math.max(1,...values.map(r=>r.length));
 s.getRange('A1').write(values.map(r=>Array.from({length:count},(_,i)=>r[i]??null)));
 const r=s.getUsedRange(); r.format.font={name:'Arial',size:10};
 s.showGridLines=false;
 widths.forEach((width,i)=>{s.getRangeByIndexes(0,i,1,1).format.columnWidth=width/9;});
 return s;
}
function header(s,range){s.getRange(range).format={fill:'#203864',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true};}
for(const month of [3,4]){
 const data=d.details[month];
 const s=sheet(`CT_ST07_T${month}`,[d.headers,...data],d.headers.map(h=>h==='Product_Name'?230:h==='Order_ID'?125:110));
 header(s,'A1:R1'); s.freezePanes.freezeRows(1);
 s.getRange(`D2:D${data.length+1}`).format.numberFormat='dd/mm/yyyy';
 s.getRange(`J2:O${data.length+1}`).format.numberFormat='#,##0';
}
const s=sheet('So_sanh',[
 ['SO SÁNH ST07 VÀ AOV'],
 ['Câu 4: số liệu từ hai sheet chi tiết; ngày có bán là ngày có ít nhất một giao dịch.'],[],
 ['Chỉ tiêu','Tháng 3','Tháng 4','Chênh lệch'],
 ['Doanh thu (VNĐ)',d.st[3].rev,d.st[4].rev],
 ['Số đơn',d.st[3].orders,d.st[4].orders],
 ['Số ngày có bán',d.st[3].days,d.st[4].days],
 ['Doanh thu/ngày có bán (VNĐ)'],
 ['Tháng 4 tăng doanh thu chủ yếu nhờ có nhiều ngày bán và nhiều đơn hơn.'],
 ['Tăng so với tháng 3 chưa đủ để khẳng định đã trở lại mức bán bình thường.'],[],
 ['Câu 5: AOV = tổng doanh thu / tổng số đơn; không lấy trung bình AOV các tháng.'],
 ['Chỉ tiêu','Quý 1','Tháng 4','Chênh lệch'],
 ['AOV (VNĐ/đơn)',d.q1aov,d.t4aov],
 ['AOV tháng 4 cao hơn quý 1 khoảng '+Math.round(d.t4aov-d.q1aov).toLocaleString('vi-VN')+' VNĐ/đơn.'],[],
 ['Câu 3: tỷ trọng Bánh kẹo','Tháng 1','Tháng 4','Chênh lệch (điểm %)'],
 ['Tỷ trọng',d.share1,d.share4]
 ],[530,150,150,190]);
for(const n of [4,13,17])header(s,`A${n}:D${n}`);
for(const n of [1,2,9,10,12,15]){s.mergeCells(`A${n}:D${n}`);s.getRange(`A${n}:D${n}`).format.wrapText=true;s.getRange(`A${n}:D${n}`).format.rowHeight=n===1?25:30;}
for(const n of [5,6,7,8,14])s.getRange(`B${n}:D${n}`).format.numberFormat='#,##0';
for(const n of [5,6,7,8,14])s.getRange(`D${n}`).formulas=[[`=C${n}-B${n}`]];
for(const [m,col] of [[3,'B'],[4,'C']]){
 const last=d.details[m].length+1;
 s.getRange(`${col}5`).formulas=[[`=SUM('CT_ST07_T${m}'!M2:M${last})`]];
 s.getRange(`${col}6`).formulas=[[`=SUM('CT_ST07_T${m}'!R2:R${last})`]];
 s.getRange(`${col}8`).formulas=[[`=${col}5/${col}7`]];
}
s.getRange('B18:C18').format.numberFormat='0.0%';
s.getRange('D18').formulas=[['=(C18-B18)*100']];s.getRange('D18').format.numberFormat='0.0';
const guide=sheet('Huong_dan_cham',[
 ['HƯỚNG DẪN CHẤM HOMEWORK BUỔI 3'],
 ['Đối chiếu số liệu ở Kiem_tra và PV_DoiSoat; kiểm thao tác, biểu đồ và insight riêng.'],
 ['Trước khi chấm: bỏ lọc Slicer, trả đủ bốn tháng, Refresh All.'],[],
 ['Câu','Tiêu chí cần kiểm','Nguồn đối chiếu'],
 ['1','12 ô Revenue, So_don, Quantity theo tháng khớp hệ thống; sửa đủ 6 ô Quantity dạng chữ; Insight!A2 nêu lỗi và cách sửa.','PV_DoiSoat; So_lieu_he_thong; Dap_an_loi'],
 ['2','Hai Pivot miền: doanh thu và % Difference From theo Thang, (previous); kết luận miền Trung +19,8%.','PV_Mien'],
 ['3','% of Column Total; tỷ trọng Bánh kẹo tháng 1, tháng 4 và chênh lệch tính bằng điểm phần trăm.','PV_Nganh; So_sanh!B18:D18'],
 ['4','Có CT_ST07_T3 và CT_ST07_T4; nêu đúng số đơn, ngày bán, doanh thu và tăng trưởng; giải thích từ chi tiết.','So_sanh!A4:D10'],
 ['5','Calculated Field AOV = Revenue/So_don; so tháng 4 với quý 1 theo tổng/tổng.','PV_AOV; So_sanh!A13:D15'],
 ['6','Ba PivotChart có tiêu đề kết luận đúng số; Slicer Region nối cả 5 Pivot miền/ngành/cửa hàng. Bấm thử từng miền, kiểm mọi Pivot và biểu đồ, rồi bỏ lọc.','pvtMienDT; pvtMienPct; pvtNganh; pvtCuaHangDT; pvtCuaHangPct'],
 ['7','Insight có số và phép so sánh, nguyên nhân khả dĩ, hàm ý; số khớp Pivot. Chấp nhận cách viết và kết luận khác nếu có bằng chứng.','Insight!A5 của học viên; Insight_mau'],[],
 ['Lưu ý','Các dấu khớp ở Kiem_tra chỉ kiểm số liệu; chưa chứng minh biểu đồ, kết nối Slicer hoặc insight đã đạt.'],
 ['Ghi nhận','Ghi Đạt / Cần sửa và dẫn vị trí cần sửa cho từng câu; không tự suy ra điểm từ số dấu khớp.']
 ],[90,760,500]);
header(guide,'A5:C5');guide.getRange('A6:C12').format.wrapText=true;guide.getRange('A6:C12').format.rowHeight=58;
for(const n of [1,2,3])guide.mergeCells(`A${n}:C${n}`);
for(const n of [14,15]){guide.mergeCells(`B${n}:C${n}`);guide.getRange(`B${n}:C${n}`).format.wrapText=true;guide.getRange(`A${n}:C${n}`).format.rowHeight=35;}
const k=sheet('Bo_sung_KT',Array.from({length:13},()=>[]),[360,20,150,150,65]);
d.checks.forEach((x,i)=>{
 const r=14+i;k.getRange(`A${r}:E${r}`).values=[[x.label,null,null,x.expected,null]];
 k.getRange(`C${r}`).formulas=[[x.formula]];
 k.getRange(`E${r}`).formulas=[[`=IF(ROUND(C${r},6)=ROUND(D${r},6),"✔","✘")`]];
 k.getRange(`C${r}:D${r}`).format.numberFormat=x.format;
});
w.recalculate();
await (await SpreadsheetFile.exportXlsx(w)).save(output);
for(const name of ['So_sanh','Huong_dan_cham']){
 const b=await w.render({sheetName:name,range:name==='So_sanh'?'A4:D18':'A5:C12',scale:1,format:'png'});
 await fs.writeFile(`${preview}/${name}.png`,new Uint8Array(await b.arrayBuffer()));
}
} catch(error) { console.error('B3 supplemental build:',error.message); process.exitCode=1; }
