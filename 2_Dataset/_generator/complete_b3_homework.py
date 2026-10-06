"""Complete only B3 homework, preserving every native Pivot/Chart/Slicer part.

Uses Artifact Tool for supplemental sheets, then transplants those worksheet parts.
Existing sheets are patched at XML level because Artifact Tool does not implement
GETPIVOTDATA and cannot reliably round-trip native PivotTable/Slicer features.
Run: python 2_Dataset/_generator/complete_b3_homework.py
Original files are backed up to a timestamped subdirectory before replacement.
"""
import copy, datetime, json, os, pathlib, re, shutil, subprocess, tempfile, zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
import openpyxl

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
ET.register_namespace('',S); ET.register_namespace('r',R)
def tag(n): return '{'+S+'}'+n
def pack(t): return ET.tostring(t,encoding='utf-8',xml_declaration=True)
def parts(p):
    with zipfile.ZipFile(p) as z: return {n:z.read(n) for n in z.namelist()}
def sheets(z):
    rel={x.attrib['Id']:x.attrib['Target'].lstrip('/') for x in ET.fromstring(z['xl/_rels/workbook.xml.rels'])}
    return {x.attrib['name']:('xl/'+rel[x.attrib['{'+R+'}id']] if not rel[x.attrib['{'+R+'}id']].startswith('xl/') else rel[x.attrib['{'+R+'}id']]) for x in ET.fromstring(z['xl/workbook.xml']).find(tag('sheets'))}
def cell(root,ref):
    sd=root.find(tag('sheetData')); rown=int(''.join(c for c in ref if c.isdigit()))
    row=next((r for r in sd if int(r.attrib['r'])==rown),None)
    if row is None: row=ET.SubElement(sd,tag('row'),{'r':str(rown)})
    c=next((c for c in row if c.attrib.get('r')==ref),None)
    return c if c is not None else ET.SubElement(row,tag('c'),{'r':ref})
def write(root,ref,value,formula=None,style=None):
    c=cell(root,ref)
    for x in list(c): c.remove(x)
    c.attrib.pop('t',None)
    if style is not None:c.set('s',str(style))
    if formula is not None: ET.SubElement(c,tag('f')).text=formula.lstrip('=')
    if isinstance(value,str):
        if formula is not None:c.set('t','str');ET.SubElement(c,tag('v')).text=value
        else:c.set('t','inlineStr');ET.SubElement(ET.SubElement(c,tag('is')),tag('t')).text=value
    else: ET.SubElement(c,tag('v')).text=str(value)
def normalize(root):
    for row in root.find(tag('sheetData')):
        row[:]=sorted(row,key=lambda c:(len(''.join(x for x in c.attrib['r'] if x.isalpha())),''.join(x for x in c.attrib['r'] if x.isalpha())))
    sd=root.find(tag('sheetData'));sd[:]=sorted(sd,key=lambda r:int(r.attrib['r']))
def merge_styles(original,donor):
    a=ET.fromstring(original);b=ET.fromstring(donor); maps={}
    for group in ['fonts','fills','borders']:
        target=a.find(tag(group));src=b.find(tag(group));maps[group]={}
        for i,x in enumerate(src):maps[group][i]=len(target);target.append(copy.deepcopy(x))
        target.set('count',str(len(target)))
    nf=a.find(tag('numFmts'))
    if nf is None:nf=ET.Element(tag('numFmts'),{'count':'0'});a.insert(0,nf)
    fm={};nextid=max([163]+[int(x.attrib['numFmtId']) for x in nf])+1
    sourceformats=b.find(tag('numFmts'))
    for x in list(sourceformats) if sourceformats is not None else []:
        old=int(x.attrib['numFmtId']);fm[old]=nextid;y=copy.deepcopy(x);y.set('numFmtId',str(nextid));nf.append(y);nextid+=1
    nf.set('count',str(len(nf)));cx=a.find(tag('cellXfs'));sm={}
    for i,x in enumerate(b.find(tag('cellXfs'))):
        y=copy.deepcopy(x)
        for attr,g in [('fontId','fonts'),('fillId','fills'),('borderId','borders')]:y.set(attr,str(maps[g][int(y.attrib.get(attr,0))]))
        n=int(y.attrib.get('numFmtId',0));y.set('numFmtId',str(fm.get(n,n)));y.set('xfId','0');sm[i]=len(cx);cx.append(y)
    cx.set('count',str(len(cx)));return pack(a),sm
def main():
    root=pathlib.Path(os.environ.get('MCI_ROOT',pathlib.Path(__file__).resolve().parents[2]))
    hw=root/'1_Sessions/B3_PivotTable/Homework';ans=hw/'B3_BTVN_MCI_Mart_T1_T4_2025_Answer.xlsx';de=hw/'B3_BTVN_MCI_Mart_T1_T4_2025.xlsx'
    if list(hw.glob('~$*.xlsx')):raise RuntimeError('Homework đang mở trong Excel; đóng file trước khi sửa.')
    w=openpyxl.load_workbook(ans,data_only=True,read_only=True);it=w['Sales'].values;headers=list(next(it));raw=[dict(zip(headers,r)) for r in it];w.close()
    # Compute directly from typed raw fields, including when Excel has not cached formulas yet.
    for x in raw:
        x['Revenue']=int(x['Quantity'])*x['Unit_Price'];x['COGS']=int(x['Quantity'])*x['Unit_Cost'];x['Gross_Margin']=x['Revenue']-x['COGS'];x['Thang']=x['Order_Date'].month;x['So_don']=int(x['Line_No']==1);x['Thu']=x['Order_Date'].strftime('%a')
    details={};st={};caches={}
    for m in (3,4):
        rr=[x for x in raw if x['Store_Code']=='ST07' and x['Thang']==m]
        details[m]=[[x[h] for h in headers] for x in rr]
        st[m]={'rev':sum(x['Revenue'] for x in rr),'orders':sum(x['So_don'] for x in rr),'days':len({x['Order_Date'] for x in rr})}
    q1=[x for x in raw if x['Thang']<=3];t4=[x for x in raw if x['Thang']==4]
    qa=sum(x['Revenue'] for x in q1)/sum(x['So_don'] for x in q1);ta=sum(x['Revenue'] for x in t4)/sum(x['So_don'] for x in t4)
    sh={m:sum(x['Revenue'] for x in raw if x['Thang']==m and x['Category']=='Bánh kẹo')/sum(x['Revenue'] for x in raw if x['Thang']==m) for m in (1,4)}
    checks=[]
    for m,col in [(3,'B'),(4,'C')]:
        for metric,r,label in [('orders',6,'Số đơn'),('days',7,'Số ngày có bán')]:checks.append({'label':f'ST07 tháng {m}: {label}','formula':f'=So_sanh!{col}{r}','expected':st[m][metric],'format':'#,##0'})
    checks += [
        {'label':'AOV quý 1 (tổng / tổng)','formula':'=So_sanh!B14','expected':qa,'format':'#,##0'},
        {'label':'AOV tháng 4 trừ AOV quý 1 (VNĐ/đơn)','formula':'=So_sanh!D14','expected':ta-qa,'format':'#,##0'},
        {'label':'Tỷ trọng Bánh kẹo tháng 1','formula':'=So_sanh!B18','expected':sh[1],'format':'0.0%'},
        {'label':'Bánh kẹo: T4 trừ T1 (điểm phần trăm)','formula':'=So_sanh!D18','expected':(sh[4]-sh[1])*100,'format':'0.0'}]
    runtime=pathlib.Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies'
    node=pathlib.Path(os.environ.get('MCI_ARTIFACT_NODE',str(runtime/'node/bin/node')))
    modules=pathlib.Path(os.environ.get('MCI_ARTIFACT_MODULES',str(runtime/'node/node_modules')))
    if not node.exists() or not modules.exists():
        raise RuntimeError('Cần runtime Artifact Tool: đặt MCI_ARTIFACT_NODE (node executable) và MCI_ARTIFACT_MODULES (thư mục node_modules của runtime Codex).')
    with tempfile.TemporaryDirectory(prefix='mci-b3-') as td:
        td=pathlib.Path(td)
        if os.name=='nt':subprocess.run(['cmd','/c','mklink','/J',str(td/'node_modules'),str(modules)],check=True,capture_output=True)
        else:(td/'node_modules').symlink_to(modules,target_is_directory=True)
        shutil.copyfile(pathlib.Path(__file__).with_suffix('.mjs'),td/'complete.mjs')
        def serial(x):
            if isinstance(x,datetime.datetime):return (x-datetime.datetime(1899,12,30)).days
            raise TypeError(type(x))
        (td/'input.json').write_text(json.dumps({'headers':headers,'details':details,'st':st,'q1aov':qa,'t4aov':ta,'share1':sh[1],'share4':sh[4],'checks':checks},ensure_ascii=False,default=serial))
        run=subprocess.run([str(node),str(td/'complete.mjs'),str(td/'input.json'),str(td/'supplement.xlsx'),str(td)],capture_output=True,text=True)
        if run.returncode:raise RuntimeError(run.stderr[-2500:])
        az=parts(ans);original_ans=dict(az);dz=parts(td/'supplement.xlsx');am=sheets(az);dm=sheets(dz)
        az['xl/styles.xml'],stylemap=merge_styles(az['xl/styles.xml'],dz['xl/styles.xml'])
        ss=ET.fromstring(dz['xl/sharedStrings.xml']) if 'xl/sharedStrings.xml' in dz else []
        strings=[''.join(t.text or '' for t in si.iter(tag('t'))) for si in ss]
        def transplant(name):
            tree=ET.fromstring(dz[dm[name]])
            for c in tree.iter(tag('c')):
                c.set('s',str(stylemap[int(c.attrib.get('s',0))]))
                if c.attrib.get('t')=='s':
                    v=strings[int(c.find(tag('v')).text)]
                    for child in list(c):c.remove(child)
                    c.set('t','inlineStr');ET.SubElement(ET.SubElement(c,tag('is')),tag('t')).text=v
            return tree
        wb=ET.fromstring(az['xl/workbook.xml']);rels=ET.fromstring(az['xl/_rels/workbook.xml.rels']);types=ET.fromstring(az['[Content_Types].xml'])
        for name in ['CT_ST07_T3','CT_ST07_T4','So_sanh','Huong_dan_cham']:
            t=transplant(name)
            if name=='So_sanh':
                for m,col in [(3,'B'),(4,'C')]:
                    last=len(details[m])+1;rng=f"'CT_ST07_T{m}'!D2:D{last}"
                    for r,val,f in [(5,st[m]['rev'],f"SUM('CT_ST07_T{m}'!M2:M{last})"),(6,st[m]['orders'],f"SUM('CT_ST07_T{m}'!R2:R{last})"),(7,st[m]['days'],f'SUMPRODUCT(1/COUNTIF({rng},{rng}))'),(8,st[m]['rev']/st[m]['days'],f'{col}5/{col}7')]:write(t,f'{col}{r}',val,f)
                for r,key in [(5,'rev'),(6,'orders'),(7,'days')]:write(t,f'D{r}',st[4][key]-st[3][key],f'C{r}-B{r}')
                write(t,'D8',st[4]['rev']/st[4]['days']-st[3]['rev']/st[3]['days'],'C8-B8')
                write(t,'B14',qa,'SUM(PV_AOV!B6:B8)/SUM(PV_AOV!C6:C8)');write(t,'C14',ta,'GETPIVOTDATA("AOV (VNĐ/đơn)",PV_AOV!$A$4,"Thang",4)');write(t,'D14',ta-qa,'C14-B14')
                for m,col in [(1,'B'),(4,'C')]:write(t,f'{col}18',sh[m],f'GETPIVOTDATA("Tỷ trọng",PV_Nganh!$A$4,"Category","Bánh kẹo","Thang",{m})')
                write(t,'D18',(sh[4]-sh[1])*100,'(C18-B18)*100')
            if name in am:path=am[name]
            else:
                sid=max(int(x.attrib['sheetId']) for x in wb.find(tag('sheets')))+1;rid='rIdB3Completion'+str(sid);path=f'xl/worksheets/b3completion{sid}.xml'
                ET.SubElement(wb.find(tag('sheets')),tag('sheet'),{'name':name,'sheetId':str(sid),'{'+R+'}id':rid})
                ET.SubElement(rels,'{'+P+'}Relationship',{'Id':rid,'Type':R+'/worksheet','Target':path[3:]})
                ET.SubElement(types,'{'+CT+'}Override',{'PartName':'/'+path,'ContentType':'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'})
            normalize(t);az[path]=pack(t)
        # Use exact name replacement without resaving unsupported extensions.
        for name,content in list(az.items()):
            if name.endswith('.xml') and (name.startswith('xl/pivotTables/') or name.startswith('xl/slicerCaches/')):az[name]=content.replace(b'name="pvtMien"',b'name="pvtMienPct"')
        kt=ET.fromstring(az[am['Kiem_tra']]);extra=transplant('Bo_sung_KT')
        for i,x in enumerate(checks,start=14):
            for c in next(r for r in extra.find(tag('sheetData')) if r.attrib['r']==str(i)):
                dest=cell(kt,c.attrib['r']);dest.attrib.clear();dest.attrib.update(c.attrib);dest[:]=[copy.deepcopy(a) for a in c]
            write(kt,f'C{i}',x['expected'],x['formula']);write(kt,f'E{i}','✔',f'IF(ROUND(C{i},6)=ROUND(D{i},6),"✔","✘")')
        write(kt,'A2','17 chỉ tiêu số liệu; bỏ lọc và Refresh All. Chấm đủ bài theo Huong_dan_cham.',style=cell(kt,'A2').attrib.get('s',0))
        kt.find(tag('dimension')).set('ref','A1:E21');normalize(kt);az[am['Kiem_tra']]=pack(kt)
        cp=wb.find(tag('calcPr'))
        if cp is None:cp=ET.SubElement(wb,tag('calcPr'))
        cp.set('fullCalcOnLoad','1');cp.set('forceFullCalc','1')
        az['xl/workbook.xml']=pack(wb);az['xl/_rels/workbook.xml.rels']=pack(rels);az['[Content_Types].xml']=pack(types)
        # Student instructions: align exact name and location with retained answer layout.
        zz=parts(de);original_de=dict(zz);zm=sheets(zz);tt=ET.fromstring(zz[zm['De_bai']]);sty=cell(tt,'B23').attrib.get('s',0)
        write(tt,'B23','Bấm dấu + tạo sheet mới, đổi tên thành PV_Mien. Tại Sales, bấm vào tblSales4T > Insert > PivotTable. Chọn Existing Worksheet, Location: PV_Mien!$A$12, OK. Mọi Pivot của bài đều dùng tblSales4T.',style=sty)
        write(tt,'B25','Tạo Pivot thứ hai từ tblSales4T: Insert > PivotTable > Existing Worksheet, Location: PV_Mien!$A$4, OK.',style=sty)
        write(tt,'C25','Pivot thứ hai xuất hiện từ A4, nằm trên Pivot doanh thu ở A12.',style=cell(tt,'C25').attrib.get('s',0))
        write(tt,'B36','Tạo Pivot thứ hai từ tblSales4T: Insert > PivotTable > Existing Worksheet, Location: PV_CuaHang!$H$4, OK. Kéo Store_Code vào Rows, Thang vào Columns, Revenue vào Values.',style=cell(tt,'B36').attrib.get('s',0))
        write(tt,'B39','Trong sheet chi tiết, chọn cột So_don và đọc Sum trên thanh trạng thái để lấy số đơn. Đếm ngày có bán: gõ =COUNTA(UNIQUE( rồi quét đúng các ô ngày trong cột Order_Date, bỏ tiêu đề và ô trống, gõ )) và Enter. Cần Excel 365; nếu không có UNIQUE, sao chép cột ngày sang vùng trống, Data > Remove Duplicates rồi đếm các ngày còn lại.',style=cell(tt,'B39').attrib.get('s',0))
        write(tt,'B55','Bấm chuột phải Slicer > Report Connections. Tick pvtMienDT, pvtMienPct, pvtNganh, pvtCuaHangDT, pvtCuaHangPct. Nếu thiếu Pivot, kiểm lại nguồn tblSales4T và Pivot cache; có thể sao chép một Pivot đã nối được rồi đổi trường để dùng chung cache.',style=cell(tt,'B55').attrib.get('s',0))
        zz[zm['De_bai']]=pack(tt)
        backup=hw/'_backup'/datetime.datetime.now().strftime('%Y%m%d_%H%M%S');backup.mkdir(parents=True)
        for p,z,original in [(ans,az,original_ans),(de,zz,original_de)]:
            # Retain prefix declarations referenced by mc:Ignorable and extension attributes.
            for n,v in list(z.items()):
                if n.endswith('.xml') and n in original and v!=original[n]:
                    oldroot=re.search(rb'<[A-Za-z_][^>]*>',original[n]).group(0)
                    newroot=re.search(rb'<[A-Za-z_][^>]*>',v).group(0)
                    for prefix,uri in re.findall(rb'xmlns:([\w]+)="([^"]+)"',oldroot):
                        if b'xmlns:'+prefix+b'=' not in newroot:
                            newroot=newroot[:-1]+b' xmlns:'+prefix+b'="'+uri+b'">'
                    z[n]=re.sub(rb'<[A-Za-z_][^>]*>',lambda _:newroot,v,count=1)
            shutil.copy2(p,backup/p.name);tmp=td/(p.name+'.tmp')
            with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as out:
                for n,v in z.items():out.writestr(n,v)
            with zipfile.ZipFile(tmp) as out:assert out.testzip() is None
            shutil.copyfile(tmp,p)
        qa=hw/'_qa';qa.mkdir(exist_ok=True)
        for n in ['So_sanh','Huong_dan_cham']:shutil.copy2(td/(n+'.png'),qa/(n+'.png'))
        print('Completed B3 homework; 17 numeric checks; backups:',backup)
if __name__=='__main__':main()
