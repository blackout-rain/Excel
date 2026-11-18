# 使い方: python3 -I patch.py <進捗確認.xlsx> <target.json> <出力.xlsx>
# ラダーシート(sheet3.xml)の数値セルを XML 直接書き換えで差し替える(プルダウン等の拡張を壊さないため)
import zipfile,re,json,sys,warnings,openpyxl,shutil
warnings.filterwarnings('ignore')
src,tj,out=sys.argv[1:4]
T=json.load(open(tj))
wb=openpyxl.load_workbook(src); ws=wb['ラダー']
want={}
def put(c,r,v): want[f'{c}{r}']=v
for k,v in T.items():
    if k=='SBプラン':
        cols=['AV','AW','AX','AY','AZ','BA']
        for i,h in enumerate(v['hdr']): put(cols[i],4,h)
        for i,row in enumerate(v['grid']):
            put('AU',5+i,row[0])
            for j,x in enumerate(row[1:]): put(cols[j],5+i,x)
        continue
    rc,sc=v['cols']
    rows=v['rows']
    for i in range(30):
        r=5+i
        if i<len(rows): put(rc,r,rows[i][0]); put(sc,r,rows[i][1])
        else: put(rc,r,None); put(sc,r,None)
    if 'all' in v:
        for i in range(30):
            put('AD',5+i, v['all'][i][1] if i<len(v['all']) else None)
changes={}
for coord,v in want.items():
    cur=ws[coord].value
    if isinstance(cur,float): cur=round(cur,6)
    if cur!=v: changes[coord]=(ws[coord].value,v)
for c,(a,b) in sorted(changes.items(),key=lambda x:(len(re.match('[A-Z]+',x[0]).group()),x[0])): print(c,a,'->',b)
print('changes',len(changes))
zin=zipfile.ZipFile(src)
xml=zin.read('xl/worksheets/sheet3.xml').decode('utf8')
for coord,(a,b) in changes.items():
    if b is None or a is None or isinstance(b,str): sys.exit('unexpected add/remove '+coord)
    pat=re.compile(r'(<c r="'+coord+r'"(?: s="\d+")?>)<v>[^<]*</v>(</c>)')
    xml,n=pat.subn(lambda m:m.group(1)+'<v>'+repr(b if not float(b).is_integer() else int(b))+'</v>'+m.group(2),xml)
    if n!=1: sys.exit('patch fail '+coord)
wbx=zin.read('xl/workbook.xml').decode('utf8')
wbx2=wbx.replace('<calcPr calcId="191029"/>','<calcPr calcId="191029" fullCalcOnLoad="1"/>')
assert wbx2!=wbx or 'fullCalcOnLoad' in wbx
zout=zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    data=zin.read(it.filename)
    if it.filename=='xl/worksheets/sheet3.xml': data=xml.encode('utf8')
    elif it.filename=='xl/workbook.xml': data=wbx2.encode('utf8')
    zout.writestr(it,data)
zout.close()
