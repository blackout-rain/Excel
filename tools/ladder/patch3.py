# 使い方: python3 -I patch3.py <入力.xlsx> <LibreOffice再計算済み.xlsx> <出力.xlsx>
# 進捗確認シートの保存済み計算結果(キャッシュ値)のうち、再計算で変わったものだけ書き換える
import zipfile,re,sys,warnings,openpyxl
warnings.filterwarnings('ignore')
src,rec,out=sys.argv[1:4]
a=openpyxl.load_workbook(src,data_only=True)['進捗確認']; b=openpyxl.load_workbook(rec,data_only=True)['進捗確認']
ch={}
for row in b.iter_rows():
    for x in row:
        va=a[x.coordinate].value
        if isinstance(x.value,(int,float)) and isinstance(va,(int,float)) and abs(x.value-va)>1e-9: ch[x.coordinate]=x.value
print(ch)
zin=zipfile.ZipFile(src); xml=zin.read('xl/worksheets/sheet1.xml').decode('utf8')
for c,v in ch.items():
    s=repr(int(v)) if float(v).is_integer() else repr(round(v,12))
    xml,n=re.subn(r'(<c r="'+c+r'"(?: s="\d+")?><f>[^<]*</f>)<v>[^<]*</v>',lambda m:m.group(1)+'<v>'+s+'</v>',xml)
    assert n==1,c
zout=zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d=zin.read(it.filename)
    if it.filename=='xl/worksheets/sheet1.xml': d=xml.encode('utf8')
    zout.writestr(it,d)
zout.close()
