# 使い方: python3 -I patch2.py <入力.xlsx> <出力.xlsx>  … ラダー!A2 の注記に出典を追記
import zipfile,sys
src,out=sys.argv[1:3]
old='<si><t>「測定項目ラダー」の表です。'
new='<si><t>【別紙】2026年11月度店舗一覧 ver0.9 の「測定項目ラダー」(SBショップラダー)の表です。'
zin=zipfile.ZipFile(src); zout=zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d=zin.read(it.filename)
    if it.filename=='xl/sharedStrings.xml':
        s=d.decode('utf8'); assert s.count(old)==1; d=s.replace(old,new).encode('utf8')
    zout.writestr(it,d)
zout.close()
