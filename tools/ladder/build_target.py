# 使い方: python3 -I build_target.py <店舗一覧.xlsx>  → カレントに target.json を出力
# 店舗一覧「測定項目ラダー」(SBショップ) → 進捗確認「ラダー」の列対応で目標値を作る
import openpyxl,warnings,json,sys
warnings.filterwarnings('ignore')
src=openpyxl.load_workbook(sys.argv[1],data_only=True)['測定項目ラダー']
def col(c): return openpyxl.utils.column_index_from_string(c)
def ladder(rc, sc, r0, r1):
    out=[]
    for r in range(r0,r1+1):
        a=src.cell(r,col(rc)).value; b=src.cell(r,col(sc)).value
        if isinstance(a,(int,float)) and isinstance(b,(int,float)): out.append((round(a,6),round(b,6)))
    return sorted(out)
# (進捗側 率列, 配点列, 店舗一覧 率列, 配点列, 行範囲)
maps={
 '新規':('A','B','D','E',6,35),'MNP':('D','E','F','G',6,35),'機変':('G','H','H','I',6,35),
 '固定':('J','K','J','K',6,35),'タブレット':('M','N','L','M',6,35),'電気':('P','Q','N','O',6,35),
 'PPCD':('S','T','F','G',48,56),'SBGold':('V','W','H','I',48,56),'SB支払設定':('AR','AS','J','K',48,56),
 '戦略機種':('Y','Z','L','M',48,56),'Watch':('AB','AC','N','O',48,56),
 'SB通話OP':('AF','AG','M','N',39,44),'SB保証':('AI','AJ','O','P',39,44),'SBセキュリティ':('AL','AM','Q','R',39,44),
}
res={}
for k,(pr,ps,sr,ss,r0,r1) in maps.items():
    res[k]={'cols':[pr,ps],'rows':ladder(sr,ss,r0,r1)}
# Watch 全メーカー得点 (P列)
res['Watch']['all']=sorted((round(src.cell(r,col('N')).value,6),src.cell(r,col('P')).value) for r in range(48,53))
# SBプラン×大容量
hdr=[round(src.cell(38,c).value,6) for c in range(col('F'),col('K')+1)]
grid=[]
for r in range(39,45):
    grid.append([round(src.cell(r,col('E')).value,6)]+[src.cell(r,c).value for c in range(col('F'),col('K')+1)])
res['SBプラン']={'hdr':hdr,'grid':sorted(grid)}
json.dump(res,open('target.json','w'),ensure_ascii=False,indent=0)
for k,v in res.items(): print(k,v.get('rows',v.get('grid')), v.get('all',''), v.get('hdr',''))
