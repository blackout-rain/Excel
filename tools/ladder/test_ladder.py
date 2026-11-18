# 使い方: RECALC=<recalc.py> python3 -I test_ladder.py <進捗確認.xlsx> [target.json]  (out/ に検証用コピーを作る)
# 境界値テスト: 目標=1000 にして累計実績を定数で入れ、全ラダー項目の J/K/L 列を Python の参照計算と照合
import openpyxl,warnings,json,subprocess,sys,os
warnings.filterwarnings('ignore')
T=json.load(open(sys.argv[2] if len(sys.argv)>2 else 'target.json'))
RECALC=os.environ.get('RECALC','recalc.py')  # xlsx スキルの scripts/recalc.py のパスを環境変数 RECALC で指定
items={6:('新規',7),7:('MNP',8),8:('固定',9),9:('タブレット',10),10:('電気',11),13:('戦略機種',14),14:('Watch',15),15:('機変',16)}
def ref(rows,rate):
    pts=[p for th,p in rows if th<=round(rate,6)+1e-12]; cur=pts[-1]
    nxt=[th for th,p in rows if th>round(rate,6)+1e-12]
    return cur,(nxt[0] if nxt else '最高段')
fails=0; n=0
# シナリオ: 各段のちょうど / 0.1%下 / 0 / 最高段超え
for scen in range(4):
    wb=openpyxl.load_workbook(sys.argv[1]); ws=wb['進捗確認']; tg=wb['目標設定']
    exp={}
    for r,(k,tr) in items.items():
        rows=[tuple(x) for x in T[k]['rows']]
        th=[x[0] for x in rows]
        if scen==0: rate=th[len(th)//2]
        elif scen==1: rate=th[len(th)//2]-0.001
        elif scen==2: rate=th[-1]+0.2
        else: rate=th[1]
        tg[f'D{tr}']=1000; ws[f'C{r}']=round(rate*1000,6)
        exp[r]=(k,rate)+ref(rows,rate)
    os.makedirs('out',exist_ok=True); f=f'out/test{scen}.xlsx'; wb.save(f)
    subprocess.run(['python3',RECALC,f,'90'],check=True,capture_output=True)
    v=openpyxl.load_workbook(f,data_only=True)['進捗確認']
    for r,(k,rate,cur,nxt) in exp.items():
        n+=1
        got=(v[f'J{r}'].value,v[f'K{r}'].value)
        ok=got[0]==cur and (got[1]==nxt or (isinstance(nxt,float) and abs(got[1]-nxt)<1e-9))
        if not ok: fails+=1
        print('OK ' if ok else 'NG ',scen,k,f'率={rate:.4f}','配点',got[0],'期待',cur,'次段',got[1],'期待',nxt)
print(f'{n-fails}/{n} passed'); sys.exit(1 if fails else 0)
