# -*- coding: utf-8 -*-
"""Rebuilds 葉耀坤郵局_現金存款 so that reversed (沖銷) entries are excluded by
formula, and the annual table is computed by SUMIFS instead of by hand."""
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SRC, FN = sys.argv[1], sys.argv[2]
F="Arial"
HDR=Font(name=F,bold=True,size=10,color="FFFFFF"); FILL=PatternFill("solid",fgColor="1F4E79")
H2=Font(name=F,bold=True,size=11,color="1F4E79")
BOLD=Font(name=F,bold=True,size=10); BODY=Font(name=F,size=10)
RED=Font(name=F,size=10,color="C00000",bold=True); NOTE=Font(name=F,size=9,color="595959")
YEL=PatternFill("solid",fgColor="FFFF00"); AMB=PatternFill("solid",fgColor="FFF2CC")
GRY=PatternFill("solid",fgColor="D9D9D9")
thin=Side(style="thin",color="BFBFBF"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)

src=load_workbook(SRC,data_only=True,read_only=True); ws0=src["交易明細"]
rows=list(ws0.iter_rows(values_only=True))
h=next(i for i,r in enumerate(rows) if r and r[0]=="頁次")
c={str(v):j for j,v in enumerate(rows[h]) if v}
def f(v):
    try: return float(v)
    except: return 0.0
REC=[]
for r in rows[h+1:]:
    if not r or not r[c["資料日期（民國）"]]: continue
    g=lambda k: r[c[k]] if k in c else None
    code=str(g("交易代號") or "")
    if code not in ("1501","1520"): continue
    inn=f(g("存款金額"))
    if inn<=0: continue
    d=str(g("資料日期（民國）"))
    REC.append((d, code, str(g("中文摘要") or ""), inn, f(g("結存金額")),
                str(g("被沖銷") or "").strip()))
src.close()
# page 1 (96年1-9月) — date column mis-registered on the original, so day is not asserted
P1=[("96","1501","現金存款",110000),("96","1501","現金存款",270000),
    ("96","1501","現金存款",60000),("96","1501","現金存款",20000),
    ("96","1501","現金存款",20000),("96","1501","現金存款",51000),
    ("96","1501","現金存款",140000),("96","1501","現金存款",20000),
    ("96","1520","無摺存款",25000),("96","1520","無摺存款",29000),
    ("96","1501","現金存款",150000),("96","1501","現金存款",54000),
    ("96","1520","無摺存款",29000)]
# pages 2-3 recovered today, already in their own sheet
P23=[("960905","1520","無摺存款",25000),("960907","1501","現金存款",90500),
     ("961005","1520","無摺存款",25000),("961005","1520","無摺存款",29000),
     ("961011","1501","現金存款",32000),("961106","1520","無摺存款",29000),
     ("961106","1520","無摺存款",25000),("961114","1501","現金存款",50000),
     ("961122","1501","現金存款",32000),("961205","1501","現金存款",19000),
     ("970105","1501","現金存款",20000),("970115","1501","現金存款",130000),
     ("970212","1501","現金存款",140000),("970310","1501","現金存款",40000),
     ("970317","1501","現金存款",20000),("970408","1501","現金存款",150000)]

wb=load_workbook(FN)
name="葉耀坤郵局_現金存款"
idx=wb.sheetnames.index(name); del wb[name]
ws=wb.create_sheet(name, idx)
ws.cell(1,1,"葉耀坤　郵局 0737273　現金類存入逐筆（民國 96.09～101.12）").font=H2
ws.cell(2,1,"含交易代號 1501 現金存款與 1520 無摺存款。原清單「被沖銷」欄標 R 者，係當日經沖銷之交易，"
            "不構成實際存入，本表以「沖銷」標示並由公式自合計排除。").font=NOTE
cols=["#","來源頁","民國年","資料日期","西元日期","交易代號","中文摘要","存入金額","存入後結存","是否沖銷","計入合計"]
for i,hh in enumerate(cols,1):
    cc=ws.cell(4,i,hh); cc.font=HDR; cc.fill=FILL
    cc.alignment=Alignment("center","center"); cc.border=BOX

ALL=[("96"+"0000",code,memo,inn,None,"","1") for _,code,memo,inn in P1] + \
    [(d,code,memo,inn,None,"",  "2-3") for d,code,memo,inn in P23] + \
    [(d,code,memo,inn,bal,rev,"4-16") for d,code,memo,inn,bal,rev in REC]
ALL.sort(key=lambda x: (x[0] if len(x[0])==7 else "0"+x[0]))

r=5; first=r
for n,(d,code,memo,inn,bal,rev,pg) in enumerate(ALL,1):
    roc = d[:3] if len(d)==7 else "0"+d[:2]
    ws.cell(r,1,n); ws.cell(r,2,pg); ws.cell(r,3,int(roc))
    if pg=="1":
        ws.cell(r,4,"96年1-9月"); ws.cell(r,5,"2007（日期欄錯位，未認定）")
    else:
        ws.cell(r,4,d); ws.cell(r,5,f"{int(roc)+1911}/{d[-4:-2]}/{d[-2:]}")
    ws.cell(r,6,code); ws.cell(r,7,memo)
    cc=ws.cell(r,8,inn); cc.number_format='#,##0'
    if bal is not None:
        cc=ws.cell(r,9,bal); cc.number_format='#,##0'
    ws.cell(r,10,"沖銷" if rev else "")
    ws.cell(r,11,f'=IF(J{r}="沖銷",0,H{r})').number_format='#,##0'
    for i in range(1,12):
        cc=ws.cell(r,i); cc.border=BOX; cc.font=RED if rev else BODY
        if rev: cc.fill=GRY
    r+=1
last=r-1
ws.cell(r,7,"合計（已排除沖銷）").font=BOLD
cc=ws.cell(r,11,f'=SUM(K{first}:K{last})'); cc.font=BOLD; cc.fill=YEL; cc.number_format='#,##0'
cc=ws.cell(r,8,f'=SUM(H{first}:H{last})'); cc.font=BODY; cc.number_format='#,##0'
ws.cell(r,10,"← 未排除者").font=NOTE
r+=2

ws.cell(r,2,"逐年彙總（公式計算，已排除沖銷）").font=H2; r+=1
for i,hh in enumerate(["民國年","筆數","現金類存入"],2):
    cc=ws.cell(r,i,hh); cc.font=HDR; cc.fill=FILL
    cc.alignment=Alignment("center","center"); cc.border=BOX
r+=1; ysum_first=r
for yr in (96,97,98,99,100,101):
    ws.cell(r,2,yr).font=BODY; ws.cell(r,2).border=BOX
    cc=ws.cell(r,3,f'=COUNTIFS($C${first}:$C${last},B{r},$J${first}:$J${last},"")')
    cc.font=BODY; cc.border=BOX
    cc=ws.cell(r,4,f'=SUMIFS($K${first}:$K${last},$C${first}:$C${last},B{r})')
    cc.font=BODY; cc.border=BOX; cc.number_format='#,##0'
    if yr in (100,101):
        for i in (2,3,4): ws.cell(r,i).fill=AMB
    r+=1
ysum_last=r-1
ws.cell(r,2,"合計").font=BOLD
for col,L in ((3,"C"),(4,"D")):
    cc=ws.cell(r,col,f'=SUM({L}{ysum_first}:{L}{ysum_last})'); cc.font=BOLD; cc.fill=YEL
    if col==4: cc.number_format='#,##0'
r+=1
ws.cell(r,2,"96～98 三年小計").font=BOLD
cc=ws.cell(r,4,f'=SUM(D{ysum_first}:D{ysum_first+2})'); cc.font=BOLD; cc.fill=YEL; cc.number_format='#,##0'
r+=1
ws.cell(r,2,"100～101 兩年小計").font=BOLD
cc=ws.cell(r,4,f'=SUM(D{ysum_first+4}:D{ysum_first+5})'); cc.font=BOLD; cc.fill=YEL; cc.number_format='#,##0'
r+=2

for t,fo in [
 ("查核結論（民國 115.11.14 二次修正，據實記載）",H2),
 ("⚠ 更正二：民國 98.02.16 之「現金存款 300,009 元」於同日經沖銷（原清單「被沖銷」欄標 R，並有對應之沖銷列），"
  "並非實際存入。先前之 98 年合計誤將其計入。本表已以公式排除，上方逐年彙總為正確數字。",RED),
 ("⚠ 更正一（前次）：97、98 兩年曾分別誤植為 593,000 與 1,287,009。",RED),
 ("",BODY),
 ("一、99 年 10、11、12 三個月，本帳戶無任何現金存入。",RED),
 ("二、99.12.23 葉耀坤代為清償房貸「之後」，本帳戶並無每月固定存入 50,000 元之形態："
  "100 年全年僅 2 筆（100.02.25 計 135,000 元、100.10.21 計 200,000 元），101 年僅 3 筆合計 118,000 元。",RED),
 ("三、密集而大額之現金存入集中於代償「之前」的 96～98 年（見上方小計）；99 年起即逐年遞減。",RED),
 ("",BODY),
 ("可能之解釋（均須另行查證，不得逕採）：",BOLD),
 ("　(1) 交付現金之起訖時間與當事人記憶有出入——96～98 年之密集存入，或即為當時按期交付之款項。",BODY),
 ("　(2) 99.12.23 以後交付之現金，葉耀坤未存入本帳戶。本帳戶有「定存淨額」入帳 10 筆合計 3,800,000 元，"
  "足見另有定存帳戶，現金可能逕存為定存或存入他行。",BODY),
 ("　(3) 第 1 頁（民國 96 年 1～8 月）及 102 年以後之紀錄仍未取得。",BODY),
 ("",BODY),
 ("★ 最具證據價值者為匯入／轉入 10 筆，合計 2,904,000 元 —— 非現金，轉出人可查：",RED),
 ("　【第1頁‧民國96年】跨行匯入 1,500,000（經辦局 0100302）／跨行轉入 450,000（經辦局 808）／入戶匯款 54,000（經辦局 010103）"
  "　※ 原件日期欄錯位，確切日期須以原件為準",RED),
 ("　96.11.02　200,000　／　96.11.02　300,000　（經辦局 808，對方帳號 0000288968134706）",RED),
 ("　97.05.05　60,000（對方帳號 0099634963401081 68800）　／　99.04.05　120,000（經辦局 808，同 0000288968134706）",RED),
 ("　100.05.09　150,000　／　100.07.24　40,000　／　100.07.25　30,000（對方帳號 0000220004009872）",RED),
 ("★★ 第 1 頁摘要欄另印有兩個人名：「葉兆雄」與「葉兆輝」。郵局清單於匯入款項之摘要欄所載人名通常即匯款人。"
  "惟原件欄位印表錯位，無法僅憑翻拍認定對應何筆，必須向中華郵政調取該三筆之匯款單／交易憑證。",RED),
 ("★ 如其中任一筆之匯款人確為葉兆輝，即為直接書證，證明力遠高於現金存款之推論。此為目前最值得立即調取者。",RED),
 ("",BODY),
 ("★ 務請注意：如逕向法院主張「每月交付現金 5 萬元」，而本帳戶明細不能支持，對造一經調閱即可反證，"
  "反傷及其餘可信之主張。宜改以「跨行轉入 7 筆可追查」及「有足額現金提領、時序與房貸期間相符」之方式陳述。",RED),
]:
    cc=ws.cell(r,1,t); cc.font=fo; cc.alignment=Alignment(vertical="top"); r+=1

for col,w in zip("ABCDEFGHIJK",[5,9,9,13,13,10,12,14,14,10,14]):
    ws.column_dimensions[col].width=w
ws.freeze_panes="A5"
wb.save(FN); print("rebuilt:", name, "rows:", last-first+1)
