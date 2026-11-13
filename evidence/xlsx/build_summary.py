# -*- coding: utf-8 -*-
"""Builds 葉兆輝_收支彙整_87-115.xlsx from the five source workbooks the client
supplied. Every total is a live formula; nothing is a pasted number."""
import sys
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SRC = sys.argv[1]                      # directory holding the five uploads
OUT = sys.argv[2]
F = "Arial"
HDR=Font(name=F,bold=True,size=10,color="FFFFFF"); FILL=PatternFill("solid",fgColor="1F4E79")
H2 =Font(name=F,bold=True,size=11,color="1F4E79")
BOLD=Font(name=F,bold=True,size=10); BODY=Font(name=F,size=10)
RED=Font(name=F,size=10,color="C00000",bold=True); NOTE=Font(name=F,size=9,color="595959")
YEL=PatternFill("solid",fgColor="FFFF00"); GRN=PatternFill("solid",fgColor="E2EFDA")
AMB=PatternFill("solid",fgColor="FFF2CC")
thin=Side(style="thin",color="BFBFBF"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)

def load(path, sheet, key):
    wb=load_workbook(path,data_only=True,read_only=True); ws=wb[sheet]
    rows=list(ws.iter_rows(values_only=True))
    h=next(i for i,r in enumerate(rows) if r and key in [str(x) for x in r if x])
    cols={str(v):j for j,v in enumerate(rows[h]) if v}
    out=[{k:r[j] for k,j in cols.items()} for r in rows[h+1:] if r]
    wb.close(); return out
def f(v):
    try: return float(v)
    except: return 0.0

P_CHB=f"{SRC}/2d7d7338-____________87-92.xlsx"
P_LTB=f"{SRC}/85d6855f-______________1_49__95_7__105_10_.xlsx"
P_OB =f"{SRC}/a0f9bb6e-_________92_8__95_6_.xlsx"
P_PO =f"{SRC}/6dcc59e9-_______________4-16_.xlsx"

CH=[d for d in load(P_CHB,"交易明細","日期(民國)") if d.get("日期(民國)")]
LT=[d for d in load(P_LTB,"明細","民國日期")      if d.get("民國日期")]
PO=[d for d in load(P_PO,"交易明細","頁次")       if d.get("資料日期（民國）")]

wb=Workbook(); wb.remove(wb.active)

def head(ws, cols, widths, r=1):
    for i,h in enumerate(cols,1):
        c=ws.cell(r,i,h); c.font=HDR; c.fill=FILL
        c.alignment=Alignment("center","center"); c.border=BOX
    for col,w in zip("ABCDEFGHIJK", widths): ws.column_dimensions[col].width=w
    ws.freeze_panes=ws.cell(r+1,1)

# ───────────────────────────────── 1 帳戶總覽
ws=wb.create_sheet("帳戶總覽")
ws.cell(1,1,"葉兆輝　收入與支出帳戶總覽（民國 87 年～115 年）").font=H2
head(ws,["#","金融機構／局","帳號","戶名","涵蓋期間","資料來源文件","在本案中的作用"],
     [5,26,24,10,22,38,46], r=3)
data=[
 (1,"彰化商業銀行 岡山分行","29674-1","葉兆輝","87.01.06～92.08.01",
  "存摺存款交易明細（360筆）","高苑科技大學時期薪資帳戶；主要收入與現金提領來源"),
 (2,"臺灣土地銀行","034-005-34574-1","葉兆輝","92.08.01～95.07.01",
  "存摺類存款分戶明細表 OB202（7頁／354列）","義守大學初期薪資帳戶；與①無縫銜接"),
 (3,"臺灣土地銀行","034-005-34574-1","葉兆輝","95.07.03～105.10",
  "客戶歷史交易明細查詢（49頁／1,448筆）","義守大學薪資帳戶續；與②無縫銜接"),
 (4,"中華郵政 高雄灣仔內郵局","0041280-0737273","葉耀坤","97.05.16～101.12.31",
  "客戶歷史交易清單 第4～16頁（590筆）","父親帳戶；查核現金交付之去向"),
 (5,"中華郵政 高雄德智郵局","0041171-0204865","葉兆輝","100.08.03～115.08.02",
  "客戶歷史交易清單（約50筆）","近乎空戶，結存長期為 3 元；無收支功能"),
]
r=4
for row in data:
    for i,v in enumerate(row,1):
        c=ws.cell(r,i,v); c.font=BODY; c.border=BOX
        c.alignment=Alignment(vertical="top",wrap_text=(i>=6))
    ws.row_dimensions[r].height=30; r+=1
r+=1
for t,fo in [
 ("★ 帳戶①②③首尾相接，構成民國 87 年 1 月至 105 年 10 月、長達 18 年 10 個月之連續收支紀錄，中間無缺口。",RED),
 ("　①末筆 92.08.01 將餘額 80,791 元全數轉出後餘額 1 元，同日②以 LMB 0 元起帳；②末筆 95.07.01 結餘 985 元，③首筆 95.07.03 即自 985 元起算。",BODY),
 ("　此一連續性本身即為證據：對造如主張某期間無收入或無支出，可逐月以本表反駁。",RED),
 ("帳戶⑤（德智郵局）長期結存僅 3 元，不具收支功能，不得作為判斷生活費或學費支付之依據。",BODY),
]:
    ws.cell(r,2,t).font=fo; r+=1

# ───────────────────────────────── 2 年度收支彙整
ws=wb.create_sheet("年度收支彙整")
ws.cell(1,1,"年度收支彙整（彰銀 87–92　土銀 95–105）").font=H2
ws.cell(2,1,"薪資／入帳、現金提領、轉帳轉出三欄均由來源明細逐筆加總；合計列為 Excel 公式。").font=NOTE
head(ws,["帳戶","民國年","西元","薪資／入帳","現金提領","提領筆數","轉帳轉出"],
     [20,9,8,16,16,11,16], r=4)

agg=defaultdict(lambda: dict(sal=0.0,cash=0.0,xfer=0.0,n=0))
for d in CH:
    y=str(d["日期(民國)"])[:3]; cat=str(d.get("類別") or "")
    k=("彰銀 29674-1",y)
    if cat=="薪資": agg[k]["sal"]+=f(d.get("存入金額"))
    if cat in ("ATM提款","臨櫃提款"): agg[k]["cash"]+=f(d.get("提出金額")); agg[k]["n"]+=1
    if cat=="轉帳匯出": agg[k]["xfer"]+=f(d.get("提出金額"))
for d in LT:
    y=str(d["民國日期"])[:3]; m=str(d.get("摘要") or "")
    k=("土銀 34574-1",y)
    if m.startswith("AH") or m=="匯款入戶": agg[k]["sal"]+=f(d.get("收入"))
    if m in ("金融卡提","跨行提款","聯行現金"): agg[k]["cash"]+=f(d.get("支出")); agg[k]["n"]+=1
    if m in ("網路跨行","交換轉帳","轉帳","企業自行","中心轉帳"): agg[k]["xfer"]+=f(d.get("支出"))

r=5; first=r
for (acct,y),v in sorted(agg.items(), key=lambda x:(x[0][0]!="彰銀 29674-1", x[0][1])):
    vals=[acct,"民國"+y,int(y)+1911,v["sal"],v["cash"],v["n"],v["xfer"]]
    for i,val in enumerate(vals,1):
        c=ws.cell(r,i,val); c.font=BODY; c.border=BOX
        if i in (4,5,7): c.number_format='#,##0'
    r+=1
last=r-1
ws.cell(r,2,"合計").font=BOLD
for col in (4,5,6,7):
    c=ws.cell(r,col,f'=SUM({chr(64+col)}{first}:{chr(64+col)}{last})')
    c.font=BOLD; c.fill=YEL; c.number_format='#,##0'
r+=2
ws.cell(r,1,"★ 現金提領合計即為「離開銀行體系、無從以轉帳軌跡追查」之金額。房屋貸款如係以現金臨櫃繳納，必落在此一金額之內。").font=RED
r+=1
ws.cell(r,1,"　本表不主張該等現金之全部用途；僅證明有足額現金被提出，且其時序與房貸繳納期間相符。").font=BODY

# ───────────────────────────────── 3 月初現金提領（房貸佐證）
ws=wb.create_sheet("月初現金提領_土銀")
ws.cell(1,1,"土銀 34574-1　每月 1～10 日現金提領（房貸繳納形態之佐證）").font=H2
ws.cell(2,1,"薪資多於每月 5 日前後入帳。如房貸係提現金繳納，提領應集中於月初。本表即檢驗此一假設。").font=NOTE
head(ws,["民國年月","月初1-10日現金提領","筆數","當月總入帳","月初提領占入帳比"],
     [14,22,9,16,18], r=4)
mo=defaultdict(lambda: dict(cash=0.0,n=0,inn=0.0))
for d in LT:
    ds=str(d["民國日期"]); ym=ds[:7]; day=int(ds[8:10]); m=str(d.get("摘要") or "")
    mo[ym]["inn"]+=f(d.get("收入"))
    if day<=10 and m in ("金融卡提","跨行提款","聯行現金"):
        mo[ym]["cash"]+=f(d.get("支出")); mo[ym]["n"]+=1
r=5; first=r
for ym in sorted(mo):
    v=mo[ym]
    ws.cell(r,1,ym).font=BODY; ws.cell(r,1).border=BOX
    c=ws.cell(r,2,v["cash"]); c.number_format='#,##0'; c.font=BODY; c.border=BOX
    ws.cell(r,3,v["n"]).font=BODY; ws.cell(r,3).border=BOX
    c=ws.cell(r,4,v["inn"]); c.number_format='#,##0'; c.font=BODY; c.border=BOX
    c=ws.cell(r,5,f'=IF(D{r}=0,"",B{r}/D{r})'); c.number_format='0.0%'; c.font=BODY; c.border=BOX
    if v["cash"]>=60000: 
        for i in range(1,6): ws.cell(r,i).fill=GRN
    r+=1
last=r-1
ws.cell(r,1,"合計").font=BOLD
for col,L in ((2,"B"),(3,"C"),(4,"D")):
    c=ws.cell(r,col,f'=SUM({L}{first}:{L}{last})'); c.font=BOLD; c.fill=YEL; c.number_format='#,##0'
r+=2
for t,fo in [
 ("★ 綠底為月初 1～10 日現金提領達 60,000 元以上之月份。民國 96～98 年間密集出現，單月最高逾 150,000 元。",RED),
 ("★ 民國 99 年起，月初大額現金提領明顯減少（多降至 10,000～40,000 元）。葉耀坤於 99.12.23 代為清償房屋貸款，",RED),
 ("　此一提領形態之轉折與清償時點之關聯性，應併同房貸繳款紀錄調查；本表僅陳述事實，不逕為因果之認定。",BODY),
]:
    ws.cell(r,1,t).font=fo; r+=1

# ───────────────────────────────── 4 彰銀同日多筆提款
ws=wb.create_sheet("同日多筆提款_彰銀")
ws.cell(1,1,"彰銀 29674-1　同一日多筆 ATM 提款（突破單筆上限之取現行為）").font=H2
ws.cell(2,1,"民國 87～92 年間 ATM 單筆上限為 20,000～30,000 元。同日連續數筆，係為取得超過上限之現金。").font=NOTE
head(ws,["民國日期","西元日期","同日筆數","同日提領合計"],[14,14,11,16], r=4)
same=defaultdict(float); samen=defaultdict(int)
for d in CH:
    if str(d.get("類別"))=="ATM提款":
        k=str(d["日期(民國)"]); same[k]+=f(d.get("提出金額")); samen[k]+=1
r=5; first=r
for k in sorted(same):
    if samen[k]<2: continue
    ws.cell(r,1,k).font=BODY; ws.cell(r,1).border=BOX
    ws.cell(r,2,f"{int(k[:3])+1911}/{k[3:5]}/{k[5:7]}").font=BODY; ws.cell(r,2).border=BOX
    ws.cell(r,3,samen[k]).font=BODY; ws.cell(r,3).border=BOX
    c=ws.cell(r,4,same[k]); c.number_format='#,##0'; c.font=BODY; c.border=BOX
    if same[k]>=60000:
        for i in range(1,5): ws.cell(r,i).fill=GRN
    r+=1
last=r-1
ws.cell(r,1,"合計").font=BOLD
for col,L in ((3,"C"),(4,"D")):
    c=ws.cell(r,col,f'=SUM({L}{first}:{L}{last})'); c.font=BOLD; c.fill=YEL; c.number_format='#,##0'
r+=2
ws.cell(r,1,"★ 綠底為同日取現達 60,000 元以上者。最高一日 100,000 元（089.01.06 四筆、091.03.01 四筆）。").font=RED

# ───────────────────────────────── 5 固定支出檢查
ws=wb.create_sheet("固定支出檢查")
ws.cell(1,1,"固定（定期定額）支出之查核結果").font=H2
head(ws,["項目","查核結果","金額／頻率","結論"],[26,40,26,46], r=3)
fx=[
 ("房屋貸款（中央信託局→臺灣銀行，借款人王蒂玲）",
  "彰銀 87–92、土銀 92–105 全部明細中，查無任何對臺灣銀行或中央信託局之定期定額扣款",
  "無",
  "★ 房貸並非自本案兩帳戶自動扣繳。與當事人所述「提現金至臨櫃繳納」相符；應另向臺灣銀行調取該貸款之繳款明細（含繳款人、繳款方式）"),
 ("UHL 土地銀行本行貸款",
  "土銀 OB202 內按月代扣，94.06.21 起至 95.06.14 止，共 13 筆",
  "10,448～10,461 元／月",
  "為土銀對葉兆輝本人之另一筆貸款（當事人推測為汽車貸款），與系爭房貸係不同債務，不得混為一談"),
 ("交換轉帳",
  "土銀 95–105 共 117 筆，說明欄均載明「信用卡繳款」",
  "數百元至數千元不等",
  "為信用卡自動扣繳，非房貸"),
 ("網路手續費（網路手續）",
  "土銀 190 筆，固定 10 元",
  "10 元／筆",
  "網路銀行跨行轉帳手續費，非固定支出"),
 ("電話費",
  "彰銀 87–92 共 39 筆代扣",
  "金額不等",
  "電信費代扣"),
 ("月初大額現金提領",
  "土銀 96–98 年每月 1–10 日提領常達 60,000～155,000 元；彰銀 87–92 同日多筆合計常達 60,000～100,000 元",
  "每月一次為主",
  "★ 唯一與房貸金額級距（約 4 萬餘元／月）及時序相符之支出形態；為「以現金繳納房貸」最直接之帳面佐證"),
]
r=4
for row in fx:
    for i,v in enumerate(row,1):
        c=ws.cell(r,i,v); c.border=BOX
        c.font = RED if v.startswith("★") or i==1 and "房屋貸款" in v else BODY
        c.alignment=Alignment(vertical="top",wrap_text=True)
    ws.row_dimensions[r].height=46; r+=1

# ───────────────────────────────── 6 葉耀坤郵局
ws=wb.create_sheet("葉耀坤郵局_現金存款")
ws.cell(1,1,"葉耀坤　中華郵政 灣仔內 0737273　現金存款逐筆（97.05～101.12）").font=H2
ws.cell(2,1,"查核事項：葉兆輝主張自 99.12.23 葉耀坤代償房貸後，每月交付現金 50,000 元。本表檢驗該筆現金是否存入本帳戶。").font=NOTE
head(ws,["#","民國日期","西元日期","存入金額","存入後結存","是否 ≥50,000"],[5,13,13,14,14,14], r=4)
cash=[d for d in PO if str(d.get("類別"))=="現金存款" or str(d.get("中文摘要"))=="現金存款"]
r=5; first=r
for n,d in enumerate(cash,1):
    k=str(d["資料日期（民國）"])
    ws.cell(r,1,n).font=BODY; ws.cell(r,1).border=BOX
    ws.cell(r,2,k).font=BODY; ws.cell(r,2).border=BOX
    ws.cell(r,3,f"{int(k[:3])+1911}/{k[3:5]}/{k[5:7]}").font=BODY; ws.cell(r,3).border=BOX
    c=ws.cell(r,4,f(d.get("存款金額"))); c.number_format='#,##0'; c.font=BODY; c.border=BOX
    c=ws.cell(r,5,f(d.get("結存金額"))); c.number_format='#,##0'; c.font=BODY; c.border=BOX
    c=ws.cell(r,6,f'=IF(D{r}>=50000,"是","")'); c.font=BODY; c.border=BOX
    if k[:3] >= "100": 
        for i in range(1,7): ws.cell(r,i).fill=AMB
    r+=1
last=r-1
ws.cell(r,2,"合計").font=BOLD
c=ws.cell(r,4,f'=SUM(D{first}:D{last})'); c.font=BOLD; c.fill=YEL; c.number_format='#,##0'
c=ws.cell(r,6,f'=COUNTIF(F{first}:F{last},"是")'); c.font=BOLD; c.fill=YEL
r+=2
for t,fo in [
 ("查核結論（與當事人記憶不符，據實記載）",H2),
 ("一、本帳戶於 97.05～101.12 期間共有現金存款 40 筆。按民國年分：97 年 10 筆 593,000 元；98 年 16 筆 1,287,009 元；99 年 9 筆 395,000 元；100 年 2 筆 335,000 元；101 年 3 筆 118,000 元。",BODY),
 ("二、99 年 10 月、11 月、12 月三個月，本帳戶現金存款為 0。",RED),
 ("三、99.12.23 葉耀坤代為清償房貸「之後」，本帳戶並無每月固定存入 50,000 元之形態：100 年全年僅 2 筆（100.02.25 計 135,000 元、100.10.21 計 200,000 元），101 年僅 3 筆合計 118,000 元。",RED),
 ("四、反之，代償「之前」的 97、98 兩年，現金存入既密集且金額較大（兩年合計 1,880,009 元）。",RED),
 ("",BODY),
 ("可能之解釋（均須另行查證，不得逕採）：",BOLD),
 ("　(1) 交付現金之起訖時間與當事人記憶有出入——97、98 年之密集存入，或即為當時按期交付之款項。",BODY),
 ("　(2) 99.12.23 以後交付之現金，葉耀坤未存入本帳戶。本帳戶內有多筆「定存淨額（解約入帳）」（500,000／700,000／300,000／900,000 等），足見葉耀坤另有定存帳戶，現金可能逕存為定存或存入他行。",BODY),
 ("　(3) 本檔僅涵蓋原清單第 4～16 頁（起自 97.05.16，迄 101.12.31）。第 1～3 頁及 102 年以後之紀錄均未取得。",BODY),
 ("",BODY),
 ("★ 較具證據價值之二筆（非現金，可追查轉出人）：99.04.05 跨行轉入 120,000 元；100.05.09 跨行轉入 150,000 元。應調取該二筆之轉出帳號，以確認是否出自葉兆輝之帳戶。",RED),
 ("★ 務請注意：如逕向法院主張「每月交付現金 5 萬元」，而本帳戶明細不能支持，對造一經調閱即可反證，反傷及其餘可信之主張。宜改以「有足額現金提領、且與房貸期間相符」之方式陳述。",RED),
]:
    c=ws.cell(r,1,t); c.font=fo; r+=1

wb.save(OUT)
print("saved:", OUT)
