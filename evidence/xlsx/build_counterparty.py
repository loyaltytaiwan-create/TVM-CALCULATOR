# -*- coding: utf-8 -*-
"""Builds the counterparty sheet from the 備註 / 轉出入帳號 columns of the
Land Bank history, which the earlier analysis never used."""
import sys
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
SRC, FN = sys.argv[1], sys.argv[2]
F="Arial"
HDR=Font(name=F,bold=True,size=10,color="FFFFFF"); FILL=PatternFill("solid",fgColor="1F4E79")
H2=Font(name=F,bold=True,size=11,color="1F4E79")
BOLD=Font(name=F,bold=True,size=10); BODY=Font(name=F,size=10)
RED=Font(name=F,size=10,color="C00000",bold=True); NOTE=Font(name=F,size=9,color="595959")
YEL=PatternFill("solid",fgColor="FFFF00"); GRN=PatternFill("solid",fgColor="E2EFDA")
ORG=PatternFill("solid",fgColor="FCE4D6")
thin=Side(style="thin",color="BFBFBF"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)

wb0=load_workbook(SRC,data_only=True,read_only=True); ws0=wb0['明細']
rows=list(ws0.iter_rows(values_only=True)); h={str(v):j for j,v in enumerate(rows[0]) if v}
def f(v):
    try: return float(v)
    except: return 0.0
D=[]
for r in rows[1:]:
    if not r or not r[h['民國日期']]: continue
    D.append(dict(d=str(r[h['民國日期']]), out=f(r[h['支出']]), inn=f(r[h['收入']]),
                  rem=str(r[h['備註']] or ''), acc=str(r[h['轉出入帳號']] or '')))
wb0.close()

wb=load_workbook(FN)
name="土銀對方身分彙總"
if name in wb.sheetnames: del wb[name]
ws=wb.create_sheet(name, 2)
ws.cell(1,1,"臺灣土地銀行 034-005-34574-1　對方身分彙總（民國 95.07～105.10，1,448 筆）").font=H2
ws.cell(2,1,"資料來源為原明細之「備註」欄（835 筆有值）與「轉出入帳號」欄（527 筆有值）。"
            "此二欄先前未納入分析，本表為首次彙整。").font=NOTE

r=4
ws.cell(r,1,"一、依「備註」彙總（原件載明之對方或用途）").font=H2; r+=1
for i,t in enumerate(["備註（原件載明）","筆數","收入合計","支出合計","性質"],1):
    c=ws.cell(r,i,t); c.font=HDR; c.fill=FILL
    c.alignment=Alignment("center","center"); c.border=BOX
r+=1
agg=defaultdict(lambda:[0,0.0,0.0])
for x in D:
    if x['rem']:
        a=agg[x['rem']]; a[0]+=1; a[1]+=x['inn']; a[2]+=x['out']
def kind(k):
    if "義守" in k: return "薪資（任職機構，直接證明在職與收入）"
    if "期貨" in k: return "★ 期貨出入金"
    if "薪資" in k or "獎金" in k or "鐘點" in k or "獎勵" in k: return "薪資／獎金"
    if "信用卡" in k or "信卡" in k: return "信用卡繳款"
    if k.startswith("700-"): return "★ 匯予葉大中（中華郵政）"
    return ""
rows_out=[(k,v) for k,v in agg.items() if v[1]+v[2]>=20000]
for k,v in sorted(rows_out, key=lambda kv:-(kv[1][1]+kv[1][2])):
    kd=kind(k)
    ws.cell(r,1,k); ws.cell(r,2,v[0])
    ws.cell(r,3,v[1] or None).number_format='#,##0'
    ws.cell(r,4,v[2] or None).number_format='#,##0'
    ws.cell(r,5,kd)
    for i in range(1,6):
        c=ws.cell(r,i); c.border=BOX; c.font=BODY
        if "義守" in k: c.fill=GRN
        if kd.startswith("★"): c.fill=ORG; c.font=RED
    r+=1
r+=1

ws.cell(r,1,"二、依「轉出入帳號」彙總（前 12 大）").font=H2; r+=1
for i,t in enumerate(["轉出入帳號","筆數","收入合計","支出合計","說明"],1):
    c=ws.cell(r,i,t); c.font=HDR; c.fill=FILL
    c.alignment=Alignment("center","center"); c.border=BOX
r+=1
agg2=defaultdict(lambda:[0,0.0,0.0])
for x in D:
    if x['acc']:
        a=agg2[x['acc']]; a[0]+=1; a[1]+=x['inn']; a[2]+=x['out']
NOTE_ACC={
 "013-0000093340020660":"★★ 本帳戶最大資金流出對象。該等列之備註欄空白，身分未經確認，必須查明",
 "700-0001010210272059":"葉大中 中華郵政帳戶（16 筆匯款 396,500 元＋16 筆手續費 160 元）",
 "013-0000095882000779":"國泰世華（013）",
 "005-0000034051036133":"土地銀行（005）—— 收入方",
 "806-0000367210718820":"元大銀行（806）—— 單筆 550,000 元收入",
}
for k,v in sorted(agg2.items(), key=lambda kv:-(kv[1][1]+kv[1][2]))[:12]:
    ws.cell(r,1,k); ws.cell(r,2,v[0])
    ws.cell(r,3,v[1] or None).number_format='#,##0'
    ws.cell(r,4,v[2] or None).number_format='#,##0'
    ws.cell(r,5,NOTE_ACC.get(k,""))
    for i in range(1,6):
        c=ws.cell(r,i); c.border=BOX; c.font=BODY
        if k.startswith("013-00000933"): c.fill=ORG; c.font=RED
        if k.startswith("700-"): c.fill=GRN
    r+=1
r+=1

ws.cell(r,1,"三、★★ 期貨入金 與 轉出至 013-0000093340020660 之逐年對照").font=H2; r+=1
for i,t in enumerate(["民國年","期貨入金（備註載明元大／日盛）","轉出至 013-…0660","差額"],1):
    c=ws.cell(r,i,t); c.font=HDR; c.fill=FILL
    c.alignment=Alignment("center","center"); c.border=BOX
r+=1
yy=defaultdict(lambda:[0.0,0.0])
for x in D:
    y=x['d'][:3]
    if any(k in x['rem'] for k in ("元大","日盛","期貨")): yy[y][0]+=x['inn']
    if x['acc']=="013-0000093340020660": yy[y][1]+=x['out']
first=r
for y in sorted(yy):
    ws.cell(r,1,"民國"+y)
    ws.cell(r,2,yy[y][0] or None).number_format='#,##0'
    ws.cell(r,3,yy[y][1] or None).number_format='#,##0'
    ws.cell(r,4,f'=N(B{r})-N(C{r})').number_format='#,##0;[Red]-#,##0'
    for i in range(1,5):
        c=ws.cell(r,i); c.border=BOX; c.font=BODY
    r+=1
last=r-1
ws.cell(r,1,"合計").font=BOLD
for col,L in ((2,"B"),(3,"C"),(4,"D")):
    c=ws.cell(r,col,f'=SUM({L}{first}:{L}{last})'); c.font=BOLD; c.fill=YEL; c.number_format='#,##0;[Red]-#,##0'
r+=2

for t,fo in [
 ("★ 有利之發現",H2),
 ("備註欄直接載明「義守大學」之入帳 133 筆、合計 11,671,555 元（民國95.07～105.10），"
  "另有逐月「10410～10509 月份薪資」及「104 年終獎金」「論文獎勵金」「鐘點費」等明細。"
  "此為任職與收入之直接書證，無須再以推論證明。",RED),
 ("備註欄亦直接載明匯予葉大中郵局帳戶 700-0001010210272059 之 16 筆、396,500 元，與另表完全相符。",RED),
 ("",BODY),
 ("★★ 必須立即處理之風險",H2),
 ("帳號 013-0000093340020660 為本帳戶最大之資金流出對象：245 筆、合計 20,635,603 元。"
  "該等列之備註欄空白，故其身分尚未確認。",RED),
 ("同期備註載明「元大期貨」「元大寶來期貨」「日盛期貨」之入金為 160 筆、合計 14,133,502 元，"
  "且其逐年消長與上開轉出高度同步（見上表）。兩者差額為 6,502,101 元。",RED),
 ("⚠ 本表不認定該差額之性質。必須先查明 013-0000093340020660 究係：",RED),
 ("　(1) 期貨商之保證金專戶 —— 若是，則差額可能為期貨交易之淨損失；或",BODY),
 ("　(2) 葉兆輝本人於國泰世華之帳戶 —— 若是，則屬同一人資金移轉，款項可能仍存在該帳戶。",BODY),
 ("　二者法律效果天差地遠。請立即調取該帳戶之開戶資料或期貨商對帳單。",RED),
 ("⚠ 於離婚剩餘財產分配爭點上，對造如先行查得此一資金流向，將極為不利。"
  "宜由己方先行查明、主動說明，不可迴避，亦不可在未查明前自行陳述其性質。",RED),
]:
    c=ws.cell(r,1,t); c.font=fo; c.alignment=Alignment(vertical="top"); r+=1

for col,w in zip("ABCDE",[30,9,16,16,56]):
    ws.column_dimensions[col].width=w
wb.save(FN); print("added:", name)
