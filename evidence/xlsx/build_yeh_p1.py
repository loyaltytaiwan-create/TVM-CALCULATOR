# -*- coding: utf-8 -*-
"""Adds 葉耀坤 postal page 1 (96年), recovered today, and the traceable-transfer register."""
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
FN=sys.argv[1]; F="Arial"
HDR=Font(name=F,bold=True,size=10,color="FFFFFF"); FILL=PatternFill("solid",fgColor="1F4E79")
H2=Font(name=F,bold=True,size=11,color="1F4E79")
BOLD=Font(name=F,bold=True,size=10); BODY=Font(name=F,size=10)
RED=Font(name=F,size=10,color="C00000",bold=True); NOTE=Font(name=F,size=9,color="595959")
YEL=PatternFill("solid",fgColor="FFFF00"); GRN=PatternFill("solid",fgColor="E2EFDA")
ORG=PatternFill("solid",fgColor="FCE4D6")
thin=Side(style="thin",color="BFBFBF"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)

# 中文摘要, 提款, 存款 — balances are recomputed in Excel from the opening figure
R=[("利息",0,11125),("利息",0,11125),("現金存款",0,110000),("現金存款",0,270000),
   ("提轉定期",500000,0),("十二敬老",0,3000),("利息",0,940),("一月敬老",0,3000),
   ("現金存款",0,60000),("現金存款",0,20000),("利息",0,940),("二月敬老",0,3000),
   ("建立磁條",0,0),("現金存款",0,20000),("現金存款",0,51000),
   ("跨行匯入",0,1500000),("利息",0,942),("提轉劃撥",1549799,0),("三月敬老",0,3000),
   ("跨行轉入",0,450000),("現金存款",0,140000),("提轉定期",700000,0),("利息",0,952),
   ("四月敬老",0,3000),("入戶匯款",0,54000),("利息",0,1333),("現金存款",0,20000),
   ("利息",0,952),("利息",0,329),("五月敬老",0,3000),("無摺存款",0,25000),
   ("無摺存款",0,29000),("利息",0,1368),("利息",0,999),("現金存款",0,150000),
   ("提轉定期",300000,0),("六月敬老",0,3000),("跨行提款",5006,0),("現金存款",0,54000),
   ("利息",0,1450),("利息",0,1035),("利息",0,621),("提轉存簿",5000,0),
   ("七月敬老",0,3000),("無摺存款",0,29000)]
OPEN=104670   # 115,795 − 11,125
PRINTED=[115795,126920,236920,506920,6920,9920,10860,13860,73860,93860,94800,97800,97800,
         117800,168800,1668800,1669742,119943,122943,572943,712943,12943,13895,16895,70895,
         72228,92228,93180,93509,96509,121509,150509,151877,152876,302876,2876,5876,870,
         54870,56320,57355,57976,52976,55976,84976]

wb=load_workbook(FN)
name="葉耀坤郵局_第1頁_96年"
if name in wb.sheetnames: del wb[name]
ws=wb.create_sheet(name, wb.sheetnames.index("葉耀坤郵局_第2-3頁_96.09-97.05"))
ws.cell(1,1,"葉耀坤　郵局 0737273　第 1 頁逐筆（民國 96 年 1 月～9 月）　※ 本次新補").font=H2
ws.cell(2,1,"原轉錄檔記載「第1～3頁未提供」。第 1 頁已由本次翻拍補回。"
            "末列結存 84,976 元與第 2 頁首列之前的餘額完全銜接，故本帳自 96.01.01 起已無缺頁。").font=NOTE
ws.cell(3,1,"⚠ 原報表之「結存金額」欄與「中文摘要」欄在印表上有一行錯位，故本表之結存係以「期初餘額＋存款－提款」"
            "逐列重算（E 欄），再與原件印出之數字（F 欄）比對；兩者一致即表示判讀無誤。").font=RED
cols=["#","中文摘要","提款金額","存款金額","重算結存","原件印出結存","差異"]
for i,h in enumerate(cols,1):
    c=ws.cell(5,i,h); c.font=HDR; c.fill=FILL
    c.alignment=Alignment("center","center"); c.border=BOX
ws.cell(6,2,"期初餘額（由首列反推）").font=BOLD
c=ws.cell(6,5,OPEN); c.number_format='#,##0'; c.font=BOLD; c.fill=YEL
r=7; first=r
for n,((memo,out,inn),pr) in enumerate(zip(R,PRINTED),1):
    ws.cell(r,1,n); ws.cell(r,2,memo)
    if out: ws.cell(r,3,out).number_format='#,##0'
    if inn: ws.cell(r,4,inn).number_format='#,##0'
    ws.cell(r,5,f'=E{r-1}+N(D{r})-N(C{r})').number_format='#,##0'
    ws.cell(r,6,pr).number_format='#,##0'
    ws.cell(r,7,f'=E{r}-F{r}').number_format='#,##0;[Red]-#,##0'
    for i in range(1,8):
        cc=ws.cell(r,i); cc.border=BOX; cc.font=BODY
    if memo in ("現金存款","無摺存款"):
        for i in range(1,8): ws.cell(r,i).fill=GRN
    if memo in ("跨行匯入","跨行轉入","入戶匯款"):
        for i in range(1,8): ws.cell(r,i).fill=ORG; ws.cell(r,i).font=RED
    r+=1
last=r-1
ws.cell(r,2,"現金類存入（現金存款＋無摺存款）").font=BOLD
c=ws.cell(r,4,f'=SUMIF($B${first}:$B${last},"現金存款",$D${first}:$D${last})'
              f'+SUMIF($B${first}:$B${last},"無摺存款",$D${first}:$D${last})')
c.font=BOLD; c.fill=YEL; c.number_format='#,##0'; r+=1
ws.cell(r,2,"匯入／轉入（可追查對方）").font=BOLD
c=ws.cell(r,4,f'=SUMIF($B${first}:$B${last},"跨行匯入",$D${first}:$D${last})'
              f'+SUMIF($B${first}:$B${last},"跨行轉入",$D${first}:$D${last})'
              f'+SUMIF($B${first}:$B${last},"入戶匯款",$D${first}:$D${last})')
c.font=BOLD; c.fill=YEL; c.number_format='#,##0'; r+=1
ws.cell(r,2,"勾稽不符列數").font=BOLD
c=ws.cell(r,7,f'=COUNTIF($G${first}:$G${last},"<>0")'); c.font=BOLD; c.fill=YEL
r+=2
for t,fo in [
 ("★★ 本頁最重要之發現：摘要欄出現「人名」",H2),
 ("原報表最右側之「摘要」欄，除帳號與序號外，本頁另印有兩個人名：「葉兆雄」與「葉兆輝」。",RED),
 ("郵局清單於匯入款項之摘要欄所載人名，通常即為匯款人。本頁同時有三筆匯入：",BODY),
 ("　跨行匯入 1,500,000 元（經辦局 0100302）／ 跨行轉入 450,000 元（經辦局 808）／ 入戶匯款 54,000 元（經辦局 010103）",RED),
 ("⚠ 但原件之欄位在印表上有行錯位，僅憑翻拍照片無法確定該二人名分別對應哪一筆匯入，"
  "故本表不作認定。此點必須向中華郵政調取該三筆之「匯款單」或「交易憑證」始能確認。",RED),
 ("★ 如其中任一筆之匯款人確為葉兆輝，即為「葉兆輝曾交付款項予葉耀坤」之直接書證，"
  "其證明力遠高於任何現金存款之推論。此為目前全案最值得立即調取之證據。",RED),
 ("★ 另一人名「葉兆雄」顯示葉耀坤同期亦自他人受領大額款項。此對本案為雙面刃："
  "一方面顯示family內有金錢往來之慣例，另一方面對造亦可能主張現金存入係他人所為。應預先掌握。",RED),
 ("摘要欄另出現帳號「8080000 0000288968134706」，與 96.11.02（兩筆）及 99.04.05 之跨行轉入為同一對方帳號。",RED),
]:
    c=ws.cell(r,1,t); c.font=fo; c.alignment=Alignment(vertical="top"); r+=1
for col,w in zip("ABCDEFG",[5,26,14,14,15,16,11]):
    ws.column_dimensions[col].width=w
ws.freeze_panes="A6"
wb.save(FN); print("added:", name, len(R), "rows")
