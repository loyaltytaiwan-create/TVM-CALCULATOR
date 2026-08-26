# -*- coding: utf-8 -*-
"""Adds the 7-page OB202 page register, the counterparty-bank decode table,
and the summary code legend to the Land Bank workbook."""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FN = "土銀034-005-34574-1_提領紀錄.xlsx"
F = "Arial"
HDR  = Font(name=F, bold=True, size=10, color="FFFFFF")
FILL = PatternFill("solid", fgColor="1F4E79")
BOLD = Font(name=F, bold=True, size=10)
BODY = Font(name=F, size=10)
RED  = Font(name=F, size=10, color="C00000", bold=True)
NOTE = Font(name=F, size=9, color="595959")
OK   = PatternFill("solid", fgColor="E2EFDA")
TODO = PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="BFBFBF")
BOX  = Border(left=thin, right=thin, top=thin, bottom=thin)

wb = load_workbook(FN)

# ------------------------------------------------------------------ 頁次總表
name = "OB202各頁狀態"
if name in wb.sheetnames: del wb[name]
ws = wb.create_sheet(name, 1)

cols = ["頁次","起訖（民國）","本頁主要內容","逐筆重建進度","尚需處理"]
for i,h in enumerate(cols,1):
    c=ws.cell(1,i,h); c.font=HDR; c.fill=FILL
    c.alignment=Alignment("center","center"); c.border=BOX

rows = [
 ("第1頁","92.08.01～93.05.28",
  "每月薪資／鐘點費入帳（AH**，代收付 7379）後數日內幾乎全額轉出；92.10.15 有 HRT 大額入帳約 599,978 元",
  "✔ 已完成，見「OB202第1頁」工作表", "3 個月區間之數字需以清晰影本覆核（該表已標黃底紅字）"),
 ("第2頁","93.06.01～93.12（末日待確認）",
  "同一收付型態延續；起始 LMB 1,097,6.00＝第1頁末餘額，續接無缺",
  "△ 僅確認起始餘額與收付型態", "全頁金額需清晰影本"),
 ("第3頁","94.01.01～94.05.25",
  "FHEO（跨行匯入）大量出現；NBFF 頻繁；94.02.04、94.03.16 有 HRT 入帳",
  "△ 僅確認日期與摘要", "全頁金額需清晰影本"),
 ("第4頁","94.06.01～94.09.30",
  "UHL 貸款代扣自 94.06.21 起按月出現；備註欄出現「日盛期貨股份有限公司」",
  "△ UHL 已另表收錄", "全頁金額需清晰影本"),
 ("第5頁","94.10.01～95.01.27",
  "UHL 續扣；日盛期貨往來持續；D1 存款利息",
  "△ UHL 已另表收錄", "全頁金額需清晰影本"),
 ("第6頁","95.02.01～95.05.15",
  "UHL 續扣；95.04.04 起出現 CD（金融卡提款）19,000／20,000 等",
  "△ 僅確認日期與摘要", "CD 提領金額需清晰影本後併入提領明細"),
 ("第7頁","95.06.01～95.07.01",
  "CD 提領多筆；UHL 95.06.14 續扣；末列 95.07.01 LMB 985 元。本頁下方有當事人手寫代碼說明",
  "△ 僅確認日期與摘要", "CD 提領金額需清晰影本後併入提領明細"),
]
r=2
for a,b,c_,d,e in rows:
    for i,v in enumerate([a,b,c_,d,e],1):
        cc=ws.cell(r,i,v); cc.border=BOX; cc.font=BODY
        cc.alignment=Alignment(vertical="top", wrap_text=(i>=3))
        cc.fill = OK if d.startswith("✔") else TODO
    ws.row_dimensions[r].height=32
    r+=1

r+=1
for t,f in [
 ("七頁為連續帳頁：每一區塊首列 LMB（上月結餘）均等於前一區塊之末列餘額，第1頁末餘額 10,976 元亦等於第2頁起始 LMB，故本七頁涵蓋民國 92.08.01～95.07.01 且無缺頁。", RED),
 ("⚠ 本七頁係原始點陣式報表之翻拍照片，金額欄字元極細。除第1頁已逐筆勾稽外，其餘各頁之金額尚不足以作為呈庭引用之依據。", RED),
 ("欲完成第2～7頁，需其一：(1) 平台式掃描 300dpi 以上之電子檔；或 (2) 每次僅拍攝一個月份區塊、令該區塊填滿畫面之照片。", NOTE),
]:
    ws.cell(r,1,t).font=f; r+=1

for col,w in zip("ABCDE",[8,22,56,34,40]):
    ws.column_dimensions[col].width=w
ws.freeze_panes="A2"

# ------------------------------------------------------- 摘要代碼與對方行庫
name2 = "OB202代碼與對方行庫"
if name2 in wb.sheetnames: del wb[name2]
ws2 = wb.create_sheet(name2, 2)

def block(ws, r, title, headers, data):
    c=ws.cell(r,1,title); c.font=BOLD; r+=1
    for i,h in enumerate(headers,1):
        cc=ws.cell(r,i,h); cc.font=HDR; cc.fill=FILL
        cc.alignment=Alignment("center","center"); cc.border=BOX
    r+=1
    for row in data:
        for i,v in enumerate(row,1):
            cc=ws.cell(r,i,v); cc.font=BODY; cc.border=BOX
            cc.alignment=Alignment(vertical="top", wrap_text=(i==3))
        r+=1
    return r+1

r = block(ws2, 1, "一、摘要代碼", ["代碼","中文意義","說明／來源"], [
 ("LMB","上月結餘","每一帳頁區塊之首列（當事人提供）"),
 ("CD","金融卡提款","當事人手寫於第7頁下方"),
 ("NBFF","網路轉帳","當事人手寫於第7頁下方"),
 ("NBFG","網路手續費","當事人手寫於第7頁下方"),
 ("UHL","土地銀行本行貸款按月代扣","當事人確認；金額約 10,448～10,461 元／月"),
 ("HRT","大額入帳","當事人於第1頁旁註記「一筆貸款」；性質待查證"),
 ("FHEO","跨行匯入","當事人手寫於第3頁旁「跨行轉入 FHEO」"),
 ("AH**","薪資／鐘點費入帳","代收付代號均為 7379；備註欄載明「專校鐘點費」「進修部鐘點費」及月份"),
 ("D1","存款利息","依金額與計息週期研判"),
 ("PINC／CDFE／CDT／FIDT／DHS","性質未明","待向土地銀行查詢代碼表"),
 ("E04/E05/E13/E17/E23/E25/R10","非交易性註記","無金額、不影響餘額（如補登、掛失、換摺等）"),
])

r = block(ws2, r, "二、備註欄「90＋行庫代號＋帳號」之解讀", ["行庫代號","行庫名稱","本帳戶往來情形"], [
 ("005","臺灣土地銀行","90838-0055928595、90833-0051788349、0117100149486、0108100323167 等"),
 ("007","第一商業銀行","0000014351514421 —— 92～93 年間多次大額轉出"),
 ("008","華南商業銀行","0000965511628208、0000244940579504"),
 ("009","彰化商業銀行","0052772100104200"),
 ("013","國泰世華商業銀行","0000095982000779（當事人於第7頁旁註記「013 國泰世華」）"),
 ("203","（待確認）","0000670605119024"),
 ("822","中國信託商業銀行","0000118530360199 —— 即葉兆輝本人帳戶，與彰銀 29674-1 之轉入帳號 B 同一"),
])

for t,f in [
 ("★ 822-118530360199 之比對結果具重要意義：彰化銀行 29674-1（民國 87～92 年）與土地銀行 034-005-34574-1（民國 92～95 年）兩帳戶，均將款項轉入同一個中國信託帳戶。該帳戶為葉兆輝本人所有，故此類轉出並非對外給付，而係本人資金之集中；主張家庭支出實際由該中信帳戶支應者，應調取該帳戶明細以資證明。", RED),
 ("★ 民國 94～95 年間備註欄多次出現「日盛期貨股份有限公司」，並有 FHEO／NBFF 雙向往來。此為期貨交易之出入金紀錄，對造於剩餘財產分配爭點上可能加以引用，宜先行掌握全貌並準備說明，不宜迴避。", RED),
]:
    ws2.cell(r,1,t).font=f
    ws2.cell(r,1).alignment=Alignment(vertical="top", wrap_text=True)
    ws2.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3)
    ws2.row_dimensions[r].height=58
    r+=1

for col,w in zip("ABC",[28,26,78]):
    ws2.column_dimensions[col].width=w

wb.save(FN)
print("added:", name, "|", name2)
