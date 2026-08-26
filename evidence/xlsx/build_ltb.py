# -*- coding: utf-8 -*-
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
F="Arial"
HDR=Font(name=F,bold=True,size=10,color="FFFFFF"); FILL=PatternFill("solid",fgColor="1F4E79")
BOLD=Font(name=F,bold=True,size=10); BODY=Font(name=F,size=10)
RED=Font(name=F,size=10,color="C00000",bold=True); NOTE=Font(name=F,size=9,color="595959")
TITLE=Font(name=F,bold=True,size=13,color="1F4E79")
TH=Side(style="thin",color="BFBFBF"); BOX=Border(left=TH,right=TH,top=TH,bottom=TH)
YEL=PatternFill("solid",fgColor="FFF2CC")
def head(ws,cols,widths):
    for i,(c,w) in enumerate(zip(cols,widths),1):
        cell=ws.cell(1,i,c); cell.font,cell.fill,cell.border=HDR,FILL,BOX
        cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A2"
def roc2ad(d):
    d=str(d); return "%d/%02d/%02d"%(int(d[:3])+1911,int(d[3:5]),int(d[5:7]))

# 提領（客戶歷史交易明細查詢，第32-48頁）
W=[("1010712","金融卡提",10000,8809,32),("1010805","跨行提款",5000,93274,32),
("1010904","金融卡提",20000,33460,32),("1010921","金融卡提",1000,1641,32),
("1010921","金融卡提",1000,641,32),("1010926","金融卡提",10000,641,32),
("1011012","金融卡提",5000,7629,33),("1011012","金融卡提",5000,2629,33),
("1011012","跨行提款",2000,629,33),("1011016","金融卡提",10000,90629,33),
("1011021","金融卡提",10000,33119,33),("1011024","金融卡提",10000,9990,33),
("1011029","金融卡提",10000,9990,33),("1011102","金融卡提",5000,3938,33),
("1011105","金融卡提",10000,9441,34),
("1020104","金融卡提",60000,70824,35),("1020104","金融卡提",60000,10824,35),
("1020106","跨行提款",5000,5824,35),("1020112","金融卡提",10000,9948,35),
("1020112","金融卡提",9900,48,35),
("1020212","跨行提款",20000,125501,36),("1020318","金融卡提",20000,41421,36),
("1020501","金融卡提",20000,118423,36),
("1020610","金融卡提",20000,32503,37),("1020610","金融卡提",1000,31503,37),
("1020705","金融卡提",20000,102086,37),("1020705","金融卡提",20000,82086,37),
("1020718","金融卡提",20000,17076,37),("1020817","金融卡提",10000,6987,37),
("1041119","金融卡提",1000,5379,45),("1041119","金融卡提",5000,379,45),
("1050520","跨行提款",20000,433981,47),("1050630","金融卡提",20000,55052,47),
("1050717","跨行提款",5000,3031,48),("1050721","金融卡提",10000,7159,48),
("1050808","金融卡提",60000,14331,48),("1050817","跨行提款",20000,25911,48)]

# UHL 貸款代扣（OB202）
U=[("0940621",10456),("0940711",10456),("0940816",10461),("0940916",10460),
("0941013",10456),("0941110",10448),("0941215",10460),("0950116",10456),
("0950218",10456),("0950310",10450),("0950410",10450),("0950510",10450),("0950614",10450)]

# 匯款予葉大中
Y=[("1010918","00:47:27",3000,32,"區間內"),("1011005","14:17:03",3000,33,"區間內"),
("1011020","22:36:42",37500,33,"區間內"),("1011102","22:04:24",3000,34,"區間內"),
("1011202","19:48:04",3000,34,"區間內"),("1011228","13:22:09",3000,34,"區間內"),
("1020204","08:55:21",3000,35,"區間內"),("1020305","18:34:15",20000,36,"區間外"),
("1020802","13:57:03",24000,37,"區間外"),("1021209","14:58:55",20000,38,"區間外"),
("1030106","08:59:00",15000,39,"區間外"),("1030607","11:24:03",50000,40,"區間外"),
("1040224","12:25:11",50000,43,"區間外"),("1040611","06:46:14",50000,44,"區間外"),
("1040831","20:38:26",50000,45,"區間外"),("1050714","14:31:57",62000,48,"區間外")]

wb=Workbook(); wb.remove(wb.active)

ws=wb.create_sheet("提領明細")
head(ws,["#","記帳日(民國)","西元日期","摘要","提領金額","提領後餘額","頁次"],[5,14,12,12,13,14,7])
r=2
for i,(d,memo,amt,bal,pg) in enumerate(W,1):
    ws.cell(r,1,i).font=BODY; ws.cell(r,1).alignment=Alignment(horizontal="center")
    ws.cell(r,2,d).font=BODY; ws.cell(r,3,roc2ad(d)).font=BODY; ws.cell(r,4,memo).font=BODY
    c=ws.cell(r,5,amt); c.font=BODY; c.number_format='#,##0'
    c=ws.cell(r,6,bal); c.font=BODY; c.number_format='#,##0'
    ws.cell(r,7,pg).font=BODY; ws.cell(r,7).alignment=Alignment(horizontal="center")
    for col in range(1,8): ws.cell(r,col).border=BOX
    r+=1
last=r-1
ws.cell(r,4,"合計").font=BOLD
c=ws.cell(r,5,"=SUM(E2:E%d)"%last); c.font=BOLD; c.number_format='#,##0'; c.fill=YEL
ws.cell(r,6,"筆數").font=BOLD
c=ws.cell(r,7,"=COUNT(E2:E%d)"%last); c.font=BOLD; c.fill=YEL
for col in range(1,8): ws.cell(r,col).border=BOX
ws.cell(r+2,1,"⚠ 本表僅涵蓋現持有之第 32～48 頁（民國101.07～105.09）。第 38～44 頁之提領明細尚未逐筆判讀，第 1～31 頁尚未取得。").font=RED
ws.auto_filter.ref="A1:G%d"%last

ws=wb.create_sheet("UHL貸款代扣")
head(ws,["#","記帳日(民國)","西元日期","摘要","扣款金額"],[5,14,12,12,13])
r=2
for i,(d,amt) in enumerate(U,1):
    ws.cell(r,1,i).font=BODY; ws.cell(r,1).alignment=Alignment(horizontal="center")
    ws.cell(r,2,d).font=BODY; ws.cell(r,3,roc2ad(d)).font=BODY; ws.cell(r,4,"UHL").font=BODY
    c=ws.cell(r,5,amt); c.font=BODY; c.number_format='#,##0'
    for col in range(1,6): ws.cell(r,col).border=BOX
    r+=1
last=r-1
ws.cell(r,4,"合計").font=BOLD
c=ws.cell(r,5,"=SUM(E2:E%d)"%last); c.font=BOLD; c.number_format='#,##0'; c.fill=YEL
for col in range(1,6): ws.cell(r,col).border=BOX
r+=2
ws.cell(r,1,"UHL＝臺灣土地銀行自本帳戶定期定額代扣之貸款（本行貸款）。起於民國94年6月，按月一次，金額 10,448～10,461 元。").font=BODY; r+=1
ws.cell(r,1,"⚠ 此為土地銀行對葉兆輝本人之另一筆貸款，與系爭房屋貸款（貸與人中央信託局、借款人王蒂玲）為不同債務，不得混為一談。").font=RED

ws=wb.create_sheet("匯款予葉大中")
head(ws,["#","記帳日(民國)","西元日期","時間","摘要","匯款金額","頁次","是否落在對造所提三頁區間"],[5,14,12,11,11,13,7,24])
r=2
for i,(d,t,amt,pg,rng) in enumerate(Y,1):
    ws.cell(r,1,i).font=BODY; ws.cell(r,1).alignment=Alignment(horizontal="center")
    ws.cell(r,2,d).font=BODY; ws.cell(r,3,roc2ad(d)).font=BODY; ws.cell(r,4,t).font=BODY
    ws.cell(r,5,"網路跨行").font=BODY
    c=ws.cell(r,6,amt); c.font=BODY; c.number_format='#,##0'
    ws.cell(r,7,pg).font=BODY; ws.cell(r,7).alignment=Alignment(horizontal="center")
    cc=ws.cell(r,8,rng); cc.font=BOLD if rng=="區間內" else BODY
    cc.alignment=Alignment(horizontal="center")
    for col in range(1,9): ws.cell(r,col).border=BOX
    r+=1
last=r-1
ws.cell(r,5,"合計").font=BOLD
c=ws.cell(r,6,"=SUM(F2:F%d)"%last); c.font=BOLD; c.number_format='#,##0'; c.fill=YEL
for col in range(1,9): ws.cell(r,col).border=BOX
r+=2
ws.cell(r,1,"區間內小計").font=BOLD
c=ws.cell(r,2,'=SUMIF(H2:H%d,"區間內",F2:F%d)'%(last,last)); c.font=BOLD; c.number_format='#,##0'; c.fill=YEL; r+=1
ws.cell(r,1,"區間外小計").font=BOLD
c=ws.cell(r,2,'=SUMIF(H2:H%d,"區間外",F2:F%d)'%(last,last)); c.font=BOLD; c.number_format='#,##0'; c.fill=YEL; r+=2
ws.cell(r,1,"受款帳號：700-0001010210272059（中華郵政　局號0101021／帳號0272059　戶名：葉大中）").font=BOLD; r+=1
ws.cell(r,1,"對造所提卷內第00044-00046頁三頁存摺，區間為民國101.09.04～102.02.06。").font=BODY

ws=wb.create_sheet("說明")
for i,w in enumerate([4,34,62],1): ws.column_dimensions[get_column_letter(i)].width=w
r=2
ws.cell(r,2,"臺灣土地銀行 帳號 034-005-34574-1　交易重建表").font=TITLE; r+=2
for k,v in [("戶　　名","葉兆輝"),("金 融 機 構","臺灣土地銀行"),("帳　　號","034-005-34574-1（併寫 034005345741）"),
            ("資料來源①","「客戶歷史交易明細查詢」　查詢期間 095/07/01～105/11/01　全份 49 頁　製表日 115/07/20"),
            ("　現　持　有","第 32～48 頁（民國 101.07 ～ 105.09）"),
            ("資料來源②","「存摺類存款分戶明細表（OB202）」　民國 92.06 ～ 95.07　全份 7 頁（已完整）")]:
    ws.cell(r,2,k).font=BOLD; ws.cell(r,3,v).font=BODY
    ws.cell(r,3).alignment=Alignment(wrap_text=True,vertical="top"); r+=1
r+=1
ws.cell(r,2,"工作表說明").font=TITLE; r+=1
for k,v in [("提領明細","現持有頁次中已逐筆判讀之金融卡提款／跨行提款。涵蓋範圍見該表下方紅字提醒。"),
            ("UHL貸款代扣","土地銀行本行貸款之按月代扣，13 筆（民國 94.06 ～ 95.06），資料來源為 OB202 全 7 頁。"),
            ("匯款予葉大中","自本帳戶經網路銀行跨行轉帳匯入葉大中中華郵政帳戶之全部紀錄，16 筆。")]:
    ws.cell(r,2,k).font=BOLD; ws.cell(r,3,v).font=BODY
    ws.cell(r,3).alignment=Alignment(wrap_text=True,vertical="top"); ws.row_dimensions[r].height=32; r+=1
r+=1
ws.cell(r,2,"⚠ 尚待補齊").font=Font(name=F,bold=True,size=12,color="C00000"); r+=1
for t in ["「客戶歷史交易明細查詢」第 20～30 頁前後 —— 涵蓋民國 099 年 12 月。"
          "民國 099.12.23 葉耀坤代為清償王蒂玲名下房屋貸款，需查明該日前後本帳戶有無大額入帳或支出，"
          "以確定實際清償人。申請時應指明文件名稱為「客戶歷史交易明細查詢」，而非 OB202。",
          "第 38～44 頁之提領明細尚未逐筆判讀（該等頁面之匯款與薪資已判讀完畢）。",
          "本表係依原始明細照片逐筆判讀重建；如需呈庭，仍應以銀行核發之原本為準。"]:
    ws.cell(r,2,"•").font=RED; ws.cell(r,3,t).font=BODY
    ws.cell(r,3).alignment=Alignment(wrap_text=True,vertical="top"); ws.row_dimensions[r].height=52; r+=1
wb.save("土銀034-005-34574-1_提領紀錄.xlsx")
print("LTB done")
