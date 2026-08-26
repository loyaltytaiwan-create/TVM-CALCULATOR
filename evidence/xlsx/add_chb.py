# -*- coding: utf-8 -*-
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
F="Arial"
HDR=Font(name=F,bold=True,size=10,color="FFFFFF"); FILL=PatternFill("solid",fgColor="8A1C1C")
BOLD=Font(name=F,bold=True,size=10); BODY=Font(name=F,size=10)
RED=Font(name=F,size=10,color="C00000",bold=True); NOTE=Font(name=F,size=9,color="595959")
TITLE=Font(name=F,bold=True,size=13,color="8A1C1C")
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

S=[("0870126",42300,1),("0870126",32140,1),("0870302",32140,1),
("0870331",31966,2),("0870430",31706,2),("0870528",31706,2),("0870630",31706,2),
("0870801",31676,3),("0870901",31691,3),("0871003",33358,3),("0871102",33358,3),("0871102",1667,3),("0871126",33394,3),
("0871229",33407,4),("0880129",33407,4),("0880210",42300,4),
("0880302",33315,5),("0880330",33263,5),("0880430",33263,5),("0880529",33263,5),
("0880629",33263,6),("0880730",33263,6),("0880831",33263,6),("0880928",33265,6),
("0881102",33265,7),("0881201",33265,7),("0881227",33265,7),
("0890129",42300,8),("0890129",33746,8),("0890301",33746,8),
("0890327",33676,9),("0890429",33676,9),("0890530",33676,9),("0890629",33676,9),
("0890801",33676,10),("0890902",33145,10),("0891003",33499,10),("0891031",33499,10),
("0891130",33499,11),("0891230",33499,11),("0900119",74424,11),("0900119",33499,11),
("0900302",34899,12),("0900330",33287,12),("0900423",12972,12),("0900502",33596,12),("0900507",8648,12),
("0900531",36366,13),("0900605",8648,13),("0900629",36366,13),("0900709",8648,13),("0900731",36366,13),
("0900807",8648,14),("0900831",52366,14),("0900928",52366,14),
("0901102",52366,15),("0901105",81075,15),("0901203",60988,15),
("0901206",32430,16),("0901211",25866,16),("0910102",60988,16),("0910103",32430,16),
("0910201",60988,17),("0910208",32430,17),("0910215",89171,17),
("0910305",74693,18),("0910329",74693,18),("0910426",15454,18),
("0910502",74453,19),("0910506",25756,19),("0910603",74669,19),("0910606",20605,19),
("0910702",69239,20),("0910708",20605,20),("0910802",69214,20),("0910806",15454,20),
("0910903",75385,21),("0911002",69676,21),("0911031",69676,21),("0911106",32035,21),
("0911202",69556,22),("0911206",16018,22),("0911211",3760,22),("0911231",70592,22),
("0920103",16018,22),("0920129",70592,22),("0920129",8009,22),("0920129",128904,22),
("0920303",70592,23),("0920402",69272,23),("0920417",26321,23),("0920430",67652,23),("0920505",26321,23),
("0920530",68800,24),("0920605",26321,24),("0920701",68800,24),("0920704",26321,24)]

wb=load_workbook("彰銀29674-1_提領紀錄.xlsx")

ws=wb.create_sheet("薪資入帳")
head(ws,["#","記帳日(民國)","西元日期","摘要","入帳金額","頁次"],[5,14,12,10,13,7])
r=2
for i,(d,amt,pg) in enumerate(S,1):
    ws.cell(r,1,i).font=BODY; ws.cell(r,1).alignment=Alignment(horizontal="center")
    ws.cell(r,2,d).font=BODY; ws.cell(r,3,roc2ad(d)).font=BODY; ws.cell(r,4,"薪水").font=BODY
    c=ws.cell(r,5,amt); c.font=BODY; c.number_format='#,##0'
    ws.cell(r,6,pg).font=BODY; ws.cell(r,6).alignment=Alignment(horizontal="center")
    for col in range(1,7): ws.cell(r,col).border=BOX
    r+=1
last=r-1
ws.cell(r,4,"合計").font=BOLD
c=ws.cell(r,5,"=SUM(E2:E%d)"%last); c.font=BOLD; c.number_format='#,##0'; c.fill=YEL
for col in range(1,7): ws.cell(r,col).border=BOX
ws.cell(r+2,1,"註：民國87年1月至92年7月，逐月均有薪資入帳，未見中斷。092年7月31日另有「存匯代號」入帳 80,413 元，交易註記為「高苑」。").font=NOTE
ws.auto_filter.ref="A1:F%d"%last

ws=wb.create_sheet("說明")
for i,w in enumerate([4,34,60],1): ws.column_dimensions[get_column_letter(i)].width=w
r=2
ws.cell(r,2,"彰化商業銀行 岡山分行 帳號 29674-1　交易重建表").font=TITLE; r+=2
for k,v in [("戶　　名","葉兆輝"),("金 融 機 構","彰化商業銀行 岡山分行（分行代號 6513）"),
            ("帳　　號","29674-1（科目 51）"),("開戶日期","民國 084 年 2 月 23 日"),
            ("查詢期間","民國 087/01/01 ～ 092/08/31"),("原始資料","「存摺存款帳號資料及交易明細查詢」全份 25 頁，360 筆"),
            ("核　　發","彰化銀行岡山分行行章，民國 115 年 8 月 6 日"),("本表涵蓋","第 1 ～ 25 頁（全份），無缺頁")]:
    ws.cell(r,2,k).font=BOLD; ws.cell(r,3,v).font=BODY; r+=1
r+=1
ws.cell(r,2,"工作表說明").font=TITLE; r+=1
for k,v in [("提領明細","全部 CD提款／現提，共 166 筆。金額尾數 007 者為跨行提款（手續費 7 元）。提領後餘額低於 1,000 元者以紅字標示。"),
            ("轉帳明細","全部 CD轉帳／電話轉帳，共 28 筆。第 18 頁以後之交易註記欄印有轉入帳號。"),
            ("薪資入帳","民國 087 年 1 月至 092 年 7 月之薪資轉存紀錄（任職高苑科技大學期間）。")]:
    ws.cell(r,2,k).font=BOLD; ws.cell(r,3,v).font=BODY
    ws.cell(r,3).alignment=Alignment(wrap_text=True,vertical="top"); ws.row_dimensions[r].height=32; r+=1
r+=1
ws.cell(r,2,"⚠ 重要提醒").font=Font(name=F,bold=True,size=12,color="C00000"); r+=1
for t in ["轉帳帳號 B（08220000118530360199，中國信託／葉兆輝本人）之 10 筆，係同一人帳戶間之資金移轉，"
          "不具給付性質，不得列入給付統計；對造可能據以主張資金移轉，應預先備妥中信帳戶明細。",
          "轉帳帳號 A（02040001400400011270，高雄第三信用合作社／葉耀坤，葉兆輝之父）之 6 筆，性質待查明。"
          "系爭房屋貸款之借款人為王蒂玲、貸與人為中央信託局，該貸款之還款不會流向葉耀坤帳戶，"
          "故不得將該 6 筆主張為房屋貸款之繳納。",
          "本表係依原始明細照片逐筆判讀重建，金額與餘額均為報表原印數字；"
          "如需呈庭，仍應以銀行核發之原本為準。"]:
    ws.cell(r,2,"•").font=RED; ws.cell(r,3,t).font=BODY
    ws.cell(r,3).alignment=Alignment(wrap_text=True,vertical="top"); ws.row_dimensions[r].height=48; r+=1
wb.save("彰銀29674-1_提領紀錄.xlsx")
print("CHB done: 薪資%d筆"%len(S))
