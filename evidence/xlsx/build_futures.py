# -*- coding: utf-8 -*-
"""期貨出入金彙總 — built from the 轉出入帳號 column plus the client's own
handwritten identification of the 001 and 013 account series."""
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
GRN=Font(name=F,size=10,color="006100",bold=True)
YEL=PatternFill("solid",fgColor="FFFF00"); ORG=PatternFill("solid",fgColor="FCE4D6")
thin=Side(style="thin",color="BFBFBF"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)

wb0=load_workbook(SRC,data_only=True,read_only=True); ws0=wb0['明細']
rows=list(ws0.iter_rows(values_only=True)); h={str(v):j for j,v in enumerate(rows[0]) if v}
def f(v):
    try: return float(v)
    except: return 0.0
D=[]
for r in rows[1:]:
    if not r or not r[h['民國日期']]: continue
    D.append(dict(d=str(r[h['民國日期']]), memo=str(r[h['摘要']] or ''), out=f(r[h['支出']]),
                  inn=f(r[h['收入']]), rem=str(r[h['備註']] or ''), acc=str(r[h['轉出入帳號']] or '')))
wb0.close()
yy=defaultdict(lambda:[0.0,0.0])
for x in D:
    y=x['d'][:3]
    if x['acc'].startswith('013-') and x['memo']!='網路手續': yy[y][0]+=x['out']
    if x['acc'].startswith('001') or any(k in x['rem'] for k in ('元大','日盛')): yy[y][1]+=x['inn']

wb=load_workbook(FN)
name="期貨出入金彙總"
if name in wb.sheetnames: del wb[name]
ws=wb.create_sheet(name, 3)
ws.cell(1,1,"期貨出入金彙總　土銀 034-005-34574-1（民國 95.07～105.10）").font=H2
ws.cell(2,1,"帳號歸類依據：當事人於原件第 1 頁手寫註記「001 是日盛期貨出金」「013 國泰世華」，"
            "並於各該列旁標註「入金」「出金」；另參酌備註欄載明之元大期貨、元大寶來期貨、日盛期貨。").font=NOTE
cols=["民國年","入金（轉出至 013 系列）","出金（自 001 系列或期貨商入帳）","當年淨額"]
for i,t in enumerate(cols,1):
    c=ws.cell(4,i,t); c.font=HDR; c.fill=FILL
    c.alignment=Alignment("center","center",wrap_text=True); c.border=BOX
r=5; first=r
for y in sorted(yy):
    ws.cell(r,1,"民國"+y)
    ws.cell(r,2,yy[y][0] or None).number_format='#,##0'
    ws.cell(r,3,yy[y][1] or None).number_format='#,##0'
    ws.cell(r,4,f'=N(C{r})-N(B{r})').number_format='#,##0;[Red]-#,##0'
    for i in range(1,5):
        c=ws.cell(r,i); c.border=BOX; c.font=BODY
    r+=1
last=r-1
ws.cell(r,1,"合計").font=BOLD
for col,L in ((2,"B"),(3,"C"),(4,"D")):
    c=ws.cell(r,col,f'=SUM({L}{first}:{L}{last})'); c.font=BOLD; c.fill=YEL
    c.number_format='#,##0;[Red]-#,##0'
r+=2
ws.cell(r,1,"帳號明細").font=H2; r+=1
for i,t in enumerate(["帳號","方向","筆數","金額","歸類依據"],1):
    c=ws.cell(r,i,t); c.font=HDR; c.fill=FILL
    c.alignment=Alignment("center","center"); c.border=BOX
r+=1
detail=[]
for a in sorted(set(x['acc'] for x in D if x['acc'].startswith('013-'))):
    s=sum(x['out'] for x in D if x['acc']==a and x['memo']!='網路手續')
    n=sum(1 for x in D if x['acc']==a and x['memo']!='網路手續')
    if s: detail.append((a,"入金（轉出）",n,s,"當事人註記 013 為入金帳號；且九年間全為轉出、從無轉入"))
for a in sorted(set(x['acc'] for x in D if x['acc'].startswith('001'))):
    s=sum(x['inn'] for x in D if x['acc']==a); n=sum(1 for x in D if x['acc']==a and x['inn'])
    if s: detail.append((a,"出金（入帳）",n,s,"當事人註記「001 是日盛期貨出金」；且全為入帳、從無轉出"))
w2=sum(x['inn'] for x in D if any(k in x['rem'] for k in ('元大','日盛')) and not x['acc'].startswith('001'))
detail.append(("（備註載明元大／日盛，無帳號）","出金（入帳）",160,w2,"原件備註欄直接印有期貨商名稱"))
for a,dr,n,amt,why in detail:
    ws.cell(r,1,a); ws.cell(r,2,dr); ws.cell(r,3,n)
    ws.cell(r,4,amt).number_format='#,##0'; ws.cell(r,5,why)
    for i in range(1,6):
        c=ws.cell(r,i); c.border=BOX; c.font=BODY
        c.alignment=Alignment(vertical="top",wrap_text=(i==5))
        if "入金" in dr: c.fill=ORG
    ws.row_dimensions[r].height=30; r+=1
r+=1
for t,fo in [
 ("⚠ 本表為第三次修正，修正經過一併記載",H2),
 ("第一次：發現轉出至 013-0000093340020660 共 20,635,603 元，與期貨入帳高度相關，提示可能為期貨淨損失。",BODY),
 ("第二次：以三項檢驗（期貨入帳之轉出人帳號、7日內配對率僅42%、行庫代號不符）認為支持度不足，"
  "改判較可能為個人帳戶，並指出若屬本人帳戶則款項可能仍存在。",BODY),
 ("第三次（本次）：當事人提供原件第 1 頁，其上有當事人親筆註記「001 是日盛期貨出金」「013 國泰世華」，"
  "並於各列旁標示「入金」「出金」。依此歸類重算，結論回到第一次之方向。",RED),
 ("",BODY),
 ("★ 重算結果：十年間入金 23,269,005 元、出金 17,262,958 元，淨流出 6,006,047 元。"
  "十一個年度中僅民國 101 年為正數（＋509,949 元），其餘各年均為負。",RED),
 ("",BODY),
 ("⚠ 本表之限制，務必一併向律師說明：",BOLD),
 ("　(1) 歸類依據為當事人之手寫註記與資金方向，並非期貨商或銀行出具之證明。",BODY),
 ("　(2) 僅計算流經本土銀帳戶者。若另有以他帳戶入金或出金，金額會變動。",BODY),
 ("　(3) 期貨損益應以期貨商之年度對帳單為準，銀行流水不等於損益。",RED),
 ("★ 應即向元大期貨（原元大寶來期貨）及日盛期貨調取歷年對帳單。此為唯一能確定損益之文件。",RED),
 ("",BODY),
 ("★ 對本案之雙面意義（兩面都要先想好）：",H2),
 ("不利面：於剩餘財產分配爭點，對造可能主張婚姻關係存續期間有大額資金投入高風險交易而減損財產。",RED),
 ("有利面：若資金確係投入期貨而虧損，則該等款項已不存在，可直接反駁「隱匿財產」之指控；"
  "且本帳戶十年間薪資收入均有紀錄、支出去向明確，並無去向不明之資金。",GRN),
]:
    c=ws.cell(r,1,t); c.font=fo; c.alignment=Alignment(vertical="top"); r+=1
for col,w in zip("ABCDE",[30,26,10,16,56]):
    ws.column_dimensions[col].width=w
wb.save(FN); print("added:", name)
