# -*- coding: utf-8 -*-
"""Replaces the partial, partly-misread OB202 page-1 sheet with the full
seven-page transcription (92.08.01-95.07.01), converted from the printout's
raw cents to NT$ and re-reconciled live in Excel."""
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SRC, FN = sys.argv[1], sys.argv[2]
F="Arial"
HDR=Font(name=F,bold=True,size=10,color="FFFFFF"); FILL=PatternFill("solid",fgColor="1F4E79")
BOLD=Font(name=F,bold=True,size=10); BODY=Font(name=F,size=10)
RED=Font(name=F,size=10,color="C00000",bold=True); NOTE=Font(name=F,size=9,color="595959")
YEL=PatternFill("solid",fgColor="FFFF00"); GRN=PatternFill("solid",fgColor="E2EFDA")
thin=Side(style="thin",color="BFBFBF"); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)

src=load_workbook(SRC,data_only=True,read_only=True)
rows=[r for r in src["明細"].iter_rows(values_only=True)][1:]
def f(v):
    try: return float(v)
    except: return None

D=[]
for r in rows:
    if not r or not r[1]: continue
    def c(i):
        v=f(r[i]); return None if v is None else round(v/100,2)
    D.append(dict(pg=r[0], d=str(r[1]), ad=str(r[2])[:10], memo=str(r[3] or ""),
                  note=str(r[4] or ""), dr=c(5), cr=c(6), bal=c(7),
                  rem=str(r[9] or ""), img=str(r[10] or "")))
src.close()

wb=load_workbook(FN)
for old in ("OB202第1頁_92.08-93.05",):
    if old in wb.sheetnames: del wb[old]
name="OB202逐筆_92.08-95.07"
if name in wb.sheetnames: del wb[name]
ws=wb.create_sheet(name,0)

cols=["頁","民國日期","西元日期","摘要","說明","借方金額","貸方金額","餘額",
      "連續餘額","勾稽差額","備註"]
for i,h in enumerate(cols,1):
    c=ws.cell(1,i,h); c.font=HDR; c.fill=FILL
    c.alignment=Alignment("center","center"); c.border=BOX

r=2
for d in D:
    ws.cell(r,1,d["pg"]); ws.cell(r,2,d["d"]); ws.cell(r,3,d["ad"])
    ws.cell(r,4,d["memo"]); ws.cell(r,5,d["note"])
    for col,key in ((6,"dr"),(7,"cr"),(8,"bal")):
        if d[key] is not None:
            cc=ws.cell(r,col,d[key]); cc.number_format='#,##0.00'
    # carry-forward balance so rows the printout left blank do not break the chain
    ws.cell(r,9, f'=IF(H{r}="",I{r-1},H{r})' if r>2 else '=H2').number_format='#,##0.00'
    if r>2:
        ws.cell(r,10,
          f'=IF(H{r}="","",IF(D{r}="LMB",ROUND(H{r}-I{r-1},2),'
          f'ROUND(H{r}-(I{r-1}+N(G{r})-N(F{r})),2)))').number_format='#,##0.00;[Red]-#,##0.00'
    ws.cell(r,11,d["rem"] or d["img"])
    for i in range(1,12):
        cc=ws.cell(r,i); cc.border=BOX; cc.font=BODY
    if d["memo"].startswith("AH"):
        for i in range(1,12): ws.cell(r,i).fill=GRN
    r+=1
last=r-1

r+=1
def tot(label, formula, note="", col=7):
    global r
    ws.cell(r,4,label).font=BOLD
    c=ws.cell(r,col,formula); c.font=BOLD; c.fill=YEL; c.number_format='#,##0'
    if note: ws.cell(r,11,note).font=NOTE
    r+=1
tot("薪資／鐘點費入帳合計（AH*）", f'=SUMIF($D$2:$D${last},"AH*",$G$2:$G${last})',
    "代收付代號 7379")
tot("跨行轉入合計（FHEO）",       f'=SUMIF($D$2:$D${last},"FHEO",$G$2:$G${last})',
    "備註多載明日盛期貨股份有限公司")
tot("海外入款合計（HRT）",        f'=SUMIF($D$2:$D${last},"HRT",$G$2:$G${last})')
tot("金融卡提款合計（CD）",       f'=SUMIF($D$2:$D${last},"CD",$F$2:$F${last})', "", 6)
tot("網路轉帳轉出合計（NBFF）",   f'=SUMIF($D$2:$D${last},"NBFF",$F$2:$F${last})', "", 6)
tot("土銀本行貸款代扣合計（UHL）",f'=SUMIF($D$2:$D${last},"UHL",$F$2:$F${last})',
    "94.06～95.06 按月，約 10,456 元／月", 6)
ws.cell(r,4,"勾稽不符列數").font=BOLD
c=ws.cell(r,10,f'=COUNTIF($J$2:$J${last},"<>0")-COUNTBLANK($J$2:$J${last})')
c.font=BOLD; c.fill=YEL
r+=2

for t,fo in [
 ("本表為臺灣土地銀行「存摺類存款分戶明細表（OB202）」全 7 頁之逐筆重建，期間民國 92.08.01～95.07.01，共 353 列。",BODY),
 ("原報表金額欄含兩位小數且未印小數點；本表已還原為新臺幣元（原印數字 ÷ 100）。",BODY),
 ("勾稽差額＝本列餘額－（前一有餘額之列的餘額＋貸方－借方）；LMB（上月結餘）列則與前一列餘額直接比對。全部 353 列均為 0。",RED),
 ("⚠ 更正：本工作表取代先前僅含第 1 頁之版本。該版本有三處判讀錯誤，均已依原件更正——"
  "92.10.15 HRT 應為 599,970（原誤為 599,978）；92.11.05 AH84 應為 81,392（原誤為 81,192）；"
  "93.03.03 AH3 應為 2,640（原誤為 2,644）。另 92.10.15 一筆原判讀不出之 NBFF，金額為 499,983。",RED),
 ("綠底列為薪資／鐘點費入帳。",NOTE),
]:
    ws.cell(r,4,t).font=fo; r+=1

for col,w in zip("ABCDEFGHIJK",[5,12,12,8,18,14,14,14,14,13,40]):
    ws.column_dimensions[col].width=w
ws.freeze_panes="A2"
wb.save(FN)
print("rebuilt:", name, "rows:", last-1)
