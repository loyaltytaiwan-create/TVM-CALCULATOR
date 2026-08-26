# -*- coding: utf-8 -*-
"""
Adds the row-by-row reconstruction of OB202 page 1 (民國92.08.01–93.05.28)
to the Land Bank workbook. Amounts are read off the original printout; the
勾稽差額 column recomputes the balance chain in Excel so any misread digit
shows up as a non-zero number instead of hiding.
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FN = "土銀034-005-34574-1_提領紀錄.xlsx"
F = "Arial"
HDR  = Font(name=F, bold=True, size=10, color="FFFFFF")
FILL = PatternFill("solid", fgColor="1F4E79")
BOLD = Font(name=F, bold=True, size=10)
BODY = Font(name=F, size=10)
RED  = Font(name=F, size=10, color="C00000", bold=True)
NOTE = Font(name=F, size=9, color="595959")
AMB  = PatternFill("solid", fgColor="FFF2CC")   # rows whose digits are not certain
YEL  = PatternFill("solid", fgColor="FFFF00")
thin = Side(style="thin", color="BFBFBF")
BOX  = Border(left=thin, right=thin, top=thin, bottom=thin)

def ad(roc):
    """民國 YYMMDD -> 西元 YYYY/MM/DD"""
    s = str(roc)
    return "%d/%s/%s" % (int(s[:3]) + 1911, s[3:5], s[5:7])

# 記帳日, 摘要, 借方, 貸方, 餘額, 對方行庫/代收付, 備註, 判讀確定?
# 借/貸/餘 = None 表示原表該欄空白
R = [
 # ---- 民國 92 年 8 月 ----
 ("0920801","LMB",  None, None, 0,      "",  "上月結餘（本頁起點）", True),
 ("0920828","E05",  None, None, 0,      "",  "原表未印金額與餘額，屬非交易性註記；餘額沿用前列以利勾稽", True),
 ("0920828","E25",  None, None, 0,      "",  "同上", True),
 # ---- 92 年 9 月 ----
 ("0920901","LMB",  None, None, 0,        "", "上月結餘", True),
 ("0920905","AH72", None, 71514, 71514,   "代收付 7379", "薪資／鐘點費入帳", True),
 ("0920905","AH72", None, 91227, 162741,  "代收付 7379", "薪資／鐘點費入帳", True),
 ("0920905","PINC", None, 0,     162741,  "8510505128", "金額 0，性質待查", True),
 ("0920916","NBFF", 160000, None, 2741,   "902-00599204-0055403201", "網路轉帳轉出", True),
 # ---- 92 年 10 月 ----
 ("0921001","LMB",  None, None, 2741,     "", "上月結餘", True),
 ("0921001","AH84", None, 77282, 80023,   "代收付 7379", "薪資／鐘點費入帳", True),
 ("0921001","R10",  None, None, 80023,    "", "非交易性註記（備註欄印 5.00）；餘額沿用前列", True),
 ("0921001","E17",  None, None, 80023,    "", "非交易性註記；餘額沿用前列", True),
 ("0921011","NBFF", 80000, None, 23,      "908007 第一銀行 0000014351514421", "網路跨行轉出", True),
 ("0921015","HRT",  None, 599978, 600001, "", "★大額入帳。原表餘額欄數字判讀不清，此處為依前後列推算值", False),
 ("0921015","NBFF", 100000, None, 500001, "908822 中國信託 0000118530360199", "轉入本人中國信託帳戶；餘額為推算值", False),
 ("0921015","NBFF", None, None, None,     "", "原表尚有一列，借方與餘額均判讀不清（餘額疑為 459,991）", False),
 # ---- 92 年 11 月 ----
 ("0921101","LMB",  None, None, 210,      "", "上月結餘。原表數字判讀不清，此處為依下列推算值", False),
 ("0921105","AH84", None, 81192, 81402,   "代收付 7379", "薪資／鐘點費入帳", False),
 ("0921111","NBFF", 70000, None, 11402,   "908007 第一銀行 0000014351514421", "網路跨行轉出", True),
 ("0921111","NBFF", 11400, None, 2,       "908822 中國信託 0000118530360199", "轉入本人中國信託帳戶", True),
 # ---- 92 年 12 月 ----
 ("0921201","LMB",  None, None, 2,        "", "上月結餘", True),
 ("0921205","AH79", None, 77282, 77284,   "代收付 7379", "薪資／鐘點費入帳", True),
 ("0921208","NBFF", 18,   None, 77266,    "908005 土地銀行 90836-0055928648", "跨行轉帳手續費 18 元", True),
 ("0921208","NBFF", 77265, None, 1,       "908005 土地銀行 90838-0055928595", "網路轉帳轉出", True),
 ("0921221","D1",   None, 49,    50,      "", "存款利息", True),
 # ---- 93 年 1 月 ----
 ("0930101","LMB",  None, None, 50,       "", "上月結餘", True),
 ("0930105","AH77", None, 77282, 77332,   "代收付 7379", "薪資／鐘點費入帳", True),
 ("0930107","NBFF", 77330, None, 2,       "908005 土地銀行 90838-0051032437", "網路轉帳轉出", True),
 ("0930107","AH77", None, 44244, 44246,   "代收付 7379", "薪資／鐘點費入帳", True),
 ("0930112","NBFF", 44000, None, 246,     "908005 土地銀行 90833-0051788349", "網路轉帳轉出", True),
 ("0930110","AH2",  None, 6900,  7146,    "", "備註欄印「036 93.1 月份專校鐘點費」；原表列序即如此", True),
 # ---- 93 年 2 月 ----
 ("0930201","LMB",  None, None, 7146,     "", "上月結餘", True),
 ("0930205","AH74", None, 77282, 84428,   "代收付 7379", "薪資／鐘點費入帳", True),
 ("0930206","NBFF", 84400, None, 28,      "908005 土地銀行 90838-0052421106", "網路轉帳轉出", True),
 # ---- 93 年 3 月 ----
 ("0930301","LMB",  None, None, 28,       "", "上月結餘", False),
 ("0930303","AH3",  None, 2644,  2664,    "", "備註欄印「036 補發 93.1 月鐘點費」；金額與餘額差 20 元，待確認", False),
 ("0930305","AH87", None, 77282, 79950,   "代收付 7379", "薪資／鐘點費入帳；餘額判讀待確認", False),
 ("0930308","NBFF", 79950, None, 0,       "908007 第一銀行 0000014351514421", "網路跨行轉出；金額判讀待確認", False),
 ("0930330","AH5",  None, 10960, 10960,   "", "備註欄印「036 93.3 月進修部鐘點費」", True),
 # ---- 93 年 4 月 ----
 ("0930401","LMB",  None, None, 10960,    "", "上月結餘", True),
 ("0930405","AH81", None, 80570, 91530,   "代收付 7379", "薪資／鐘點費入帳", True),
 ("0930405","NBFF", 91500, None, 30,      "908007 第一銀行 0000014351514421", "網路跨行轉出", True),
 ("0930410","AH6",  None, 10960, 10990,   "", "備註欄印「036 93.4 月進修部鐘點費」", True),
 # ---- 93 年 5 月 ----
 ("0930501","LMB",  None, None, 10990,    "", "上月結餘", True),
 ("0930505","AH79", None, 78926, 89916,   "代收付 7379", "薪資／鐘點費入帳", True),
 ("0930508","NBFF", 89900, None, 16,      "908822 中國信託 0000118530360199", "轉入本人中國信託帳戶", True),
 ("0930528","AH1",  None, 10960, 10976,   "", "備註欄印「036 93.5 月進修部鐘點費」（本頁最後一列）", True),
]

wb = load_workbook(FN)
name = "OB202第1頁_92.08-93.05"
if name in wb.sheetnames:
    del wb[name]
ws = wb.create_sheet(name, 0)

cols = ["#","記帳日(民國)","西元日期","摘要","借方金額","貸方金額","餘額",
        "對方行庫／代收付代號","備註","勾稽差額"]
for i, h in enumerate(cols, 1):
    c = ws.cell(1, i, h); c.font = HDR; c.fill = FILL
    c.alignment = Alignment("center", "center"); c.border = BOX

r = 2
for n, (d, memo, dr, cr, bal, cp, note, sure) in enumerate(R, 1):
    ws.cell(r, 1, n)
    ws.cell(r, 2, d)
    ws.cell(r, 3, ad(d))
    ws.cell(r, 4, memo)
    if dr is not None: ws.cell(r, 5, dr).number_format = '#,##0'
    if cr is not None: ws.cell(r, 6, cr).number_format = '#,##0'
    if bal is not None: ws.cell(r, 7, bal).number_format = '#,##0'
    ws.cell(r, 8, cp)
    ws.cell(r, 9, note)
    # balance chain: 本列餘額 - (前列餘額 + 貸方 - 借方); blank when either balance missing
    ws.cell(r, 10, '=IF(OR(G%d="",G%d=""),"",ROUND(G%d-(G%d+N(F%d)-N(E%d)),2))'
                   % (r, r-1, r, r-1, r, r)).number_format = '#,##0;[Red]-#,##0'
    for i in range(1, 11):
        cc = ws.cell(r, i); cc.border = BOX
        cc.font = BODY if sure else RED
        if not sure: cc.fill = AMB
    r += 1

ws.cell(2, 10, "")          # first row has no predecessor
last = r - 1

r += 1
ws.cell(r, 4, "薪資／鐘點費入帳合計").font = BOLD
c = ws.cell(r, 6, '=SUMIF($D$2:$D$%d,"AH*",$F$2:$F$%d)' % (last, last))
c.font = BOLD; c.number_format = '#,##0'; c.fill = YEL
ws.cell(r, 9, "摘要以 AH 起首者，即代收付代號 7379 之薪資／鐘點費入帳").font = NOTE
r += 1
ws.cell(r, 4, "轉出（NBFF）合計").font = BOLD
c = ws.cell(r, 5, '=SUMIF($D$2:$D$%d,"NBFF",$E$2:$E$%d)' % (last, last))
c.font = BOLD; c.number_format = '#,##0'; c.fill = YEL
r += 1
ws.cell(r, 4, "勾稽不符列數").font = BOLD
c = ws.cell(r, 10, '=COUNTIF($J$2:$J$%d,"<>0")-COUNTBLANK($J$2:$J$%d)' % (last, last))
c.font = BOLD; c.fill = YEL
r += 2

for t, f in [
 ("本表為土地銀行「存摺類存款分戶明細表（OB202）」第 1 頁之逐筆重建，期間民國 92.08.01～93.05.28。", NOTE),
 ("勾稽差額＝本列餘額－（前列餘額＋貸方－借方）。數值為 0 表示該列判讀與原表餘額相符；非 0 者即為判讀存疑，已以黃底紅字標示。", NOTE),
 ("⚠ 黃底紅字各列之數字係原印表過細、無法確認，部分為依前後列推算所得，不得逕行引用；呈庭前應以銀行核發之原本或重新掃描之清晰影本核對。", RED),
 ("本頁重大事項：92.10.15 摘要 HRT 有一筆約 599,978 元之大額入帳（當事人註記為「一筆貸款」），入帳當日即分筆轉出，其中一筆 100,000 元轉入本人中國信託 118530360199 帳戶。", RED),
 ("本頁結構性事實：92.09～93.05 每月薪資／鐘點費入帳後，均於數日內幾乎全額轉出至第一銀行 14351514421、中國信託 118530360199 及土銀 0055928xxx 等帳戶，月底餘額多僅剩個位數至數百元。", RED),
 ("上開結構表示：本帳戶為「薪資中轉戶」，家庭支出與學費若由受款帳戶支付，本帳戶自不會出現該等支出紀錄。此與郵局帳戶不會出現學費之理由相同。", RED),
]:
    ws.cell(r, 4, t).font = f
    r += 1

for col, w in zip("ABCDEFGHIJ", [5, 13, 12, 8, 13, 13, 13, 34, 52, 11]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A2"

wb.save(FN)
print("sheet added:", name, "rows:", last - 1)
