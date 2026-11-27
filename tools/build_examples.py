import random, sys, os
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment

OUT = sys.argv[1]
os.makedirs(OUT, exist_ok=True)
random.seed(2026)

FONT = "微軟正黑體"
F = lambda **k: Font(name=FONT, size=k.pop("size", 11), **k)
HEAD_FILL = PatternFill("solid", fgColor="1F5F6B")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
SUM_FILL = PatternFill("solid", fgColor="E8F1F2")
thin = Side(style="thin", color="B7C4C7")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center")
WRAP = Alignment(wrap_text=True, vertical="top")


def header(ws, row, titles, widths=None):
    for i, t in enumerate(titles, 1):
        c = ws.cell(row=row, column=i, value=t)
        c.font = F(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = CENTER
        c.border = BORDER
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.row_dimensions[row].height = 22


def body(ws, r1, r2, c1, c2, center_cols=()):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            if cell.font is None or cell.font.name != FONT:
                cell.font = F()
            if c in center_cols:
                cell.alignment = CENTER


def guide_sheet(wb, title, level, goals, prompt, checks, legend):
    ws = wb.active
    ws.title = "說明"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 90
    ws["B2"] = title
    ws["B2"].font = F(size=18, bold=True, color="1F5F6B")
    ws["B3"] = f"難度：{level}　｜　這是老師用 AI 產生的範例檔，可以拿來和你自己做的版本比較"
    ws["B3"].font = F(size=10, color="5A6B6E")
    r = 5
    for label, content in [("學習重點", goals), ("給 Codex 的提示詞", prompt),
                           ("驗收方式", checks), ("怎麼使用", legend)]:
        ws.cell(row=r, column=2, value=label).font = F(bold=True, color="1F5F6B")
        ws.cell(row=r, column=2).alignment = Alignment(vertical="top")
        c = ws.cell(row=r, column=3, value=content)
        c.font = F()
        c.alignment = WRAP
        if label == "給 Codex 的提示詞":
            c.fill = PatternFill("solid", fgColor="EEF3F4")
        lines = content.count("\n") + 1 + len(content) // 70
        ws.row_dimensions[r].height = max(20, 17 * lines)
        r += 2
    ws.cell(row=r, column=2, value="圖例").font = F(bold=True, color="1F5F6B")
    ws.cell(row=r, column=3, value="淺黃色底的儲存格＝可以自己修改的輸入格；其他欄位都是公式，會自動計算").fill = INPUT_FILL
    ws.cell(row=r, column=3).font = F()
    return ws


SURN = list("陳林黃張李王吳劉蔡楊許鄭謝郭洪曾邱廖賴周葉蘇莊呂江何")
GIVEN = list("宇恩宥辰睿承翔彥廷柏妤欣涵芷琪佳穎婕雅晴安品子語心詠浩傑哲瑜庭")


def names(n):
    out = set()
    while len(out) < n:
        out.add(random.choice(SURN) + random.choice(GIVEN) + random.choice(GIVEN))
    return list(out)


def fix_chart(ch, axes=True):
    ch.roundedCorners = False
    if ch.legend is not None:
        ch.legend.overlay = False
    if ch.title is not None:
        ch.title.overlay = False
    if axes:
        ch.x_axis.delete = False
        ch.y_axis.delete = False


# ---------------------------------------------------------------- 1. 成績表
def build_grades():
    wb = Workbook()
    guide_sheet(
        wb, "專案 1｜班級成績計算表", "入門",
        "SUM、AVERAGE、ROUND、RANK、IF、COUNTIF 等基本公式\n條件式格式（不及格整列變色）、資料驗證（分數只能填 0–100）、凍結窗格、長條圖",
        "建立一個 Excel 檔，包含 30 位學生（隨機中文姓名）的國文、英文、數學成績（40–100 分）。"
        "加上總分、平均（取到小數第一位）、班排名、等第欄位，平均低於 60 分的整列標成淺紅色。"
        "表格下方加上各科平均、最高分、最低分、及格人數。第一列凍結，表頭用粗體加底色，"
        "另外畫一張各科平均的長條圖。所有計算都要用 Excel 公式，不要直接填數字。",
        "1. 把某位同學的數學改成 30，看總分、平均、排名、等第有沒有自動更新\n"
        "2. 他的整列有沒有變成淺紅色？\n3. 試著在分數欄輸入 120，Excel 應該會擋下來\n4. 點總分欄的儲存格，上方應該看到 =SUM(...) 公式，而不是數字",
        "只修改黃色的成績欄（C～E 欄），其他欄位會自動算好。")
    ws = wb.create_sheet("成績表")
    header(ws, 1, ["座號", "姓名", "國文", "英文", "數學", "總分", "平均", "排名", "等第"],
           [7, 11, 8, 8, 8, 9, 9, 8, 8])
    N = 30
    for i, nm in enumerate(names(N), start=2):
        ws.cell(row=i, column=1, value=i - 1)
        ws.cell(row=i, column=2, value=nm)
        for c in (3, 4, 5):
            v = max(40, min(100, int(random.gauss(76, 13))))
            cell = ws.cell(row=i, column=c, value=v)
            cell.fill = INPUT_FILL
        ws.cell(row=i, column=6, value=f"=SUM(C{i}:E{i})")
        ws.cell(row=i, column=7, value=f"=ROUND(AVERAGE(C{i}:E{i}),1)").number_format = "0.0"
        ws.cell(row=i, column=8, value=f"=RANK(F{i},$F$2:$F${N+1})")
        ws.cell(row=i, column=9, value=f'=IF(G{i}>=90,"優",IF(G{i}>=80,"甲",IF(G{i}>=70,"乙",IF(G{i}>=60,"丙","丁"))))')
    last = N + 1
    body(ws, 2, last, 1, 9, center_cols=range(1, 10))
    # 統計列
    stats = [("各科平均", "ROUND(AVERAGE({c}2:{c}%d),1)" % last, "0.0"),
             ("最高分", "MAX({c}2:{c}%d)" % last, "0"),
             ("最低分", "MIN({c}2:{c}%d)" % last, "0"),
             ("及格人數", 'COUNTIF({c}2:{c}%d,">=60")' % last, "0")]
    r = last + 2
    for label, f, fmt in stats:
        ws.cell(row=r, column=1, value=label)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        for col in "CDEFG":
            if col == "F" and label == "及格人數":
                continue  # 總分沒有「及格」的意義
            cell = ws[f"{col}{r}"]
            cell.value = "=" + f.format(c=col)
            cell.number_format = fmt
        for c in range(1, 8):
            ws.cell(row=r, column=c).fill = SUM_FILL
        body(ws, r, r, 1, 7, center_cols=range(1, 8))
        ws.cell(row=r, column=1).font = F(bold=True)
        r += 1
    # 條件格式
    red = PatternFill("solid", fgColor="F8D7D3")
    ws.conditional_formatting.add(f"A2:I{last}", FormulaRule(formula=[f"$G2<60"], fill=red))
    ws.conditional_formatting.add(f"C2:E{last}", CellIsRule(operator="lessThan", formula=["60"], font=Font(color="C0392B", bold=True)))
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="100",
                        showErrorMessage=True, errorTitle="分數超出範圍", error="請輸入 0 到 100 的整數")
    ws.add_data_validation(dv)
    dv.add(f"C2:E{last}")
    ws.freeze_panes = "C2"
    # 圖表
    ch = BarChart()
    ch.type = "col"
    ch.title = "各科平均"
    ch.y_axis.title = "分數"
    ch.y_axis.scaling.min = 0
    ch.y_axis.scaling.max = 100
    data = Reference(ws, min_col=3, max_col=5, min_row=last + 2)
    cats = Reference(ws, min_col=3, max_col=5, min_row=1)
    ch.add_data(data, from_rows=True, titles_from_data=False)
    ch.set_categories(cats)
    ch.legend = None
    ch.series[0].graphicalProperties.solidFill = "2A8C9A"
    ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
    fix_chart(ch)
    ch.height, ch.width = 7.5, 12
    ws.add_chart(ch, "K2")
    wb.active = 1
    wb.save(os.path.join(OUT, "專案1_班級成績計算表.xlsx"))


# ---------------------------------------------------------------- 2. 記帳本
def build_budget():
    wb = Workbook()
    guide_sheet(
        wb, "專案 2｜個人記帳本", "入門",
        "日期格式、下拉選單（資料驗證）、SUMIF／COUNTIF 分類加總、跨工作表參照、圓餅圖",
        "做一個記帳 Excel，第一個工作表「記帳」欄位有日期、類別、項目、金額、備註，先放 20 筆 11 月的範例資料。"
        "「類別」用下拉選單限定為：餐飲、交通、娛樂、學用品、其他；金額只能輸入正整數。"
        "另開一個工作表「統計」，用 SUMIF 算出每個類別的總支出、筆數和占比，"
        "加上一格可以自己填的「每月預算」，算出剩餘金額，超支時顯示「超支！」。最後畫一張各類別支出的圓餅圖。",
        "1. 在「記帳」新增一筆：類別選「娛樂」、金額 500，回到「統計」看娛樂的金額和圓餅圖有沒有變\n"
        "2. 在類別欄打「吃飯」，應該會被擋下來\n3. 把每月預算改成 2000，看狀態是否變成「超支！」",
        "在「記帳」工作表往下填新的一列即可（已預留到第 200 列）。「統計」工作表裡只有黃色的「每月預算」需要自己填。")
    ws = wb.create_sheet("記帳")
    header(ws, 1, ["日期", "類別", "項目", "金額（元）", "備註"], [14, 10, 22, 12, 24])
    items = [("餐飲", "早餐蛋餅＋奶茶", 65), ("交通", "捷運", 30), ("學用品", "原子筆兩支", 40),
             ("餐飲", "便當", 95), ("娛樂", "電影票", 290), ("餐飲", "手搖飲", 60),
             ("交通", "公車儲值", 200), ("學用品", "參考書", 380), ("餐飲", "麥當勞", 159),
             ("其他", "理髮", 250), ("餐飲", "早餐", 55), ("娛樂", "遊戲點數", 150),
             ("餐飲", "滷味", 120), ("交通", "Ubike", 20), ("學用品", "筆記本", 45),
             ("餐飲", "便利商店", 78), ("其他", "生日禮物", 350), ("餐飲", "拉麵", 210),
             ("娛樂", "KTV 分攤", 180), ("餐飲", "早餐", 70)]
    notes = {4: "和同學去看", 9: "週末", 16: "送給好朋友"}
    d = date(2026, 11, 2)
    for i, (cat, it, amt) in enumerate(items, start=2):
        ws.cell(row=i, column=1, value=d).number_format = "yyyy/mm/dd"
        ws.cell(row=i, column=2, value=cat)
        ws.cell(row=i, column=3, value=it)
        ws.cell(row=i, column=4, value=amt).number_format = "#,##0"
        ws.cell(row=i, column=5, value=notes.get(i - 1, ""))
        d += timedelta(days=random.choice([0, 1, 1, 2]))
    body(ws, 2, 200, 1, 5, center_cols=(1, 2))
    for r in range(2, 201):
        ws.cell(row=r, column=1).number_format = "yyyy/mm/dd"
        ws.cell(row=r, column=4).number_format = "#,##0"
    cats = ["餐飲", "交通", "娛樂", "學用品", "其他"]
    dv = DataValidation(type="list", formula1='"' + ",".join(cats) + '"', showErrorMessage=True,
                        errorTitle="類別錯誤", error="請從下拉選單選擇類別")
    ws.add_data_validation(dv); dv.add("B2:B200")
    dv2 = DataValidation(type="whole", operator="greaterThan", formula1="0", showErrorMessage=True,
                         errorTitle="金額錯誤", error="金額請輸入正整數")
    ws.add_data_validation(dv2); dv2.add("D2:D200")
    ws.freeze_panes = "A2"

    st = wb.create_sheet("統計")
    header(st, 1, ["類別", "支出總額", "筆數", "占比"], [12, 12, 8, 10])
    for i, c in enumerate(cats, start=2):
        st.cell(row=i, column=1, value=c)
        st.cell(row=i, column=2, value=f"=SUMIF(記帳!$B$2:$B$200,A{i},記帳!$D$2:$D$200)").number_format = "#,##0"
        st.cell(row=i, column=3, value=f"=COUNTIF(記帳!$B$2:$B$200,A{i})")
        st.cell(row=i, column=4, value=f"=IF($B$7=0,0,B{i}/$B$7)").number_format = "0.0%"
    st["A7"] = "合計"; st["B7"] = "=SUM(B2:B6)"; st["C7"] = "=SUM(C2:C6)"; st["D7"] = "=SUM(D2:D6)"
    st["B7"].number_format = "#,##0"; st["D7"].number_format = "0.0%"
    body(st, 2, 7, 1, 4, center_cols=(1, 3))
    for c in range(1, 5):
        st.cell(row=7, column=c).fill = SUM_FILL
        st.cell(row=7, column=c).font = F(bold=True)
    st["A9"] = "每月預算"; st["B9"] = 3500; st["B9"].fill = INPUT_FILL; st["B9"].number_format = "#,##0"
    st["B9"].comment = Comment("範例數值，請改成你自己的預算", "老師")
    st["A10"] = "剩餘金額"; st["B10"] = "=B9-B7"; st["B10"].number_format = "#,##0;[Red]-#,##0"
    st["A11"] = "狀態"; st["B11"] = '=IF(B10<0,"超支！",IF(B10<B9*0.2,"快用完了","還很充裕"))'
    body(st, 9, 11, 1, 2, center_cols=(1, 2))
    for r in (9, 10, 11):
        st.cell(row=r, column=1).font = F(bold=True)
    st.conditional_formatting.add("B11", FormulaRule(formula=['$B$10<0'], font=Font(color="C0392B", bold=True), fill=PatternFill("solid", fgColor="F8D7D3")))
    pie = PieChart()
    pie.title = "各類別支出占比"
    pie.add_data(Reference(st, min_col=2, min_row=1, max_row=6), titles_from_data=True)
    pie.set_categories(Reference(st, min_col=1, min_row=2, max_row=6))
    pie.dataLabels = DataLabelList(); pie.dataLabels.showPercent = True
    fix_chart(pie, axes=False)
    pie.height, pie.width = 8, 11
    st.add_chart(pie, "F1")
    wb.active = 1
    wb.save(os.path.join(OUT, "專案2_個人記帳本.xlsx"))


# ---------------------------------------------------------------- 3. 圖書借閱
def build_library():
    wb = Workbook()
    guide_sheet(
        wb, "專案 5｜圖書借閱紀錄", "中級",
        "日期加減、TODAY()、巢狀 IF 判斷狀態、MAX 避免負數、條件式格式提醒、把「設定值」獨立成一格讓公式引用",
        "做一個班級圖書角的借閱紀錄 Excel。最上方放三個設定格：今天日期、借閱天數（14 天）、每日逾期罰款（2 元）。"
        "表格欄位有編號、書名、借閱人、借出日期、應還日期、歸還日期、狀態、逾期天數、罰款。"
        "應還日期＝借出日期＋借閱天數；狀態依情況顯示「已歸還」「逾期」「即將到期」（3 天內）或「借閱中」。"
        "逾期的整列標紅色、即將到期的標黃色、已歸還的字變灰色。另開「統計」工作表，計算各狀態的本數和罰款總額。"
        "公式都要引用上方的設定格，不要把 14 天直接寫在公式裡。",
        "1. 把「借閱天數」從 14 改成 21，看應還日期和狀態是否全部跟著改變\n"
        "2. 在某本逾期書的「歸還日期」填上日期，狀態應變成「已歸還」，字變灰色\n"
        "3. 把「今天日期」改成 =TODAY()，表格就會用真正的今天來判斷",
        "黃色的設定格和「歸還日期」欄可以修改。今天日期目前固定為 2026/11/25 方便課堂示範，實際使用時請改成 =TODAY()。")
    ws = wb.create_sheet("借閱紀錄")
    ws["A1"] = "今天日期"; ws["B1"] = date(2026, 11, 25); ws["B1"].number_format = "yyyy/mm/dd"
    ws["B1"].comment = Comment("為了課堂示範固定成 2026/11/25。改成 =TODAY() 就會用電腦的今天日期。", "老師")
    ws["A2"] = "借閱天數"; ws["B2"] = 14
    ws["A3"] = "每日罰款"; ws["B3"] = 2; ws["B3"].number_format = '0" 元"'; ws["B2"].number_format = '0" 天"'
    for r in (1, 2, 3):
        ws.cell(row=r, column=1).font = F(bold=True, color="1F5F6B")
        ws.cell(row=r, column=2).fill = INPUT_FILL
        ws.cell(row=r, column=2).font = F(bold=True)
        ws.cell(row=r, column=2).border = BORDER
        ws.cell(row=r, column=2).alignment = CENTER
    H = 5
    header(ws, H, ["編號", "書名", "借閱人", "借出日期", "應還日期", "歸還日期", "狀態", "逾期天數", "罰款（元）"],
           [11, 24, 10, 14, 14, 14, 11, 10, 11])
    books = ["小王子", "哈利波特：神秘的魔法石", "追風箏的孩子", "解憂雜貨店", "被討厭的勇氣", "原子習慣",
             "夜市人生", "動物農莊", "少年Pi的奇幻漂流", "偷書賊", "蘇菲的世界", "牧羊少年奇幻之旅",
             "人類大歷史", "我的阿勒泰", "灌籃高手 1", "三國演義（少年版）", "西遊記（少年版）", "福爾摩斯探案集"]
    people = names(len(books))
    base = date(2026, 11, 25)
    # 借出日距今天幾天、是否已歸還（設計成各種狀態都有）
    plan = [(30, 12), (25, None), (20, None), (19, None), (13, None), (12, None), (10, None),
            (8, None), (5, None), (3, None), (28, 15), (40, 10), (17, None), (2, None),
            (22, 20), (12, None), (1, None), (35, None)]
    first = H + 1
    for k, (bk, (ago, ret)) in enumerate(zip(books, plan)):
        r = first + k
        ws.cell(row=r, column=1, value=f"B{k+1:03d}")
        ws.cell(row=r, column=2, value=bk)
        ws.cell(row=r, column=3, value=people[k])
        ws.cell(row=r, column=4, value=base - timedelta(days=ago)).number_format = "yyyy/mm/dd"
        ws.cell(row=r, column=5, value=f"=D{r}+$B$2").number_format = "yyyy/mm/dd"
        rc = ws.cell(row=r, column=6, value=(base - timedelta(days=ago) + timedelta(days=ret)) if ret else None)
        rc.number_format = "yyyy/mm/dd"; rc.fill = INPUT_FILL
        ws.cell(row=r, column=7, value=f'=IF(F{r}<>"","已歸還",IF($B$1>E{r},"逾期",IF(E{r}-$B$1<=3,"即將到期","借閱中")))')
        ws.cell(row=r, column=8, value=f'=IF(F{r}<>"",MAX(0,F{r}-E{r}),MAX(0,$B$1-E{r}))')
        ws.cell(row=r, column=9, value=f"=H{r}*$B$3").number_format = "#,##0"
    lastr = first + len(books) - 1
    body(ws, first, lastr, 1, 9, center_cols=(1, 3, 4, 5, 6, 7, 8, 9))
    for r in range(first, lastr + 1):
        ws.cell(row=r, column=6).fill = INPUT_FILL
    rng = f"A{first}:I{lastr}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$G{first}="逾期"'], fill=PatternFill("solid", fgColor="F8D7D3"), font=Font(color="9B2C1F")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$G{first}="即將到期"'], fill=PatternFill("solid", fgColor="FCEBC0")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$G{first}="已歸還"'], font=Font(color="9AA5A8")))
    ws.freeze_panes = f"A{first}"

    st = wb.create_sheet("統計")
    header(st, 1, ["狀態", "本數"], [14, 10])
    for i, s in enumerate(["借閱中", "即將到期", "逾期", "已歸還"], start=2):
        st.cell(row=i, column=1, value=s)
        st.cell(row=i, column=2, value=f'=COUNTIF(借閱紀錄!$G${first}:$G${lastr},A{i})')
    st["A6"] = "合計"; st["B6"] = "=SUM(B2:B5)"
    st["A8"] = "罰款總額（元）"; st["B8"] = f"=SUM(借閱紀錄!$I${first}:$I${lastr})"
    st["A9"] = "逾期最久"; st["B9"] = f'=MAX(借閱紀錄!$H${first}:$H${lastr})&" 天"'
    body(st, 2, 6, 1, 2, center_cols=(1, 2)); body(st, 8, 9, 1, 2, center_cols=(1, 2))
    for c in (1, 2):
        st.cell(row=6, column=c).fill = SUM_FILL; st.cell(row=6, column=c).font = F(bold=True)
    for r in (8, 9):
        st.cell(row=r, column=1).font = F(bold=True)
    st.column_dimensions["A"].width = 16
    bar = BarChart(); bar.type = "bar"; bar.title = "借閱狀態"
    bar.add_data(Reference(st, min_col=2, min_row=1, max_row=5), titles_from_data=True)
    bar.set_categories(Reference(st, min_col=1, min_row=2, max_row=5))
    bar.legend = None
    bar.series[0].graphicalProperties.solidFill = "2A8C9A"
    bar.dataLabels = DataLabelList(); bar.dataLabels.showVal = True
    fix_chart(bar); bar.height, bar.width = 7, 11
    st.add_chart(bar, "D1")
    wb.active = 1
    wb.save(os.path.join(OUT, "專案5_圖書借閱紀錄.xlsx"))


# ---------------------------------------------------------------- 4. 查詢表
def build_lookup():
    wb = Workbook()
    guide_sheet(
        wb, "專案 8｜學生資料查詢表", "進階",
        "INDEX＋MATCH 查表、IFERROR 處理查無資料、COUNTIFS 算班排名、下拉選單直接選學號、圖表跟著查詢結果變動",
        "建立一個學生資料庫工作表「學生資料」：30 位學生，欄位有學號（11301 起）、姓名、班級（301～303）、"
        "國文、英文、數學、自然、社會、總分、平均、校排名、班排名。"
        "再做一個「查詢」工作表：在黃色儲存格選擇或輸入學號後，下方自動顯示該生的姓名、班級、各科成績、平均和排名，"
        "並畫一張長條圖比較該生各科成績和全體平均。查表請用 INDEX 和 MATCH，學號不存在時顯示「查無此學號」。",
        "1. 在「查詢」的學號格換成別的學號，所有資料和長條圖都要跟著換\n"
        "2. 輸入一個不存在的學號（例如 99999），應顯示「查無此學號」而不是 #N/A\n"
        "3. 回到「學生資料」改某人的成績，再查他一次，排名應該會變",
        "到「查詢」工作表，在黃色的學號格用下拉選單選學號即可。學生資料表裡黃色的成績欄也可以修改。\n"
        "延伸：新版 Excel 也可以用 XLOOKUP 取代 INDEX＋MATCH，可以請 Codex 改寫一版比較看看。")
    ws = wb.create_sheet("學生資料")
    cols = ["學號", "姓名", "班級", "國文", "英文", "數學", "自然", "社會", "總分", "平均", "校排名", "班排名"]
    header(ws, 1, cols, [9, 10, 7, 7, 7, 7, 7, 7, 8, 8, 8, 8])
    N = 30; last = N + 1
    for k, nm in enumerate(names(N)):
        r = k + 2
        ws.cell(row=r, column=1, value=11301 + k)
        ws.cell(row=r, column=2, value=nm)
        ws.cell(row=r, column=3, value=301 + k // 10)
        for c in range(4, 9):
            ws.cell(row=r, column=c, value=max(35, min(100, int(random.gauss(74, 14))))).fill = INPUT_FILL
        ws.cell(row=r, column=9, value=f"=SUM(D{r}:H{r})")
        ws.cell(row=r, column=10, value=f"=ROUND(AVERAGE(D{r}:H{r}),1)").number_format = "0.0"
        ws.cell(row=r, column=11, value=f"=RANK(I{r},$I$2:$I${last})")
        ws.cell(row=r, column=12, value=f'=COUNTIFS($C$2:$C${last},C{r},$I$2:$I${last},">"&I{r})+1')
    body(ws, 2, last, 1, 12, center_cols=range(1, 13))
    for r in range(2, last + 1):
        for c in range(4, 9):
            ws.cell(row=r, column=c).fill = INPUT_FILL
    ws.freeze_panes = "C2"

    q = wb.create_sheet("查詢")
    q.sheet_view.showGridLines = False
    q.column_dimensions["A"].width = 3; q.column_dimensions["B"].width = 12
    q.column_dimensions["C"].width = 16; q.column_dimensions["D"].width = 12
    q["B2"] = "學生成績查詢"; q["B2"].font = F(size=16, bold=True, color="1F5F6B")
    q["B4"] = "請選擇學號"; q["B4"].font = F(bold=True)
    q["C4"] = 11307; q["C4"].fill = INPUT_FILL; q["C4"].font = F(size=13, bold=True); q["C4"].alignment = CENTER
    q["C4"].border = Border(left=Side("medium", "1F5F6B"), right=Side("medium", "1F5F6B"), top=Side("medium", "1F5F6B"), bottom=Side("medium", "1F5F6B"))
    dv = DataValidation(type="list", formula1=f"學生資料!$A$2:$A${last}", showErrorMessage=False)
    q.add_data_validation(dv); dv.add("C4")
    q["D4"] = "← 點一下會出現下拉選單"; q["D4"].font = F(size=9, color="5A6B6E")

    def look(colletter):
        return f'=IFERROR(INDEX(學生資料!${colletter}$2:${colletter}${last},MATCH($C$4,學生資料!$A$2:$A${last},0)),"查無此學號")'
    info = [("姓名", "B"), ("班級", "C"), ("總分", "I"), ("平均", "J"), ("校排名", "K"), ("班排名", "L")]
    r = 6
    for label, col in info:
        q.cell(row=r, column=2, value=label).font = F(bold=True, color="1F5F6B")
        c = q.cell(row=r, column=3, value=look(col))
        if label == "平均":
            c.number_format = "0.0"
        r += 1
    body(q, 6, 11, 2, 3, center_cols=(2, 3))
    # 成績比較表
    q["B13"] = "科目"; q["C13"] = "該生成績"; q["D13"] = "全體平均"
    for c in "BCD":
        q[f"{c}13"].font = F(bold=True, color="FFFFFF"); q[f"{c}13"].fill = HEAD_FILL; q[f"{c}13"].alignment = CENTER
    subj = [("國文", "D"), ("英文", "E"), ("數學", "F"), ("自然", "G"), ("社會", "H")]
    for i, (s, col) in enumerate(subj, start=14):
        q.cell(row=i, column=2, value=s)
        q.cell(row=i, column=3, value=f'=IFERROR(INDEX(學生資料!${col}$2:${col}${last},MATCH($C$4,學生資料!$A$2:$A${last},0)),0)')
        q.cell(row=i, column=4, value=f"=ROUND(AVERAGE(學生資料!{col}2:{col}{last}),1)").number_format = "0.0"
    body(q, 13, 18, 2, 4, center_cols=(2, 3, 4))
    q.conditional_formatting.add("C14:C18", FormulaRule(formula=["C14<D14"], font=Font(color="C0392B", bold=True)))
    q.conditional_formatting.add("C14:C18", FormulaRule(formula=["C14>=D14"], font=Font(color="1E7B4A", bold=True)))
    q["B19"] = "綠色＝高於全體平均，紅色＝低於全體平均"; q["B19"].font = F(size=9, color="5A6B6E")
    ch = BarChart(); ch.type = "col"; ch.title = "各科成績 vs 全體平均"
    ch.add_data(Reference(q, min_col=3, max_col=4, min_row=13, max_row=18), titles_from_data=True)
    ch.set_categories(Reference(q, min_col=2, min_row=14, max_row=18))
    ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 100
    ch.series[0].graphicalProperties.solidFill = "2A8C9A"
    ch.series[1].graphicalProperties.solidFill = "C9D3D5"
    ch.legend.position = "b"
    fix_chart(ch); ch.height, ch.width = 8, 13
    q.add_chart(ch, "F4")
    wb.active = 2
    wb.save(os.path.join(OUT, "專案8_學生資料查詢表.xlsx"))


# ---------------------------------------------------------------- 5. 測驗卷
def build_quiz():
    wb = Workbook()
    guide_sheet(
        wb, "專案 11｜自動批改測驗卷", "進階",
        "題目資料與答案分開、用公式算出正確答案、作答後即時顯示 ✔／✘、COUNTIF 統計答對題數、分數與評語",
        "做一份 20 題的兩位數加、減、乘法練習卷。「測驗卷」工作表顯示題號、題目和黃色的作答欄，"
        "學生輸入答案後，旁邊自動顯示 ✔ 或 ✘（沒作答就留白）。右上角統計已作答題數、答對題數、得分（每題 5 分）和評語。"
        "「答案」工作表列出每題的正確答案，正確答案要用公式算，不要直接寫數字。減法題要確保答案不是負數。",
        "1. 隨便作答幾題，看 ✔／✘ 和得分是否立刻更新\n2. 把某題的數字改掉，「答案」工作表的正確答案也要跟著變\n"
        "3. 請 Codex「重新出 20 題新的題目」，看它只改題目、公式保持不變",
        "只在黃色的「作答」欄填答案。已先填入 6 題示範作答（其中 1 題故意答錯），可以刪掉重新練習。")
    ws = wb.create_sheet("測驗卷")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "兩位數計算練習"; ws["A1"].font = F(size=16, bold=True, color="1F5F6B")
    ws["A2"] = "姓名："; ws["B2"].fill = INPUT_FILL; ws.merge_cells("B2:C2")
    H = 4
    titles = ["題號", "數字 1", "運算", "數字 2", "", "作答", "批改"]
    header(ws, H, titles, [7, 9, 7, 9, 5, 10, 8])
    ops = ["+", "-", "×"] * 7
    random.shuffle(ops)
    first = H + 1
    for k in range(20):
        r = first + k
        op = ops[k]
        if op == "×":
            a, b = random.randint(11, 39), random.randint(2, 9)
        else:
            a, b = random.randint(20, 99), random.randint(10, 99)
            if op == "-" and b > a:
                a, b = b, a
        ws.cell(row=r, column=1, value=k + 1)
        ws.cell(row=r, column=2, value=a)
        ws.cell(row=r, column=3, value=op)
        ws.cell(row=r, column=4, value=b)
        ws.cell(row=r, column=5, value="=")
        ws.cell(row=r, column=6).fill = INPUT_FILL
        ws.cell(row=r, column=7, value=f'=IF(F{r}="","",IF(F{r}=答案!C{r-H+1},"✔","✘"))')
    lastr = first + 19
    body(ws, first, lastr, 1, 7, center_cols=range(1, 8))
    for r in range(first, lastr + 1):
        ws.cell(row=r, column=6).fill = INPUT_FILL
        ws.cell(row=r, column=6).font = F(bold=True, color="1F4E9C")
        ws.cell(row=r, column=2).font = F(size=12); ws.cell(row=r, column=4).font = F(size=12)
        ws.cell(row=r, column=3).font = F(size=12, bold=True)
    ws.conditional_formatting.add(f"G{first}:G{lastr}", CellIsRule(operator="equal", formula=['"✔"'], font=Font(color="1E7B4A", bold=True)))
    ws.conditional_formatting.add(f"G{first}:G{lastr}", CellIsRule(operator="equal", formula=['"✘"'], font=Font(color="C0392B", bold=True), fill=PatternFill("solid", fgColor="F8D7D3")))
    # 成績欄
    ws.column_dimensions["I"].width = 12; ws.column_dimensions["J"].width = 14
    summ = [("已作答", f'=COUNTA(F{first}:F{lastr})&" / 20"'),
            ("答對題數", f'=COUNTIF(G{first}:G{lastr},"✔")'),
            ("得分", f"=J{H+2}*5"),
            ("評語", f'=IF(COUNTA(F{first}:F{lastr})=0,"尚未作答",IF(COUNTA(F{first}:F{lastr})<20,"還有 "&(20-COUNTA(F{first}:F{lastr}))&" 題沒寫",IF(J{H+3}>=90,"太厲害了！",IF(J{H+3}>=60,"不錯，再檢查錯的題目","多練習幾次就會進步"))))')]
    for i, (lab, f) in enumerate(summ):
        r = H + 1 + i
        ws.cell(row=r, column=9, value=lab).font = F(bold=True, color="1F5F6B")
        ws.cell(row=r, column=10, value=f)
        for c in (9, 10):
            ws.cell(row=r, column=c).border = BORDER; ws.cell(row=r, column=c).alignment = CENTER
            ws.cell(row=r, column=c).fill = SUM_FILL
    ws.cell(row=H + 3, column=10).font = F(size=14, bold=True, color="1F5F6B")
    # 示範作答：前 6 題，第 4 題故意錯
    ans = wb.create_sheet("答案")
    header(ans, 1, ["題號", "題目", "正確答案"], [7, 20, 11])
    for k in range(20):
        r = k + 2; tr = first + k
        ans.cell(row=r, column=1, value=f"=測驗卷!A{tr}")
        ans.cell(row=r, column=2, value=f'=測驗卷!B{tr}&" "&測驗卷!C{tr}&" "&測驗卷!D{tr}&" ="')
        ans.cell(row=r, column=3, value=f'=IF(測驗卷!C{tr}="+",測驗卷!B{tr}+測驗卷!D{tr},IF(測驗卷!C{tr}="-",測驗卷!B{tr}-測驗卷!D{tr},測驗卷!B{tr}*測驗卷!D{tr}))')
    body(ans, 2, 21, 1, 3, center_cols=(1, 2, 3))
    for k in range(6):
        r = first + k
        a = ws.cell(row=r, column=2).value; b = ws.cell(row=r, column=4).value; op = ws.cell(row=r, column=3).value
        v = a + b if op == "+" else a - b if op == "-" else a * b
        ws.cell(row=r, column=6, value=v + (1 if k == 3 else 0))
    wb.active = 1
    wb.save(os.path.join(OUT, "專案11_自動批改測驗卷.xlsx"))


build_grades(); build_budget(); build_library(); build_lookup(); build_quiz()
print("ok")
