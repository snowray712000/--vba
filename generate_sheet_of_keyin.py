#%
# 產生清單
from OneKeyin import OneKeyin


import typing as t


def generate_sheet_of_keyin(sh, rows: t.List[OneKeyin]):
    # 日期	姓名 金額 主科目 次科目	備註
    sh.cell(row=1, column=1).value = '日期'
    sh.cell(row=1, column=2).value = '姓名'
    sh.cell(row=1, column=3).value = '金額'
    sh.cell(row=1, column=4).value = '主科目'
    sh.cell(row=1, column=5).value = '次科目'
    sh.cell(row=1, column=6).value = '備註'
    sh.cell(row=1, column=7).value = '科目代碼'
    sh.cell(row=1, column=8).value = '部門'

    row = 2
    for a1 in rows:
        sh.cell(row=row, column=1).value = a1.date
        sh.cell(row=row, column=2).value = a1._who
        sh.cell(row=row, column=3).value = a1.money
        sh.cell(row=row, column=4).value = a1.subject1str
        sh.cell(row=row, column=5).value = a1.subject2str
        sh.cell(row=row, column=6).value = a1.memo

        r1 = a1.subjectNumber
        sh.cell(row=row, column=7).value = r1[0]
        sh.cell(row=row, column=8).value = r1[1]
        row += 1