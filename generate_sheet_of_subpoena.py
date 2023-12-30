from OneDataOfAcc import OneDataOfAcc


import typing as t


def generate_sheet_of_subpoena(sh, rows: t.List[OneDataOfAcc]):

    sh.cell(row=1, column=1).value = '傳票編號'
    sh.cell(row=1, column=2).value = '傳票日期'
    sh.cell(row=1, column=3).value = '部門'
    sh.cell(row=1, column=4).value = '科目'
    sh.cell(row=1, column=5).value = '摘要'
    sh.cell(row=1, column=6).value = '金額'

    row = 2
    for i in rows:
        sh.cell(row=row, column=1).value = i.subpoena
        sh.cell(row=row, column=2).value = i.date
        sh.cell(row=row, column=3).value = i.department
        sh.cell(row=row, column=4).value = i.subject
        sh.cell(row=row, column=5).value = i.memo
        sh.cell(row=row, column=6).value = i.money
        row += 1