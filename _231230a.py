#%%
# 取得所有資料，兩個都會用到
from ExportCompare import ExportCompare
from OneDataOfAcc import OneDataOfAcc
from generate_dict_subpoena_date import generate_dict_subpoena_date
from get_all_data_of_acc import get_all_data_of_acc
from get_all_data_of_keyin import get_all_data_of_keyin

data_acc = get_all_data_of_acc('2023傳票_20231216匯出.xlsx')
data_keyin_sunday, data_keyin_transfer = get_all_data_of_keyin('2023 奉獻輸入資料.xlsx')

#%%
# 產生 主日 的比較
from filter_subpoena_sunday_and_income import filter_subpoena_not_sunday_and_income, filter_subpoena_sunday_and_income
from sort_acc import sort_acc
from sort_keyin import sort_keyin
from grouping_keyin import grouping_keyin

data_acc_sunday = sort_acc( filter_subpoena_sunday_and_income(data_acc) )
data_keyin_sunday = sort_keyin( grouping_keyin(data_keyin_sunday) )

#%

import openpyxl
wb = openpyxl.Workbook()
wb.remove(wb["Sheet"])

sh = wb.create_sheet('主日_比較')
exportCompare = ExportCompare()
exportCompare.main(sh, data_acc_sunday, data_keyin_sunday)

wb.save('程式測試_比較.xlsx')
wb.close()

#%%
# 產生 非主日/轉帳 的比較
from sort_keyin import sort_keyin

data_not_sunday_acc = filter_subpoena_not_sunday_and_income(data_acc)

deta_not_sunday_keyin = sort_keyin( data_keyin_transfer )

# 列出所有存在的日期，並且排序。
import linque as lq
dates = lq.linque(data_not_sunday_acc).select(lambda x: x.date).concat(lq.linque(deta_not_sunday_keyin).select(lambda x: x.date)).distinct().sort().to_list()

# 準備輸出
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from OneDataOfAcc import OneDataOfAcc
from OneKeyin import OneKeyin

import typing as t
import re
if 'wb' not in locals() or wb is None:
    wb = openpyxl.Workbook()
    wb.remove(wb["Sheet"])
    
sh = wb.create_sheet('轉帳_比較')

def write_header(sh: Worksheet):    
    sh.cell(1, 1).value = '傳票號碼'
    sh.cell(1, 2).value = '日期'
    sh.cell(1, 3).value = '科目'
    sh.cell(1, 4).value = '部門'
    sh.cell(1, 5).value = '金額'
    sh.cell(1, 6).value = '摘要'
    sh.cell(1, 8).value = '日期'
    sh.cell(1, 7).value = ''
    sh.cell(1, 9).value = '科目'
    sh.cell(1, 10).value = '部門'
    sh.cell(1, 11).value = '金額'
    sh.cell(1, 12).value = '摘要'
    sh.cell(1, 13).value = '奉獻者' # 要比較奉獻者正確嗎

def write_data(sh: Worksheet):
    def _chk_if_acc_in_keyins(acc: OneDataOfAcc, keyins: t.List[OneKeyin])->bool:
        for keyin in keyins:
            if acc.date == keyin.date and acc.money == keyin.money and acc.subject == keyin.subjectNumber[0] and acc.department == keyin.subjectNumber[1]:
                return True
        return False
    def _chk_if_keyin_in_accs(keyin: OneKeyin, accs: t.List[OneDataOfAcc])->bool:
        for acc in accs:
            if acc.date == keyin.date and acc.money == keyin.money and acc.subject == keyin.subjectNumber[0] and acc.department == keyin.subjectNumber[1]:
                return True
        return False
    def _chk_keyin_name_in_accs(keyin: OneKeyin, accs: t.List[OneDataOfAcc])->bool:
        if '尾碼' in keyin.who:
            # `尾碼07321` 這種格式時，要取出 07321。
            who = re.search(r'尾碼(\d+)', keyin.who).group(1)
            
            for acc in accs:
                if acc.date == keyin.date and acc.money == keyin.money and acc.subject == keyin.subjectNumber[0] and acc.department == keyin.subjectNumber[1] and who in acc.memo: 
                    return True
            return False
        else:
            # `阿張三` 或 `16` 這種格式的
            for acc in accs:
                if acc.date == keyin.date and acc.money == keyin.money and acc.subject == keyin.subjectNumber[0] and acc.department == keyin.subjectNumber[1] and keyin.who in acc.memo: 
                    return True
            return False
    row = 2
    for date in dates:
        # 取得這個日期的 acc 與 keyin
        acc2 = lq.linque(data_not_sunday_acc).where(lambda x: x.date == date).to_list()
        keyin2 = lq.linque(deta_not_sunday_keyin).where(lambda x: x.date == date).to_list()
        
        max_row = max(len(acc2), len(keyin2))
        for i in range(max_row):
            acc3: OneDataOfAcc = acc2[i] if i < len(acc2) else None
            keyin3: OneKeyin = keyin2[i] if i < len(keyin2) else None
            
            if acc3 is not None:
                sh.cell(row, 1).value = acc3.subpoena
                sh.cell(row, 2).value = acc3.date
                sh.cell(row, 3).value = acc3.subject
                sh.cell(row, 4).value = acc3.department
                sh.cell(row, 5).value = acc3.money
                sh.cell(row, 6).value = acc3.memo
                
                if _chk_if_acc_in_keyins(acc3, keyin2):
                    sh.cell(row, 5).font = openpyxl.styles.Font(color='008800')
            if keyin3 is not None:
                sh.cell(row, 8).value = keyin3.date
                sh.cell(row, 9).value = keyin3.subjectNumber[0]
                sh.cell(row, 10).value = keyin3.subjectNumber[1]
                sh.cell(row, 11).value = keyin3.money
                sh.cell(row, 12).value = keyin3.memo
                sh.cell(row, 13).value = keyin3.who
                
                if _chk_if_keyin_in_accs(keyin3, acc2):
                    sh.cell(row, 11).font = openpyxl.styles.Font(color='008800')
                if _chk_keyin_name_in_accs(keyin3, acc2):
                    sh.cell(row, 13).font = openpyxl.styles.Font(color='008800')
            row += 1
        row += 1 # 空一行
# set date format
def set_date_column_format(sh: Worksheet):
    for i in range(1, sh.max_row+1):
        cell = sh.cell(i, 2)
        cell.number_format = 'm/d'
        cell = sh.cell(i, 8)
        cell.number_format = 'm/d'
def set_money_column_format(sh: Worksheet):
    for i in range(1, sh.max_row+1):
        cell = sh.cell(i, 5)
        cell.number_format = '#,##0'
        cell = sh.cell(i, 11)
        cell.number_format = '#,##0'
        
write_header(sh)
write_data(sh)
set_date_column_format(sh)
set_money_column_format(sh)

wb.save('程式測試_比較.xlsx')
wb.close()