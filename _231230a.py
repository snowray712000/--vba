#%%
# 列出主日，以作比較
from OneDataOfAcc import OneDataOfAcc
from generate_dict_subpoena_date import generate_dict_subpoena_date
from get_all_data_of_acc import get_all_data_of_acc
from get_all_data_of_keyin import get_all_data_of_keyin

data_acc = get_all_data_of_acc('2023傳票_20231216匯出.xlsx')
data_keyin_sunday, data_keyin_transfer = get_all_data_of_keyin('2023傳票_20231216匯出.xlsx')

#%%
from filter_subpoena_sunday_and_income import filter_subpoena_sunday_and_income
from sort_acc import sort_acc
from sort_keyin import sort_keyin
from grouping_keyin import grouping_keyin

data_acc_sunday = sort_acc( filter_subpoena_sunday_and_income(data_acc) )
data_keyin_sunday = sort_keyin( grouping_keyin(data_keyin_sunday) )

# 1/1雙福建堂奉獻存入土銀建堂
from datetime import datetime
import typing as t
import linque as lq
dict_date_subpoena = generate_dict_subpoena_date(data_acc_sunday)

#%%
from OneKeyin import OneKeyin

# 列出所有 key, 最後再把沒輸出的輸出出來
keyBool_acc = lq.linque(data_acc_sunday).select(lambda a1: a1.subpoena).distinct().to_dict(lambda a1: a1, lambda a1: False)
keyBool_keyin = lq.linque(data_keyin_sunday).select(lambda a1: a1.date).distinct().to_dict(lambda a1: a1, lambda a1: False)

print(len(keyBool_acc))
print(len(keyBool_keyin))
import openpyxl
class ExportCompare:
    def __init__(self, sh):
        self.row = 2     
        self.sh = sh
    def write_header(self):
        sh = self.sh
        
        sh.cell(row=1, column=1).value = '傳票號碼'
        sh.cell(row=1, column=2).value = '日期'
        sh.cell(row=1, column=3).value = '科目'
        sh.cell(row=1, column=4).value = '部門'
        sh.cell(row=1, column=5).value = '金額'
        sh.cell(row=1, column=6).value = '摘要'
        sh.cell(row=1, column=7).value = ''
        sh.cell(row=1, column=8).value = '日期'
        sh.cell(row=1, column=9).value = '科目'
        sh.cell(row=1, column=10).value = '部門'
        sh.cell(row=1, column=11).value = '金額'
        sh.cell(row=1, column=12).value = '摘要'
        
    def _chk_is_fit_any_keyin(self, acc: OneDataOfAcc, keyins: t.List[OneKeyin])->bool:
        '''其中一個 acc, 是否在 keyins 中有找到, 為了上色(綠色，表示有)
        - 科目，部門，金額要對。
        '''
        for a1 in keyins:
            if a1.money != acc.money:
                continue
            r1 = a1.subjectNumber
            if r1[0]!=acc.subject or r1[1]!=acc.department:
                continue
            return True            
        return False
    def _chk_is_fit_any_acc(self, keyin: OneKeyin, accs: t.List[OneDataOfAcc])->bool:
        '''其中一個 keyin, 是否在 accs 中有找到, 為了上色(綠色，表示有)
        - 科目，部門，金額要對。
        '''
        for a1 in accs:
            if a1.money != keyin.money:
                continue
            if a1.subject!=keyin.subjectNumber[0] or a1.department!=keyin.subjectNumber[1]:
                continue
            return True
        return False
    
    def do_one_pair(self,acckey: str, keyinkey: datetime):
        r1: t.List[OneDataOfAcc] = lq.linque(data_acc_sunday).where(lambda a1: a1.subpoena==acckey).to_list()
        r2: t.List[OneKeyin] = lq.linque(data_keyin_sunday).where(lambda a1: a1.date==keyinkey).to_list()
        
        # acc 傳票號碼 日期 科目 部門 金額 摘要 ; ; keyin 日期 科目 部門 金額 摘要
        countRow = max(len(r1), len(r2))
        row = self.row
        sh = self.sh
        for i in range(countRow):
            acc = r1[i] if i<len(r1) else None
            keyin = r2[i] if i<len(r2) else None
            if acc is not None:
                self._print_acc(acc, row)
                if self._chk_is_fit_any_keyin(acc, r2):
                    cell = sh.cell(row=row, column=5)
                    # 文字顏色 綠色
                    cell.font = openpyxl.styles.Font(color='208800')
            
            if keyin is not None:
                self._print_keyin(keyin, row)
                if self._chk_is_fit_any_acc(keyin, r1):
                    cell = sh.cell(row=row, column=11)
                    # 文字顏色 綠色
                    cell.font = openpyxl.styles.Font(color='008800')
                        
            row += 1
            self.row += 1
        self.row += 1 # 空白一行
        pass
    def _print_acc(self, acc: OneDataOfAcc, row: int):
        sh = self.sh
        
        sh.cell(row=row, column=1).value = acc.subpoena
        sh.cell(row=row, column=2).value = acc.date
        sh.cell(row=row, column=3).value = acc.subject
        sh.cell(row=row, column=4).value = acc.department
        sh.cell(row=row, column=5).value = acc.money
        sh.cell(row=row, column=6).value = acc.memo
        
    def _print_keyin(self, keyin: OneKeyin, row: int):
        sh = self.sh
        
        sh.cell(row=row, column=8).value = keyin.date
        sh.cell(row=row, column=9).value = keyin.subjectNumber[0]
        sh.cell(row=row, column=10).value = keyin.subjectNumber[1]
        sh.cell(row=row, column=11).value = keyin.money
        sh.cell(row=row, column=12).value = keyin.memo
        
        
    def do_residue(self, acc2: t.List[OneDataOfAcc], keyin2: t.List[OneKeyin]):
        sh = self.sh
        row = self.row
        countRow = max(len(acc2), len(keyin2))
        
        for i in range(countRow):
            acc = acc2[i] if i<len(acc2) else None
            keyin = keyin2[i] if i<len(keyin2) else None
            
            if acc is not None:
                self._print_acc(acc, row)
            
            if keyin is not None:
                self._print_keyin(keyin, row)
                        
            row += 1
            self.row += 1            
        pass
        
wb = openpyxl.Workbook()
wb.remove(wb["Sheet"])
sh = wb.create_sheet('主日_比較')
exportCompare = ExportCompare(sh)
exportCompare.write_header()

# 以 keyin 鍵為主，從 acc 中找，如果沒找到，就記載
for a1 in lq.linque(data_keyin_sunday).group(lambda a1: a1.date):
    date: datetime = a1[0]
    r1: lq.Linque = a1[1]
    if date in dict_date_subpoena:
        subpoena = dict_date_subpoena[date]
        keyBool_acc[subpoena] = True
        keyBool_keyin[date] = True    
            
        exportCompare.do_one_pair(subpoena, date)
    else:
        print(date)

# 剩下沒有輸出的 acc
acc2: t.List[OneDataOfAcc] = []
for a1 in lq.linque(keyBool_acc).where(lambda a1: keyBool_acc[a1]==False):
    r1 = lq.linque(data_acc_sunday).where(lambda a2: a2.subpoena==a1).to_list()
    acc2.extend(r1)

# 剩下沒有輸出的 keyin
keyin2: t.List[OneKeyin] = []
for a1 in lq.linque(keyBool_keyin).where(lambda a1: keyBool_keyin[a1]==False):
    r1 = lq.linque(data_keyin_sunday).where(lambda a2: a2.date==a1).to_list()
    keyin2.extend(r1)

exportCompare.do_residue(acc2, keyin2)

# 修正格式 (非必要)
# sheet column B format
row = sh.max_row
for i in range(2, row+1):
    cell = sh.cell(row=i, column=2)
    cell.number_format = 'm/d'
    cell = sh.cell(row=i, column=8)
    cell.number_format = 'm/d'
# E, K 金額，加上千分位
for i in range(2, row+1):
    cell = sh.cell(row=i, column=5)
    cell.number_format = '#,##0'
    cell = sh.cell(row=i, column=11)
    cell.number_format = '#,##0'

wb.save('程式測試_比較.xlsx')
wb.close()