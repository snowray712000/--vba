#%%
# 取得所有資料，兩個都會用到
from ExportCompare import ExportCompare
from get_all_data_of_acc import get_all_data_of_acc
from get_all_data_of_keyin import get_all_data_of_keyin
from write_transfer_sheet import write_transfer_sheet

data_acc = get_all_data_of_acc('2023傳票_20231216匯出.xlsx')
data_keyin_sunday, data_keyin_transfer = get_all_data_of_keyin('2023 奉獻輸入資料.xlsx')

#%%
# 產生 主日 的比較
from filter_subpoena_sunday_and_income import filter_subpoena_not_sunday_and_income, filter_subpoena_sunday_and_income
from sort_acc import sort_acc
from sort_keyin import sort_keyin
from grouping_keyin import grouping_keyin

# 產生 主日 的比較
data_acc_sunday = sort_acc( filter_subpoena_sunday_and_income(data_acc) )
data_keyin_sunday = sort_keyin( grouping_keyin(data_keyin_sunday) )

# 產生 非主日/轉帳 的比較
data_not_sunday_acc = sort_acc(filter_subpoena_not_sunday_and_income(data_acc))
deta_not_sunday_keyin = sort_keyin( data_keyin_transfer )

#%
import openpyxl
wb = openpyxl.Workbook()
wb.remove(wb["Sheet"])

sh = wb.create_sheet('主日_比較')
exportCompare = ExportCompare()
exportCompare.main(sh, data_acc_sunday, data_keyin_sunday)

sh = wb.create_sheet('轉帳_比較')
write_transfer_sheet(sh, data_not_sunday_acc, deta_not_sunday_keyin)

wb.save('程式測試_比較.xlsx')
wb.close()