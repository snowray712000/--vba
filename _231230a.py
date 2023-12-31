#%%
# 列出主日，以作比較
from ExportCompare import ExportCompare
from generate_dict_subpoena_date import generate_dict_subpoena_date
from get_all_data_of_acc import get_all_data_of_acc
from get_all_data_of_keyin import get_all_data_of_keyin

data_acc = get_all_data_of_acc('2023傳票_20231216匯出.xlsx')
data_keyin_sunday, data_keyin_transfer = get_all_data_of_keyin('2023 奉獻輸入資料.xlsx')

#%%
from filter_subpoena_sunday_and_income import filter_subpoena_sunday_and_income
from sort_acc import sort_acc
from sort_keyin import sort_keyin
from grouping_keyin import grouping_keyin

data_acc_sunday = sort_acc( filter_subpoena_sunday_and_income(data_acc) )
data_keyin_sunday = sort_keyin( grouping_keyin(data_keyin_sunday) )

#%%

import openpyxl
wb = openpyxl.Workbook()
wb.remove(wb["Sheet"])

sh = wb.create_sheet('主日_比較')
exportCompare = ExportCompare()
exportCompare.main(sh, data_acc_sunday, data_keyin_sunday)

wb.save('程式測試_比較.xlsx')
wb.close()