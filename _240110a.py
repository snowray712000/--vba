'''
- 產生要上傳給 奉獻系統 的檔案。
- 注意
    - 因為上傳必需奉獻者都是會員，所以要先確認 當年上傳的人員 是否都在會員名冊中。
    - 奉獻類別，是其它，就不用上傳。代轉，也不用上傳。
    - 小心，不要重複上傳。
    - 匯入一次只能 500 筆
'''
#%%
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
import linque as lq
import typing as t

#%%
from get_all_data_of_keyin import get_all_data_of_keyin
data_keyin_sunday, data_keyin_transfer = get_all_data_of_keyin('2023 奉獻輸入資料.xlsx')

#%%
from OneKeyin import OneKeyin

r1 = lq.linq(data_keyin_sunday).count(lambda a1: a1.isUpload)
print(f'已經上傳的筆數: {r1}')
# 
# wb2 = openpyxl.load_workbook('2023 奉獻輸入資料.xlsx')
# sh2a = wb2['輸入原始資料 2023']
# sh2b = wb2['輸入原始資料_2023_轉帳']
