#%%
from __future__ import annotations

import typing as t
import PySimpleGUI as sg
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from export_and_upload_xlsx import export_and_upload_xlsx
from get_all_data_of_keyin import get_all_data_of_keyin
from get_person_and_to_dict import get_person_and_to_dict

dictPerson = get_person_and_to_dict('雙福教會會友列表2023.xlsx')
    
#%%
fileBrowse = sg.FileBrowse('keyin檔案', target='textFile',
    file_types=(('keyin excel', '*.xlsx'),('all files', '*.*')))

sliderCount = sg.Slider(range=(1, 500), default_value=20, orientation='h', key='sliderCount')
layout: t.List[t.List[sg.Element]] = [
    [sg.InputText('2023 奉獻輸入資料.xlsx',key='textFile'), fileBrowse],
    [sg.Button('載入keyin資料',key='btnLoad')],
    [sg.Radio('主日','tp1',default=True, key='tp1'), sg.Radio('轉帳','tp1')],
    [sg.T('上傳筆數'), sliderCount],
    [sg.Button('產生',key='btnExport')],
]

win = sg.Window('240110b', layout=layout)

#%
# callbacks 

def choose_sheet_name(tp: t.Literal['主日','轉帳']) -> str:
    if tp == '主日':
        return '輸入原始資料 2023'
    else:
        return '輸入原始資料_2023_轉帳'
    
while True:
    event, values = win.read(timeout=20)
    if event == 'Exit' or event == sg.WIN_CLOSED:
        break
    if event == 'btnLoad':
        print(values['textFile'])
        try:
            data_keyin_sunday, data_keyin_transfer = get_all_data_of_keyin(values['textFile'])
            sg.popup_ok(f'載入成功 {len(data_keyin_sunday)} 筆資料 {len(data_keyin_transfer)} 筆資料')
        except Exception as e:
            print(e)
            sg.popup_error(e)
            continue
    if event == 'btnExport':
        # backup .xlsx
        if 'OK' == sg.popup_ok_cancel('是否備份 keyin 檔案'):
            # copy file to backup ... .bak.240113_235901.xlsx  , 包含時間
            import shutil
            import os
            import datetime
            r1,r2 = os.path.splitext(values['textFile'])
            now = datetime.datetime.now().strftime('%y%m%d_%H%M%S')
            shutil.copyfile(values['textFile'], r1 + '.bak.' + now + r2)
        
        # sheet for upload 
        wbRe = Workbook()
        shRe: Worksheet = wbRe.active
        
        # sheet of keyin 
        pathExcelKeyin = values['textFile']
        tp = '主日' if values['tp1'] else '轉帳'
        wbKeyin = openpyxl.load_workbook(pathExcelKeyin)
        shKeyin = wbKeyin[choose_sheet_name(tp)]
        
        # data
        data = data_keyin_sunday if tp == '主日' else data_keyin_transfer
        
        # count
        cntTake = int(values['sliderCount'])
        
        # main, modify shRe and shKeyin
        export_and_upload_xlsx(shRe, shKeyin, data, cntTake, tp, dictPerson)
        
        # save
        now = datetime.datetime.now().strftime('%y%m%d_%H%M%S')
        wbRe.save(f"程式測試_上傳資料_{now}.xlsx")
        
        while True:
            try:
                wbKeyin.save(pathExcelKeyin)
                break 
            except Exception as e:
                sg.popup_error(e)
                sg.popup_ok('請關閉 Keyin Excel 再按確定')
        sg.popup_ok('完成')
    if event == 'btnUpdate':
        pass
    if event == 'btnOK':
        print(values['textFile'])

win.close()
win = None

#%%
# 開發演算法時用
