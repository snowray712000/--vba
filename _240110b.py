#%%
import PySimpleGUI as sg
import typing as t
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
import linque as lq
import typing as t
from get_all_data_of_keyin import get_all_data_of_keyin
from OneKeyin import OneKeyin

#%%
fileBrowse = sg.FileBrowse('keyin檔案', target='textFile',
    file_types=(('keyin excel', '*.xlsx'),('all files', '*.*')))

layout: t.List[t.List[sg.Element]] = [
    [sg.InputText('2023 奉獻輸入資料.xlsx',key='textFile'), fileBrowse],
    [sg.Button('載入keyin資料',key='btnLoad')]
]

win = sg.Window('240110b', layout=layout)

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
    if event == 'btnOK':
        print(values['textFile'])

win.close()
win = None