from OneDataOfAcc import OneDataOfAcc
from get_all_sunday_income_subpoena import get_all_sunday_income_subpoena


import linque as lq


import typing as t


def filter_subpoena_sunday_and_income(rows: t.List[OneDataOfAcc])->t.List[OneDataOfAcc]:
    """ 準備給比對 主日奉獻 用的資料，另外還將會有 轉帳奉獻 """
    names = get_all_sunday_income_subpoena(rows)
    def fn_1(a1: OneDataOfAcc)->bool:
        return a1.subpoena in names
    def fn_2(a1:OneDataOfAcc)->bool:
        return a1.subject[0] == '4'
    r3 = lq.linque(rows).where(fn_1).where(fn_2).to_list()    
    return r3

def filter_subpoena_not_sunday_and_income(rows: t.List[OneDataOfAcc])->t.List[OneDataOfAcc]:
    """ 準備給比對 非主日奉獻 用的資料，另外還將會有 轉帳奉獻 """
    names = get_all_sunday_income_subpoena(rows)
    def fn_1(a1: OneDataOfAcc)->bool:
        return a1.subpoena not in names
    def fn_2(a1:OneDataOfAcc)->bool:
        return a1.subject[0] == '4'
    r3 = lq.linque(rows).where(fn_1).where(fn_2).to_list()    
    return r3
