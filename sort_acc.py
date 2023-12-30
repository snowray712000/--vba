from OneDataOfAcc import OneDataOfAcc


import linque as lq


import typing as t


def sort_acc(rows: t.List[OneDataOfAcc])->t.List[OneDataOfAcc]:
    ''' sort by subpoena, subject, department '''
    def fn_sort_key(a1: OneDataOfAcc):
        return (a1.subpoena, a1.subject, a1.department)
    return lq.linque(rows).sort(fn_sort_key).to_list()