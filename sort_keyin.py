from OneKeyin import OneKeyin
import linque as lq
import typing as t
from datetime import datetime


def sort_keyin(rows: t.List[OneKeyin])->t.List[OneKeyin]:
    def fn_sort_by_date_subject_department(a1: OneKeyin)->t.Tuple[datetime,str,str]:
        r1 = a1.subjectNumber
        return (a1.date, r1[0], r1[1])
    return lq.linque(rows).sort(fn_sort_by_date_subject_department).to_list()