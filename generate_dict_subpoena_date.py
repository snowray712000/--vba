from OneDataOfAcc import OneDataOfAcc


import linque as lq


import re
import typing as t
from datetime import datetime


def generate_dict_subpoena_date(rows: t.List[OneDataOfAcc])->t.Dict[datetime,str]:
    """ 正航系統 傳票日期 通常不等於 主日日期 ，因為通常是星期二才會輸入
    - 分析摘要，取得 主日日期；若型如 `5/12.....收入..`，就得到 5/12，並且得到主日日期對應傳票號碼
    - 前提假設) 承上，這樣設計，就是一個主日日期，只能有一個傳票號碼。
    """
    # 字串，若包含 (\d+)/(\d+) 並且 '收入' 字眼，就可以解析得到日期 m/d
    def get_date_from_str(s):
        """ 字串，若包含 (\d+)/(\d+) 並且 '收入' 字眼，就可以解析得到日期 m/d
        - 年, 使用傳票編號的年份即可
        - 若 None, 後面可以使用傳票日期
        - 同一張出現多個日期，選最多的那個日期。(有可能是輸入錯，也有可能是2週作在一起，例如過年)
        """
        m = re.search(r'(\d+)/(\d+)', s)
        if m and '收入' in s:
            return int(m.group(1)),int(m.group(2))
        else:
            return None

    dict_date_subpoena: {datetime:str} = {}
    for a1 in lq.linque(rows).group(lambda a1: a1.subpoena):
        r1: lq.Linque = a1[1]
        date: datetime = r1.first().date

        r2 = r1.select(lambda a2: get_date_from_str(a2.memo)).where(lambda a2: a2 is not None).to_list()
        if len(r2)==0:
            dict_date_subpoena[date] = a1[0]
        else:
            # print(r2)
            # 例子 [(5, 21), (5, 21), (5, 21), (5, 21), (5, 23)]
            # 取得其中，最多次的。上例就是 (5, 21)
            r3 = lq.linque(r2).distinct().to_list()

            # count 出現次數
            # print ( r2.count(r3[0]) )
            r4 = lq.linque(r3).sort(lambda a1: r2.count(a1), True).first()
            date2 = datetime(date.year, r4[0], r4[1])
            dict_date_subpoena[date2] = a1[0]
    return dict_date_subpoena