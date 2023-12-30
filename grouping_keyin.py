from OneKeyin import OneKeyin


import linque as lq


import typing as t


def grouping_keyin(rows: t.List[OneKeyin]) -> t.List[OneKeyin]:
    ''' group by 日期 then by 科目 then by 部門 '''
    re: t.List[OneKeyin] = []
    for a1 in lq.linque(rows).group(lambda a1: a1.date):
        r1: lq.Linque = a1[1]
        for a2 in lq.linque(r1).group(lambda a1: a1.subjectNumber[0]):
            r2: lq.Linque = a2[1]
            for a3 in lq.linque(r2).group(lambda a1: a1.subjectNumber[1]):
                r3: lq.Linque = a3[1]

                money = r3.select(lambda a1: a1.money).sum()

                subject1 = r3.first().subject1
                subject2 = r3.first().subject2

                who = ' ; '.join(r3.select(lambda a1: a1.who).where(lambda a1: a1 is not None and len(a1)>0).to_list())
                memo = ' ; '.join(r3.select(lambda a1: a1.memo).where(lambda a1: a1 is not None and len(a1)>0).to_list())

                re.append(OneKeyin(money, a1[0], subject1, subject2, who, memo))
    return re