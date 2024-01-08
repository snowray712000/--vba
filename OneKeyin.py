#%%
# 取得 B column 的 row 數

import typing as t
from datetime import datetime


class OneKeyin:
    dictSubject: t.Dict[str,str] = {"1":"代轉奉獻", "3":"專案奉獻", "4":"什一奉獻", "5":"主日奉獻", "6":"感恩奉獻", "7":"初熟果子", "8":"主日學", "9":"團契奉獻", "10":"設備購置", "11":"場地維護", "12":"建堂修繕", "13":"宣教事工", "14":"開拓植堂", "15":"神學培育", "16":"獎助學金", "17":"愛心救助", "18":"堂會奉獻款", "19":"堂會奉獻款-外會", "21":"特別奉獻", "22":"牧區愛心基金", "243":"學生中心", "244":"購地", "253":"喜樂班", "331":"雙福恩典團契","x":"其他","x2":"租金"}
    def __init__(self, money: int, date: datetime, subject1: str, subject2: str, who: str, memo: str):
        self.money = money
        self.date = date
        self.subject1 = subject1
        self.subject2 = subject2
        self.who = who
        self.memo = memo
    def __repr__(self) -> str:
        return f'{self.money} {self.date} {self.subject1str} {self.subject2str} {self.who}'
    @property
    def subject2str(self)->str:
        # check        
        if self.subject1 is None:
            return ''
        if self.subject2 not in OneKeyin.dictSubject:
            return ''

        return OneKeyin.dictSubject[self.subject2]
    @property
    def subject1str(self)->str:
        # check exist key
        if self.subject1 is None:
            return ''
        if self.subject1 not in OneKeyin.dictSubject:
            return ''
        return OneKeyin.dictSubject[self.subject1]
    @property
    def subjectNumber(self)->t.Tuple[str,str]:
        r1 = self.subject1str
        if r1 == '主日奉獻':
            return '4111000', 'A01'
        elif r1 == '什一奉獻':
            return '4112000', 'A01'
        elif r1 == '感恩奉獻':
            return '4113000', 'A01'
        elif r1 == '代轉奉獻':
            return '代轉', ''
        elif r1 == '神學培育':
            return '4314100', 'A04'
        elif r1 == '建堂修繕':
            return '4312100', 'A02'
        elif r1 == '宣教事工':
            return '4313100', 'A03'
        elif r1 == '愛心救助':
            return '4315100', 'A05'
        elif r1 == '特別奉獻':
            r2 = self.subject2str
            if r2 == '購地':
                return '4114000', 'A02'
            elif r2 == '學生中心':
                return '4114000', 'B01'
            else:
                return '4114000', 'A01'
        elif r1 == '其他':
            return '4131000', 'A01'
        else:
            return '代轉', ''