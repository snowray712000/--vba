from datetime import datetime


class OneDataOfAcc:
    def __init__(self):
        self.subject: str = None
        " 科目 e.g.: 411000 "
        self.subpoena: str = None
        """ JC20230101002 """
        self.department: str = None
        """A01 A02 A03 A04 A05 B01 B02"""
        self.memo: str = None
        """ 摘要 """
        self.money: int = None
        """ 金額 """
    def set(self,subject: str, subpoena: str, department: str, memo: str, money: int):
        self.subject = subject
        self.subpoena = subpoena
        self.department = department
        self.memo = memo
        self.money = money
    @property
    def date(self)->datetime:
        return datetime.strptime(self.subpoena[2:10], '%Y%m%d')
    def __repr__(self) -> str:
        return f'{self.subject} {self.money} {self.subpoena} {self.department} {self.memo}'
    def __str__(self) -> str:
        return self.__repr__()