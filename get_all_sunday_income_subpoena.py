#%%
# 輸出所有主日傳票，來比對用。

from OneDataOfAcc import OneDataOfAcc


import linque as lq


import typing as t


def get_all_sunday_income_subpoena(rows: t.List[OneDataOfAcc])-> t.List[str]:
    """ 回傳是可能是主日的傳票編號。 
    - Notes:
        - 通常是先呼叫 get_all_data_of_acc 得到 rows, 再呼叫此函式。        
    - Steps:
        - 1. 先找出所有 4 開頭的，它會是收入
        - 2. 主日，通常那張傳票，會同時有 4111000(主日奉獻) 和 4112000(什一奉獻)。所以要先 group by 傳票編號，再 where 每個 group。
        - 3. 符合2的 group，再 select 出傳票編號。
    - Return:
        - 例如 ['JC20230101002', 'JC20230101003'] 
    """
    r1 = lq.linque(rows).where(lambda a1: a1.subject[0] == '4')
    def fn_where(a1)->bool:
        r1: lq.Linque = a1[1]
        r2 = r1.select(lambda a1: a1.subject).to_list()
        # 一定要有主日，月定。
        if '4111000' not in r2:
            return False
        if '4112000' not in r2:
            return False

        return True
    r2 = r1.group(lambda a1: a1.subpoena).where(fn_where)
    return r2.select(lambda a1: a1[0]).to_list()