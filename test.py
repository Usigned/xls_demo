from typing import List
import pandas as pd


DEFAULT_COL_VAL = {
    '申报部门': '研发部',
    '奖励金额（元）': '1000',
    '奖励类型': '季度安全奖',
    '季度': '一季度',
    '人员所在部门': '研发部'
}

DEFAULT_RESULT_COLS = ['申报部门', '人员所在部门', '人员', '次数', '奖励类别', '申请事项简题', '奖励金额（元）', '奖励类型', '季度', '分类']


def process_data(df: pd.DataFrame, name_dict, in_cols=['奖励对象名单', '申请事项简题', '奖励类别细分'],
                 name_split_fn=lambda x: x.split('：')[-1].split('、')):
    '''
    in_cols的第一个元素应为奖励对象名单的列名
    '''
    result = []
    for names, *cols in df[in_cols].values:
        for name in name_split_fn(names):
            result.append([name_dict[name], *cols])
    return result

def generate_name_dict(df: pd.DataFrame, name_col='姓名', type_col='性质'):
    name_dict = {}
    for name, type in zip(df[name_col], df[type_col]):
        name = name.split('（')[0]
        name_dict[name] = f'{name}（{type}）'
    return name_dict


def wrap_result_list(result: List, col_names=[], col_val_dict=DEFAULT_COL_VAL):
    pass


if __name__ == '__main__':
    sheet_names = ['安全生产突出贡献奖', '安全生产管理奖']
    df1 = pd.read_excel(r'C:\Users\84371\Desktop\playground\xls_demo\研发部一季度申报 - 副本.xlsx', sheet_name='安全生产突出贡献奖')
    df2 = pd.read_excel(r'C:\Users\84371\Desktop\playground\xls_demo\研发部一季度申报 - 副本.xlsx', sheet_name='安全生产管理奖')
    p_df = pd.read_excel(r'C:\Users\84371\Desktop\playground\xls_demo\员工基础信息表1.xlsx', sheet_name='基础信息表')

    print(process_data(df1, generate_name_dict(p_df)))

    # name_dict = {}

    # for name, type in zip(p_df['姓名'], p_df['性质']):
    #     name = name.split('（')[0]
    #     name_dict[name] = f'{name}（{type}）'


    # result = []

    # for names, title, type in zip(df1['奖励对象名单'], df1['申请事项简题'], df1['奖励类别细分']):
    #     names = names.split('：')[-1].split('、')
    #     for name in names:
    #         result.append([name_dict[name], title, type, '安全生产突出贡献奖'])

    # for names, title, type in zip(df2['奖励对象名单'], df2['申请事项简题'], df2['奖励类别']):
    #     names = names.split('：')[-1].split('、')
    #     for name in names:
    #         result.append([name_dict[name], title, type, '安全生产管理奖'])


    # result = sorted(sorted(result, key=lambda x: x[0]), key=lambda x: x[0].split('（')[-1], reverse=True)

    # result_df = pd.DataFrame(result, columns=['人员', '申请事项简题', '奖励类别', '分类'])

    # count_df = result_df['人员'].value_counts().sort_values(ascending=False)