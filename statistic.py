from typing import Dict, List
import pandas as pd
import argparse
import os

DEFAULT_COL_VAL = {
    '申报部门': '研发部',
    '奖励金额（元）': '1000',
    '奖励类型': '季度安全奖',
    '季度': '一季度',
    '人员所在部门': '研发部'
}

DEFAULT_RESULT_COLS = ['申报部门', '人员所在部门', '人员',
                       '奖励类别', '申请事项简题', '奖励金额（元）', '奖励类型', '季度', '分类']

DEFAULT_VAL_MAP = {
    '安全生产突出贡献奖': '突出贡献奖',
    '安全生产管理奖': '生产管理奖',
}


def process_data(df: pd.DataFrame, name_dict, in_cols=['奖励对象名单', '申请事项简题', '奖励类别细分'],
                 out_cols=['人员', '申请事项简题', '奖励类别'],
                 name_split_fn=lambda x: x.split('：')[-1].split('、')):
    '''
    in_cols的第一个元素应为奖励对象名单的列名
    out_cols和in_cols数量应该一致
    '''
    result = []
    for names, *cols in df[in_cols].values:
        for name in name_split_fn(names):
            name = name.strip()
            if name not in name_dict:
                print(name, *cols)
            else:
                result.append([name_dict[name], *cols])
    return pd.DataFrame(result, columns=out_cols)


def generate_name_dict(df: pd.DataFrame, name_col='姓名', type_col='性质'):
    name_dict = {}
    for name, type in zip(df[name_col], df[type_col]):
        name = name.split('（')[0]
        name_dict[name] = f'{name}（{type}）'
    return name_dict


def _defaul_post_process_fn(result: pd.DataFrame):
    result = result.sort_values(
        by='人员', 
        key=lambda x: [r[-1].strip()[:-1]+r[0] for r in x.str.split('（')], 
        ascending=False)
    result.index = [i for i in range(1, len(result)+1)]
    return result


def wrap_result_list(results: List[pd.DataFrame], col_names=DEFAULT_RESULT_COLS, col_val_dict: Dict = DEFAULT_COL_VAL, val_map=DEFAULT_VAL_MAP, post_process_fn=_defaul_post_process_fn) -> pd.DataFrame:
  
    # 0. stack every df
    result = pd.concat(results)
    # 1. add default col
    for col in col_names:
        if col not in result.columns:
            result[col] = col_val_dict[col]
    # 2. val mapping
    result.replace(val_map, inplace=True)
    return post_process_fn(result)[col_names]


def sheet_process(sheet_name, wrap_fn, df: pd.DataFrame, *args, **kwargs):
    df = process_data(df, *args, **kwargs)
    return wrap_fn(df, sheet_name)


def _wrap_fn(df, sheet_name):
    df['分类'] = sheet_name
    return df


def _c_name_split_fn(names: str):
    lines = [line for line in names.split('\n') if '研发部' in line]
    result = []
    for line in lines:
        result += line.split('：')[-1].split('、')
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--data', default=os.path.join(os.path.abspath(os.path.curdir), '公司2024年一季度安全奖申报表-研发部.xlsx'), help='data file path')
    parser.add_argument('-cd', '--cdata', default=os.path.join(os.path.abspath(os.path.curdir), '公司2024年一季度安全奖申报表.xls'), help='data file path')
    parser.add_argument('-p', '--person', default=os.path.join(os.path.abspath(os.path.curdir), '员工基础信息表.xlsx'), help='person file path')
    parser.add_argument('-o', '--output', help='output file path')
    args = parser.parse_args()
    sheet_names = ['安全生产突出贡献奖', '安全生产管理奖']
    in_cols_lst = [['奖励对象名单', '申请事项简题', '奖励类别细分'],
                   ['奖励对象名单', '申请事项简题', '奖励类别']]
    d_fpath = args.data
    cd_fpath = args.cdata
    p_fpath = args.person
    name_dict = generate_name_dict(pd.read_excel(p_fpath, sheet_name='基础信息表'))
    dfs = [sheet_process(
        sheet_name=sheet_name,
        wrap_fn=_wrap_fn,
        df=pd.read_excel(d_fpath, sheet_name=sheet_name),
        name_dict=name_dict,
        in_cols=in_cols,
    ) for sheet_name, in_cols in zip(sheet_names, in_cols_lst)]
    cdfs = [sheet_process(
        sheet_name=sheet_name,
        wrap_fn=_wrap_fn,
        df=pd.read_excel(cd_fpath, sheet_name=sheet_name)[:-1],
        name_dict=name_dict,
        in_cols=in_cols,
        name_split_fn=_c_name_split_fn
    ) for sheet_name, in_cols in zip(sheet_names, in_cols_lst)]
    if args.output:
        with pd.ExcelWriter(args.output, mode='w') as writer:
            dfs = wrap_result_list(dfs)
            cdfs = wrap_result_list(cdfs)
            dfs.to_excel(writer, index_label='序号', sheet_name='部门表')
            cdfs.to_excel(writer, index_label='序号', sheet_name='公司表')
            counts = pd.concat([dfs['人员'], cdfs['人员']]).value_counts().sort_values(ascending=False)
            counts.name = '次数'
            counts.to_excel(writer, index_label='人员', sheet_name='次数')
    else:
        print(wrap_result_list(dfs).to_csv())
        print(wrap_result_list(cdfs).to_csv())