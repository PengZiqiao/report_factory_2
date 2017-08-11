from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import numpy as np
from time import sleep
from winsun.date import Week
from winsun.func import percent
from winsun.tools import GisSpider, QUANSHI_BUHAN_LIGAO


class GetData:
    """
    目前依赖winsun.tools.GisSpider从gis系统爬数据，将来可能改写
    """

    def __init__(self):
        self.gis = GisSpider()
        sleep(3)

    def trend_gxj(self, wuye):
        """周度供销走势"""
        # 选项
        option = {
            # 周度供销
            'by': 'week',
            # 选择物业类型
            'usg': wuye,
            # 板块为全市不含溧水高淳
            'plate': QUANSHI_BUHAN_LIGAO,
            # 查询10周的数据
            'period': 10
        }
        df = self.gis.trend_gxj(**option)

        # 日期
        w = Week()
        wN = w.N
        wList = list(w.history(i) for i in range(10))
        wList.reverse()

        # x轴标签为“第n周”
        # df.index = list(f'第{i}周' for i in range(wN - 9, wN + 1))

        # x轴标签为日期
        df.index = list(f'{each[0].strftime("%m/%d")}-{each[1].strftime("%m/%d")}' for each in wList)
        # 系列名称
        df.columns = ['上市面积(万㎡)', '成交面积(万㎡)', '成交均价(元/㎡)']

        return df

    def plate_gxj(self, wuye):
        """当周供销价"""
        option = {
            # 周度供销
            'by': 'week',
            # 选择物业类型
            'usg': wuye,
            # 板块为全市不含溧水高淳
            'plate': QUANSHI_BUHAN_LIGAO,
        }

        df = self.gis.current_gxj(**option)
        # 删除合计行
        df = df[:-1]
        # 填充空行，修改列名
        index = ['城中', '城东', '城南', '河西', '城北', '仙林', '江宁', '浦口', '六合']
        df = pd.concat([pd.DataFrame(index=index), df], axis=1)
        # df.iloc[:,:2] = df.iloc[:,:2].fillna(0)
        df.columns = ['上市面积(万㎡)', '成交面积(万㎡)', '成交均价(元/㎡)']
        return df.reindex(index)

    def rank(self, wuye):
        """排行榜
        :param wuye: 物业类型
        :return: df组成的列表，分别为面积、套数、金额、均价排行榜
        """
        return self.gis.rank(by='week', usg=wuye, plate=QUANSHI_BUHAN_LIGAO)

    def zhuzhai_rank(self):
        dfs = self.rank(['住宅'])
        # 遍历面积、套数榜
        for i in [0, 1]:
            df = dfs[i]
            # 替换仙西为仙林
            df = df.apply(lambda x: x.replace('仙西', '仙林'))
            # 0:案名,1:推广名,2:板块,3:面积,4:套数,5:金额,6:均价
            # i + 3， i=0时为3即面积，i=1时为4即套数
            df = df.iloc[:, [1, 2, i + 3]]
            # 在最左插入一列‘排名’列
            df.insert(0, '排名', df.index)
            dfs[i] = df
        return dfs[:2]

    def rank_shuoli(self, wuye):
        # 获得面积排行榜
        df = self.gis.rank(by='week', usg=wuye, plate=QUANSHI_BUHAN_LIGAO)[0]
        # 替换仙西为仙林
        df = df.apply(lambda x: x.replace('仙西', '仙林'))
        # 找出排名前3
        text = ''
        for i in range(3):
            # 0:案名,1:推广名,2:板块,3:面积,4:套数,5:金额,6:均价
            name, plate, area, sets, price = df.iloc[i, [1, 2, 3, 4, 6]]
            # 面积除以10000后保留两位小数
            text += f'{plate}{name}（{round(area/1e4,2)}万㎡，{sets}套，{price}元/㎡）、'
        # 最后一个顿号改成句号
        return text[:-1] + '。'

    def df_liangjia(self, wuye):
        dfs = dict()

        print(f'>>> 正在查询{wuye}走势数据...')
        dfs['走势'] = self.trend_gxj(wuye)

        print(f'>>> 正在查询{wuye}板块数据...')
        dfs['板块'] = self.plate_gxj(wuye)

        return dfs

class ShuoLi():
    """
    df 传入一个供销走势的DataFrame
    ss, cj, jj 分别代表上市、成交、均价
    ss_, cj_, jj_ 加下划线代表对应环比
    """

    def __init__(self, df, degree=0):
        # 计算环比只需要近2周的数据
        df = df[-2:]
        # 计算出环比
        huanbi = df.pct_change()
        # 调整
        self.df = pd.concat([df.iloc[-1:], huanbi.iloc[-1:]])
        self.df.index = ['value', 'change']
        self.df.columns = ['ss', 'cj', 'jj']

    def shangshi(self):
        value = self.df.at['value', 'ss']
        change = self.df.at['change', 'ss']
        # 上市量为0或为空
        if value == 0 or np.isnan(value):
            return '无上市'
        # 环比值为空
        elif np.isnan(change):
            return f'上市{value:.2f}万㎡'
        else:
            return f'上市{value:.2f}万㎡，环比{percent(change,0)}'

    def chengjiao(self):
        value = self.df.at['value', 'cj']
        change = self.df.at['change', 'cj']
        # 成交量为0或为空
        if value == 0 or np.isnan(value):
            return '无成交'
        # 环比值为空
        elif np.isnan(change):
            return f'成交{value:.2f}万㎡'
        else:
            return f'成交{value:.2f}万㎡，环比{percent(change,0)}'

    def junjia(self):
        value = self.df.at['value', 'jj']
        change = self.df.at['change', 'jj']

        # 没有成交即没有成交均价
        v_cj = self.df.at['value', 'cj']
        if v_cj == 0 or np.isnan(v_cj):
            return None
        # 环比值为空
        elif np.isnan(change):
            return f'成交均价{value:.0f}元/㎡'
        else:
            return f'成交均价{value:.0f}元/㎡，环比{percent(change,0)}'

    def all(self):
        ss = self.shangshi()
        cj = self.chengjiao()
        jj = self.junjia()
        return f'{ss}；{cj}，{jj}。'


class Excel:
    def __init__(self, input_file='template.xlsx'):
        self.wb = load_workbook(input_file)

    def df2sheet(self, df, sheet_name):
        """将df贴到指定sheet上
        :param df: DataFrame对象
        :param sheet_name:sheet名称
        """
        ws = self.wb.get_sheet_by_name(sheet_name)
        for row in dataframe_to_rows(df, header=False):
            ws.append(row)


    def save(self, output_name='data.xlsx'):
        self.wb.save(output_name)
