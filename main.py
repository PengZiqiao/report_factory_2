from ppt import PPT
from get_data import GetData, ShuoLi, Excel
from winsun.date import Week
from winsun.tools import BIE_SHU


def chart_data(g):
    """
    :param g: GetData实例，获取数据逻辑改写后可改写此段
    :return: df_dict ，4 * 2 共8张表的字典
    """
    df = dict()
    excel = Excel()
    # 遍历4种物业类型
    for wuye in wuyes:
        # dfs 为走势、板块两张表
        dfs = g.df_liangjia(wuyes[wuye])
        # 遍历两张表
        for item in ['走势', '板块']:
            key = f'{wuye}{item}'
            # 分别写入df和excel文件
            df[key] = dfs[item]
            excel.df2sheet(df[key], key)
    excel.save()
    return df


def shuoli_data(g, dfs):
    shuoli = dict()
    for wuye in wuyes:
        s = ShuoLi(dfs[f'{wuye}走势'])
        if wuye == '住宅':
            shuoli[wuye] = f'本周{wuye}市场{s.shangshi()}。\r{s.chengjiao()}。\r{s.junjia()}。'
        elif wuye == '别墅':
            shuoli[wuye] = f'本周{wuye}市场{s.all()}\r补充说理。'
        else:
            print('>>> 正在查询{wuye}排行...')
            rank_info = g.rank_shuoli(wuye)
            shuoli[wuye] = f'本周{wuye}市场{s.all()}\r本周成交面积榜单前三：{rank_info}'
    return shuoli


class Report(PPT):
    # 日期
    w = Week()
    wN = w.N
    year = w.sunday.year
    monday = w.monday
    monday = f'{monday.month}月{monday.day}日'
    sunday = w.sunday
    sunday = f'{sunday.month}月{sunday.day}日'
    date_text = f'{year}年第{wN}周（{monday}-{sunday}）'

    def title(self):
        print('>>> 生成标题页...')
        # 标题
        txt = f'{self.year}年第{self.wN}周周报'
        self.text(txt, 0, 0)
        # 小标题
        txt = f'{self.year}年{self.monday}-{self.sunday}'
        self.text(txt, 0, 1)

    def liangjia(self, wuye, shuoli, page_idx):
        print(f'>>> 生成{wuye}量价页...')
        # 右图标题
        txt = f'{self.year}年第{self.wN}周南京（不含高淳溧水）{wuye}市场分板块供销量价'
        self.text(txt, page_idx, 3)
        # 说理
        self.text(shuoli[wuye], page_idx, 4)

    def paihang(self, dfs, page_idx):
        print(f'>>> 生成住宅排行页...')
        # 左表标题
        txt = f'{self.date_text}成交面积排行榜'
        self.text(txt, page_idx, 2)
        # 右表标题
        txt = f'{self.date_text}成交套数排行榜'
        self.text(txt, page_idx, 3)
        # 左表
        self.df2table(dfs['住宅面积榜'], page_idx, 5)
        # 右表
        self.df2table(dfs['住宅套数榜'], page_idx, 6)


if __name__ == '__main__':
    # 准备数据
    wuyes = {'住宅': ['住宅'], '商业': ['商业'], '办公': ['办公'], '别墅': BIE_SHU}
    g = GetData()
    dfs = chart_data(g)
    dfs['住宅面积榜'], dfs['住宅套数榜'] = g.zhuzhai_rank()
    shuoli = shuoli_data(g, dfs)

    # 生成ppt报告
    rpt = Report()
    rpt.title()
    rpt.liangjia('住宅', shuoli, 1)
    rpt.paihang(dfs, 2)
    rpt.liangjia('商业', shuoli, 3)
    rpt.liangjia('办公', shuoli, 4)
    rpt.liangjia('别墅', shuoli, 5)
    print('>>> 正在保存...')
    rpt.save(f'e:/周报测试/{rpt.date_text}周报.pptx')
    print('>>> 保存成功！')