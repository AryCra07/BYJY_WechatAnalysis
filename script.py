import pandas as pd
import os


class Content:

    def __init__(self, _type):
        """ 初始化 """
        self.type = _type  # 分析的内容类型： 可选（1 : 内容分析 2: 篇数分析）
        if self.type == 1:
            self.name = '内容分析'
        else:
            self.name = '篇数分析'
        self.file = None  # 合并后的文件

    def output(self, s):
        output_file = open(f'{self.name}结果分析.txt', "w")
        output_file.write(s)
        output_file.close()

    def merge_files(self):
        """ 合并文档 """
        # 获取当前脚本所在目录
        dir_path = os.path.dirname(os.path.realpath(__file__))

        # 获取source文件夹路径
        source_path = os.path.join(dir_path, f'src/{self.name}')

        # 获取source文件夹下所有xlsx文件路径
        xlsx_files = [os.path.join(source_path, f) for f in os.listdir(source_path) if f.endswith('.xls')]

        # 读取所有xlsx文件
        dfs = [pd.read_excel(f) for f in xlsx_files]

        # 合并所有文件
        merged_df = pd.concat(dfs, ignore_index=True)

        # 保存合并后的文件
        merged_df.to_excel(f'{self.name}汇总.xlsx', index=False)

        self.file = merged_df

    def analysis(self):
        """ 进行分析 """
        if self.type == 1:
            # 进行内容分析
            self.analysis_1()
        else:
            # 进行篇数分析
            self.analysis_2()

    def analysis_1(self):
        """ 进行内容分析 """
        row_count = self.file.shape[0]
        # 获取合并后的文件的所有列名
        columns = list(self.file.columns)
        # 进行统计
        columns = columns[1:-1]
        s = "统计信息：\n"
        for column in columns:
            s += f'{column}的总数为：{self.file[column].sum()}\n'

        for i in range(1, 4):
            s += f'日均{columns[i]}为：{self.file[columns[i]].sum() / row_count}\n'
        self.output(s)
        pass

    def analysis_2(self):
        """ 进行篇数分析 """
        # 获取合并后的文件的所有列名
        columns = list(self.file.columns)
        # 进行统计
        columns = columns[2:-10]
        pd.set_option('display.max_columns', None)
        # pd.set_option('display.max_colwidth', 100)
        pd.set_option('display.max_rows', None)

        s = "统计信息：\n"
        s += "发表时间最早的文章是 \n" + str(
            self.file[self.file['发表时间'] == self.file["发表时间"].min()][["内容标题", "发表时间"]]) + "\n\n"

        s += "发表时间最晚的文章是 \n" + str(
            self.file[self.file['发表时间'] == self.file["发表时间"].max()][["内容标题", "发表时间"]]) + "\n\n"
        s += "阅读总数最多的文章是 \n" + str(
            self.file[self.file['总阅读次数'] == self.file["总阅读次数"].max()][["内容标题","发表时间"]]) + "\n\n"

        s += "总篇数为：" + str(self.file.shape[0]) + "\n"
        for column in columns:
            s += f'总数为 {column} 的总数为 : {self.file[column].sum()}'
        self.output(s)
        pass


if __name__ == '__main__':
    print("==================== 欢迎使用公众号快捷分析 ====================")
    print("使用说明如下：")
    print("将需要进行内容分析的文件放入到source/内容分析文件夹下面")
    print("将需要进行篇数分析的文件放入到source/篇数分析文件夹下面")
    print("请确保文件名不重复")
    print("结果将会保存在当前目录下")
    print("功能代码： 1 : 群发分析 2: 篇数分析")
    print("输入功能代码：")
    _type = int(input())
    content = Content(_type)
    print(f"==================== {content.name} ====================")
    content.merge_files()
    content.analysis()
    print("分析完成，结果保存在当前目录下")
    print("=========================================================")
