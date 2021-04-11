# coding=utf-8
import argparse

import pandas as pd


class ExcelHandler:
    def __init__(self,old_file,new_file,sheet_name=0):
        self.old_file = old_file
        self.new_file = new_file
        self.sheet_name = sheet_name

    def read_data(self):
        """"
            函数功能：获取两个excel表里的数据
            函数参数：第一个为旧表的数据，第二个为新表的数据用于更新新表
            返回值：返回一个元祖（包含旧表的数据和新表的数据）
            """
        old_data = pd.read_excel(self.old_file, self.sheet_name)
        new_data = pd.read_excel(self.new_file, self.sheet_name)
        return (old_data, new_data)

    def extend_data(self,olddata, newdata, key_index,keep="last"):
        """
        该函数用于通过新数据更新旧数据表的场景：
        1）新旧数据的主键有相同索引时，用新数据替换旧有数据
        2）当某一个索引只存在于新数据中，则将该索引的数据添加到旧数据中
        :param olddata: 旧数据
        :param newdata: 新数据
        :param key_index: 指定的主键
        :return:返回更新后的数据
        """
        if key_index in olddata.columns.values:
            extend_result = pd.concat([olddata, newdata], keys=['a', 'b']).drop_duplicates(subset=key_index,
                                                                                           ignore_index=True,
                                                                                           keep=keep)
            return extend_result
        else:
            raise ValueError("%s 不是主键"%key_index)

def get_argparses():
    parse = argparse.ArgumentParser(description="This is a tool to deal data of excel.")
    parse.add_argument("-old",type=argparse.FileType(),metavar="old_file",help="指定旧excel文件路径",dest="old_file")
    parse.add_argument("-new",type=argparse.FileType(),metavar="new_file",help="指定新excel文件路径",dest="new_file")
    parse.add_argument("-key",metavar="key_index",help="指定一个主键",dest="key_index")
    parse.add_argument("--keep",choices=['first','last'],default='last',help="指定旧excel文件路径",dest="keep")

    return parse.parse_args()

def main():
    # 获取命令行参数
    parse_arg = get_argparses()
    old_file = parse_arg.old_file.name
    new_file = parse_arg.new_file.name
    for path in (old_file,new_file):
        if path.rpartition(".")[-1] not in ("xls","xlsx"):
            raise ValueError("%s not is excel file"%path)
    keep = parse_arg.keep
    key_index = parse_arg.key_index
    # 开始处理数据
    handler = ExcelHandler(old_file,new_file)
    result = handler.read_data()
    result_update = handler.extend_data(result[0],result[1],key_index,keep=keep)
    # print(result_update) #保存数据
    result_update.to_excel("result_update.xlsx",index=False)

if __name__ == '__main__':
    main()