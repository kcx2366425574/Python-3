#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
@File       : xlwt_util.py
@Time       :2021/5/6 16:59
@Author     :kuang congxian
@Contact    :kuangcx@inspur.com
@Description : 写入数据，格式为xls，尽量不使用
"""
import ast
import xlwt


def json_to_excel(json_path):
    with open(json_path, "r") as f:
        data = ast.literal_eval(f.read())
        # 创建新的workbook（其实就是创建新的excel）
        workbook = xlwt.Workbook(encoding='ascii')

        # 创建新的sheet表
        worksheet = workbook.add_sheet("My new Sheet")

        # 往表格写入标题
        i = 0
        for key in data[0].keys():
            worksheet.write(0, i, key)
            i += 1

        if type(data) == list:
            i = 1
            for d in data:
                j = 0
                for key, value in d.items():
                    worksheet.write(i, j, value)
                    j += 1
                i += 1

        # 保存
        workbook.save("data.xls")


if __name__ == '__main__':
    json_to_excel("data.json")
