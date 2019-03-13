# -*- coding: utf-8 -*-
# @Time    : 3/11/2019 21:46
# @Author  : MARX·CBR
# @File    : csv 转 排序后的 csv.py

import csv
import os
import sys
import time
import xlrd
import xlwt


class marx:
    def __init__(self, cn):
        self.mydict = {}
        self.csv_name = cn
        # self.xlsx_name = xn

    def csv_to_xlsx(self):
        with open(self.csv_name, 'r', encoding='gbk') as f:
            read = csv.reader(f)
            workbook = xlwt.Workbook()
            sheet = workbook.add_sheet('data')  # 创建一个sheet表格
            l = 0
            for line in read:
                print(line)
                r = 0
                for i in line:
                    print(i)
                    sheet.write(l, r, i)  # 一个一个将单元格数据写入
                    r = r + 1
                l = l + 1

            workbook.save('Temp.xlsx')  # 保存Excel

    def transform_to_csv(self):
        workbook = xlrd.open_workbook('Temp.xlsx')
        table = workbook.sheet_by_index(0)

        with open(self.csv_name, 'w', encoding='utf-8') as f:
            f.write('文件名,最后访问时间,文件偏移位置,文件数据区未压缩时大小,文件数据区压缩后大小,zip文件大小\n')
            write = csv.writer(f)
            for row_num in range(table.nrows):
                row_value = table.row_values(row_num)
                mytemp=','.join(row_value)
                mytemp+='\n'
                # write.writerow(row_value)
                f.write(mytemp)

    def write_new_data(self, dic):
        with open('Temp.xlsx', 'w+', encoding='utf-8') as f:
            # read = csv.reader(f)
            workbook = xlwt.Workbook()
            sheet = workbook.add_sheet('data')  # 创建一个sheet表格
            l = 0
            mydic = dict(dic.copy())
            for key in mydic:
                r = 0
                content = mydic[key]
                lines = content.split("|^|")
                print(lines)
                for i in lines:
                    print(i)
                    sheet.write(l, r, i)  # 一个一个将单元格数据写入
                    r = r + 1
                l = l + 1

            workbook.save('Temp.xlsx')  # 保存Excel

    def mysort(self):
        # 按key排序
        dic = self.mydict
        dict = sorted(dic.items(), key=lambda d: int(d[0]))
        print(dict)
        return dict

    def read_data(self, name='Temp.xlsx'):
        workbook = xlrd.open_workbook(name)
        sheet_name = workbook.sheet_names()
        print(sheet_name)

        sheet = workbook.sheet_by_index(0)  # 选择第一张sheet
        print(sheet.nrows)
        print(sheet.ncols)
        for row in range(1, sheet.nrows):  # 第一个for循环遍历所有行
            print("reading:", row)
            content = ""
            number = ""
            for col in range(sheet.ncols):
                # print("%7s" % sheet.row(row)[col].value, '\t', end='')
                content += (sheet.row(row)[col].value + "|^|")
                if col == 2:
                    #排第三列
                    number = sheet.row(row)[col].value
            # print(number,content)
            t = {number: content}
            self.mydict.update(t)


start = time.time()
if __name__ == '__main__':
    print('Automatically Convert CSV To Sorted CSV\n Input: BC.exe xxxx.csv')
    if len(sys.argv) == 1:
        print('Please input parameter')
        sys.exit()
    else:
        cn = sys.argv[1]
        # xn = sys.argv[2]
        # run = marx(cn=cn, xn=xn)
        run = marx(cn=cn)
        run.csv_to_xlsx()
        run.read_data()
        new_data = run.mysort()
        run.write_new_data(new_data)
        run.transform_to_csv()
        os.remove('Temp.xlsx')
        # with open('t.xlsx','w+') as f:
        #     f.write('Temp')
        #     f.close()
        print('End of conversion thanks for using')
