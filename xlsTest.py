#-*- coding: UTF-8 -*-
import os 
from xlutils.copy import copy 
import xlrd as ExcelRead
import xlwt
import urllib
import urllib2
import requests
import re
import codecs
import time

def wwExcel(file_name):
    datas = [["Ann", "woman", 21, "UK"],["Ann", "woman", 22, "UK"],["Ann", "woman", 23, "UK"]]
    r_xls = ExcelRead.open_workbook(file_name)
    r_sheet = r_xls.sheet_by_index(0)
    rows = r_sheet.nrows
    w_xls = copy(r_xls)
    sheet_write = w_xls.get_sheet(0)
    # num = 5
    # for ii in range(2, num):
    #     print ii
    for i in range(len(datas)):
        for j in range(len(datas[i])):
            sheet_write.write(rows+i, j, datas[i][j])
    w_xls.save(file_name)
    print u'爬虫报告：会刊目录已保存至当前文件夹下"huikan.xls"文件'
    print u'加载完成，按Enter键退出爬虫程序'

if __name__ == "__main__":
    wwExcel("huikan.xls")