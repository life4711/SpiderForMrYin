#-*- coding: UTF-8 -*-
import os
from xlutils.copy import copy
import xlrd as ExcelRead
import urllib
import urllib2
import requests
import re
import codecs
import time
import xlwt

#作用就是将html文件里的一些标签去掉，只保留文字部分
class HTML_Tool:
    BgnCharToNoneRex = re.compile("(\t|\n| |<a.*?>|<img.*?>)")
    EndCharToNoneRex = re.compile("<.*?>")
    BgnPartRex = re.compile("<p.*?>")
    CharToNewLineRex = re.compile("<br/>|</p>|<tr>|<div>|</div>")
    CharToNextTabRex = re.compile("<td>")

    def Replace_Char(self,x):
        x=self.BgnCharToNoneRex.sub("",x)
        x=self.BgnPartRex.sub("\n   ",x)
        x=self.CharToNewLineRex.sub("\n",x)
        x=self.CharToNextTabRex.sub("\t",x)
        x=self.EndCharToNoneRex.sub("",x)
        return x

#爬虫类
class Spider:
    def __init__(self,headers,id,num_start,num_end):
        self.id = id
        self.num_start = num_start
        self.num_end = num_end+1
        self.headers = headers
        self.s = requests.session()
        self.login = ""
        self.datas = []
        self.myTool = HTML_Tool()
        #s = time.strftime("%Y-%m-%d(%H_%M_%S)")
        self.file_name = 'huikan.xls'
        print u'会刊网爬虫程序已启动，正在开始加载...'

    def Solve(self):
        self.get_data()
        print u'爬虫报告：会刊目录已保存至当前文件夹下"huikan.xls"文件'
        print u'加载完成，按Enter键退出爬虫程序'
        raw_input()

    def wExcel(self):
        r_xls = ExcelRead.open_workbook(self.file_name)
        r_sheet = r_xls.sheet_by_index(0)
        rows = r_sheet.nrows
        w_xls = copy(r_xls)
        sheet_write = w_xls.get_sheet(0)
        for i in range(len(self.datas)):
            for j in range(len(self.datas[i])):
                sheet_write.write(rows + i, j, self.datas[i][j].decode('utf8'))
        w_xls.save(self.file_name)

    def get_data(self):
        for i in range(self.num_start,self.num_end):
            print u'爬虫报告：第%d页正在加载...' %i
            afterURL = "http://www.huikan.net/viphuikan.asp?scck=1&id="
            response = self.s.get(afterURL+str(self.id)+"&page="+str(i),headers = headers)
            mypage = response.content
            self.deal_data(mypage.decode('utf-8'))
            self.wExcel()
            #print mypage

    def deal_data(self,mypage):
        myItems = re.findall('<table cellpadding="0" cellspacing="0" style="background:#EEF6FF">.*?<td height="30" width="100%">&nbsp;(.*?)&nbsp;&nbsp;.*?<td height="30" width="100%">&nbsp;(.*?)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(.*?)</td>.*?<td height="30" width="100%">&nbsp;(.*?)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(.*?)</td>.*?<td height="30" width="100%">&nbsp;.*?>(.*?)</SPAN>.*?target="_blank">(.*?)</a>.*?<td height="30" width="100%">&nbsp;(.*?)</td>.*?</table>',mypage,re.S)
        #print myItems
        for item in myItems:
            items = []
            #print item[0]+' '+item[1]+' '+item[2]+' '+item[3]+' '+item[4]+' '+item[5]+' '+item[6]+' '+item[7]
            items.append(item[0].replace("\n",""))
            items.append(item[1].replace("\n",""))
            items.append(item[2].replace("\n", ""))
            items.append(item[3].replace("\n", ""))
            items.append(item[4].replace("\n", ""))
            items.append(item[5].replace("\n", ""))
            items.append(item[6].replace("\n", ""))
            items.append(item[7].replace("\n", ""))
            #print items
            mid = []
            for i in items:
                mid.append(self.myTool.Replace_Char(i.replace("\n","").encode('utf-8')))
            #print mid
            self.datas.append(mid)
        #print self.datas

if __name__ == "__main__":
    #----------- 程序的入口处 -----------
    print u"""
    *-----------------------------------------------------------------------------------------------
    * 程序：网络爬虫
    * 版本：V03
    * 作者：lvshubao
    * 日期：2016-05-08
    * 语言：Python 2.7
    * 功能：将会刊网中指定ID的会刊参展人信息以文件追加写入的方式保存到当前目录下的"huikan.xls"文件
    * 操作：按Enter键开始执行程序，加载时间可能会很长，请耐心等待
    * 提示：1、程序运行之初请保证当前目录下存在"huikan.xls"文件，并保证没有其他程序正在使用该文件
    *       2、程序运行过程中需要连接互联网，期间允许中断，中断后请按当前执行进度重新指定起止页码
    *
    *-----------------------------------------------------------------------------------------------
    """
    raw_input()
    headers = {'Cache-Control': 'max-age=0','Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8','Origin': ' http://www.huikan.net','Upgrade-Insecure-Requests': '1','User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.110 Safari/537.36','Accept-Language': 'zh-CN,zh;q=0.8','Cookie': 'pgv_pvi=7133927424; news=2014%u4E2D%u56FD%u56FD%u9645%u4FE1%u606F%u901A%u4FE1%u5C55%u89C8%u4F1A%A0%u7B2C1%u9875%204%u67083%u65E5%2012%3A52^http%3A//www.huikan.net/viphuikan.asp%3Fscck%3D1%26id%3D2588%26page%3D1$2014%u4E2D%u56FD%u56FD%u9645%u4FE1%u606F%u901A%u4FE1%u5C55%u89C8%u4F1A%A0%u7B2C1%u9875 4月3日 12:53^http%3A//www.huikan.net/viphuikan.asp%3Fscck%3D1%26id%3D2588%26page%3D1$|; ASPSESSIONIDSAQTDTSB=FHCEDBJAHJGPCOKONMNLEOGH; ASPSESSIONIDQATQARQC=DGKDMPABKDHMOHEAKJCLLJFB; pgv_si=s1220992000; PPst%5FLevel=3; PPst%5FUserName=handwudi; ehk%5Fname=handwudi; PPst%5FLastIP=222%2E171%2E12%2E123; PPst%5FLastTime=2016%2F4%2F10+18%3A27%3A38; visits=18; AJSTAT_ok_pages=2; AJSTAT_ok_times=4'}
    print u'请输入会刊ID:'
    id = raw_input()
    print u'请输入需爬取的起始页码：'
    num_start = raw_input()
    print u'请输入需爬取的截止页码：'
    num_end = raw_input()
    MySpider = Spider(headers,int(id),int(num_start),int(num_end))
    MySpider.Solve()