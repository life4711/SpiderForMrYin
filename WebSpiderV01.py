#-*- coding: UTF-8 -*-
import urllib
import urllib2
import requests
import re
import codecs
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
    def __init__(self,parll,headerss,num):
        self.num = num+1
        self.par1 = parll
        self.headers = headerss
        self.s = requests.session()
        self.login = ""
        self.datas = []
        self.myTool = HTML_Tool()
        print u'会刊网爬虫程序已启动,正尝试登陆...'

    def Solve(self):
        self.Login()
        self.get_data()
        # self.save_data(3)
        print u'爬虫报告：加载全部完成，正在写入xls表格...'
        self.wExcel('huikan.xls')

    def wExcel(self, outEfile):
        w = xlwt.Workbook()
        sheet = w.add_sheet('sheet1')
        for i in range(len(self.datas)):
            for j in range(len(self.datas[i])):
                # print buf[i]
                sheet.write(i, j, self.datas[i][j].decode('utf8'))
        w.save(outEfile)
        print u'爬虫报告：会刊目录已保存至当前文件夹下huikan.xls文件'
        print u'加载完成，按任意键退出爬虫程序'
        raw_input()
    # def save_data(self,endpage):
    #     f = codecs.open('huikan.xls','w+', "utf-8")
    #     for i in self.datas:
    #         str = u""
    #         for j in i:
    #             str=str+j.decode('utf-8')
    #         f.writelines(str+u'\n')
    #     f.close()
    #     print u'爬虫报告：会刊目录已保存至当前文件夹下huikan.xls文件'
    #     print u'加载完成，按任意键退出爬虫程序'
    #     #raw_input()

    def Login(self):
        loginURL = "http://www.huikan.net/Login.asp?Action=Login"
        self.login = self.s.post(loginURL,data = self.par1,headers = self.headers)
        print u'爬虫报告：模拟浏览器登陆成功，正在爬取资料...'

    def get_data(self):
        for i in range(1,self.num):
            print u'爬虫报告：爬虫%d号正在加载中...' %i
            afterURL = "http://www.huikan.net/allehuikan.asp?page="
            response = self.s.get(afterURL+str(i), cookies = self.login.cookies,headers = headers)
            mypage = response.content
            self.deal_data(mypage.decode('utf-8'))
            #print mypage

    def deal_data(self,mypage):
        myItems = re.findall('<tr>.*?<td height="30" class="tdcs" align="center" width="70">(.*?)</td>.*?<td height="30" class="tdcs" align="center" width="90">(.*?)</td>.*?<td height="9" align="center" class="tdcs" width="50">(.*?)</td>.*?<td height="9" class="tdcs">.*?&nbsp;(.*?)</a></td>.*?<td height="9" class="tdcs" width="170" align="center">&nbsp;(.*?)</td>.*?<td height="9" align="center" class="tdcs" width="100">(.*?)</td>.*?<td height="9" class="tdcs" width="40" align="center">(.*?)</td>.*?</tr>',mypage,re.S)
        #print myItems
        for item in myItems:
            items = []
            #print item[0]+' '+item[1]+' '+item[2]+' '+item[3]+' '+item[4]+' '+item[5]+' '+item[6]
            items.append(item[0].replace("\n",""))
            items.append(item[1].replace("\n",""))
            items.append(item[2].replace("\n", ""))
            items.append(item[3].replace("\n", ""))
            items.append(item[4].replace("\n", ""))
            items.append(item[5].replace("\n", ""))
            items.append(item[6].replace("\n", ""))
            #print items
            mid = []
            for i in items:
                mid.append(self.myTool.Replace_Char(i.replace("\n","").encode('utf-8')))
            #print mid
            self.datas.append(mid)
        #print self.datas

#----------- 程序的入口处 -----------
print u"""
----------------------------------------------------------
   程序：网络爬虫
   版本：V01
   作者：lvshubao
   日期：2016-04-11
   语言：Python 2.7
   功能：将会刊网中会刊目录保存到当前目录下的huikan.xls表格
   操作：按任意键开始执行程序，加载时间可能会很长，请耐心等待
----------------------------------------------------------
"""
raw_input()
headers = {'Cache-Control': 'max-age=0','Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8','Origin': ' http://www.huikan.net','Upgrade-Insecure-Requests': '1','User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.110 Safari/537.36','Accept-Language': 'zh-CN,zh;q=0.8','Cookie': 'pgv_pvi=7133927424; news=2014%u4E2D%u56FD%u56FD%u9645%u4FE1%u606F%u901A%u4FE1%u5C55%u89C8%u4F1A%A0%u7B2C1%u9875%204%u67083%u65E5%2012%3A52^http%3A//www.huikan.net/viphuikan.asp%3Fscck%3D1%26id%3D2588%26page%3D1$2014%u4E2D%u56FD%u56FD%u9645%u4FE1%u606F%u901A%u4FE1%u5C55%u89C8%u4F1A%A0%u7B2C1%u9875 4月3日 12:53^http%3A//www.huikan.net/viphuikan.asp%3Fscck%3D1%26id%3D2588%26page%3D1$|; ASPSESSIONIDSAQTDTSB=FHCEDBJAHJGPCOKONMNLEOGH; ASPSESSIONIDQATQARQC=DGKDMPABKDHMOHEAKJCLLJFB; pgv_si=s1220992000; PPst%5FLevel=3; PPst%5FUserName=handwudi; ehk%5Fname=handwudi; PPst%5FLastIP=222%2E171%2E12%2E123; PPst%5FLastTime=2016%2F4%2F10+18%3A27%3A38; visits=18; AJSTAT_ok_pages=2; AJSTAT_ok_times=4'}
data = {"Username":"handwudi","Passwrod":"x2008bzzs","B1":"登陆"}
print u'请输入爬虫需要爬取的页数:'
num = raw_input()
myspider = Spider(data,headers,int(num))
myspider.Solve()