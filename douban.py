# -*- coding: UTF-8 -*-
import sys
from bs4 import BeautifulSoup #网页解析，获取数据
import re #正则表达式，进行文字匹配
import xlwt #进行excel操作
import urllib.request, urllib.error #制定url，获取网页数据

def main():
    #爬取网页
    baseurl="https://movie.douban.com/top250?start="
    datalist=getData(baseurl)
    savepath='豆瓣电影Top250.xls'
    saveData(datalist,savepath)
#影片详情链接的规则
findLink=re.compile(r'<a href="(.*?)">') #创建正则表达式对象，表示规则（字符串的模式）（字母r表示忽视所有特殊符合，如\）(.表示一个字符，*表示有多个字符)
#影片图片
findImgSrc=re.compile(r'<img.*src="(.*?)"',re.S) #re.S让换行符包含在字符中
#影片片名
findTitle=re.compile(r'<span class="title">(.*)</span>')
#影片评分
findRating=re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')
#找到评价人数
findJudge=re.compile(r'<span>(\d*)人评价</span>') #\d表示数字
#找到概况
findInq=re.compile(r'<span class="inq">(.*)</span>')
#找到影片的相关内容
findBd=re.compile(r'<p class="">(.*?)</p>',re.S)

def getData(baseurl):
    datalist=[]
    for i in range(0,10):
        url=baseurl+str(i*25)
        html=askURL(url)
        #开始逐一解析数据
        soup=BeautifulSoup(html,"html.parser")
        for item in soup.find_all('div',class_="item"): #查找符合要求的字符串，形成列表
            data=[] #保存一部电影的所有信息
            item=str(item) 
            link=re.findall(findLink,item)[0] #re库用来通过正则表达式查找指定的字符串
            data.append(link)
            imgSrc=re.findall(findImgSrc,item)[0] 
            data.append(imgSrc)
            titles=re.findall(findTitle,item)
            if len(titles)>=2: 
                ctitle=titles[0] #中文名
                data.append(ctitle)
                ftitle=titles[1].replace("/","") #去掉无关符号
                data.append(ftitle) #外文名
            else:
                data.append(titles[0])
                data.append(' ') #外文名留空
            rating=re.findall(findRating,item)[0]
            data.append(rating)
            judgeNum=re.findall(findJudge,item)[0]
            data.append(judgeNum)
            inq=re.findall(findInq,item)
            if len(inq)!=0: 
                inq = inq[0].replace("。","")
                data.append(inq)
            else:
                data.append(" ")
            bd=re.findall(findBd,item)[0]
            bd=re.sub('<br(\s+)?/>(\s+)?'," ",bd) #替换掉<br/>
            bd=re.sub('/'," ",bd)
            data.append(bd.strip()) #bd.strip()去掉前后空格
            datalist.append(data) #把处理好的一部电影信息放入datalist
    return datalist

def saveData(datalist,savepath):
    book=xlwt.Workbook(encoding="utf-8",style_compression=0) #创建workbook对象，理解为一个文件
    sheet=book.add_sheet('豆瓣电影top250',cell_overwrite_ok=True) #创建工作表
    col=("Link","Image","Chinese Name","Foreign Name","Rating","Judge Number","Info","Related")
    for i in range(0,8):
        sheet.write(0,i,col[i]) #第0行，第i列，写入col里第i个字符串
    for i in range(0,250):
        data = datalist[i] #第i个电影的data
        for j in range(0,8):
            sheet.write(i+1,j,data[j]) #第i+1行（第0行为表头），第j列，写入第i个电影的data的第j个信息

    book.save(savepath) #保存数据表


def askURL(url):
    head={"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36"}
    request=urllib.request.Request(url,headers=head)
    html=''
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode('utf-8')
        #print(html)
    except urllib.error.URLError as e:
        if hasattr(e,"code"):
            print(e.code)
        if hasattr(e,"reason"):
            print(e.reason)
    return html


if __name__ == '__main__': #加入main函数用于测试程序 #协调入口（从哪个代码开始执行）
    main()




    

 