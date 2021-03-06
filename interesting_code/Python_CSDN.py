#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   csdn.py
@Time    :   2019/08/26 09:54:47
@Author  :   xiao ming 
@Version :   1.0
@Contact :   xiaoming3526@gmail.com
@Desc    :   None
@github  :   https://github.com/aimi-cn/AILearners
'''

# 导入相关爬虫库和解析xml库即可
import time
from pyquery import PyQuery as pq
import requests
from bs4 import BeautifulSoup
import random  
from fake_useragent import UserAgent 
from lxml import etree
import ssl 
ssl._create_default_https_context = ssl._create_unverified_context 

# 爬取csdn类
class ScrapyMyCSDN:
    ''' class for csdn'''
    def __init__(self,blogname):
        '''init 类似于构造函数 param[in]:blogname:博客名'''
        csdn_url = 'https://blog.csdn.net/' #常规csdnurl
        self.blogurl = csdn_url+blogname #拼接字符串成需要爬取的主页url

    ''' Func:获取写了多少篇原创文章 '''
    ''' return:写了多少篇原创文章'''
    def getOriginalArticalNums(self,proxies):
        main_response = requests.get(self.blogurl,proxies=proxies)
        # 判断是否成功获取 (根据状态码来判断)
        if main_response.status_code == 200:
            print('获取成功')
            self.main_html = main_response.text
            main_doc = pq(self.main_html)
            mainpage_str = main_doc.text() #页面信息去除标签信息
            origin_position = mainpage_str.index('原创') #找到原创的位置
            end_position = mainpage_str.index('原创',origin_position+1) #最终的位置,即原创底下是数字多少篇博文
            self.blog_nums = ''
            # 获取写的博客数目
            for num in range(3,10):
                #判断为空格 则跳出循环
                if mainpage_str[end_position + num].isspace() == True:
                    break
                self.blog_nums += mainpage_str[end_position + num]
            print(type(str(self.blog_nums)))
            cur_blog_nums = (int((self.blog_nums))) #获得当前博客文章数量
            return cur_blog_nums #返回博文数量
        else:
            print('爬取失败')
            return 0 #返回0 说明博文数为0或者爬取失败

    ''' Func：分页'''
    ''' param[in]:nums:博文数 '''
    ''' return: 需要爬取的页数'''
    def getScrapyPageNums(self,nums):
        self.blog_original_nums = nums
        if nums == 0:
            print('它没写文章，0页啊！')
            return 0
        else:
            print('现在开始计算')
            cur_blog = nums/20 # 获得精确的页码
            cur_read_page = int(nums/20) #保留整数
            # 进行比对
            if cur_blog > cur_read_page:
                self.blog_original_nums = cur_read_page + 1
                print('你需要爬取 %d'%self.blog_original_nums + '页')
                return self.blog_original_nums #返回的数字
            else:
                self.blog_original_nums = cur_read_page
                print('你需要爬取 %d'%self.blog_original_nums + '页')
            return self.blog_original_nums

    '''Func:开始爬取，实际就是刷浏览量hhh'''
    '''param[in]:page_num:需要爬取的页数'''
    '''return:0:浏览量刷失败'''
    def beginToScrapy(self,page_num,proxies):
        if page_num == 0:
            print('连原创博客都不写 爬个鬼!')
            return 0
        else:
            for nums in range(1,page_num+1):
                self.cur_article_url = self.blogurl + '/article/list/%d'%nums+'?t=1&'  #拼接字符串
                article_doc = requests.get(self.cur_article_url,proxies=proxies) #访问该网站
                # 先判断是否成功访问
                if article_doc.status_code == 200:
                    print('成功访问网站%s'%self.cur_article_url)
                    #进行解析
                    cur_page_html = article_doc.text
                    #print(cur_page_html)
                    soup = BeautifulSoup(cur_page_html,'html.parser')
                    for link in soup.find_all('p',class_="content"):
                        #print(link.find('a')['href'])
                        requests.get(link.find('a')['href'],proxies=proxies) #进行访问
                else:
                    print('访问失败')
        print('访问结束')

 
# IP地址取自国内髙匿代理IP网站：http://www.xicidaili.com/nn/  
  
#功能：爬取IP存入ip_list列表  
def get_ip_list(url, headers):  
    web_data = requests.get(url, headers=headers)  
    soup = BeautifulSoup(web_data.text, 'lxml')  
    ips = soup.find_all('tr')  
    ip_list = []  
    for i in range(1, len(ips)):  
        ip_info = ips[i]  
        tds = ip_info.find_all('td') #tr标签中获取td标签数据  
        if not tds[8].text.find('天')==-1:  
            ip_list.append(tds[1].text + ':' + tds[2].text)  
    return ip_list  
  
#功能：1,将ip_list中的IP写入IP.txt文件中  
#      2,获取随机IP，并将随机IP返回  
def get_random_ip(ip_list):  
    proxy_list = []  
    for ip in ip_list:  
        proxy_list.append(ip)  
        f=open('IP.txt','a+',encoding='utf-8')  
        f.write('http://' + ip)  
        f.write('\n')  
        f.close()  
    proxy_ip = random.choice(proxy_list)  
    proxies = {'http':proxy_ip}  
    return proxies  
  
if __name__ == '__main__':  
    # range(1,3)是页数 一共很多页 所以我们取很多很多页  
    for i in range(1,3):
        url = 'http://www.xicidaili.com/wt/{}'.format(i) 
        headers = {  
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36'  
        } 
        
        ip_list = get_ip_list(url, headers=headers)  
        proxies = get_random_ip(ip_list)  
        print(proxies)  
        #如何调用该类
        mycsdn = ScrapyMyCSDN('baidu_31657889') #初始化类 参数为博客名
        cur_write_nums = mycsdn.getOriginalArticalNums(proxies) #得到写了多少篇文章
        cur_blog_page = mycsdn.getScrapyPageNums(cur_write_nums) #cur_blog_page:返回需要爬取的页数

        mycsdn.beginToScrapy(cur_blog_page,proxies)
        time.sleep(20) # 给它休息时间 还是怕被封号的