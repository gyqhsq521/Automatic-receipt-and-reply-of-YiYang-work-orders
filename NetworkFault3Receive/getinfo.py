import datetime
import re
from selenium.webdriver.common.by import By
import requests
from lxml import etree

def get_html_str(pagesource):
    """
    定义一个函数, 新建一个空变量html_str， 请求网页获取网页源码，如果请求成功，则返回结果，如果失败则返回空值
    url: 入参参数, 指的是我们普通浏览器中的访问网址
    """
    html_str = ""
    try:
        """使用三方库etree把网页源码字符串转换成HTML格式"""
        html_str = etree.HTML(pagesource)
    except Exception as e:
        print(e)
    return html_str

#获取网络故障3.0待办工单页面工单列表中第8列数据
def get_processing_stage(html_str):
    data_list = []
    try:
        option = html_str.xpath('//*[@id="app"]/section/section/main/div[1]/div[2]/div/div/div[2]/div/div[3]/table/tbody/tr')

        for op in option:
            try:
                col8 = re.findall('[\\u4e00-\\u9fa5]+',op.xpath("./td[8]/div/text()")[0])[0]               

            except Exception as e:        
                print(e)
                col8 = ""
                
            data_list.append(col8)
            
    except Exception as e:
        print(e)
        
    return data_list
